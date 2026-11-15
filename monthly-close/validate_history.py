"""Independent reconciliation from delivered CSVs and cached real-Excel results.
No import from the generator; never recalculates or modifies Office documents.
"""
from pathlib import Path
from decimal import Decimal
from collections import defaultdict
from datetime import datetime
import argparse,csv,json,hashlib,zipfile
from openpyxl import load_workbook
from pptx import Presentation
ROOT=Path(__file__).parent;DRIVE=ROOT/'Alder Mechanical - Finance'
def cents(x):return int(Decimal(str(x or 0))*100)
def read(p):return list(csv.DictReader(p.open()))
def check(condition,msg):
    if not condition:raise AssertionError(msg)
def close(a,b,msg):check(abs(cents(a)-cents(b))<=1,f'{msg}: {a} vs {b}')
def validate(require_excel=False):
    manifest=json.loads((ROOT/'manifest.json').read_text());prior={(r['Entity'],r['Department'],r['Account']):cents(r['Ending']) for r in read(DRIVE/'00 Reference'/'Opening TB 2023-08.csv')};monthly={};results=[];formula_count=0
    coa={r['Account']:r for r in read(DRIVE/'00 Reference'/'Reporting COA.csv')}
    support=read(DRIVE/'00 Reference'/'Supporting monthly history 2022-Aug2023.csv')
    for r in support:monthly.setdefault(r['Period'][:7],defaultdict(int))[r['Line']]+=cents(r['USD'])
    folders=sorted(DRIVE.glob('FY20??/[0-1][0-9] *'))
    for folder in folders:
        if not (folder/'01 ERP exports').exists():continue
        period=f'{folder.parent.name[2:]}-{folder.name[:2]}'
        gl=read(next((folder/'01 ERP exports').glob('GL detail*.csv')));tb=read(next((folder/'01 ERP exports').glob('Trial balance*.csv')))
        sums=defaultdict(int);debits=defaultdict(int);credits=defaultdict(int);journals=defaultdict(int);facts=defaultdict(int)
        for r in gl:
            key=(r['Entity'],r['Department'],r['GLAccount']);v=cents(r['Debit'])-cents(r['Credit']);sums[key]+=v;debits[key]+=cents(r['Debit']);credits[key]+=cents(r['Credit']);journals[(r['Entity'],r['JournalID'])]+=v
            if coa[r['GLAccount']]['Type']=='PL' and r['Source'] not in ['Year close','Opening']:
                line=coa[r['GLAccount']]['Reporting line'];facts[line]+= -v if line in ['Contract revenue','Demand revenue','IC recovery'] else v
        check(all(v==0 for v in journals.values()),period+' unbalanced journal')
        check(sum(cents(r['Ending']) for r in tb)==0,period+' TB not balanced')
        now={}
        for r in tb:
            k=(r['Entity'],r['Department'],r['Account']);op=cents(r['Opening']);cl=cents(r['Ending']);now[k]=cl
            check(op+sums[k]==cl,period+' GL to TB '+str(k));check(debits[k]==cents(r['Debits']) and credits[k]==cents(r['Credits']),period+' movements '+str(k))
            if prior:check(prior.get(k,0)==op,period+' opening to preceding close '+str(k))
        for k,v in prior.items():check(k in now or v==0,period+' missing continuing account '+str(k))
        for kind,a in [('AR','1100'),('AP','2000')]:
            aging=read(next((folder/'02 Subledgers').glob(f'{kind} aging*.csv')))
            for e in ['100','200','300']:
                check(sum(cents(r['Open USD']) for r in aging if r['Entity']==e)==abs(sum(v for (en,dp,ac),v in now.items() if en==e and ac==a)),period+' '+kind+' aging '+e)
        for r in read(next((folder/'02 Subledgers').glob('Bank rec*.csv'))):
            e=r['Entity'];bank=cents(r['Bank statement'])+cents(r['Deposits in transit'])-cents(r['Outstanding checks']);cash=sum(v for (en,dp,a),v in now.items() if en==e and a=='1000')
            check(cash>=0,period+' unfunded entity cash balance '+e)
            check(bank==cash and cents(r['GL cash'])==cash and cents(r['Difference'])==0,period+' bank rec '+e)
        for r in read(next((folder/'02 Subledgers').glob('Payroll summary*.csv'))):
            e=r['Entity'];wages=sum(cents(x['Debit'])-cents(x['Credit']) for x in gl if x['Entity']==e and x['GLAccount'] in ['5100','5200','6000'] and x['Source']!='Year close')
            check(wages==cents(r['Gross plus benefits USD']),period+' payroll expense '+e)
            check(cents(r['Closing payroll accrual'])==-sum(v for (en,dp,a),v in now.items() if en==e and a=='2100'),period+' payroll accrual '+e)
        for a in ['1500','2500']:check(sum(v for (e,dp,ac),v in now.items() if ac==a)==0,period+' IC elimination '+a)
        check(facts['IC recovery']==facts['IC allocation'],period+' IC P&L elimination')
        if period.endswith('-01'):
            for e in ['100','200','300','ELIM']:
                transferred=sum(v for (en,dp,a),v in prior.items() if en==e and coa[a]['Type']=='PL')
                change=sum(v for (en,dp,a),v in now.items() if en==e and a=='3100')-sum(v for (en,dp,a),v in prior.items() if en==e and a=='3100')
                check(transferred==change,period+' retained earnings annual close '+e)
        monthly[period]=facts;prior=now
        rec={'period':period,'journals':len(journals),'gl_rows':len(gl),'tb_rows':len(tb),'source_controls':'PASS'}
        info=next((m for m in manifest if m['period']==period),None)
        if info:
            p=DRIVE/info['workbook'];wf=load_workbook(p,data_only=False);wv=load_workbook(p,data_only=True)
            n=sum(c.data_type=='f' for ws in wf for row in ws for c in row);formula_count+=n;check(n>300,period+' formula coverage')
            check(not wf._external_links,period+' external links');check(len(wf.worksheets)>=20,period+' sheets')
            errors=[f'{ws.title}!{c.coordinate}: {c.value}' for ws in wv for row in ws for c in row if c.data_type=='e'];check(not errors,period+str(errors[:8]))
            calculated=wv['PL Group']['B16'].value is not None
            if require_excel:check(calculated,period+' not recalculated in Excel')
            if calculated:
                for i,line in enumerate(['Contract revenue','Demand revenue','IC recovery','Parts','Field labor','Office labor','Occupancy','Fleet','Warranty','Insurance','Depreciation','IC allocation','Interest','Tax'],2):
                    close(wv['PL Group'].cell(i,2).value,facts[line]/100,period+' P&L '+line)
                    py=f'{int(period[:4])-1}{period[4:]}'
                    close(wv['PL Group'].cell(i,6).value,monthly.get(py,{}).get(line,0)/100,period+' PY '+line)
                    expected=sum(v.get(line,0) for per,v in monthly.items() if per[:4]==period[:4] and per<=period)
                    close(wv['YTD'].cell(i,wv['YTD'].max_column).value,expected/100,period+' YTD '+line)
                for row in list(wv['Checks'].values)[1:]:check(row[2]=='PASS',period+' check '+str(row))
                close(wv['Balance Sheet']['B2'].value,sum(v for (e,dp,a),v in now.items() if a=='1000')/100,period+' cash')
            prs=Presentation(DRIVE/info['deck']);check(len(prs.slides)>=4,period+' slide count')
            charts=[sh.chart for sl in prs.slides for sh in sl.shapes if sh.has_chart];check(len(charts)>=2,period+' native charts')
            for ch in charts:check(ch.part.chart_workbook.xlsx_part is not None,period+' editable embedded chart data')
            rev=facts['Contract revenue']+facts['Demand revenue']+facts['IC recovery'];close(charts[0].series[0].values[-1],rev/100000,period+' deck revenue chart')
            table=next(sh.table for sh in prs.slides[1].shapes if sh.has_table)
            revenue=facts['Contract revenue']+facts['Demand revenue']+facts['IC recovery']
            ebitda=revenue-sum(facts[l] for l in ['Parts','Field labor','Office labor','Occupancy','Fleet','Warranty','Insurance','IC allocation'])
            ni=ebitda-facts['Depreciation']-facts['Interest']-facts['Tax']
            for ri,value in [(1,revenue/100),(2,ebitda/100),(3,ni/100),(4,sum(v for (e,dp,a),v in now.items() if a=='1000')/100)]:
                actual=float(table.cell(ri,1).text.replace('$','').replace(',',''));check(abs(actual-value)<=.51,period+' deck table row '+str(ri))
            fulltext=' '.join(sh.text for sl in prs.slides for sh in sl.shapes if sh.has_text_frame)
            check(f'{datetime.fromisoformat(period+"-01"):%B %Y}' in fulltext,period+' deck date')
            rec.update(formulas=n,sheets=len(wf.worksheets),excel_cached=calculated,slides=len(prs.slides),editable_charts=len(charts))
        results.append(rec)
    check(len(manifest)==37,'Expected 37 closed months')
    out={'company':'Alder Mechanical Services Group — fictional sample','closed_months':len(manifest),'period_start':'2023-09','period_end':'2026-09','october_inputs':'2026-10','formula_count':formula_count,'independent_source_validation':'PASS','desktop_excel_validation':'PASS' if require_excel else 'PENDING','months':results}
    (ROOT/'validation.json').write_text(json.dumps(out,indent=2));return out
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--require-excel',action='store_true');a=p.parse_args();r=validate(a.require_excel);print(json.dumps({k:v for k,v in r.items() if k!='months'},indent=2))

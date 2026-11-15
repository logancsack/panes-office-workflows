"""Reconcile the recorded October outputs directly to the delivered source exports."""
import csv,json,calendar,zipfile,hashlib
from pathlib import Path
from decimal import Decimal
from datetime import datetime
from collections import defaultdict
from openpyxl import load_workbook
from pptx import Presentation
ROOT=Path(__file__).parent;DRIVE=ROOT/'Alder Mechanical - Finance';OUT=ROOT/'outputs'
def read(p):return list(csv.DictReader(p.open()))
def money(v):return round(Decimal(str(v or 0)),2)
def equal(a,b,label):
    if abs(float(a or 0)-float(b or 0))>.011:raise AssertionError(f'{label}: {a} != {b}')
def validate():
    w=load_workbook(OUT/'October 2026 close.xlsx',data_only=True);wf=load_workbook(OUT/'October 2026 close.xlsx',data_only=False)
    coa={r['Account']:r for r in read(DRIVE/'00 Reference'/'Reporting COA.csv')}
    expected=defaultdict(lambda:defaultdict(Decimal));dims=defaultdict(Decimal);cash={}
    files=sorted(DRIVE.glob('FY20??/?? ???/01 ERP exports/GL detail*.csv'))
    for p in files:
        per=p.stem[-7:]
        for r in read(p):
            ac=coa[r['GLAccount']];v=money(r['Debit'])-money(r['Credit'])
            if ac['Type']=='PL' and r['Source'] not in ['Year close','Opening']:
                line=ac['Reporting line'];v=-v if line in ['Contract revenue','Demand revenue','IC recovery'] else v;expected[per][line]+=v
                if per=='2026-10':dims[(r['Entity'],r['Department'],line)]+=v
        tb=read(p.with_name('Trial balance '+per+'.csv'));cash[per]=sum(money(r['Ending']) for r in tb if r['Account']=='1000')
    for r in read(DRIVE/'00 Reference'/'Supporting monthly history 2022-Aug2023.csv'):expected[r['Period'][:7]][r['Line']]+=money(r['USD'])
    cells=0
    for i in range(2,16):
        line=w['PL Group'].cell(i,1).value
        equal(w['PL Group'].cell(i,2).value,expected['2026-10'][line],line);cells+=1
        equal(w['PL Group'].cell(i,6).value,expected['2025-10'][line],line+' PY');cells+=1
        for sheet in ['Trailing 12','YTD']:
            ws=w[sheet]
            for j in range(2,ws.max_column):
                period=ws.cell(1,j).value.strftime('%Y-%m');equal(ws.cell(i,j).value,expected[period][line],sheet+' '+line+' '+period);cells+=1
        for sheet,dim in [('PL Entity',0),('PL Department',1)]:
            ws=w[sheet]
            for j in range(2,ws.max_column+1):
                key=str(ws.cell(1,j).value);value=sum(v for k,v in dims.items() if k[dim]==key and k[2]==line)
                equal(ws.cell(i,j).value,value,sheet+' '+key+' '+line);cells+=1
    budget=read(DRIVE/'FY2026/00 Planning/FY2026 Board budget approved.csv')
    forecast=read(DRIVE/'FY2026/00 Planning/FY2026 July reforecast v2.csv')
    for i in range(2,16):
        line=w['PL Group'].cell(i,1).value
        equal(w['PL Group'].cell(i,3).value,next(r['Budget USD'] for r in budget if r['Period']=='2026-10-31' and r['Line']==line),'original budget '+line)
        equal(w['PL Group'].cell(i,8).value,next(r['Reforecast USD'] for r in forecast if r['Period']=='2026-10-31' and r['Line']==line),'reforecast '+line);cells+=2
    payroll=read(DRIVE/'FY2026/10 Oct/02 Subledgers/Payroll summary 2026-10.csv')
    equal(w['KPI']['B5'].value,sum(int(r['Ending FTE']) for r in payroll),'ending headcount');cells+=1
    for row in list(w['Checks'].values)[1:]:
        if row[2]!='PASS':raise AssertionError(str(row))
    errors=[f'{s.title}!{c.coordinate} {c.value}' for s in w for row in s for c in row if c.data_type=='e']
    assert not errors,errors
    assert w['Control']['B2'].value==datetime(2026,10,31)
    assert w['YTD'].max_column==12
    assert w['Trailing 12']['B1'].value==datetime(2025,11,30)
    assert w['Trailing 12']['M1'].value==datetime(2026,10,31)
    assert len([r for r in w['Mapping'].values if r[0]=='6265'])==1
    assert not wf._external_links
    revenue=lambda per:sum(expected[per][l] for l in ['Contract revenue','Demand revenue','IC recovery'])
    ebitda=lambda per:revenue(per)-sum(expected[per][l] for l in ['Parts','Field labor','Office labor','Occupancy','Fleet','Warranty','Insurance','IC allocation'])
    equal(w['PL Group']['B16'].value,revenue('2026-10'),'total revenue');equal(w['PL Group']['B17'].value,ebitda('2026-10'),'EBITDA');equal(w['PL Group']['B18'].value,ebitda('2026-10')-sum(expected['2026-10'][l] for l in ['Depreciation','Interest','Tax']),'net income');equal(w['Balance Sheet']['B2'].value,cash['2026-10'],'cash');cells+=4
    periods=[f'{2025+(m-1)//12}-{(m-1)%12+1:02}' for m in range(11,23)]
    prs=Presentation(OUT/'October 2026 review.pptx');charts=[s.chart for sl in prs.slides for s in sl.shapes if s.has_chart]
    assert len(prs.slides)==6 and len(charts)==3
    for i,per in enumerate(periods):
        equal(charts[0].series[0].values[i],revenue(per)/1000,'deck revenue '+per)
        equal(charts[0].series[1].values[i],ebitda(per)/1000,'deck EBITDA '+per)
        equal(charts[1].series[0].values[i],ebitda(per)/revenue(per),'deck margin '+per)
        equal(charts[2].series[0].values[i],cash[per]/1000,'deck cash '+per)
        cells+=4
    expected_labels=[datetime.fromisoformat(per+'-01').strftime('%b %y') for per in periods]
    for ch in charts:
        assert [c.label for c in ch.plots[0].categories]==expected_labels
        assert ch.part.chart_workbook.xlsx_part is not None
    with zipfile.ZipFile(OUT/'October 2026 review.pptx') as z:
        assert not any(b'TargetMode="External"' in z.read(n) for n in z.namelist() if n.endswith('.rels'))
    desktop=json.loads((OUT/'october-desktop-validation.json').read_text(encoding='utf-8-sig'))
    assert not desktop['workbook']['formulaErrors']
    assert all(r['result']=='PASS' for r in desktop['workbook']['checks'])
    assert desktop['presentation']['editableCharts']==3 and desktop['presentation']['linkedCharts']==0
    text=' '.join(s.text for sl in prs.slides for s in sl.shapes if s.has_text_frame)
    assert 'September 2026' not in text
    r={'status':'PASS','application_versions':{'Excel':desktop['workbook']['version'],'PowerPoint':desktop['presentation']['version']},'independently_reconciled_cells':cells,'workbook_checks':14,'formula_errors':0,'formula_count':desktop['workbook']['formulaCount'],'slides':6,'editable_embedded_charts':3,'external_links':0,'october':{'revenue':float(revenue('2026-10')),'ebitda':float(ebitda('2026-10')),'net_income':round(w['PL Group']['B18'].value,2),'cash':float(cash['2026-10'])},'sha256':{p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in OUT.glob('*') if p.suffix in ['.xlsx','.pptx']}}
    (ROOT/'october-validation.json').write_text(json.dumps(r,indent=2));return r
if __name__=='__main__':print(json.dumps(validate(),indent=2))

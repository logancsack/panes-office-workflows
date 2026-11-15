"""Deterministic fictional finance shared drive. Never touches production data.
Amounts are integer cents in the journal, converted to USD only at file boundaries.
"""
from __future__ import annotations
import calendar, csv, json, random, shutil, hashlib
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path
from copy import copy
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.chart import BarChart, LineChart, Reference
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION

ROOT=Path(__file__).parent
DRIVE=ROOT/'Alder Mechanical - Finance'
RNG=random.Random(42073)
ENT={'100':'Alder Service LLC','200':'Alder Shared Services LLC','300':'Summit Field Services LLC','ELIM':'Consolidation eliminations'}
AC={
'1000':('Operating cash','Cash','BS'), '1100':('Trade receivables','AR','BS'),
'1200':('Service parts inventory','Inventory','BS'),'1300':('Prepaid insurance','Prepaids','BS'),
'1400':('Fleet and equipment cost','PPE','BS'),'1490':('Accumulated depreciation','Accum dep','BS'),
'1500':('Intercompany receivable','IC receivable','BS'),'1600':('Acquisition goodwill','Goodwill','BS'),
'2000':('Trade payables','AP','BS'),'2100':('Accrued payroll','Payroll accrual','BS'),
'2150':('Warranty reserve','Warranty accrual','BS'),'2200':('Income taxes payable','Tax payable','BS'),
'2500':('Intercompany payable','IC payable','BS'),'2600':('Term debt','Debt','BS'),
'3000':('Member capital','Capital','BS'),'3100':('Retained earnings','RE','BS'),'3200':('Member distributions','Capital','BS'),
'4000':('Service contracts','Contract revenue','PL'),'4100':('Demand service','Demand revenue','PL'),
'4900':('Shared service recovery','IC recovery','PL'),'5000':('Parts and subcontractors','Parts','PL'),
'5100':('Field wages and benefits','Field labor','PL'),'6000':('Office payroll and benefits','Office labor','PL'),
'6100':('Occupancy and facilities','Occupancy','PL'),'6200':('Fleet operating costs - legacy','Fleet','PL'),'6205':('Fleet fuel and maintenance','Fleet','PL'),
'6250':('Customer warranty costs','Warranty','PL'),'6265':('Third-party warranty callouts','Warranty','PL'),
'6300':('Insurance amortization','Insurance','PL'),'6400':('Depreciation','Depreciation','PL'),
'6500':('Shared service allocation','IC allocation','PL'),'6600':('Interest expense','Interest','PL'),
'6700':('Income tax expense','Tax','PL'),'5200':('Field payroll - legacy','Field labor','PL')}
LINES=['Contract revenue','Demand revenue','IC recovery','Parts','Field labor','Office labor','Occupancy','Fleet','Warranty','Insurance','Depreciation','IC allocation','Interest','Tax']
SEASON=[1.00,0.92,0.90,0.95,1.10,1.24,1.36,1.29,1.08,0.97,1.06,1.17]

def months():
    return [datetime(y,m,1) for y in range(2022,2027) for m in range(1,13) if (y,m)<=(2026,10)]
def end(d): return datetime(d.year,d.month,calendar.monthrange(d.year,d.month)[1])
def dollars(c): return round(c/100,2)
def csvout(p,headers,rows):
    p.parent.mkdir(parents=True,exist_ok=True)
    with p.open('w',newline='') as f:
        w=csv.writer(f);w.writerow(headers);w.writerows(rows)
def dept(d,k): return ('Maintenance' if k=='Contract' else 'Service') if d<datetime(2025,4,1) else ('Planned Maintenance' if k=='Contract' else 'Field Response')
def narrative(d):
    notes=[f'{d:%B} demand follows the heating/cooling service cycle; contract pricing increased 4% each January.']
    if d.year>=2025:notes+=['Summit included from January 2025; acquisition funded with term debt and member capital.']
    if d==datetime(2024,7,1):notes+=['Parts inventory write-down of $28,500 for discontinued refrigerant components.']
    if d==datetime(2026,6,1):notes+=['Restated June: $42,000 warranty provision recorded after vendor claim review; see R1 bridge.']
    if d==datetime(2026,10,1):notes+=['New GL 6265: $18,750 third-party warranty callouts mapped to Warranty; $24,600 payroll accrual true-up.']
    if d>=datetime(2025,4,1):notes+=['Service Operations split into Planned Maintenance and Field Response in April 2025; historical departments retained.']
    return notes

def ledger():
    balances=defaultdict(int); rows=[]; snapshots={}; schedules={}; facts=[]; metrics={}
    def post(d,e,a,b,amount,dep,memo,source='ERP',doc=None):
        nonlocal rows
        n=len(rows)//2+1; jid=doc or f'{d:%Y%m}-{n:06}'
        for account,value in [(a,amount),(b,-amount)]:
            key=(e,dep if AC[account][2]=='PL' else 'Balance Sheet',account)
            balances[key]+=value
            rows.append([end(d).strftime('%Y-%m-%d'),jid,e,key[1],account,AC[account][0],dollars(max(value,0)),dollars(max(-value,0)),memo,source])
    for d in months():
        start=len(rows);opening=dict(balances)
        # Annual closing entries are included in January GL, never silently reset.
        if d.month==1:
            for (e,dp,a),v in list(balances.items()):
                if AC[a][2]=='PL' and v:post(d,e,'3100',a,v,dp,'Prior-year earnings transfer','Year close')
        if d==datetime(2022,1,1):
            for e in ['100','200']:
                post(d,e,'1000','3000',180000000 if e=='100' else 25000000,'Corporate','Opening member capital','Opening')
                if e=='100':
                    post(d,e,'1400','2600',65000000,'Corporate','Opening fleet / term debt','Opening')
                    post(d,e,'1200','3000',18000000,'Corporate','Opening parts inventory','Opening')
                    post(d,e,'1300','1000',12000000,'Corporate','Annual insurance premium')
        if d==datetime(2025,1,1):
            post(d,'300','1000','3000',100000000,'Corporate','Summit closing equity','Acquisition')
            post(d,'300','1400','2600',48000000,'Corporate','Summit fleet / assumed debt','Acquisition')
            post(d,'300','1600','2600',62000000,'Corporate','Summit purchase goodwill','Acquisition')
            post(d,'300','1200','1000',11500000,'Corporate','Summit opening parts stock','Acquisition')
        active=['100','200']+(['300'] if d.year>=2025 else [])
        hc={}; capex={}; amort={}; depreciation={}; pay={}
        for e in active:
            operational=e!='200'; scale=(1 if e=='100' else .32 if e=='300' else .18)
            growth=1.075**(d.year-2022)*(1+.002*(d.month-1));noise=RNG.uniform(.976,1.026)
            rev=int(116000000*scale*growth*SEASON[d.month-1]*noise) if operational else 0
            head=round((68 if e=='100' else 24 if e=='300' else 12)*(1+.075*(d.year-2022)))+(2 if operational and d.month>=4 else 0)+(1 if operational and d.month>=7 else 0)
            hc[e]=head
            if operational:
                contract=round(rev*.55);demand=rev-contract
                for a,v,k in [('4000',contract,'Contract'),('4100',demand,'Demand')]:
                    dp=dept(d,k)
                    # Customer invoice batches, rather than a single monthly revenue line.
                    n=24 if e=='100' else 12; rem=v
                    for i in range(n):
                        part=rem if i==n-1 else round(v/n*RNG.uniform(.6,1.4));rem-=part
                        post(d,e,'1100',a,part,dp,f'{k} invoice batch {i+1:02} / branch '+(['Cedar Rapids','Madison','Rockford'][i%3] if e=='100' else 'Des Moines'))
                receivable=sum(v for (en,_,a),v in balances.items() if en==e and a=='1100')
                target=round(rev*RNG.uniform(1.05,1.23))
                post(d,e,'1000','1100',max(0,receivable-target),'Corporate','Customer lockbox receipts')
                for k,f in [('Contract',.55),('Demand',.45)]:
                    dp=dept(d,k)
                    parts=round(rev*f*(.237 if k=='Contract' else .29))
                    
                    remparts=parts
                    for j in range(12):
                        piece=remparts if j==11 else round(parts/12*RNG.uniform(.7,1.3));remparts-=piece
                        post(d,e,'5000','2000',piece,dp,f'Supplier invoice / parts and subcontractors batch {j+1:02}')
                    labor=round(head*670000*f*(1.035**(d.year-2022)))
                    account='5200' if d<datetime(2026,1,1) else '5100'
                    post(d,e,account,'2100',labor,dp,'Biweekly field payroll and employer benefits')
                if d==datetime(2024,7,1) and e=='100':post(d,e,'5000','1200',2850000,dept(d,'Demand'),'Obsolete R-22 parts write-down','Close JE')
                # Stock build and cash conversion are seasonal.
                stock=sum(v for (en,_,a),v in balances.items() if en==e and a=='1200')
                targetstock=round(rev*.20)
                if targetstock>stock:post(d,e,'1200','2000',targetstock-stock,'Corporate','Seasonal van-stock replenishment')
                else:post(d,e,'2000','1200',stock-targetstock,'Corporate','Inventory returns / supplier credits')
            office=round((head if not operational else max(4,round(head*.12)))*760000*1.035**(d.year-2022))
            post(d,e,'6000','2100',office,'Corporate','Office payroll, benefits and taxes')
            payroll=sum(v for (en,_,a),v in balances.items() if en==e and a=='2100')
            targetpay=round((office+(head*670000*1.035**(d.year-2022) if operational else 0))*.27)
            pay[e]=targetpay
            post(d,e,'2100','1000',-payroll-targetpay,'Corporate','Payroll cash settlement')
            if d==datetime(2026,10,1) and e=='100':post(d,e,'5100','2100',2460000,dept(d,'Demand'),'Accrued field overtime true-up: approved hours after payroll cut-off','Close JE','JE-202610-017')
            post(d,e,'6100','2000',round(1900000*scale*growth),'Corporate','Branch rent and utilities')
            post(d,e,'6200' if d<datetime(2024,1,1) else '6205','2000',round(1600000*scale*growth*SEASON[d.month-1]),dept(d,'Demand') if operational else 'Corporate','Fuel, fleet maintenance and telematics')
            if operational:
                war=round(rev*.008)
                if d==datetime(2026,6,1) and e=='100':war+=4200000
                post(d,e,'6250','2150',war,dept(d,'Demand'),'Warranty provision'+(' / R1 vendor claim reserve' if d==datetime(2026,6,1) else ''),'Close JE')
                oldwar=-sum(v for (en,_,a),v in opening.items() if en==e and a=='2150')
                post(d,e,'2150','1000',round(max(0,oldwar)*.28),'Corporate','Warranty cash claims')
                if d==datetime(2026,10,1) and e=='100':post(d,e,'6265','2000',1875000,dept(d,'Demand'),'Emergency subcontractor warranty callouts','ERP','AP-WAR-2610-084')
            premium=round(12000000*scale*1.05**(d.year-2022))
            if d.month==1 and d!=datetime(2022,1,1):post(d,e,'1300','1000',premium,'Corporate','Annual insurance renewal')
            elif d==datetime(2022,1,1) and e=='200':post(d,e,'1300','1000',premium,'Corporate','Annual insurance premium')
            amort[e]=round(premium/12)
            post(d,e,'6300','1300',amort[e],'Corporate','Monthly insurance amortization','Close JE')
            capex[e]=round((850000 if operational else 110000)*growth*(2.2 if d.month in [3,6,9] else .5))
            post(d,e,'1400','1000',capex[e],'Corporate','Fleet replacement and tools additions')
            ppe=sum(v for (en,_,a),v in balances.items() if en==e and a=='1400')
            depreciation[e]=round(ppe/84)
            post(d,e,'6400','1490',depreciation[e],'Corporate','Straight-line fleet / tools depreciation, 84 months','Close JE')
            debt=-sum(v for (en,_,a),v in balances.items() if en==e and a=='2600')
            if debt:
                post(d,e,'6600','1000',round(debt*.0675/12),'Corporate','Term debt interest, 6.75%')
                post(d,e,'2600','1000',min(debt,850000 if e=='100' else 650000),'Corporate','Scheduled debt principal')
            ap=-sum(v for (en,_,a),v in balances.items() if en==e and a=='2000')
            post(d,e,'2000','1000',round(ap*.67),'Corporate','Weekly vendor payment runs')
        # Cost-plus shared-services allocation. Entity profit + balances eliminated in consolidation.
        sharedcost=sum(round(float(r[6])*100)-round(float(r[7])*100) for r in rows[start:] if r[2]=='200' and AC[r[4]][2]=='PL')
        alloc=round(sharedcost*1.04)
        for e,weight in [('100',.76 if '300' in active else 1),('300',.24 if '300' in active else 0)]:
            if not weight:continue
            v=round(alloc*weight)
            post(d,e,'6500','2500',v,'Corporate','Central finance, HR and dispatch allocation')
            post(d,'200','1500','4900',v,'Corporate','Cost-plus shared service recovery')
            post(d,'ELIM','4900','6500',v,'Corporate','Eliminate intercompany P&L','Consolidation')
            post(d,'ELIM','2500','1500',v,'Corporate','Eliminate intercompany balances','Consolidation')
            prior_due=-sum(x for (en,_,a),x in opening.items() if en==e and a=='2500')
            settle=round(v*.8)+prior_due
            post(d,e,'2500','1000',settle,'Corporate','Intercompany settlement')
            post(d,'200','1000','1500',settle,'Corporate','Intercompany settlement')
            post(d,'ELIM','1500','2500',settle,'Corporate','Reverse settled IC elimination','Consolidation')
        for e in active:
            profit=-sum(round(float(r[6])*100)-round(float(r[7])*100) for r in rows[start:] if r[2]==e and AC[r[4]][2]=='PL' and r[9]!='Year close')
            tax=max(0,round(profit*.25))
            post(d,e,'6700','2200',tax,'Corporate','Monthly tax provision at 25%','Close JE')
            if d.month%3==0:
                due=-sum(v for (en,_,a),v in balances.items() if en==e and a=='2200')
                post(d,e,'2200','1000',due,'Corporate','Quarterly estimated income tax')
        if d.month in [3,6,9,12]:
            for e in ['100','200']+(['300'] if d.year>=2025 else []):
                available=sum(v for (en,_,a),v in balances.items() if en==e and a=='1000')
                distribution=max(0,round((available-(150000000 if e=='100' else 30000000))*.55))
                if distribution:post(d,e,'3200','1000',distribution,'Corporate','Quarterly member distribution approved by board','Treasury')
        monthly=defaultdict(int)
        for r in rows[start:]:
            if r[9] not in ['Year close','Opening'] and AC[r[4]][2]=='PL':monthly[(r[2],r[3],AC[r[4]][1])]+=round(float(r[6])*100)-round(float(r[7])*100)
        for (e,dp,line),v in monthly.items():facts.append([end(d),e,dp,line,dollars(-v if 'revenue' in line or line=='IC recovery' else v)])
        snapshots[d]=dict(balances)
        schedules[d]={'headcount':hc,'capex':capex,'amort':amort,'depreciation':depreciation,'opening':opening}
        net=lambda line:sum(x[4] for x in facts if x[0]==end(d) and x[3]==line)
        rev=net('Contract revenue')+net('Demand revenue')+net('IC recovery')
        ebitda=rev-sum(net(l) for l in LINES[3:10])-net('IC allocation')
        metrics[d]={'revenue':round(rev,2),'ebitda':round(ebitda,2),'net_income':round(ebitda-net('Depreciation')-net('Interest')-net('Tax'),2),'cash':dollars(sum(v for (e,dp,a),v in balances.items() if a=='1000')),'headcount':sum(hc.values())}
    return rows,snapshots,schedules,facts,metrics

NAVY='183949';TEAL='007F7B';NUM='#,##0;[Red](#,##0);–'
def tab(wb,name,headers,rows=(),table=None):
    ws=wb.create_sheet(name);ws.append(headers)
    for r in rows:ws.append(r)
    if table:
        t=Table(displayName=table,ref=f'A1:{ws.cell(1,len(headers)).column_letter}{max(ws.max_row,2)}');t.tableStyleInfo=TableStyleInfo(name='TableStyleMedium2',showRowStripes=True);ws.add_table(t)
    ws.freeze_panes='B2';ws.sheet_view.showGridLines=False
    return ws

def style(wb):
    for ws in wb:
        ws.sheet_properties.pageSetUpPr.fitToPage=True
        ws.page_setup.orientation='landscape';ws.page_setup.paperSize=ws.PAPERSIZE_A3;ws.page_setup.fitToWidth=1;ws.page_setup.fitToHeight=0
        ws.print_options.horizontalCentered=True;ws.print_title_rows='1:1'
        ws.sheet_properties.outlinePr.summaryRight=False
        for c in ws[1]:c.fill=PatternFill('solid',fgColor=NAVY);c.font=Font(name='Aptos',size=11,bold=True,color='FFFFFF');c.alignment=Alignment(wrap_text=True,vertical='center')
        ws.row_dimensions[1].height=32
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.font=Font(name='Aptos',size=10,color=TEAL if c.data_type=='f' else '2454A6' if isinstance(c.value,(float,int)) else '253846')
                if c.data_type=='f' or isinstance(c.value,(float,int)):c.number_format=NUM
                if isinstance(c.value,datetime):c.number_format='mmm-yy'
        for col in ws.columns:
            letter=col[0].column_letter
            size=min(46,max(16,max(len(str(c.value or '')) if c.data_type!='f' else 16 for c in list(col)[:80])+2))
            ws.column_dimensions[letter].width=size
        ws.sheet_properties.tabColor=TEAL if ws.title in ['PL Group','Checks','Flux','Dashboard'] else NAVY

def sumfact(line,period='CloseMonth',entity=None,department=None):
    f=f'SUMIFS(Facts[Amount],Facts[Line],{line},Facts[Period],{period}'
    if entity:f+=f',Facts[Entity],{entity}'
    if department:f+=f',Facts[Department],{department}'
    return f+')'

def workbook(d,path,rows,snapshots,schedules,facts,metrics,previous=None):
    # A persisted workbook is the seed for the next close; retain legacy tabs and change log.
    wb=load_workbook(previous) if previous else Workbook()
    preserved=[]
    if previous:
        preserved=list(wb['Change Log'].values)[1:]
    for ws in list(wb):wb.remove(ws)
    for n in list(wb.defined_names):del wb.defined_names[n]
    s=snapshots[d];sch=schedules[d]; mr=[r for r in rows if r[0]==end(d).strftime('%Y-%m-%d')]
    currentfacts=[r for r in facts if r[0]<=end(d)]
    controls=tab(wb,'Control',['ALDER MECHANICAL SERVICES GROUP | Monthly close','Value','Owner / source'],[
        ['Reporting month',end(d),'Calendar FY / USD'],['Company','Fictional sample company','No real company data'],['Prepared by','Finance team','FP&A'],['Status','Reviewed','Controller'],['Source folder',f'FY{d.year}/{d:%m %b}','ERP + subledgers'],['Revision','R1 restated' if d==datetime(2026,6,1) else 'FINAL','June 2026 reserve revision'],['Budget scenario',f'FY{d.year} Board budget','Original plan; reforecast shown separately'],['Model convention','Blue inputs / teal formulas','Debits positive; reporting revenue positive']])
    wb.defined_names.add(DefinedName('CloseMonth',attr_text="'Control'!$B$2"))
    wb.defined_names.add(DefinedName('YearStart',attr_text='DATE(YEAR(CloseMonth),1,1)'))
    mapping=[[a,*v,('FY26 payroll consolidation' if a in ['5100','5200'] and d>=datetime(2026,1,1) else 'FY24 fleet reporting' if a in ['6200','6205'] and d>=datetime(2024,1,1) else 'Original reporting COA')] for a,v in AC.items() if a!='6265' and (a!='6205' or d>=datetime(2024,1,1)) and (a!='5100' or d>=datetime(2026,1,1))]
    tab(wb,'Mapping',['Account','Description','Line','Type','Mapping note'],mapping,'Map')
    tb=[]
    for e,dp,a in sorted(set(s)|set(sch['opening'])):
        matches=[r for r in mr if r[2:5]==[e,dp,a]]
        debit=round(sum(r[6] for r in matches),2);credit=round(sum(r[7] for r in matches),2)
        v=s.get((e,dp,a),0);op=sch['opening'].get((e,dp,a),0)
        if not any([debit,credit,v,op]):continue
        n=len(tb)+2
        tb.append([e,dp,a,AC[a][0],dollars(op),debit,credit,dollars(v),f'=E{n}+F{n}-G{n}-H{n}',f'=IFERROR(INDEX(Map[Line],MATCH(C{n},Map[Account],0)),"UNMAPPED")',f'=IFERROR(INDEX(Map[Type],MATCH(C{n},Map[Account],0)),"UNMAPPED")'])
    tab(wb,'TB Import',['Entity','Department','Account','Description','Opening','Debits','Credits','Ending','Roll check','Line','Type'],tb,'TB')
    tab(wb,'GL Import',['Date','Journal','Entity','Department','Account','Description','Debit','Credit','Memo','Source'],mr,'GL')
    tab(wb,'Monthly History',['Period','Entity','Department','Line','Amount'],currentfacts,'Facts')
    # Fixed yearly budget set at planning date; never backsolved to actual variances.
    bud=[]
    for m in range(1,13):
        p=datetime(d.year,m,1)
        for line in LINES:
            base=(1180000*1.072**(d.year-2022)*(1.30 if d.year>=2025 else 1))*SEASON[m-1]
            weights={'Contract revenue':.55,'Demand revenue':.45,'IC recovery':0,'Parts':.264,'Field labor':.385/SEASON[m-1],'Office labor':.125/SEASON[m-1],'Occupancy':.021/SEASON[m-1],'Fleet':.015,'Warranty':.009,'Insurance':.009/SEASON[m-1],'Depreciation':.013/SEASON[m-1],'IC allocation':0,'Interest':.003,'Tax':.032}
            bud.append([end(p),line,round(base*weights[line],2),round(base*weights[line]*(1.035 if m>=7 else 1),2)])
    tab(wb,'Budget Import',['Period','Line','Budget','Reforecast'],bud,'Plan')
    pl=[]
    for i,line in enumerate(LINES,2):
        pl.append([line,'='+sumfact(f'A{i}'),f'=SUMIFS(Plan[Budget],Plan[Line],A{i},Plan[Period],CloseMonth)',f'=B{i}-C{i}',f'=IFERROR(D{i}/ABS(C{i}),0)', '='+sumfact(f'A{i}','EOMONTH(CloseMonth,-12)'),f'=B{i}-F{i}',f'=SUMIFS(Plan[Reforecast],Plan[Line],A{i},Plan[Period],CloseMonth)'])
    pl += [['Revenue','=SUM(B2:B4)','=SUM(C2:C4)','=B16-C16','=IFERROR(D16/C16,0)','=SUM(F2:F4)','=B16-F16','=SUM(H2:H4)'],['EBITDA','=B16-SUM(B5:B11)-B13','=C16-SUM(C5:C11)-C13','=B17-C17','=IFERROR(D17/ABS(C17),0)','=F16-SUM(F5:F11)-F13','=B17-F17','=H16-SUM(H5:H11)-H13'],['Net income','=B17-B12-B14-B15','=C17-C12-C14-C15','=B18-C18','=IFERROR(D18/ABS(C18),0)','=F17-F12-F14-F15','=B18-F18','=H17-H12-H14-H15']]
    ws=tab(wb,'PL Group',['Reporting line','Actual','Budget','Variance $','Variance %','Prior year','YoY $','Reforecast'],pl)
    for i in range(2,19):ws.cell(i,5).number_format='0.0%'
    for name,dim,vals in [('PL Entity','Entity',list(ENT)),('PL Department','Department',sorted(set(r[2] for r in currentfacts)))]:
        data=[]
        for line in LINES:
            i=len(data)+2;data.append([line]+['='+sumfact(f'$A{i}',entity=f'{chr(66+j)}$1' if dim=='Entity' else None,department=f'{chr(66+j)}$1' if dim=='Department' else None) for j in range(len(vals))])
        tab(wb,name,['Reporting line']+vals,data)
    for name,periods in [('Trailing 12',[end(months()[months().index(d)-11+i]) for i in range(12)]),('YTD',[end(datetime(d.year,m,1)) for m in range(1,d.month+1)])]:
        data=[]
        for line in LINES:
            i=len(data)+2; data.append([line]+['='+sumfact(f'$A{i}',f'{chr(66+j)}$1') for j in range(len(periods))]+[f'=SUM(B{i}:{chr(65+len(periods))}{i})'])
        tab(wb,name,['Reporting line']+periods+['Total'],data)
    tab(wb,'Prior Year',['Line','Current YTD','Prior YTD','Change'],[[l,f'=SUMIFS(Facts[Amount],Facts[Line],A{i},Facts[Period],">="&YearStart,Facts[Period],"<="&CloseMonth)',f'=SUMIFS(Facts[Amount],Facts[Line],A{i},Facts[Period],">="&EDATE(YearStart,-12),Facts[Period],"<="&EDATE(CloseMonth,-12))',f'=B{i}-C{i}'] for i,l in enumerate(LINES,2)])
    bslines=['Cash','AR','Inventory','Prepaids','PPE','Accum dep','IC receivable','Goodwill','AP','Payroll accrual','Warranty accrual','Tax payable','IC payable','Debt','Capital','RE']
    bs=[]
    for i,l in enumerate(bslines,2):bs.append([l,f'=SUMIFS(TB[Ending],TB[Line],A{i})',f'=SUMIFS(TB[Opening],TB[Line],A{i})',f'=B{i}-C{i}'])
    bs += [['Current-year earnings','=SUMIFS(TB[Ending],TB[Type],"PL")','=SUMIFS(TB[Opening],TB[Type],"PL")','=B18-C18'],['Balance check','=SUM(B2:B18)','=SUM(C2:C18)','=B19-C19']]
    tab(wb,'Balance Sheet',['Debits positive / credits negative','Closing','Opening','Movement'],bs)
    # Indirect cash flow from actual account changes; initial funding disclosed separately.
    cf=[['Net income',"='PL Group'!B18"],['Depreciation',"='PL Group'!B12"],['Receivables',"=-'Balance Sheet'!D3"],['Inventory',"=-'Balance Sheet'!D4"],['Prepaids',"=-'Balance Sheet'!D5"],['Trade payables',"=-'Balance Sheet'!D10"],['Payroll accrual',"=-'Balance Sheet'!D11"],['Warranty reserve',"=-'Balance Sheet'!D12"],['Tax payable',"=-'Balance Sheet'!D13"],['Cash from operations','=SUM(B2:B10)'],['Fleet capital expenditure',"=-'Balance Sheet'!D6"],['Acquisition goodwill',"=-'Balance Sheet'!D9"],['Debt and capital funding',"=-'Balance Sheet'!D15-'Balance Sheet'!D16"],['Change in cash','=SUM(B11:B14)'],['Opening cash',"='Balance Sheet'!C2"],['Closing cash','=B15+B16'],['Cash flow tie',"=B17-'Balance Sheet'!B2"]]
    tab(wb,'Cash Flow',['Indirect cash flow','USD'],cf)
    tab(wb,'KPI',['Metric','Value','Definition'],[['Revenue',"='PL Group'!B16",'External service revenue; IC eliminated'],['EBITDA',"='PL Group'!B17",'Revenue less operating costs before D&A, interest and tax'],['EBITDA margin','=IFERROR(B3/B2,0)','EBITDA / revenue'],['Headcount',metrics[d]['headcount'],'Payroll roster; month-end FTE'],['Revenue per FTE','=IFERROR(B2/B5,0)','Monthly revenue / ending FTE'],['DSO',"=IFERROR('Balance Sheet'!B3/B2*DAY(CloseMonth),0)",'Month-end receivables / monthly revenue × days'],['Cash',"='Balance Sheet'!B2",'Reconciled operating cash'],['Net leverage',"=IFERROR((-'Balance Sheet'!B15-B8)/(B3*12),0)",'Debt less cash / annualized current EBITDA']])
    flux=[]
    for i,l in enumerate(LINES,2):
        comment=next((n for n in narrative(d) if ('Warranty'==l and ('warranty' in n.lower())) or (l=='Field labor' and ('payroll' in n.lower()))),narrative(d)[0] if 'revenue' in l else 'Timing and run-rate movements; reconcile to supporting schedule / GL detail.')
        av=sum(r[4] for r in currentfacts if r[0]==end(d) and r[3]==l)
        bv=next(r[2] for r in bud if r[0]==end(d) and r[1]==l)
        py=sum(r[4] for r in currentfacts if r[0].year==d.year-1 and r[0].month==d.month and r[3]==l)
        baseline=f'Actual ${av:,.0f}; budget ${bv:,.0f}; prior year ${py:,.0f}. '
        if l=='Field labor':comment=f'{metrics[d]["headcount"]} ending FTE across the group; field wage rates grow 3.5% annually, with spring/summer staffing steps. '+(narrative(d)[-2] if d==datetime(2026,10,1) else '')
        elif l=='Depreciation':comment='84-month straight-line fleet and tool pools; additions and accumulated depreciation reconcile in Fixed Assets.'
        elif l=='Insurance':comment='Annual January premiums amortized monthly; closing balance reconciles in Prepaids.'
        elif l=='Interest':comment='6.75% term debt rate and scheduled principal reduction; acquisition financing included from January 2025.' if d.year>=2025 else '6.75% term debt rate and scheduled principal reduction; see debt account 2600.'
        elif l=='Tax':comment='25% current provision on positive entity pretax profit; quarterly estimated payments in March, June, September and December.'
        elif l in ['IC recovery','IC allocation']:comment='Shared services allocated at cost plus 4%; equal opposite elimination journals remove the group impact.'
        flux_comment=baseline+comment
        flux.append([l,f"='PL Group'!D{i}",f"='PL Group'!G{i}",f'=IF(ABS(B{i})>25000,"Review","Within threshold")',flux_comment,'FP&A','Reviewed'])
    tab(wb,'Flux',['Line','vs budget','vs PY','Threshold $25k','Finance commentary','Owner','Review'],flux)
    pre=[];fa=[];accr=[]
    for e in sch['headcount']:
        def bal(a,opening=False):return dollars(sum(v for (en,_,aa),v in (sch['opening'] if opening else s).items() if en==e and aa==a))
        n=len(pre)+2; add=bal('1300')-bal('1300',True)+dollars(sch['amort'][e]);pre.append([e,'Annual insurance',bal('1300',True),round(add,2),dollars(sch['amort'][e]),f'=C{n}+D{n}-E{n}',f'=F{n}-SUMIFS(TB[Ending],TB[Entity],A{n},TB[Account],"1300")'])
        n=len(fa)+2;fa.append([e,'Fleet and tools / 84 months',bal('1400',True),round(bal('1400')-bal('1400',True),2),f'=C{n}+D{n}',-bal('1490',True),dollars(sch['depreciation'][e]),f'=F{n}+G{n}',f'=E{n}-H{n}',f'=E{n}-SUMIFS(TB[Ending],TB[Entity],A{n},TB[Account],"1400")',f'=H{n}+SUMIFS(TB[Ending],TB[Entity],A{n},TB[Account],"1490")'])
        for a in ['2100','2150']:
            n=len(accr)+2;rr=[r for r in mr if r[2]==e and r[4]==a];accr.append([e,a,AC[a][0],-bal(a,True),round(sum(r[7] for r in rr),2),round(sum(r[6] for r in rr),2),f'=D{n}+E{n}-F{n}',f'=G{n}+SUMIFS(TB[Ending],TB[Entity],A{n},TB[Account],B{n})','Approved payroll cut-off / vendor claims'])
    tab(wb,'Prepaids',['Entity','Contract','Opening','Additions','Amortization','Closing','Tie'],pre)
    tab(wb,'Fixed Assets',['Entity','Pool','Opening cost','Additions','Closing cost','Opening accum dep','Depreciation','Closing accum dep','NBV','Cost tie','Dep tie'],fa)
    tab(wb,'Accruals',['Entity','Account','Description','Opening reserve','Provision','Settled','Closing reserve','Tie','Evidence'],accr)
    tab(wb,'Journal Log',['Journal','Description','Debit USD','Credit USD','Approval'],[[r[1],r[8],r[6],r[7],'Controller'] for r in mr if r[9] in ['Close JE','Consolidation','Year close']])
    changes=preserved+[[end(d),'Monthly roll-forward',f'{d:%b %Y} GL/TB import, schedule roll, flux and deck refresh','Finance team']]
    if d==datetime(2024,1,1):changes.append([end(d),'Mapping change','New fleet GL 6205 mapped to Fleet; legacy 6200 retained for comparisons. No change to reporting line totals.','Controller'])
    if d==datetime(2025,1,1):changes.append([end(d),'Acquisition','Summit acquired January 1; opening purchase journals and headcount added.','Controller'])
    if d==datetime(2025,4,1):changes.append([end(d),'Reorganization','Service renamed Field Response; Maintenance renamed Planned Maintenance; old department labels retained in historical facts.','FP&A'])
    if d==datetime(2026,1,1):changes.append([end(d),'COA migration','5200 replaced by 5100, both map to Field labor; legacy payroll tab hidden.','Controller'])
    if d==datetime(2026,6,1):changes.append([end(d),'R1 restatement','Warranty reserve increased $42,000; deferred tax effect $10,500; net income reduced $31,500. Superseded pack retained.','Controller'])
    tab(wb,'Change Log',['Date','Event','Details','Owner'],changes)
    legacy=tab(wb,'Legacy Dept Pivot',['Legacy department','Replacement','Effective'],[['Service','Field Response','Apr 2025'],['Maintenance','Planned Maintenance','Apr 2025']]);legacy.sheet_state='hidden'
    if d<datetime(2025,4,1):wb.remove(legacy)
    legacy=tab(wb,'Payroll 5200 old',['Old account','Current account','Effective'],[['5200','5100','Jan 2026']]);legacy.sheet_state='hidden'
    if d<datetime(2026,1,1):wb.remove(legacy)
    checks=[['TB debits less credits','=SUM(TB[Ending])'],['TB movement roll','=SUMPRODUCT(ABS(TB[Roll check]))'],['GL debits less credits','=SUM(GL[Debit])-SUM(GL[Credit])'],['Unmapped accounts','=COUNTIF(TB[Line],"UNMAPPED")'],['Balance sheet',"='Balance Sheet'!B19"],['Cash flow',"='Cash Flow'!B18"],['Prepaid schedule','=SUMPRODUCT(ABS(Prepaids!G2:G4))'],['Fixed asset cost','=SUMPRODUCT(ABS(\'Fixed Assets\'!J2:J4))'],['Accum depreciation','=SUMPRODUCT(ABS(\'Fixed Assets\'!K2:K4))'],['Accrual schedule','=SUMPRODUCT(ABS(Accruals!H2:H7))'],['IC receivable elimination','=SUMIFS(TB[Ending],TB[Account],"1500")'],['IC payable elimination','=SUMIFS(TB[Ending],TB[Account],"2500")'],['IC profit elimination',"='PL Group'!B4-'PL Group'!B13"],['P&L to TB YTD',f'=SUMIFS(TB[Ending],TB[Type],"PL")+SUM(\'YTD\'!{chr(66+d.month)}2:{chr(66+d.month)}4)-SUM(\'YTD\'!{chr(66+d.month)}5:{chr(66+d.month)}15)']]
    # Formula above includes all 14 fact lines; revenue rows 2:4, expenses 5:15.
    tab(wb,'Checks',['Control | zero = pass','Difference','Result'],[[l,f,f'=IF(ABS(B{i})<0.01,"PASS","FAIL")'] for i,(l,f) in enumerate(checks,2)])
    chart=BarChart();chart.title='Monthly revenue and cost mix';chart.add_data(Reference(wb['PL Group'],min_col=2,max_col=3,min_row=1,max_row=15),titles_from_data=True);chart.set_categories(Reference(wb['PL Group'],min_col=1,min_row=2,max_row=15));chart.width=25;chart.height=12;wb['PL Group'].add_chart(chart,'J2')
    style(wb)
    # Date headers / percentages after common style pass.
    for ws in [wb['Trailing 12'],wb['YTD']]:
        for c in ws[1]:
            if isinstance(c.value,datetime):c.number_format='mmm-yy'
    
    for i in range(2,19):wb['PL Group'].cell(i,5).number_format='0.0%'
    wb['PL Group'].column_dimensions['A'].width=26
    for row in wb['PL Group'].iter_rows(min_row=16,max_row=18):
        for c in row:c.font=Font(name='Aptos',size=10,color=TEAL,bold=True)
    wb['KPI']['B4'].number_format='0.0%';wb['Control']['B2'].number_format='mmmm yyyy'
    wb['Flux'].column_dimensions['E'].width=95
    for ws in [wb['Flux'],wb['Change Log']]:
        for row in ws.iter_rows(min_row=2):
            ws.row_dimensions[row[0].row].height=42
            for c in row:c.alignment=Alignment(vertical='center',wrap_text=True)
    wb.active=wb.sheetnames.index('PL Group');wb.calculation.fullCalcOnLoad=True
    path.parent.mkdir(parents=True,exist_ok=True);wb.save(path)
    return bud

def deck(d,path,metrics):
    prs=Presentation();prs.slide_width=Inches(13.333);prs.slide_height=Inches(7.5)
    fresh=d>=datetime(2026,2,1);color='183949' if fresh else '234B66'
    def text(sl,x,y,w,h,t,size=18,bold=False,col=color):
        tx=sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h));tx.text_frame.word_wrap=True;p=tx.text_frame.paragraphs[0];p.text=t;p.font.name='Aptos';p.font.size=Pt(size);p.font.bold=bold;p.font.color.rgb=RGBColor.from_string(col);return tx
    def slide(title,subtitle):
        sl=prs.slides.add_slide(prs.slide_layouts[6]);text(sl,.5,.25,12,.3,'ALDER MECHANICAL  /  MONTHLY OPERATING REVIEW',11,True);text(sl,.5,.85,12,.7,title,30,True);text(sl,.5,1.6,12,.55,subtitle,16);text(sl,.5,7.08,12,.25,f'{d:%B %Y}  |  Fictional sample company  |  Finance / Controller reviewed  |  {len(prs.slides)}',10);return sl
    m=metrics[d];hist=[p for p in metrics if p<=d][-12:]
    sl=slide(f'{d:%B %Y}: operating performance',f'Revenue ${m["revenue"]/1e6:,.2f}m   |   EBITDA ${m["ebitda"]/1e3:,.0f}k   |   EBITDA margin {m["ebitda"]/m["revenue"]:.1%}')
    cd=CategoryChartData();cd.categories=[p.strftime('%b %y') for p in hist];cd.add_series('Revenue ($000)',[metrics[p]['revenue']/1000 for p in hist]);cd.add_series('EBITDA ($000)',[metrics[p]['ebitda']/1000 for p in hist])
    sl.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(.5), Inches(2.4), Inches(8), Inches(4.3),cd)
    text(sl,9,2.5,3.6,3.7,'\n\n'.join(narrative(d)[:2]),17)
    sl=slide('Monthly profit and working capital','Management view • USD • consolidated including elimination entries')
    tbl=sl.shapes.add_table(6,3, Inches(.65), Inches(2.5), Inches(6.6), Inches(3.5)).table
    for i,row in enumerate([['Measure','Current month','Finance lens'],['Revenue',f'${m["revenue"]:,.0f}','External customers'],['EBITDA',f'${m["ebitda"]:,.0f}','Before D&A, interest, tax'],['Net income',f'${m["net_income"]:,.0f}','After tax'],['Cash',f'${m["cash"]:,.0f}','Bank reconciliation'],['Ending headcount',str(m['headcount']),'Month-end payroll roster']]):
        for j,v in enumerate(row):tbl.cell(i,j).text=v
    text(sl,7.8,2.5,4.8,3.8,'\n\n'.join(narrative(d)[-2:]),17)
    sl=slide('Close judgments and follow-ups','The month-end pack keeps its audit trail and supporting schedules.')
    text(sl,.7,2.5,7,3.8,'\n\n'.join(narrative(d)),18)
    text(sl,8.5,2.5,4,3.5,'Controller review\n\nTB and statement tie-outs\nIntercompany eliminations\nPayroll and warranty cut-off\nBank and subledger ties',18)
    sl=slide('Trailing performance and capacity','Seasonality is visible across a full year; acquisition and reorg remain in the history.')
    cd=CategoryChartData();cd.categories=[p.strftime('%b %y') for p in hist];cd.add_series('EBITDA margin',[metrics[p]['ebitda']/metrics[p]['revenue'] for p in hist]);sl.shapes.add_chart(XL_CHART_TYPE.LINE,Inches(.6),Inches(2.4),Inches(8),Inches(4.2),cd)
    text(sl,9,2.6,3.6,3.4,f'{m["headcount"]} employees at month end.\n\nBudget, prior year, T12 and YTD are retained in the close model.\n\n'+('February 2026 template refresh.' if fresh else 'Finance template v3 / September 2023.'),17)
    sl=slide('Cash and capital allocation','Cash is reconciled to the bank; member distributions follow the quarterly treasury review.')
    cd=CategoryChartData();cd.categories=[p.strftime('%b %y') for p in hist];cd.add_series('Closing cash ($000)',[metrics[p]['cash']/1000 for p in hist]);sl.shapes.add_chart(XL_CHART_TYPE.LINE,Inches(.6),Inches(2.4),Inches(8),Inches(4.2),cd)
    text(sl,9,2.5,3.5,3.6,'Quarter-end distributions preserve operating liquidity.\n\nFleet investment and scheduled debt service are shown separately in the cash flow.\n\nBank recs retain deposits in transit and outstanding checks.',17)
    sl=slide('Decisions for next month','Owner actions stay with the monthly pack and are revisited in the next close.')
    tx=sl.shapes.add_table(5,3,Inches(.65),Inches(2.5), Inches(12), Inches(3.6)).table
    for i,row in enumerate([['Owner','Action','Evidence'],['Treasury','Review 13-week liquidity and quarterly distributions','Bank rec + cash flow'],['Service operations','Match warranty claims to subcontractor invoices','Warranty provision / accrual roll'],['Finance','Review payroll cut-off and overtime','Payroll summary + close JE'],['FP&A','Review seasonal staffing against plan','KPI, YTD and trailing 12']]):
        for j,v in enumerate(row):tx.cell(i,j).text=v
    for sl in prs.slides:
        for sh in sl.shapes:
            if sh.has_chart:
                sh.chart.has_legend=True;sh.chart.legend.position=XL_LEGEND_POSITION.BOTTOM
                if fresh:
                    for ix,ser in enumerate(sh.chart.series):
                        ser.format.fill.solid();ser.format.fill.fore_color.rgb=RGBColor.from_string('183949' if ix==0 else '007F7B')
                        ser.format.line.color.rgb=RGBColor.from_string('183949' if ix==0 else '007F7B')
            if sh.has_table and fresh:
                for cell in sh.table.rows[0].cells:
                    cell.fill.solid();cell.fill.fore_color.rgb=RGBColor.from_string('183949')
                    for paragraph in cell.text_frame.paragraphs:paragraph.font.color.rgb=RGBColor.from_string('FFFFFF')
    path.parent.mkdir(parents=True,exist_ok=True);prs.save(path)


def main():
    rows,snapshots,schedules,facts,metrics=ledger();DRIVE.mkdir(exist_ok=True)
    support=DRIVE/'00 Reference';support.mkdir(exist_ok=True)
    csvout(support/'Reporting COA.csv',['Account','Description','Reporting line','Type'],[[a,*v] for a,v in AC.items()])
    csvout(support/'Supporting monthly history 2022-Aug2023.csv',['Period','Entity','Department','Line','USD'],[[r[0].strftime('%Y-%m-%d'),*r[1:]] for r in facts if r[0]<datetime(2023,9,1)])
    (support/'Company and close policies.md').write_text('# Alder Mechanical Services Group\n\nFictional sample company; all figures are synthetic. Calendar financial year; USD. Commercial HVAC maintenance, emergency repair and parts across Midwest branches. 100 Alder Service, 200 Shared Services, 300 Summit (acquired Jan 2025), ELIM consolidation. Reporting uses debit-positive TB and positive revenue in P&L.\n\nMonthly exports include balanced double-entry journals, January prior-year earnings closing, support schedules and an explicit intercompany elimination entity. The archive is a constructed history of a finance process, not a claim that these files existed in 2023.\n\nBlue numbers: inputs. Teal formulas: calculations. All reporting facts are reconciled to monthly GL. Acquisition is an asset purchase: no parent investment account. Straight-line 84-month pooled fleet depreciation. Monthly tax provision 25%; quarterly cash payments. Cost-plus shared services at 4%. Annual insurance premiums amortized monthly.\n')
    anchor=snapshots[datetime(2023,8,1)]
    csvout(support/'Opening TB 2023-08.csv',['Entity','Department','Account','Ending'],[[e,dp,a,dollars(v)] for (e,dp,a),v in sorted(anchor.items()) if v])
    csvout(support/'Supporting cash history 2022-Aug2023.csv',['Period','Entity','Cash USD'],[[d.strftime('%Y-%m'),e,dollars(sum(v for (en,_,a),v in snap.items() if en==e and a=='1000'))] for d,snap in snapshots.items() if d<datetime(2023,9,1) for e in ['100','200']])
    manifest=[];previous=None
    for d in months():
        if d<datetime(2023,9,1):continue
        folder=DRIVE/f'FY{d.year}'/d.strftime('%m %b'); inp=folder/'01 ERP exports';out=folder/'04 Close and review';schdir=folder/'02 Subledgers'
        mr=[r for r in rows if r[0]==end(d).strftime('%Y-%m-%d')]
        csvout(inp/f'GL detail {d:%Y-%m}.csv',['PostingDate','JournalID','Entity','Department','GLAccount','Description','Debit','Credit','Memo','Source'],mr)
        tb=[]
        for (e,dp,a),v in sorted(snapshots[d].items()):
            op=schedules[d]['opening'].get((e,dp,a),0);match=[r for r in mr if r[2:5]==[e,dp,a]]
            if v or op or match:tb.append([e,dp,a,AC[a][0],dollars(op),round(sum(r[6] for r in match),2),round(sum(r[7] for r in match),2),dollars(v)])
        csvout(inp/f'Trial balance {d:%Y-%m}.csv',['Entity','Department','Account','Description','Opening','Debits','Credits','Ending'],tb)
        for kind,a in [('AR','1100'),('AP','2000')]:
            aging=[]
            customers=['Prairie Ridge Logistics','Westhaven Medical Campus','Cedar Industrial Parks','Northline Warehousing','Meadowcrest Foods','Lakeside Property Group','Elmstone Packaging','Crossroads Retail Partners']
            vendors=['Prairie HVAC Supply','Central Parts Distribution','Midwest Trade Services','Summit Fleet Supply','Cedar Controls Supply','Northstar Refrigeration','Valley Mechanical Parts','Regional Building Services']
            for e in schedules[d]['headcount']:
                openlots=[]
                for r in rows:
                    if r[0]>end(d).strftime('%Y-%m-%d') or r[2]!=e or r[4]!=a:continue
                    amount=round((r[6]-r[7])*(1 if kind=='AR' else -1)*100)
                    if amount>0:openlots.append([r[0],r[1],amount])
                    elif amount<0:
                        remaining=-amount
                        for lot in openlots:
                            applied=min(lot[2],remaining);lot[2]-=applied;remaining-=applied
                            if not remaining:break
                        assert remaining==0,(kind,e,d,'overpayment')
                for i,(date,doc,amount) in enumerate(openlots):
                    if not amount:continue
                    dt=datetime.fromisoformat(date);age=(end(d)-dt).days;overdue=max(0,age-30)
                    band='Current' if age<=30 else '1-30 overdue' if overdue<=30 else '31-60 overdue' if overdue<=60 else '61+ overdue'
                    aging.append([e,doc,(customers if kind=='AR' else vendors)[i%8],date,band,dollars(amount)])
            csvout(schdir/f'{kind} aging {d:%Y-%m}.csv',['Entity','Document','Counterparty','Invoice date','Aging band','Open USD'],aging)
        csvout(schdir/f'Payroll summary {d:%Y-%m}.csv',['Entity','Ending FTE','Gross plus benefits USD','Closing payroll accrual'],[[e,h,round(sum(r[6]-r[7] for r in mr if r[2]==e and r[4] in ['5100','5200','6000'] and r[9]!='Year close'),2),dollars(-sum(v for (en,_,a),v in snapshots[d].items() if en==e and a=='2100'))] for e,h in schedules[d]['headcount'].items()])
        bank=[]
        for e in schedules[d]['headcount']:
            v=dollars(sum(x for (en,_,a),x in snapshots[d].items() if en==e and a=='1000'));checks=round(9500*(1 if e=='100' else .4),2);deposits=round(17600*(1 if e=='100' else .4),2)
            bank.append([e,'Operating ****'+('4072' if e=='100' else '9231' if e=='200' else '5386'),round(v+checks-deposits,2),deposits,checks,v,0,'Reviewed / cash ledger'])
        csvout(schdir/f'Bank rec {d:%Y-%m}.csv',['Entity','Masked account','Bank statement','Deposits in transit','Outstanding checks','GL cash','Difference','Review'],bank)
        if d==datetime(2026,10,1):
            (folder/'READ ME - October close task.md').write_text('Roll forward the September FINAL workbook and review deck. Import the October ERP exports. Map new account 6265 Third-party warranty callouts to Warranty, validate JE-202610-017 payroll true-up ($24,600), update history/schedules/flux and refresh the deck. Confirm all checks and subledger ties. Do not overwrite September.\n')
            continue
        suffix='FINAL_R1_CM' if d==datetime(2026,6,1) else 'v2_FINAL_CM' if d.month in [3,9,12] else 'FINAL_CM'
        wp=out/f'AMSG {d:%Y-%m} Close {suffix}.xlsx'; pp=out/f'AMSG {d:%Y-%m} Management review {suffix}.pptx'
        bud=workbook(d,wp,rows,snapshots,schedules,facts,metrics,previous);previous=wp
        deck(d,pp,metrics)
        # Current forecast files retain original budget and separate July reforecast.
        plan=DRIVE/f'FY{d.year}'/'00 Planning'
        if d.month==9 or not plan.exists():
            csvout(plan/f'FY{d.year} Board budget approved.csv',['Period','Line','Budget USD'],[[r[0].strftime('%Y-%m-%d'),r[1],r[2]] for r in bud])
            csvout(plan/f'FY{d.year} July reforecast v2.csv',['Period','Line','Reforecast USD'],[[r[0].strftime('%Y-%m-%d'),r[1],r[3]] for r in bud])
        manifest.append({'period':d.strftime('%Y-%m'),'workbook':str(wp.relative_to(DRIVE)),'deck':str(pp.relative_to(DRIVE)),'metrics':metrics[d]})
        print(d.strftime('%Y-%m'),wp.stat().st_size,flush=True)
    # Preserve the actual superseded June revision, including an explicit $42k bridge.
    d=datetime(2026,6,1);folder=DRIVE/'FY2026'/'06 Jun'/'03 Review trail';folder.mkdir(parents=True,exist_ok=True)
    source=next((DRIVE/'FY2026'/'06 Jun'/'04 Close and review').glob('*.xlsx'))
    wb=load_workbook(source);ws=wb['TB Import']
    # Reverse reserve and related tax provision in the superseded TB / activity facts.
    for row in ws.iter_rows(min_row=2):
        e,dp,a=[c.value for c in row[:3]]
        delta={('100','2150'):42000,('100','6250'):-42000,('100','2200'):-10500,('100','6700'):10500}.get((e,a),0)
        if delta:
            row[7].value+=delta
            if delta>0:row[5].value+=delta
            else:row[6].value-=delta
    ws=wb['Monthly History']
    for row in ws.iter_rows(min_row=2):
        if row[0].value==end(d) and row[1].value=='100':
            if row[3].value=='Warranty':row[4].value-=42000
            if row[3].value=='Tax':row[4].value+=10500
    # Reverse only the extra reserve provision and related tax in the old source snapshot.
    for row in wb['GL Import'].iter_rows(min_row=2):
        if row[2].value!='100':continue
        account=row[4].value
        if row[8].value=='Warranty provision / R1 vendor claim reserve':
            if account=='6250':row[6].value-=42000
            elif account=='2150':row[7].value-=42000
        if row[8].value=='Monthly tax provision at 25%':
            if account=='6700':row[6].value+=10500
            elif account=='2200':row[7].value+=10500
        if row[8].value=='Quarterly estimated income tax':
            if account=='2200':row[6].value+=10500
            elif account=='1000':row[7].value+=10500
    # June tax settlement was paid in the original close. R1 cash tax is lower by $10,500.
    for row in wb['TB Import'].iter_rows(min_row=2):
        if row[0].value=='100' and row[2].value=='2200':
            row[7].value-= -10500;row[5].value+=10500
        if row[0].value=='100' and row[2].value=='1000':
            row[7].value-=10500;row[6].value+=10500
    for row in wb['Accruals'].iter_rows(min_row=2):
        if row[0].value=='100' and row[1].value=='2150':row[4].value-=42000
    wb['Control']['B7']='SUPERSEDED - use R1'
    wb['Control']['B8']='v1 before vendor claim review'
    for row in wb['Flux'].iter_rows(min_row=2):
        if row[0].value=='Warranty':row[4].value='Routine warranty provision; additional vendor claim review not yet completed.'
    wb.save(folder/'AMSG 2026-06 Close v1 SUPERSEDED.xlsx')
    oldmetrics={k:dict(v) for k,v in metrics.items()};oldmetrics[d]['ebitda']+=42000;oldmetrics[d]['net_income']+=31500;oldmetrics[d]['cash']-=10500
    deck(d,folder/'AMSG 2026-06 Management review v1 SUPERSEDED.pptx',oldmetrics)
    csvout(folder/'June restatement bridge.csv',['Measure','v1 USD','R1 USD','Change USD'],[['Incremental warranty provision',0,42000,42000],['Net income',metrics[d]['net_income']+31500,metrics[d]['net_income'],-31500],['June cash tax paid',10500,0,-10500],['Closing cash',metrics[d]['cash']-10500,metrics[d]['cash'],10500]])
    (ROOT/'manifest.json').write_text(json.dumps(manifest,indent=2))
    (ROOT/'model.json').write_text(json.dumps({'metrics':{d.strftime('%Y-%m'):v for d,v in metrics.items()}},indent=2))
    print('Generated 37 monthly packs and October source exports.',flush=True)
if __name__=='__main__':
    main()
    from finalize_sources import finalize
    finalize()

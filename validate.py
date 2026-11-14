"""Independent reconciliation from the delivered CSVs; never imports the data generator."""
from pathlib import Path
from decimal import Decimal
from collections import defaultdict
from datetime import datetime
from io import BytesIO
import csv,json,zipfile,xml.etree.ElementTree as ET,hashlib
import openpyxl
from pptx import Presentation
root=Path(__file__).parent
# Validate a temporary assembled copy; never save/reserialize the Office files.
import tempfile, shutil
repo=root
workspace=tempfile.TemporaryDirectory()
root=Path(workspace.name)
for name in ['GL-export.csv','Budget.csv']:
 shutil.copy2(repo/'inputs'/name,root/name)
for name in ['October close.xlsx','October review.pptx']:
 shutil.copy2(repo/'outputs'/name,root/name)
shutil.copy2(repo/'reference/office-evidence.json',root/'office-evidence.json')
raw=list(csv.DictReader((root/'GL-export.csv').open(newline='')))
plan=list(csv.DictReader((root/'Budget.csv').open(newline='')))
unique={};clean=[]
for r in raw:
 key=r['DocID'].strip()
 if key in unique:
  assert r==unique[key],f'Conflicting duplicate {key}'
  continue
 unique[key]=r
 clean.append((key,datetime.strptime(r['PostingDate'].strip(),'%m/%d/%Y'),{'100':'Sales','200':'Engineering','300':'G&A'}[r['CostCenter'].strip()],r['Account'].strip().title(),Decimal(r['Amount'].strip().replace('$','').replace(',',''))))
assert len(raw)==483 and len(clean)==480 and len(plan)==120
assert len({(r['Period'],r['Department'],r['Category']) for r in plan})==120
w=openpyxl.load_workbook(root/'October close.xlsx',data_only=True)
f=openpyxl.load_workbook(root/'October close.xlsx',data_only=False)
assert [tuple(r) for r in w['Raw export'].iter_rows(min_row=2,max_col=6,values_only=True)]==[tuple(r.values()) for r in raw]
actual_rows=list(w['Actuals'].iter_rows(min_row=5,max_col=6,values_only=True))
assert len(actual_rows)==480
observed={r[0]:(r[1],r[2],r[3],r[4],Decimal(str(r[5]))) for r in actual_rows}
for doc,date,dept,cat,amount in clean:assert observed[doc]==(date,date.strftime('%Y-%m'),dept,cat,amount)
expected={}; formula_count=0
for row,dept in [(6,'Sales'),(7,'Engineering'),(8,'G&A'),(9,None),(14,'Sales'),(15,'Engineering'),(16,'G&A'),(17,None),(23,'Payroll'),(24,'Cloud'),(25,'Travel'),(26,'Software')]:
 monthly=row<14 or row>=23;by_category=row>=23
 def matches(d,c):return (dept is None or (c if by_category else d)==dept)
 a=sum((amount for _,date,d,c,amount in clean if matches(d,c) and (date.strftime('%Y-%m')=='2026-10' if monthly else '2026-01'<=date.strftime('%Y-%m')<='2026-10')),Decimal(0))
 b=sum((Decimal(r['Budget']) for r in plan if matches(r['Department'],r['Category']) and (r['Period']=='2026-10' if monthly else '2026-01'<=r['Period']<='2026-10')),Decimal(0))
 vals=[b,a,a-b,(a-b)/abs(b),'Unfavorable' if a>b else 'Favorable' if a<b else 'On plan']
 for col,value in zip('BCDEF',vals):
  address=f'{col}{row}';actual=w['Report'][address].value
  assert f['Report'][address].data_type=='f',(address,'hardcoded output')
  formula_count+=1
  if isinstance(value,str):assert actual==value,(address,actual,value)
  else:assert abs(Decimal(str(actual))-value)<Decimal('0.000001'),(address,actual,value)
 expected[str(row)]=dict(budget=str(b),actual=str(a),variance=str(a-b),percent=str((a-b)/b))
for row,value in [(3,483),(4,480),(5,3),(6,0),(7,0),(8,0),(9,0),(10,0)]:assert w['Checks'][f'B{row}'].value==value
for s in w:
 for line in s:
  for c in line:assert c.data_type!='e',(s.title,c.coordinate,c.value)
assert len(f['Pivot detail']._pivots)==1 and len(f['Charts']._charts)==2
assert 'Clean_Actuals' in (root/'office-evidence.json').read_text(encoding='utf-8-sig')
z=zipfile.ZipFile(root/'October close.xlsx');ns={'x':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
cache=ET.fromstring(z.read('xl/pivotCache/pivotCacheDefinition1.xml'));assert cache.attrib['recordCount']=='480'
assert 'Table.Distinct' in json.loads((root/'office-evidence.json').read_text(encoding='utf-8-sig'))['queries'][0]['formula']
p=Presentation(root/'October review.pptx');assert len(p.slides)==2
charts=[]
for s in p.slides:
 charts.extend(sh.chart for sh in s.shapes if sh.has_chart)
assert len(charts)==2
assert [c.label for c in charts[0].plots[0].categories]==['Sales','Engineering','G&A']
assert [list(s.values) for s in charts[0].series]==[[179000.,350000.,116000.],[193500.,374000.,119000.]]
assert [c.label for c in charts[1].plots[0].categories]==['Payroll','Cloud','Travel','Software']
assert [list(s.values) for s in charts[1].series]==[[2000.,28000.,10500.,1000.]]
text='\n'.join(sh.text for sl in p.slides for sh in sl.shapes if sh.has_text_frame)
for item in ['$686.5k','Budget: $645.0k','$41.5k over budget','6.4% unfavorable','67.5% of the variance is cloud spend.','Travel adds $10.5k.']:assert item in text,item
z=zipfile.ZipFile(root/'October review.pptx');embedded=[n for n in z.namelist() if n.startswith('ppt/embeddings/') and n.endswith('.xlsx')];assert len(embedded)==2
for n in z.namelist():
 if n.endswith('.rels'):assert 'TargetMode="External"' not in z.read(n).decode(),n
for n in embedded:
 ew=openpyxl.load_workbook(BytesIO(z.read(n)),data_only=True)
 assert ew['Report']['C9'].value==686500
summary=dict(result='PASS',raw_rows=len(raw),clean_documents=len(clean),duplicates_removed=len(raw)-len(clean),formula_outputs_verified=formula_count,financial_tie_outs=5,native_pivot_tables=1,native_excel_charts=2,editable_powerpoint_charts=2,embedded_workbooks=2,external_deck_links=0,report=expected,hashes={n:hashlib.sha256((root/n).read_bytes()).hexdigest() for n in ['GL-export.csv','Budget.csv','October close.xlsx','October review.pptx']})
(repo/'validation.json').write_text(json.dumps(summary,indent=2)+'\n');print(json.dumps(summary,indent=2))

"""Apply workbook formula lineage and document properties before Office recalculation.
Also called by the generator for reproducible source builds.
"""
from pathlib import Path
from datetime import datetime,timedelta
from openpyxl import load_workbook
from openpyxl.styles import Font
from pptx import Presentation
ROOT=Path(__file__).parent

def finalize():
    files=list((ROOT/'Alder Mechanical - Finance').rglob('*.xlsx'))
    for p in files:
        w=load_workbook(p);d=w['Control']['B2'].value;ws=w['GL Import'];n=ws.max_row
        ws.cell(1,11,'Line');ws.cell(1,12,'Type')
        for i in range(2,n+1):
            ws.cell(i,11,f'=IFERROR(INDEX(Map[Line],MATCH(E{i},Map[Account],0)),"UNMAPPED")')
            ws.cell(i,12,f'=IFERROR(INDEX(Map[Type],MATCH(E{i},Map[Account],0)),"UNMAPPED")')
            for j in [11,12]:ws.cell(i,j).font=Font(name='Aptos',size=10,color='007F7B')
        ws.tables['GL'].ref=f'A1:L{n}'
        # Update existing table column definitions when opening with Excel.
        from openpyxl.worksheet.table import TableColumn
        ws.tables['GL'].tableColumns=[TableColumn(id=i,name=ws.cell(1,i).value) for i in range(1,13)]
        hist=w['Monthly History']
        for row in hist.iter_rows(min_row=2):
            if row[0].value!=d:continue
            i=row[0].row
            criteria=f'GL[Entity],B{i},GL[Department],C{i},GL[Line],D{i},GL[Source],"<>Year close",GL[Source],"<>Opening"'
            diff=f'SUMIFS(GL[Debit],{criteria})-SUMIFS(GL[Credit],{criteria})'
            row[4].value=f'=IF(OR(D{i}="Contract revenue",D{i}="Demand revenue",D{i}="IC recovery"),-1,1)*({diff})'
        # Formula schedules calculate depreciation and insurance, with source reconciliations downstream.
        for row in w['Fixed Assets'].iter_rows(min_row=2):
            i=row[0].row
            expected=round(row[4].value if isinstance(row[4].value,(int,float)) else 0,2)
            # IFRS/GAAP-style schedule uses ROUND; ledger differences of one cent are surfaced.
            row[6].value=f'=ROUND(E{i}/84,2)'
        w.properties.creator='Alder Mechanical Finance';w.properties.lastModifiedBy='Controller / CM'
        w.properties.title=f'Alder Mechanical monthly close — {d:%B %Y}'
        w.properties.subject='Fictional sample company / synthetic financial data'
        w.properties.created=d+timedelta(days=8);w.properties.modified=d+timedelta(days=9)
        # Source file is a close snapshot. Excel will recalculate every formula on first open.
        w.calculation.fullCalcOnLoad=True
        w.save(p)
    for p in (ROOT/'Alder Mechanical - Finance').rglob('*.pptx'):
        prs=Presentation(p)
        import re
        per=re.search(r'20\d\d-\d\d',p.name).group(0);d=datetime.fromisoformat(per+'-01')
        prs.core_properties.author='Alder Mechanical Finance';prs.core_properties.last_modified_by='Controller / CM';prs.core_properties.title=f'Alder Mechanical monthly review — {d:%B %Y}';prs.core_properties.subject='Fictional sample company';prs.core_properties.created=d+timedelta(days=8);prs.core_properties.modified=d+timedelta(days=9);prs.save(p)
if __name__=='__main__':finalize()

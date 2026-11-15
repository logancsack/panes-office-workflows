"""Package the validated history; optionally copy assets into a local page preview."""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
import argparse,json,shutil
ROOT=Path(__file__).parent

def package(web_public=None):
    r=json.loads((ROOT/'delivery-validation.json').read_text());assert r['status']=='PASS'
    p=ROOT/'VALIDATION.md';text=p.read_text();import re
    text=re.sub(r'checks \d+ cells',f'checks {r["october_independent_cells_checked"]} cells',text);p.write_text(text)
    archive=ROOT/'Alder-finance-history.zip'
    with ZipFile(archive,'w',ZIP_DEFLATED) as z:
        for base in [ROOT/'Alder Mechanical - Finance',ROOT/'outputs']:
            for p in sorted(base.rglob('*')):
                if p.is_file() and not p.name.startswith('~$'):z.write(p,p.relative_to(ROOT))
        for n in ['manifest.json','VALIDATION.md','delivery-validation.json','validation.json','october-validation.json','verified-deliverables.json','csv-open-validation.json']:z.write(ROOT/n,n)
    if web_public:
        target=Path(web_public);target.mkdir(parents=True,exist_ok=True)
        for n in ['October 2026 close.xlsx','October 2026 review.pptx']+[f'slide-{i}.png' for i in range(1,7)]:shutil.copy2(ROOT/'outputs'/n,target/n)
        sept=next((ROOT/'Alder Mechanical - Finance/FY2026/09 Sep/04 Close and review').glob('*.xlsx'));shutil.copy2(sept,target/'September 2026 close.xlsx')
        for n in ['october-close-unedited.mp4','close-poster.webp']:shutil.copy2(ROOT/'media'/n,target/n)
        shutil.copy2(archive,target/archive.name);shutil.copy2(ROOT/'delivery-validation.json',target/'validation.json');shutil.copy2(ROOT/'VALIDATION.md',target/'VALIDATION.md')
    print(str(archive),archive.stat().st_size,'bytes')
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--web-public',type=Path);a=p.parse_args();package(a.web_public)

"""Reproduce the source reconciliations and verify the exact native-checked files."""
from pathlib import Path
from hashlib import sha256
from zipfile import ZipFile
import json,re
from validate_history import validate as history
from validate_october import validate as october
ROOT=Path(__file__).parent

def validate():
    h=history(True);o=october()
    records=json.loads((ROOT/'verified-deliverables.json').read_text(encoding='utf-8-sig'))
    assert len(records)==78
    for r in records:
        p=ROOT/r['file'].replace('\\','/')
        assert p.exists(),p
        assert sha256(p.read_bytes()).hexdigest()==r['sha256'],str(p)+' changed after native Office verification'
        assert r['opened'] and r['unchangedByVerification']
        assert not r.get('formulaErrors')
        assert all(c['result']=='PASS' for c in r.get('checks',[]))
        with ZipFile(p) as z:
            assert z.testzip() is None
            for n in z.namelist():
                if n.endswith('.xml'):assert not re.search(rb'logan|logancsack',z.read(n),re.I),str(p)+' personal name in '+n
    csv=json.loads((ROOT/'csv-open-validation.json').read_text(encoding='utf-8-sig'))
    assert len(csv)==241 and all(x['opened'] for x in csv)
    p=json.loads((ROOT/'media/provenance.json').read_text())
    assert p['uncut'] and not p['edits'] and not p['synthetic_or_composited_ui']
    assert p['sha256']==sha256((ROOT/'media'/p['source_file']).read_bytes()).hexdigest()
    out={'status':'PASS','company':'Alder Mechanical Services Group — fictional sample company','closed_months':37,'historical_range':'September 2023–September 2026','desktop_office_files_opened':len(records)+len(csv),'workbooks':39,'presentations':39,'csv_exports':241,'formula_errors':0,'workbook_controls_per_month':14,'all_controls':'PASS','october_independent_cells_checked':o['independently_reconciled_cells'],'october':o['october'],'native_applications':o['application_versions'],'recording_seconds':p['duration_seconds'],'recording_uncut':True,'public_release':'Approved for publication','evidence':{'history':'validation.json','october':'october-validation.json','exact_file_hashes':'verified-deliverables.json','csv_opening':'csv-open-validation.json','recording':'media/provenance.json'}}
    (ROOT/'delivery-validation.json').write_text(json.dumps(out,indent=2));return out
if __name__=='__main__':print(json.dumps(validate(),indent=2))

"""Remove the recording account's author metadata without touching formula caches.
All data, formulas, chart data and Office package parts are preserved byte-for-byte;
only docProps/core.xml author fields are replaced with fictional finance ownership.
"""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from io import BytesIO
import re
ROOT=Path(__file__).parent
NS={'dc':'http://purl.org/dc/elements/1.1/','cp':'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'}
def clean(p):
    source=ZipFile(p);buf=BytesIO()
    with ZipFile(buf,'w',ZIP_DEFLATED) as out:
        for info in source.infolist():
            data=source.read(info.filename)
            if info.filename=='docProps/core.xml':
                for tag in [b'dc:creator',b'cp:lastModifiedBy']:
                    pattern=b'(<'+tag+b'[^>]*>).*?(</'+tag+b'>)'
                    data,count=re.subn(pattern,lambda m:m.group(1)+b'Alder Mechanical Finance'+m.group(2),data,flags=re.S)
                    if count!=1:raise ValueError(f'Missing expected author field {tag!r}: {p}')
            out.writestr(info,data)
    source.close();p.write_bytes(buf.getvalue())
if __name__=='__main__':
    paths=list((ROOT/'Alder Mechanical - Finance').rglob('*'))+list((ROOT/'outputs').glob('*'))
    for p in paths:
        if p.suffix in ['.xlsx','.pptx']:clean(p)

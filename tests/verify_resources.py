from pathlib import Path
from io import BytesIO
import hashlib
import json
import struct
from PIL import Image
import pefile

ROOT = Path(__file__).resolve().parent.parent
expected = {16,20,24,32,40,48,64,96,128,256}
checks = []
with Image.open(ROOT/'assets'/'OfficePDF.ico') as ico:
    assert ico.ico.sizes() == {(s,s) for s in expected}
gui_candidates = list(ROOT.glob('Office资料一键合并PDF_*.exe'))
if not gui_candidates:
    raise FileNotFoundError("Cannot find Office资料一键合并PDF_*.exe in " + str(ROOT))
for executable in [gui_candidates[0], ROOT/'OfficePdfWorker.exe']:
    pe=pefile.PE(str(executable))
    types={entry.id:entry.directory for entry in pe.DIRECTORY_ENTRY_RESOURCE.entries}
    assert 3 in types and 14 in types
    groups=types[14].entries
    assert len(groups)==1
    data=groups[0].directory.entries[0].data.struct
    group=pe.get_data(data.OffsetToData,data.Size)
    reserved,kind,count=struct.unpack_from('<HHH',group)
    assert reserved==0 and kind==1 and count==10
    sizes=set()
    for index in range(count):
        width,height,colors,unused,planes,bits,length,rid=struct.unpack_from('<BBBBHHIH',group,6+14*index)
        size=width or 256; assert size==(height or 256); sizes.add(size)
        entry=next(e for e in types[3].entries if e.id==rid).directory.entries[0].data.struct
        payload=pe.get_data(entry.OffsetToData,entry.Size)
        with Image.open(BytesIO(payload)) as image:
            assert image.size==(size,size) and image.mode=='RGBA'
            # Lanczos filtering can leave alpha=1 at the extreme corner.
            assert image.getpixel((0,0))[3]<8 and image.getchannel('A').getextrema()[0]==0
    assert sizes==expected
    checks.append({'exe':executable.name,'icon_sizes':sorted(sizes),'sha256':hashlib.sha256(executable.read_bytes()).hexdigest()})
(ROOT/'tests'/'icon-resource-report.json').write_text(json.dumps(checks,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(checks,ensure_ascii=False,indent=2))

from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
import re,sys
from assemble import paths

def normalise(source,dest):
 with ZipFile(source) as z,ZipFile(dest,'w',ZIP_DEFLATED,compresslevel=6) as zz:
  _,pp=paths(z);parts=set(pp.values())
  for part in z.infolist():
   data=z.read(part.filename)
   if part.filename in parts:
    target=b'1' if part.filename==pp['EXEC V06'] else b'0'
    data=re.sub(rb'(<sheetView\b[^>]*\btabSelected=")(?:0|1|false|true)(")',lambda m:m[1]+target+m[2],data)
   zz.writestr(part.filename,data)
 print('Single selected dashboard',dest)
if __name__=='__main__':normalise(Path(sys.argv[1]),Path(sys.argv[2]))

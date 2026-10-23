"""Controlled version labels only; numeric cells and engine parts remain intact."""
from bank_r6 import *
from zip_copy import rewrite

def label(file,version):
 changes={}
 with ZipFile(file) as z:
  for part in z.namelist():
   if part=='xl/sharedStrings.xml' or part.startswith('xl/worksheets/') and part.endswith('.xml'):
    data=z.read(part);new=data.replace(b'R5 CHECKPOINT',version.encode())
    if new!=data:changes[part]=new
 if changes:
  tmp=Path(file).with_suffix('.labels.xlsx');rewrite(file,tmp,changes);tmp.replace(file)
 return len(changes)
if __name__=='__main__':print('Version labels updated',label(BM,'R6.1 BANK CHECKPOINT'))

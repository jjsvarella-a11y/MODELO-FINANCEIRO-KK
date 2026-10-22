"""Equivalent explicit fee references in draw algebra; no cached results inserted."""
from engine import ROOT
import zipfile,re,os
P=ROOT/'entrega/KK01_Modelo_Integrado_V04.xlsx';count=0
with zipfile.ZipFile(P) as zi,zipfile.ZipFile(P.with_suffix('.draw'),'w',zipfile.ZIP_DEFLATED) as zo:
 for inf in zi.infolist():
  data=zi.read(inf.filename)
  if inf.filename.startswith('xl/worksheets/') and inf.filename.endswith('.xml'):
   text=data.decode('utf-8')
   def fix(m):
    global count
    cell=m.group();rr=m.group(1);pattern=f'SUM(T{rr}:V{rr})'
    if pattern in cell:
     cell=cell.replace(pattern,f'(T{rr}+U{rr}+V{rr})');cell=re.sub(r'<v>.*?</v>','',cell);count+=1
    return cell
   text=re.sub(r'<c\b[^>]*\br="X(\d+)"[^>]*>.*?</c>',fix,text);data=text.encode('utf-8')
  zo.writestr(inf,data)
os.replace(P.with_suffix('.draw'),P);print(count,'draw formula dependency expressions updated; full recalc required.')

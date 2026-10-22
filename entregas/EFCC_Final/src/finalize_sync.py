"""Refresh verified calculation caches while preserving original formula text and styles."""
import json,re,hashlib
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
from assemble import inject_caches
from validate_final import read
from premium_ui import ROOT,WORK

def main():
 source=ROOT/'KK01_V06_Executive_Master.xlsx';calculated=WORK/'premium_final_recalc/out/KK01_V06_Executive_Master.xlsx'
 target=source.with_suffix('.cached.xlsx');count=inject_caches(source,calculated,target);target.replace(source)
 with ZipFile(calculated) as z:cells,errors,_=read(z)
 assert not errors
 target=source.with_suffix('.charts.xlsx');ns={'c':'http://schemas.openxmlformats.org/drawingml/2006/chart'};C=lambda n:'{'+ns['c']+'}'+n;points=0
 with ZipFile(source) as z,ZipFile(target,'w',ZIP_DEFLATED,compresslevel=6) as out:
  for item in z.infolist():
   data=z.read(item.filename)
   if re.fullmatch(r'xl/charts/chart\d+.xml',item.filename):
    root=E.fromstring(data)
    for ref in root.findall('.//c:numRef',ns):
     formula=ref.findtext('c:f',namespaces=ns) or '';match=re.fullmatch(r"(?:'([^']+)'|([^!]+))!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)",formula)
     assert match,formula
     sn=match[1] or match[2];assert match[3]==match[5],formula
     cache=ref.find(C('numCache'))
     if cache is None:cache=E.SubElement(ref,C('numCache'))
     for old in list(cache):cache.remove(old)
     E.SubElement(cache,C('formatCode')).text='General';E.SubElement(cache,C('ptCount'),val=str(int(match[6])-int(match[4])+1))
     for i,row in enumerate(range(int(match[4]),int(match[6])+1)):
      value=cells[sn][match[3]+str(row)];assert isinstance(value,(int,float)),(sn,row,value)
      point=E.SubElement(cache,C('pt'),idx=str(i));E.SubElement(point,C('v')).text=str(value);points+=1
    data=E.tostring(root,encoding='UTF-8',xml_declaration=True,standalone=True)
   out.writestr(item,data)
 target.replace(source)
 proof=json.loads((ROOT/'Integridade_Premium.json').read_text());proof.update(sha256=hashlib.sha256(source.read_bytes()).hexdigest(),all_existing_engine_caches_unchanged=False,cache_update='Required correction of Base/F&B/MICE scenario propagation; original engine formula expressions preserved',formula_caches_refreshed=count,chart_points_refreshed=points,financial_reference='Referencia_Python_V06_SYNC.json.gz');(ROOT/'Integridade_Premium.json').write_text(json.dumps(proof,indent=2,ensure_ascii=False));print(count,points)
if __name__=='__main__':main()

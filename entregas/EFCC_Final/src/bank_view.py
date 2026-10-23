"""One active tab; exact percent units for banking inputs. No financial changes."""
from bank_r6 import *
from zip_copy import rewrite

def apply(file=None,percent=True):
 file=Path(file or BM);changes={}
 with ZipFile(file) as z:
  wb,pp=paths(z);names=[n.get('name') for n in wb.find(T('sheets'))]
  for view in wb.find(T('bookViews')):view.set('activeTab',str(names.index('EXEC V06')));view.set('firstSheet',str(names.index('EXEC V06')))
  changes['xl/workbook.xml']=xml(wb)
  for sn,p in pp.items():
   raw=z.read(p)
   if sn=='EXEC V06':
    # Set/add only the normal worksheet view's selected flag.
    root=E.fromstring(raw)
    for v in root.findall(T('sheetViews')+'/'+T('sheetView')):v.set('tabSelected','1')
    changes[p]=xml(root)
   elif b'tabSelected="1"' in raw or b'tabSelected="true"' in raw:changes[p]=raw.replace(b'tabSelected="1"',b'tabSelected="0"').replace(b'tabSelected="true"',b'tabSelected="0"')
  if percent:
   st=E.fromstring(z.read('xl/styles.xml'));ss=[''.join(n.itertext()) for n in E.fromstring(z.read('xl/sharedStrings.xml'))];style={}
   for c in E.fromstring(z.read(pp[ACT])).find(T('sheetData'))[0]:
    if c.get('t')=='s':style[ss[int(c.findtext(T('v')))]]=int(c.get('s'))
   xfs=st.find(T('cellXfs'));fmt=copy.deepcopy(xfs[style['input_pct']]);fmt.set('numFmtId','10');fmt.set('applyNumberFormat','1');pct=len(xfs);xfs.append(fmt);wholefmt=copy.deepcopy(xfs[style['input_number']]);wholefmt.set('numFmtId','1');wholefmt.set('applyNumberFormat','1');whole=len(xfs);xfs.append(wholefmt);xfs.set('count',str(len(xfs)));changes['xl/styles.xml']=xml(st)
   s=Sheet(E.fromstring(changes.get(pp[BN],z.read(pp[BN]))),{})
   for key,pt,en,rv,addr,kind in PARAMS:
    if kind=='pct':s.cells[addr].set('s',str(pct))
    elif kind=='whole':s.cells[addr].set('s',str(whole))
   changes[pp[BN]]=s.finish()
 tmp=file.with_suffix('.view.xlsx');rewrite(file,tmp,changes);tmp.replace(file);print('Active tab and bank input units checked',flush=True)
if __name__=='__main__':apply()

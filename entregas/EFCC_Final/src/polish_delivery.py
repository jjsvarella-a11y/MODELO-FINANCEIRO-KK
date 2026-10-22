from reference_io import load_reference
"""Compact nine dashboards and chart annual aggregates without changing the engine.
Only presentation rows, print areas and chart series are changed; original monthly
engine remains the source. Add a formula-only annual chart projection.
"""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import json,sys,copy,re
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name as col
from assemble import paths,T,xml
sys.path.insert(0,str(Path(__file__).resolve().parents[2]/'V06/src'))
from merge import merge
ROOT=Path(__file__).resolve().parents[1];WORK=Path('/workspace/kk01_efcc');C='http://schemas.openxmlformats.org/drawingml/2006/chart';CT=lambda n:'{'+C+'}'+n
DASHES=['EXEC V06','DASH 01 V06','DASH 02 V06','DASH 03 V06','DASH 04A V06','DASH 04B V06','DASH 05 V06','DASH 06 V06','DASH 07 V06','DASH 10 V06']
KEYS=['cash','reserve','capex','construction_insurance','vat_credit','tax_paid','debt','receipts','deed_sales','owner_distribution','owner_target','revenue','ebitda','equity','investor_outlay','dividend']
STOCK={'cash','reserve','vat_credit','debt'}
def polish(source,output):
 ref=load_reference(ROOT);cc=ref['main_columns'];rr=ref['main_rows'];helper=WORK/(Path(output).stem+'_UI_Annual.xlsx');wb=xlsxwriter.Workbook(str(helper));ws=wb.add_worksheet('V06 UI Anual');ws.write_row(3,0,['Year']+KEYS);vals={};fmt=wb.add_format({'num_format':'#,##0.00'})
 for y in range(1,16):
  ws.write(y+3,0,y)
  for j,key in enumerate(KEYS,1):
   expr=f"'V06 Motor'!{col(cc[key])}{y*12+4}" if key in STOCK else f"SUM('V06 Motor'!{col(cc[key])}{(y-1)*12+5}:{col(cc[key])}{y*12+4})";value=rr[y*12-1][key] if key in STOCK else sum(r[key] for r in rr[(y-1)*12:y*12]);ws.write_formula(y+3,j,'='+expr,fmt,value);vals.setdefault(key,[]).append(value)
 wb.close();merged=Path(output).with_suffix('.polish.xlsx');merge(source,helper,merged)
 with ZipFile(merged) as z,ZipFile(output,'w',ZIP_DEFLATED,compresslevel=6) as zz:
  w,pp=paths(z);sheets=w.find(T('sheets'));order={s.get('name'):s for s in sheets};oldnames=[s.get('name') for s in sheets];newnames=DASHES+[n for n in oldnames if n not in DASHES];sheets[:]=[order[n] for n in newnames]
  for dn in w.findall(T('definedNames')+'/'+T('definedName')):
   if dn.get('localSheetId') is not None:dn.set('localSheetId',str(newnames.index(oldnames[int(dn.get('localSheetId'))])))
  for v in w.find(T('bookViews')):v.set('activeTab','0');v.set('firstSheet','0')
  changes={'xl/workbook.xml':xml(w)}
  for sn in DASHES[:9]:
   root=E.fromstring(z.read(pp[sn]));sd=root.find(T('sheetData'))
   for row in sd:
    r=int(row.get('r'));h=10.5
    # Single-row input controls need enough height for text. Labels/cards merge.
    if sn=='DASH 05 V06' and r in [23,25,27,29,32,34,36]:h=14
    if sn in ['DASH 04B V06','DASH 07 V06'] and r in [23,25]:h=14
    if r in [5,20,34,46]:h=12
    if r in [2,3,4]:h=16
    row.set('ht',str(h));row.set('customHeight','1')
   for v in root.findall(T('sheetViews')+'/'+T('sheetView')):v.set('tabSelected','1' if sn=='EXEC V06' else '0')
   ps=root.find(T('pageSetup'));ps.set('orientation','landscape');ps.set('paperSize','9');ps.set('fitToWidth','1');ps.set('fitToHeight','1')
   changes[pp[sn]]=xml(root)
  # Make charts fit the compact rows: 175px height; keep their existing anchors.
  for part in z.namelist():
   if part.startswith('xl/drawings/drawing') and part.endswith('.xml'):
    data=z.read(part);root=E.fromstring(data);ns={'xdr':'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing','c':C}
    changed=False
    for anchor in root:
     chart=anchor.find('.//'+CT('chart'));fr=anchor.find('{'+ns['xdr']+'}from');to=anchor.find('{'+ns['xdr']+'}to')
     if chart is not None and fr is not None and to is not None:
      row=int(fr.find('{'+ns['xdr']+'}row').text)
      # Only new dashboard E35 / E47 chart positions, never legacy drawings.
      if row in [34,46] and int(fr.find('{'+ns['xdr']+'}col').text)==4:
       to.find('{'+ns['xdr']+'}row').text=str(row+12);to.find('{'+ns['xdr']+'}rowOff').text='0'
       xf=anchor.find('{'+ns['xdr']+'}graphicFrame/{'+ns['xdr']+'}xfrm')
       if xf is not None:
        ext=xf.find('{http://schemas.openxmlformats.org/drawingml/2006/main}ext')
        if ext is not None:ext.set('cy',str(126*12700))
       changed=True
    if changed:changes[part]=xml(root)
   if part.startswith('xl/charts/chart') and part.endswith('.xml'):
    root=E.fromstring(z.read(part));changed=False
    for ser in root.iter(CT('ser')):
     vf=ser.find(CT('val')+'/'+CT('numRef')+'/'+CT('f'))
     if vf is None or vf.text is None:continue
     m=re.fullmatch(r"'?V06 Motor'?!(?:\$)?([A-Z]+)\$?5:\$?\1\$?184",vf.text)
     if not m:continue
     key=next((k for k in KEYS if col(cc[k])==m[1]),None)
     if key is None:continue
     vf.text=f"'V06 UI Anual'!${col(KEYS.index(key)+1)}$5:${col(KEYS.index(key)+1)}$19"
     valref=ser.find(CT('val')+'/'+CT('numRef'));cache=valref.find(CT('numCache'))
     if cache is not None:valref.remove(cache)
     cache=E.SubElement(valref,CT('numCache'));E.SubElement(cache,CT('formatCode')).text='General';E.SubElement(cache,CT('ptCount'),val='15')
     for i,v in enumerate(vals[key]):pt=E.SubElement(cache,CT('pt'),idx=str(i));E.SubElement(pt,CT('v')).text=str(v)
     cat=ser.find(CT('cat'))
     if cat is not None:
      for child in list(cat):cat.remove(child)
      cr=E.SubElement(cat,CT('numRef'));E.SubElement(cr,CT('f')).text="'V06 UI Anual'!$A$5:$A$19";cache=E.SubElement(cr,CT('numCache'));E.SubElement(cache,CT('formatCode')).text='0';E.SubElement(cache,CT('ptCount'),val='15')
      for i in range(15):pt=E.SubElement(cache,CT('pt'),idx=str(i));E.SubElement(pt,CT('v')).text=str(i+1)
     changed=True
    if changed:
     for skip in root.iter(CT('tickLblSkip')):skip.set('val','1')
     for skip in root.iter(CT('tickMarkSkip')):skip.set('val','1')
     changes[part]=xml(root)
  for part in z.infolist():zz.writestr(part.filename,changes.get(part.filename,z.read(part.filename)))
 merged.unlink();print('Compact UI + annual charts',output)
if __name__=='__main__':polish(Path(sys.argv[1]),Path(sys.argv[2]))

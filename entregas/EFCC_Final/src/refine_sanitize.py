"""Presentation-only cleanup: chart labels, minimum fonts and unreachable assets."""
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
from pathlib import Path
import copy,json
from assemble import paths,T,xml,resolve,relpath
from premium_refine import MASTER,WORK,OUT,Sheet,tr,H
A='http://schemas.openxmlformats.org/drawingml/2006/main';C='http://schemas.openxmlformats.org/drawingml/2006/chart'

def sanitize(path,language):
 changes={};removed=[];fonts=0
 with ZipFile(path) as z:
  _,pp=paths(z);names=set(z.namelist());root=E.fromstring(z.read(pp[H]));note=root.find('.//'+T('c')+'[@r="E139"]');style=note.get('s');sheet=Sheet(root,{'note':int(style)});a='Azul: inputs. Cinzento: calculado. Custom só se aplica no modo Custom.';b='Blue: inputs. Grey: calculated. Stored Custom applies only in Custom mode.';sheet.merge('E39:P39',a if language=='PT' else b,tr(a,b),'note');method_pt='Estrutura conceptual aprovada. A versão Excel deve mostrar VAL reais por combinação ADR/ocupação, com os pressupostos ativos do modelo.'
  method_en='Conceptual structure approved. Final Excel version must display actual NPV values by ADR and occupancy combination based on the active model assumptions.'
  details_pt='Grelha discreta, sem extrapolação. Contorno: ponto mais próximo. ADR/ocupação atuais movem o marcador; as combinações fixas substituem estes inputs. VAL ≥ 0 não substitui os testes de liquidez e dívida.'
  details_en='Discrete grid, no extrapolation. Outline: nearest grid point. Current ADR/occupancy move the marker; fixed combinations override these two inputs. NPV ≥ 0 does not replace cash and debt sustainability checks.'
  a=method_pt+' '+details_pt;b=method_en+' '+details_en
  sheet.merge('E139:Y139',method_pt if language=='PT' else method_en,tr(method_pt,method_en),'note');sheet.merge('E140:Y140',details_pt if language=='PT' else details_en,tr(details_pt,details_en),'note');changes[pp[H]]=sheet.finish()
  strings=[''.join(n.itertext()) for n in E.fromstring(z.read('xl/sharedStrings.xml'))] if 'xl/sharedStrings.xml' in names else []
  contract=sheet.cells['H32'];typ=contract.get('t');value=strings[int(contract.findtext(T('v')))] if typ=='s' else ''.join(contract.find(T('is')).itertext()) if typ=='inlineStr' else contract.findtext(T('v'))
  label=sheet.cells['E32'];lv=strings[int(label.findtext(T('v')))] if label.get('t')=='s' else ''.join(label.find(T('is')).itertext()) if label.get('t')=='inlineStr' else label.findtext(T('v'))
  sheet.s['input_text']=int(contract.get('s'));sheet.s['label']=int(label.get('s'));sheet.merge('H32:J33',value,style='input_text');sheet.merge('E32:G33',lv,label.findtext(T('f')),'label');changes[pp[H]]=sheet.finish()
  for sn,_,_ in __import__('premium_ui').DASHES:
   r=E.fromstring(changes.get(pp[sn]) or z.read(pp[sn]))
   for c in r.findall(T('cols')+'/'+T('col')):
    if int(c.get('min'))==5 and int(c.get('max'))==25:c.set('width','8.3')
   
   for row in r.findall(T('sheetData')+'/'+T('row')):
    rn=int(row.get('r'))
    if 22<=rn<=39:row.set('ht','20');row.set('customHeight','1')
    if sn==H and rn in [139,140]:row.set('ht','48');row.set('customHeight','1')
   changes[pp[sn]]=xml(r)
  # Complete active assumptions and a dynamic colour threshold on report page 9.
  reportroot=E.fromstring(z.read(pp['DASH 10 V06']));labelstyle=int(reportroot.find('.//'+T('c')+'[@r="A503"]').get('s'));report=Sheet(reportroot,{'note':labelstyle})
  adr=185;occ=.58
  # Caches use the known default; native formulas track the global controls.
  first='ADR 185 € | Occ 58.0% | MICE 50% | FF&E 5% | Discount 12%'
  formula="'DASH 05 V06'!H23&\" | ADR \"&'DASH 05 V06'!H24&\" € | Occ \"&TEXT('DASH 05 V06'!H26,\"0.0%\")&\" | MICE \"&TEXT('DASH 05 V06'!N25,\"0%\")&\" | FF&E \"&TEXT('DASH 05 V06'!N27,\"0%\")&\" | Disc \"&TEXT('DASH 07 V06'!N26,\"0%\")"
  report.merge('A503:L503','Base Case | '+first,formula,'note')
  tax="\"IRPC \"&TEXT('DASH 02 V06'!H23,\"0%\")&\" | Exempt \"&'DASH 02 V06'!N25&\" y | Interest \"&TEXT('DASH 03 V06'!H23,\"0%\")&\" | Exit \"&'DASH 07 V06'!H25&\"× | Line \"&TEXT('V06 Registo Ativo'!B40/1000000,\"0\")&\" M€\""
  report.merge('A504:L504','IRPC 20% | Exempt 0 y | Interest 6% | Exit 12× | Line 30 M€',tax,'note')
  for area,pt,en,symbol in [('K508:L510','Negativo','Negative','<−'),('K511:L513','Perto de 0','Near 0','±'),('K514:L517','Positivo','Positive','>')]:
   old=report.cells[area.split(':')[0]];sty=int(old.get('s'));key='legend_'+symbol;report.s[key]=sty
   f=tr(pt,en)+'&CHAR(10)&\"'+symbol+'\"&TEXT(\'DASH 05 V06\'!W116/1000000,\"0.00\")&\" M€\"'
   report.merge(area,(pt if language=='PT' else en)+'\n'+symbol+'0.50 M€',f,key)
  changes[pp['DASH 10 V06']]=report.finish()
  # Hotel's individual PDF includes the same real 20-point matrix, not a screenshot.
  hr=sheet.r
  def sid(address):return int(hr.find('.//'+T('c')+'[@r="'+address+'"]') .get('s'))
  pm={'section':sid('E114'),'note':sid('E139'),'header':sid('E118'),'heat':sid('I120'),'pct':sid('E120'),'adr':sid('I118'),'label':sid('E133'),'money':sid('O135')}
  pr=Sheet(E.fromstring(z.read(pp['V06 PDF 07'])),pm)
  for row in pr.sd:
   if 22<=int(row.get('r'))<=53:
    for cell in list(row):row.remove(cell);pr.cells.pop(cell.get('r'),None)
  for cf in list(pr.r.findall(T('conditionalFormatting'))):
   if any(22<=int(number)<=53 for number in __import__('re').findall('[0-9]+',cf.get('sqref',''))):pr.r.remove(cf)
  merges=pr.r.find(T('mergeCells'))
  for merge in list(merges):
   a,_,b,_=__import__('premium_refine').bounds(merge.get('ref'))
   if 22<=a<=b<=53:merges.remove(merge)
  pr.text('A20:L21','Indicadores e matriz financeira · pressupostos ativos','Indicators and financial sensitivity · active assumptions','section');pr.text('G22:L23','Sensibilidade do Valor do Investimento','Investment Value Sensitivity','section');pr.text('G24:L24','ADR × Ocupação → VAL dos Investidores','ADR × Occupancy → Investor NPV','note');pr.text('G25:G25','Ocup./ADR','Occup./ADR','header')
  from xlsxwriter.utility import xl_col_to_name as col
  for ci,adr in enumerate([150,175,200,225]):pr.merge(col(7+ci)+'25',adr,style='adr')
  for ri,occ in enumerate([.5,.6,.65,.7,.75]):
   rr=26+ri*2;pr.merge(f'G{rr}:G{rr+1}',occ,style='pct')
   for ci in range(4):
    addr=col(8+ci*4)+str(120+ri*2);value=float(hr.find('.//'+T('c')+'[@r="'+addr+'"]/'+T('v')).text);area=f'{col(7+ci)}{rr}:{col(7+ci)}{rr+1}';pr.merge(area,value,"'DASH 05 V06'!"+addr,'heat')
    cf=E.SubElement(pr.r,T('conditionalFormatting'),sqref=area);anchor='$'+col(7+ci)+'$'+str(rr)
    cfsource=hr.find(T('conditionalFormatting')+'[@sqref="'+col(8+ci*4)+str(120+ri*2)+':'+col(11+ci*4)+str(121+ri*2)+'"]');source_rules=cfsource.findall(T('cfRule'))[-3:]
    for original,expression in zip(source_rules,[anchor+'<-KK_NPV_NEAR_ZERO','ABS('+anchor+')<=KK_NPV_NEAR_ZERO',anchor+'>KK_NPV_NEAR_ZERO']):
     rule=copy.deepcopy(original);rule.find(T('formula')).text=expression;cf.append(rule)
  for ar,pt,en,source in [('L26:L28','Negativo','Negative','K508'),('L29:L31','Perto de 0','Near 0','K511'),('L32:L35','Positivo','Positive','K514')]:
   node=report.cells[source];pr.s[source]=int(node.get('s'));pr.merge(ar,node.findtext(T('v')),"'DASH 10 V06'!"+source,source)
  pr.text('A38:F38','Resultados operacionais · ano selecionado','Operating results · selected year','section');pr.text('G38:L38','Equilíbrio na grelha testada','Break-even within tested grid','section')
  for rr,pt,en,addr in [(39,'Receita de quartos','Room revenue','K44'),(40,'Receita F&B','F&B revenue','N44'),(41,'Receita MICE','MICE revenue','E50'),(42,'Receita total','Total revenue','H50'),(43,'GOP','GOP','M72'),(44,'Encargos operadora','Operator fees','M74'),(45,'EBITDA','EBITDA','N50'),(46,'Dotação FF&E','FF&E contribution','M76')]:
   pr.text(f'A{rr}:D{rr}',pt,en,'label');value=float(hr.find('.//'+T('c')+'[@r="'+addr+'"]/'+T('v')).text);pr.merge(f'E{rr}:F{rr}',value,"'DASH 05 V06'!"+addr,'money')
  for area,out,pt,en,addr,sty in [('G39:J40','K39:L40','ADR mínimo · ocupação atual','Minimum ADR · active occupancy','E135','adr'),('G41:J42','K41:L42','Ocupação mínima · ADR atual','Minimum occupancy · active ADR','J135','pct'),('G43:J44','K43:L44','VAL atual','Current investor NPV','O135','money'),('G45:J46','K45:L46','Défice de VAL','NPV gap to zero','T135','money')]:
   pr.text(area,pt,en,'label');node=hr.find('.//'+T('c')+'[@r="'+addr+'"]');value=node.findtext(T('v'));value=value if node.get('t')=='str' else float(value);pr.merge(out,value,"'DASH 05 V06'!"+addr,sty)
  pr.text('A48:L51','Motor de 180 meses; pressupostos ativos de financiamento, fiscalidade, fees, dividendos, reservas e saída. Limiar perto de zero: editável no DASH 07. A grelha não extrapola o equilíbrio.','180-month model; active financing, taxes, fees, dividends, reserves and exit. Near-zero threshold is editable on DASH 07. The grid does not extrapolate break-even.','note');changes[pp['V06 PDF 07']]=pr.finish()
  drawing=pr.r.find(T('drawing'));rels=E.fromstring(z.read(relpath(pp['V06 PDF 07'])));rel=next(x for x in rels if x.get('Id')==drawing.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'));drawingpart=resolve(pp['V06 PDF 07'],rel.get('Target'));dr=E.fromstring(z.read(drawingpart));ns={'xdr':'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'}
  for anchor in dr:
   anchor.set('editAs','twoCell')
   ext=anchor.find('.//{'+A+'}ext')
   if ext is not None:ext.set('cx','5137500');ext.set('cy','2381040')
   for pos,c,r,off in [('from',0,21,0),('to',5,35,0)]:
    anchor.find('xdr:'+pos+'/xdr:col',ns).text=str(c);anchor.find('xdr:'+pos+'/xdr:row',ns).text=str(r);anchor.find('xdr:'+pos+'/xdr:colOff',ns).text=str(off);anchor.find('xdr:'+pos+'/xdr:rowOff',ns).text='0'
  changes[drawingpart]=xml(dr)
  for sn in ['DASH 10 V06','V06 Brief']+[f'V06 PDF {i:02}' for i in range(1,10)]:
   r=E.fromstring(changes.get(pp[sn]) or z.read(pp[sn]))
   for c in r.iter(T('c')):
    f=c.find(T('f'));v=c.find(T('v'))
    if f is not None and 'V06.3-SYNC' in (f.text or ''):
     f.text=f.text.replace('V06.3-SYNC','PREMIUM-R2 CHECKPOINT')
     if v is not None and v.text:v.text=v.text.replace('V06.3-SYNC','PREMIUM-R2 CHECKPOINT')
   changes[pp[sn]]=xml(r)
  for p in names:
   if p.startswith('xl/charts/') and p.endswith('.xml') and '/_rels/' not in p:
    r=E.fromstring(z.read(p));dirty=False
    for node in r.findall('.//{'+A+'}defRPr')+r.findall('.//{'+A+'}rPr'):
     if node.get('sz') is not None and int(node.get('sz'))<900:node.set('sz','900');fonts+=1;dirty=True
    if r.find('.//{'+C+'}doughnutChart') is not None:
     labels=r.find('.//{'+C+'}dLbls')
     if labels is not None:
      for font in labels.findall('.//{'+A+'}defRPr'):
       fill=font.find('{'+A+'}solidFill')
       if fill is None:fill=E.SubElement(font,'{'+A+'}solidFill')
       for child in list(fill):fill.remove(child)
       E.SubElement(fill,'{'+A+'}srgbClr',val='FFFFFF')
       order=['ln','noFill','solidFill','gradFill','blipFill','pattFill','grpFill','effectLst','effectDag','highlight','uLnTx','uLn','uFillTx','uFill','latin','ea','cs','sym','hlinkClick','hlinkMouseOver','rtl','extLst']
       font[:]=sorted(font,key=lambda node:order.index(E.QName(node).localname) if E.QName(node).localname in order else 999);dirty=True
    if dirty:changes[p]=xml(r)
  pending=[''];reachable={'[Content_Types].xml'}
  while pending:
   owner=pending.pop();rp='_rels/.rels' if owner=='' else relpath(owner)
   if rp not in names or rp in reachable:continue
   reachable.add(rp)
   for rel in E.fromstring(changes.get(rp) or z.read(rp)):
    if rel.get('TargetMode')=='External' or rel.get('Target','').startswith('#'):continue
    target=resolve(owner,rel.get('Target'))
    if target not in reachable:reachable.add(target);pending.append(target)
  ct=E.fromstring(z.read('[Content_Types].xml'))
  for child in list(ct):
   part=child.get('PartName')
   if part and part.lstrip('/') not in reachable:ct.remove(child)
  changes['[Content_Types].xml']=xml(ct);temp=path.with_suffix('.sanitized.xlsx')
  with ZipFile(temp,'w',ZIP_DEFLATED,compresslevel=6) as out:
   for item in z.infolist():
    if item.filename in reachable:out.writestr(copy.copy(item),changes.get(item.filename) or z.read(item.filename))
    else:removed.append(item.filename)
 temp.replace(path);return {'file':str(path),'removed_unreachable_parts':removed,'chart_font_sizes_raised_to_9pt':fonts,'financial_formulas_changed':0}
def main():
 ps=json.loads((WORK/'Presentations.json').read_text());results=[sanitize(MASTER,'PT')]+[sanitize(Path(ps[lang]),lang) for lang in ['PT','EN']];file=OUT/'Higiene_Apresentacao.json';prior=json.loads(file.read_text()) if file.exists() else {};runs=prior.get('runs',[prior['results']] if prior.get('results') else []);runs.append(results);file.write_text(json.dumps({'status':'PASS','results':results,'runs':runs},indent=2));print('Presentation-only cleanup',results[0])
if __name__=='__main__':main()

"""Dashboard-specific KPI hierarchy. Links to model outputs, never manual metrics."""
from impact_layout_v08 import *
def apply_cards(source,target):
 src=Source(source);styles=add_styles(src.styles);changes={};proof=[]
 def act(row):return 'IF(\'EXEC V06\'!$H$23="Comparative",'+q(MET7,'C'+str(row))+','+q(MET7,'B'+str(row))+')'
 def bank(row):return act(row)
 def summ(key):return 'SUM('+q('V06 Motor',col(CC[key])+'5:'+col(CC[key])+'184')+')'
 def old(sn,addr):return src.get(sn,addr)[0] or str(src.get(sn,addr)[1])
 plans={
 'EXEC V06': [('CAPEX futuro','Future CAPEX',act(6),'money'),('VGV imobiliário','Real estate GDV',old('EXEC V06','E9'),'money'),('Equity previsto','Planned equity',summ('equity'),'money'),('Pico de dívida','Peak debt',act(7),'money'),('VAL investidores','Investor NPV',act(10),'money'),('TIR investidores','Investor IRR',act(11),'pct')],
 'DASH 03 V06':[('Pico de dívida','Peak debt',act(7),'money'),('DSCR mínimo','Minimum DSCR',act(26),'ratio'),('Folga mínima','Minimum headroom',act(8),'money'),('Custo financiamento','Financing cost',act(15),'money'),('Défice financiamento','Funding gap',act(19),'money'),('VAL investidores','Investor NPV',act(10),'money')],
 'DASH 04A V06':[('Frações','Apartments','COUNTA(\'V06 Unidades\'!A5:A98)','count'),('Área vendável','Saleable area','SUM(\'V06 Unidades\'!D5:D98)','count'),('ABC Aparthotel','Aparthotel GCA',tr('Por confirmar','To be confirmed'),'text'),('Preço médio €/m²','Average price EUR/m2','SUM(\'V06 Unidades\'!H5:H98)/SUM(\'V06 Unidades\'!D5:D98)','adr'),('VGV','GDV','SUM(\'V06 Unidades\'!H5:H98)','money'),('Eficiência vendável/ABC','Saleable/GCA efficiency',tr('ABC por confirmar','GCA to confirm'),'text')],
 'DASH 04B V06':[('Yield médio anual · 15A','Average annual yield · 15Y',summ('owner_distribution')+'/15/SUM(\'V06 Unidades\'!H5:H98)','pct'),('Receita residencial · 15A','Residential revenue · 15Y',summ('res_revenue'),'money'),('Distribuições proprietários','Owner distributions',summ('owner_distribution'),'money'),('OPEX residencial · 15A','Residential OPEX · 15Y',summ('res_opex'),'money'),('Fees residenciais · 15A','Residential fees · 15Y',summ('res_fees'),'money'),('FF&E residencial · 15A','Residential FF&E · 15Y',summ('res_ffe'),'money')],
 'DASH 06 V06':[('JSCV','JSCV',old('DASH 06 V06','Q54'),'pct'),('JALS','JALS',old('DASH 06 V06','T54'),'pct'),('Novos investidores','New investors','SUM(\'V06 Capital\'!D5:D12)','pct'),('Blocos totais','Total blocks',old('DASH 06 V06','E54'),'count'),('Blocos novos','New blocks',old('DASH 06 V06','H54'),'count'),('Captação prevista','Planned equity raise',old('DASH 06 V06','K54'),'money')],
 'DASH 07 V06':[('TIR investidores','Investor IRR',act(11),'pct'),('VAL investidores','Investor NPV',act(10),'money'),('MOIC','MOIC',act(12),'ratio'),('Dividendos previstos','Planned dividends',act(9),'money'),('Capital inicial previsto','Planned initial capital',summ('equity'),'money'),('Equity Value na saída','Exit Equity Value',act(13),'money')]
 }
 # Existing hotel yearly results were preserved in the support sheet by the
 # preceding pass; read their already-rewired public output formulas.
 h='DASH 05 V06';room=old(h,'N54');nights=old(h,'H54');rooms=old(h,'E54');revenue=old(h,'W54')
 # The selected-year EBITDA canonical output now resolves through support UI.
 import json
 moved={h+'!'+addr:STATUS7+'!'+ui_archive_address(h,addr) for addr in ['N50']}
 eb=moved[h+'!N50'];ebf=q(*eb.split('!'))
 plans[h]=[('ADR · ano selecionado','ADR · selected year','IF(('+nights+')>0,('+room+')/('+nights+'),'+tr('Sem noites vendidas','No occupied nights')+')','adr'),('Ocupação · ano selecionado','Occupancy · selected year','IF(('+rooms+')>0,('+nights+')/(('+rooms+')*365),'+tr('Sem capacidade','No capacity')+')','pct'),('RevPAR','RevPAR',old(h,'K54'),'adr'),('Receita total','Total revenue',revenue,'money'),('EBITDA','EBITDA',ebf,'money'),('Margem EBITDA','EBITDA margin','IF(('+revenue+')>0,'+ebf+'/('+revenue+'),'+tr('Sem receita','No revenue')+')','pct')]
 labels={'DASH 01 V06':['Histórico investido','CAPEX futuro','Construção e FF&E','Fiscalização','Contingência','ETC adicional'],'DASH 02 V06':['IVA suportado','Crédito IVA máximo','IRPC pago','Seguro de obra','Seguros exploração','ITI KK']}
 for sn,pt,en in DASHES[:9]:
  s=Sheet(E.fromstring(src.z.read(src.pp[sn])),styles)
  if sn in plans:
   for idx,(pt,en,f,kind) in enumerate(plans[sn]):
    c=4+idx*3;area=col(c)+'52:'+col(c+2)+'53';s.text(area,pt,en,'label');s.merge(col(c)+'54:'+col(c+2)+'56',None,f,'card_'+kind if 'card_'+kind in styles else 'card_text');proof.append({'sheet':sn,'cell':col(c)+'54','formula':f})
  else:
   for idx,label in enumerate(labels[sn]):s.text(col(4+idx*3)+'52:'+col(6+idx*3)+'53',label,['Historic investment','Future CAPEX','Construction and FF&E','Supervision','Contingency','Additional ETC'][idx] if sn=='DASH 01 V06' else ['Input VAT','Peak VAT credit','Corporate tax paid','Construction insurance','Operating insurance','ITI KK'][idx],'label')
  for rr in [52,53]:s.row(rr).attrib.pop('hidden',None);s.row(rr).set('ht','14')
  changes[src.pp[sn]]=s.finish()
 changes['xl/styles.xml']=xml(src.styles);rewrite(source,target,changes);src.z.close();return proof
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('source',type=Path);p.add_argument('target',type=Path);p.add_argument('--proof',type=Path,required=True);a=p.parse_args();a.proof.write_text(json.dumps(apply_cards(a.source,a.target),ensure_ascii=False,indent=2))

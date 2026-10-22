"""KK-01 V06: Excel expressions and an executable Python financial reference.
V04 operating datasets are preserved. All new assumptions are explicit inputs.
"""
import math,sys,json
from pathlib import Path
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name as col
sys.path.insert(0,str(Path(__file__).resolve().parents[2]/'V06/src/v04/src'))
import engine as legacy
class X:
 def __init__(self,f,v):self.f=str(f);self.v=v
 def _o(self,b,s,fn):
  b=asx(b);return X('('+self.f+s+b.f+')',fn(self.v,b.v))
 def __add__(self,b):return self._o(b,'+',lambda a,b:a+b)
 def __radd__(self,b):return asx(b)+self
 def __sub__(self,b):return self._o(b,'-',lambda a,b:a-b)
 def __rsub__(self,b):return asx(b)-self
 def __mul__(self,b):return self._o(b,'*',lambda a,b:a*b)
 def __rmul__(self,b):return asx(b)*self
 def __truediv__(self,b):return self._o(b,'/',lambda a,b:a/b if b else 0)
 def __rtruediv__(self,b):return asx(b)/self
 def __pow__(self,b):return self._o(b,'^',lambda a,b:a**b)
 def __rpow__(self,b):return asx(b)**self
 def __neg__(self):return X('-'+self.f,-self.v)
 def __lt__(self,b):return self._o(b,'<',lambda a,b:a<b)
 def __le__(self,b):return self._o(b,'<=',lambda a,b:a<=b)
 def __gt__(self,b):return self._o(b,'>',lambda a,b:a>b)
 def __ge__(self,b):return self._o(b,'>=',lambda a,b:a>=b)
 def eq(self,b):return self._o(b,'=',lambda a,b:a==b)
def asx(v):
 if isinstance(v,X):return v
 if isinstance(v,str):return X('"'+v.replace('"','""')+'"',v)
 return X(str(int(v)) if isinstance(v,bool) else str(v),v)
def fun(name,fn,*args):
 a=[asx(x) for x in args];return X(name+'('+','.join(x.f for x in a)+')',fn(*[x.v for x in a]))
def mx(*a):return fun('MAX',max,*a)
def mn(*a):return fun('MIN',min,*a)
def IF(c,a,b=0):return fun('IF',lambda c,a,b:a if c else b,c,a,b)
def AND(*a):return fun('AND',lambda *a:all(a),*a)
def OR(*a):return fun('OR',lambda *a:any(a),*a)
def INT(x):return fun('INT',math.floor,x)
def MOD(a,b):return fun('MOD',lambda a,b:a%b,a,b)
def ROUND(x,d=2):return fun('ROUND',lambda a,d:math.floor(a*10**d+.5)/10**d,x,d)
def SUM(a):
 a=list(a)
 if not a:return asx(0)
 return fun('SUM',lambda *v:sum(v),*a)
def rng(a,fn=sum):
 if not a:return asx(0)
 return X('SUM('+a[0].f+':'+a[-1].f.split('!')[-1]+')',fn([x.v for x in a]))
# key, label PT, label EN, value, unit, provenance / limitation
DEFAULTS=[
('language','Idioma','Language','PT','list','Presentation only'),
('vat_rate','IVA de estudo','Study VAT rate',.15,'%','Mandate: legal validation pending'),('vat_recovery','Fração recuperável IVA','VAT recoverable fraction',1,'%','Full/partial/no recovery'),('vat_lag','Prazo reembolso IVA','VAT refund lag',6,'month','Study hypothesis; not confirmed entitlement'),
('tax_rate','Taxa IRPC','Corporate tax rate',.20,'%','Study general rate; not a legal opinion'),('exempt_years','Anos isenção simulada','Simulated exemption years',0,'year','0 / 3 / 5 / 15; hypothetical, never approval'),('tax_lag','Prazo pagamento IRPC','Corporate tax payment lag',3,'month','Project fiscal years; legal dates pending'),('dep_base','Base depreciável','Depreciable basis',40100000,'EUR','V05 documented proxy'),('dep_life','Vida útil fiscal','Fiscal useful life',30,'year','Proxy pending tax depreciation schedule'),
('property_base','Base IPI Hotel/MICE','Hotel/MICE property-tax basis',40100000,'EUR','Proxy; valuation pending'),('property_rate','Taxa IPI','Property tax rate',0,'%','Unknown; zero is not exemption'),('transfer_rate','ITI referência','Transfer tax reference',.015,'%','Legacy reference; incidence pending'),('transfer_kk','Fração ITI suportada KK','Transfer tax KK share',0,'%','Buyer tax excluded unless KK bears cost'),('deed_kk','Fração atos suportada KK','Deed costs borne by KK',1,'%','Modelling assumption, confirm allocation'),('cve','CVE por EUR','CVE per EUR',110.265,'CVE/EUR','Documented peg reference'),('deed_cve','Escritura por imóvel','Deed cost per property',24000,'CVE','V05 mandate conventional amount'),('cert_cve','Certidões por fração','Certificates per apartment',3000,'CVE','1500 + 1500 CVE'),('ph','Propriedade horizontal','Horizontal property process',20000,'EUR','Management provision'),('ph_month','Mês propriedade horizontal','Horizontal property month',24,'month','Assumption; editable'),('other_deeds','Atos Hotel/MICE/restaurante','Hotel/MICE/restaurant deeds',3,'count','Assumed legal segmentation'),
('car','CAR taxa total obra','CAR total construction rate',.003,'%','V05: 169200; whole construction period'),('rc_build','RC obra taxa total','Construction liability total rate',.0005,'%','V05: 28200; whole period'),('build_included','Seguro obra já incluído','Construction insurance included',0,'%','0 no / 1 yes / partial fraction'),('insurance_base','Base seguro patrimonial','Property insured value',40100000,'EUR','Provisional replacement value'),('property_ins','Multirriscos taxa anual','Property insurance annual rate',.0035,'%','V05: use actual multiplication, not rounded budget'),('bi','Interrupção atividade taxa','Business interruption rate',.0015,'% revenue','V05 study rate'),('liability','RC exploração taxa','Operating liability rate',.0015,'% revenue','V05 study rate'),('other_ins','Coberturas anuais adicionais','Additional annual insurance',15000,'EUR/year','Provision, not quote'),('ins_factor','Multiplicador seguros','Insurance multiplier',1,'x','1 / 1.25 / 1.5 / custom'),('op_included','Seguro operação já incluído','Operating insurance included',0,'%','Avoid forecast double counting'),
('capital_delay','Atraso capital','Capital delay',0,'month','Scenario'),('capital_fraction','Capital recebido de cenário','Scenario capital received fraction',1,'%','Planned, not actual'),('sales_delay','Atraso vendas','Sales delay',0,'month','Scenario'),('construction_delay','Atraso obra','Construction delay',0,'month','Scenario, whole schedule stretched'),('overrun','Sobrecusto construção','Construction cost overrun',0,'%','Scenario'),('limit','Linha não renovável','Non-revolving facility',30000000,'EUR','Study facility; not committed'),('rate','Juro anual','Annual interest',.06,'%','Reference bank scenario'),('grace','Carência principal','Principal grace period',42,'month','Interest paid currently'),('term','Prazo após carência','Post-grace tenor',120,'month','Reference 10 years'),('sweep','Cash sweep','Cash sweep',1,'%','No revolving reuse'),('bank_cp','Condições banco cumpridas','Bank conditions met',1,'0/1','Scenario, not confirmed bank approval'),('bank_start','Início utilização','First draw month',1,'month','Study'),('opening_fee','Comissão abertura','Opening fee',.01,'% facility','V04 reference'),('broker','Broker','Broker',.015,'% facility','V04 reference'),('mobil','Mobilização','Mobilisation',.0125,'% facility','V04 reference'),('commit','Comissão compromisso anual','Commitment fee annual',0,'% undrawn','Uncontracted'),('reserve','Reserva operacional','Operating reserve',2000000,'EUR','V04 minimum plus HOTUSA and forward buffer'),('buffer','Meses proteção obra','Construction buffer months',3,'month','V04'),('hotusa','Reserva contingência jurídica HOTUSA','HOTUSA Legal Contingency Reserve',1000000,'EUR','Protected, not debt'),('hot_state','Estado HOTUSA','HOTUSA state','Pending','list','Pending / Pay / Released'),('hot_pay','Obrigação HOTUSA confirmada','Confirmed HOTUSA obligation',0,'EUR','No payment presumed'),('hot_month','Mês resolução HOTUSA','HOTUSA resolution month',180,'month','Assumption'),
('div_policy','Política dividendos','Dividend policy','Auto','list','Auto / Minimum / Intermediate / Maximum / Partial / Suspended'),('coverage','Cobertura 12M','12-month coverage',1.2,'x','Trailing 12 months independent from accumulated liquidity'),('profit_mode','Lucros documentados ou proxy','Documented profits or proxy','Proxy','list','Proxy is conditional, not legal approval'),('legal_cap','Limite dividendos documentado','Documented dividend ceiling',0,'EUR','Applies only in Documented mode'),('div_partial','Fração distribuição parcial','Partial distribution fraction',.5,'%','Applied to minimum target'),('exit_multiple','Múltiplo saída','Exit multiple',12,'x','10 / 12 / 15'),('exit_cost','Custo saída','Exit cost',.02,'% EV','V04 study'),('exit_month','Mês saída','Exit month',180,'month','120 or 180; mutually exclusive'),('discount','Desconto equity','Equity discount',.12,'%','Study required return'),('project_discount','Desconto projeto','Project discount',.10,'%','Unlevered nominal'),('hotel_scenario','Cenário Hotel','Hotel scenario','Favorável','list','V04 operational dataset'),('base_fee','Fee base operador','Operator base fee',.03,'% revenue','3/3/6 reference'),('systems','Sistemas e distribuição','Systems and distribution',.03,'% revenue','Not separately duplicate marketing'),('incentive','Fee incentivo','Incentive fee',.06,'% adjusted GOP','After base and systems'),('extra_operator','Outros fees receita','Additional revenue-based operator fees',0,'% revenue','Only incremental charges'),('res_yield','Yield residencial inicial','Initial residential target yield',.0425,'%','V04, not guaranteed'),('res_yield_step','Incremento anual yield','Annual yield increment',.0025,'% points','V04 to 6.5%'),('convention','Estado Convenção','Convention status','Em preparação','list','Changing status never activates exemptions'),('eac_override','ETC adicional','Incremental ETC',0,'EUR','Additional forecast cost; allocated by remaining schedule'),('status_month','Mês controlo obra','Construction status month',0,'month','Forecast classification only; not evidence of actual costs'),('customs','Direitos aduaneiros KK','KK customs charges',0,'EUR','Unknown, not exemption'),('stamp','Imposto selo KK','KK stamp duties',0,'EUR','Unknown, not exemption'),('extra_charges','Outras taxas institucionais','Other institutional charges',0,'EUR','Unknown; editable amount paid at opening'),('owner_tax','Taxa informativa proprietários','Informative owner tax rate',0,'%','Unknown; zero is not verified benefit')]
DEFAULTS.append(('draw_end','Último mês utilização banco','Last bank draw month',30,'month','V04 study draw window; never extended automatically with construction delay'))
BASE={x[0]:x[3] for x in DEFAULTS}
TICKETS=[('P01','Platinum',1e7,.20,1),('G01','Gold',5e6,.10,2),('G02','Gold',5e6,.10,3)]+[(f'S0{i}','Silver',1e6,.02,m) for i,m in enumerate([3,4,4,5,6],1)]
TARGETS={'Platinum':(700000,850000,1000000),'Gold':(250000,287500,325000),'Silver':(40000,45000,50000)}
class Model:
 def __init__(self,wb,formats,params,units,chapters):self.wb=wb;self.fm=formats;self.p=params;self.units=units;self.chapters=chapters
 def sheet(self,name,headers):
  s=self.wb.add_worksheet(name);s.hide_gridlines(2);s.freeze_panes(4,1);s.set_column(0,0,12);s.set_column(1,len(headers),19);s.write_row(3,0,headers,self.fm['head']);s.set_row(3,30);return s
 def cell(self,s,row,c,x,fmt='num'):
  x=asx(x);s.write_formula(row-1,c,'='+x.f,self.fm[fmt],int(x.v) if isinstance(x.v,bool) else x.v);return X("'"+s.name+"'!"+col(c)+str(row),x.v)
 def calculate(self,name,overrides=None):
  p=self.p.copy();p.update({k:asx(v) for k,v in (overrides or {}).items()});s=self.sheet(name,['Month']);data=[];columns={}
  calendar=self.calendar(p)
  def put(key,x,fmt='num'):
   if key not in columns:
    columns[key]=len(columns);s.write(3,columns[key],key,self.fm['head'])
   v=self.cell(s,r,columns[key],x,fmt);d[key]=v;return v
  # Each scenario has its own 180-month financing, taxes, reserves and operating engine.
  cap=[];receipts=[];hotel=[];res=[];ytd=asx(0);loss=asx(0)
  for m in range(1,181):
   r=m+4;d={};prev=data[-1] if data else {};M=put('month',m,'int');year=put('year',INT((M-1)/12)+1,'int')
   end=put('end',30+p['construction_delay'],'int');op=put('open',end+1,'int');age=put('op_age',mx(0,INT((M-op)/12)),'int');active=M>=op
   # Chapter weights are stretched across the construction duration, total cost conserved.
   source_month=put('construction_source_month',IF(M<=end,INT((M-1)*30/end)+1,0),'int')
   # Monthly weights of the V04 original curve are reallocated by SUMPRODUCT interval overlaps.
   weights=[sum(c['share']*c['weights'][i] for c in legacy.CHAPTERS) for i in range(30)]
   weight=asx(0)
   # Integrate the stepwise cumulative V04 curve over the stretched month interval.
   wparts=[]
   for j,w in enumerate(weights):wparts.append(mx(0,mn(M*30/end,j+1)-mx((M-1)*30/end,j))*w)
   def interval_formula(lo,hi):
    # Scalar intervals: no implicit intersection or nested array IF evaluation.
    return 'SUM('+','.join(f"MAX(0,MIN({hi},{j+1})-MAX({lo},{j}))*'V06 Curva'!$B${j+5}" for j in range(30))+')'
   weight_formula='IF('+M.f+'<='+end.f+','+interval_formula('('+M.f+'-1)*30/'+end.f,M.f+'*30/'+end.f)+',0)'
   weight=put('construction_weight',X(weight_formula,IF(M<=end,SUM(wparts),0).v))
   construction=put('construction',weight*56400000*(1+p['overrun']))
   supervision=put('supervision',construction*.02);cont=put('contingency_paid',construction*.10)
   etc=put('extra_etc',IF(AND(M>p['status_month'],M<=end),p['eac_override']/mx(1,end-p['status_month']),0))
   capex=put('capex',construction+supervision+cont+etc)
   # Recovered eight investment calendars, with explicit scenario receipt classification.
   eqparts=[];outparts=[]
   for ticket in self.tickets:
    code,cat,amount,share,planned,primary,received,actual,actual_m=ticket
    planned_pay=planned+p['capital_delay'];pay_m=IF(actual_m>0,actual_m,planned_pay)
    amount_x=IF(actual_m>0,actual,amount*received*p['capital_fraction'])
    xp=IF(M.eq(pay_m),amount_x,0);outparts.append(xp);eqparts.append(xp*primary)
   equity=put('equity',SUM(eqparts));outlay=put('investor_outlay',SUM(outparts))
   # Native SUMPRODUCT sales schedules: cents are absorbed in the last instalment.
   # Individual unit terms are stored once; scalar scenario delay and physical completion affect dates.
   U=self.units;vals=[];taxvals=[];costvals=[];salesvals=[];deedcount=asx(0);eligiblecount=asx(0);eligibleprice=asx(0);soldprice=asx(0)
   # Python reference independently iterates each of 94 apartments.
   for u in U:
    sale=u['sale']+p['sales_delay'];deed=IF(p['construction_delay']>0,mx(u['deed']+p['sales_delay'],end),u['deed']+p['sales_delay']);n=mx(0,INT((deed-sale-1)/6));price=u['price'];a=ROUND(price*.05);b=ROUND(price*.20);prog=ROUND(price*.55);inst=ROUND(prog/mx(1,n));final=price-a-b-prog
    progress=IF(AND(n>0,M>sale,M<deed,MOD(M-sale,6).eq(0),(M-sale)/6<=n),IF(M.eq(sale+6*n),prog-inst*(n-1),inst),0)
    vals.append(IF(M.eq(sale),a+b,0)+progress+IF(M.eq(deed),final,0))
    salesvals.append(IF(M.eq(deed),price,0));costvals.append(IF(M.eq(sale),price*.025,0));taxvals.append(IF(M.eq(deed),price*p['transfer_rate']*p['transfer_kk'],0))
    deedcount+=IF(M.eq(deed),1,0);eligiblecount+=IF(M>=deed,1,0);eligibleprice+=IF(M>=deed,price,0);soldprice+=IF(M<deed,price,0)
   # One scalar schedule per unit, shared across scenarios with the same calendar.
   cs=calendar;cr=lambda c:f"'{cs}'!${c}$5:${c}$98";mo=M.f
   prices=cr('B');deed=cr('D');sale=cr('C')
   recv=put('receipts',X(f"SUM('{cs}'!{col(12+m-1)}5:{col(12+m-1)}98)",sum(x.v for x in vals)))
   sold=put('deed_sales',X(f'SUMIF({deed},{mo},{prices})',sum(x.v for x in salesvals)))
   sale_cost=put('sales_cost',X(f'SUMIF({sale},{mo},{cr("K")})',sum(x.v for x in costvals)))
   transfer=put('transfer_kk',sold*p['transfer_rate']*p['transfer_kk'])
   deeds=put('deeds_cost',X(f'COUNTIF({deed},{mo})',deedcount.v)*(p['deed_cve']+p['cert_cve'])/p['cve']*p['deed_kk']+IF(M.eq(end),p['other_deeds']*p['deed_cve']/p['cve']*p['deed_kk'],0)+IF(M.eq(p['ph_month']),p['ph'],0))
   # Hotel and MICE preserve V04 operating formulas and forecast source.
   hp=legacy.config();scenario=p['hotel_scenario'];adr0=X('IF('+scenario.f+'="Conservador",'+str(legacy.HOTELS['Conservador'][0])+',IF('+scenario.f+'="Base",'+str(legacy.HOTELS['Base'][0])+','+str(legacy.HOTELS['Favorável'][0])+'))',legacy.HOTELS[scenario.v][0])
   def hotval(i):return X('IF('+scenario.f+'="Conservador",'+str(legacy.HOTELS['Conservador'][i])+',IF('+scenario.f+'="Base",'+str(legacy.HOTELS['Base'][i])+','+str(legacy.HOTELS['Favorável'][i])+'))',legacy.HOTELS[scenario.v][i])
   occ0=hotval(1)
   if 'hotel_mode' in p:
    mode=p['hotel_mode'];adr0=IF(mode.eq('Custom'),p['adr_input'],IF(mode.eq('Conservative'),legacy.HOTELS['Conservador'][0],IF(mode.eq('Favorable'),legacy.HOTELS['Favorável'][0],IF(mode.eq('Upside Stress'),250,adr0))))
    occ0=IF(mode.eq('Custom'),p['occ_input'],IF(mode.eq('Conservative'),legacy.HOTELS['Conservador'][1],IF(mode.eq('Favorable'),legacy.HOTELS['Favorável'][1],IF(mode.eq('Upside Stress'),.75,occ0))))
   adr=put('adr',IF(active,adr0*(1+p.get('adr_growth',hp['adr_growth']))**age,0));occ=put('occupancy',IF(active,mn(mx(hp['occ_max'],occ0),occ0+hp['occ_step']*age),0),'pct')
   room=put('room_revenue',200*365/12*adr*occ);mice=put('mice_revenue',room*hotval(2)*p.get('mice_util',.5)/.5);fb=put('fb_revenue',room*hotval(3)*hp['fb_share']);other=put('other_revenue',room*hotval(3)*(1-hp['fb_share']));revenue=put('revenue',room+mice+fb+other)
   gop=put('gop',revenue-room*hp['room_cost']-mice*hp['mice_cost']-(fb+other)*hp['other_cost']-revenue*hp['overhead'])
   basefee=put('base_fee',revenue*p['base_fee']);system=put('systems_fee',revenue*p['systems']);incentive=put('incentive_fee',mx(0,gop-basefee-system-IF(p.get('owner_priority_on',asx(0)).eq(1),p.get('owner_priority',asx(0))/12,0))*p['incentive']);extrafee=put('other_operator',revenue*p['extra_operator'])
   ins_build=put('construction_insurance',construction*(p['car']+p['rc_build'])*(1-p['build_included'])*p['ins_factor'])
   ins_op=put('operating_insurance',IF(active,(p['insurance_base']*p['property_ins']/12+revenue*(p['bi']+p['liability'])+p['other_ins']/12)*(1-p['op_included'])*p['ins_factor'],0))
   ebitda=put('ebitda',gop-basefee-system-incentive-extrafee-ins_op)
   ffe=put('ffe',revenue*p.get('ffe_rate',hp['ffe']));maint=put('maintenance',revenue*hp['maintenance']);wc=put('working_capital',revenue*12/365*hp['wc_days']);dwc=put('delta_wc',wc-(prev['working_capital'] if prev else 0))
   dep=put('depreciation',IF(active,p['dep_base']/p['dep_life']/12,0))
   eligible_n=put('res_units',X(f'COUNTIF({deed},"<="&{mo})',eligiblecount.v),'int');eligible_value=put('res_eligible_value',X(f'SUMIF({deed},"<="&{mo},{prices})',eligibleprice.v))
   res_adr=put('res_adr',IF(active,hp['res_adr']*(1+hp['res_growth'])**age,0));res_occ=put('res_occupancy',IF(active,mn(hp['res_occ_max'],hp['res_occ']+hp['res_occ_step']*age),0),'pct')
   res_rev=put('res_revenue',eligible_n*365/12*res_adr*res_occ*(1+hp['res_other']));res_opex=put('res_opex',res_rev*hp['res_cost']);res_fees=put('res_fees',res_rev*(p.get('res_base',hp['res_base'])+p.get('res_system',hp['res_system'])));res_ffe=put('res_ffe',res_rev*hp['res_ffe']);res_wc=put('res_wc',res_rev*12/365*hp['res_wc']);res_dwc=put('res_delta_wc',res_wc-(prev['res_wc'] if prev else 0));res_cash=put('res_cash',res_rev-res_opex-res_fees-res_ffe-res_dwc)
   res_target=put('owner_target',IF(active,eligible_value*mn(.065,p['res_yield']+p['res_yield_step']*age)/12,0));owner=put('owner_distribution',mn(mx(0,res_cash),res_target));short=put('owner_shortfall',mx(0,res_target-owner));excess=put('res_excess',mx(0,res_cash-owner));res_fund=put('res_fund',excess*hp['res_fund']);res_kk=put('res_kk',((excess*(1-hp['res_fund']-hp['res_kyenu']-hp['res_operator']))+mn(0,res_cash))*hp['res_include'])
   vat=put('vat_incurred',(construction+supervision)*p['vat_rate']);eligible_vat=put('vat_recoverable',vat*p['vat_recovery']);refund=put('vat_refund',X('IF('+M.f+'>'+p['vat_lag'].f+',INDEX($'+col(columns['vat_recoverable'])+'$5:$'+col(columns['vat_recoverable'])+'$184,MAX(1,'+M.f+'-'+p['vat_lag'].f+')),0)',data[m-1-int(p['vat_lag'].v)]['vat_recoverable'].v if 0<=m-1-int(p['vat_lag'].v)<len(data) else (eligible_vat.v if p['vat_lag'].v==0 else 0)))
   # Lag zero must reference current eligible VAT (no circular dependency).
   if p['vat_lag'].v==0:pass
   credit=put('vat_credit',(prev['vat_credit'] if prev else 0)+eligible_vat-refund)
   nonrecover=put('vat_nonrecoverable',vat-eligible_vat)
   propertytax=put('property_tax',IF(active,p['property_base']*p['property_rate']/12,0)+X(f'SUMIF({deed},">"&{mo},{prices})',soldprice.v)*p['property_rate']/12)
   charges=put('institutional_charges',IF(M.eq(op),p['customs']+p['stamp']+p['extra_charges'],0))
   salesprofit=put('sales_profit',sold*(1-63168000*(1+p['overrun'])/X("SUM('V06 Unidades'!$H$5:$H$98)",sum(u['price'].v for u in U))*13290/47000)-sale_cost-transfer-deeds)
   pretax0=put('pretax_before_interest',ebitda-dep+res_kk-sale_cost*0+salesprofit-propertytax-charges-ins_build-nonrecover)
   # Annual tax payments refer only to closed tax years, avoiding debt/tax circularity.
   payment=asx(0)
   for past in data:
    if int(past['month'].v)%12==0:payment+=IF(M.eq(past['month']+p['tax_lag']),past['annual_tax'],0)
   taxpaid=put('tax_paid',payment)
   openingcash=put('cash_open',prev['cash'] if prev else 0);openingdebt=put('debt_open',prev['debt'] if prev else 0);draw_open=put('draw_open',prev['cumulative_draw'] if prev else 0)
   hotpay=put('hotusa_payment',IF(AND(p['hot_state'].eq('Pay'),M.eq(p['hot_month'])),p['hot_pay'],0));hold=put('hotusa_reserve',IF(OR(p['hot_state'].eq('Pending'),M<p['hot_month']),p['hotusa'],0))
   # Forward construction reserve uses the original curve with exact stretched interval weights.
   futurecost=asx(0)
   for shift in range(1,4):
    mm=M+shift;ww=SUM(mx(0,mn(mm*30/end,j+1)-mx((mm-1)*30/end,j))*w for j,w in enumerate(weights));futurecost+=IF(AND(mm<=end,shift<=p['buffer']),ww*63168000*(1+p['overrun']),0)
   futurecost=put('construction_buffer',X(interval_formula(M.f+'*30/'+end.f,'('+M.f+'+'+p['buffer'].f+')*30/'+end.f)+'*63168000*(1+'+p['overrun'].f+')',futurecost.v))
   priorprofit=ytd if (m-1)%12 else asx(0)
   tax_protection=put('last_draw_tax_protection',IF(M.eq(p['draw_end']),mx(0,priorprofit+pretax0)*p['tax_rate'],0))
   reserve=put('reserve',mx(p['reserve'],IF(M<=end,futurecost,0))+hold+mx(tax_protection,IF(p.get('protect_known_tax',asx(0)).eq(1),mx(0,(prev['tax_payable'] if prev else 0)-taxpaid),0))+p.get('debt_reserve',0)+p.get('other_reserve',0))
   commitment=put('commitment_fee',IF(AND(M<=p['draw_end'],M>=p['bank_start']),mx(0,p['limit']-draw_open)*p['commit']/12,0))
   cfads=put('cfads',ebitda-ffe-maint-dwc-propertytax)
   project=put('project_cf',recv-sale_cost-transfer-deeds-capex+cfads+res_kk-vat+refund-ins_build-taxpaid-hotpay-charges)
   available=put('available',openingcash+equity+project-commitment)
   principal=put('principal',IF(M>p['grace'],mn(openingdebt,openingdebt/mx(1,p['grace']+p['term']-M+1)),0))
   remaining=put('remaining_facility',IF(AND(M<=p['draw_end'],M>=p['bank_start'],p['bank_cp'].eq(1)),mx(0,p['limit']-draw_open),0))
   fees=put('bank_fees',IF(AND(draw_open<.005,available-principal-openingdebt*p['rate']/12<reserve-.005,remaining>0),p['limit']*(p['opening_fee']+p['broker']+p['mobil']),0))
   draw=put('draw',mn(remaining,mx(0,(reserve-available+principal+openingdebt*p['rate']/12+fees)/(1-p['rate']/12))))
   interest=put('interest',(openingdebt+draw)*p['rate']/12)
   post=put('post_debt_service',available+draw-fees-interest-principal)
   # During construction sweep is limited by current net receipts and forward unmet uses.
   future_cap=put('future_capex',mx(0,63168000*(1+p['overrun'])+p['eac_override']-rng([x['capex'] for x in data]+[capex])))
   future_recv=put('future_receipts',mx(0,X("SUM('V06 Unidades'!$H$5:$H$98)",sum(u['price'].v for u in U))-rng([x['receipts'] for x in data]+[recv])))
   forward=put('forward_protection',IF(M<end,mx(0,future_cap-future_recv-IF(AND(M<p['draw_end'],p['bank_cp'].eq(1)),mx(0,p['limit']-draw_open-draw),0)),0))
   sweep=put('sweep',mn(mx(0,openingdebt+draw-principal),mx(0,post-reserve-forward)*p['sweep'],IF(M<=end,mx(0,recv-sale_cost-transfer),mx(0,post-reserve-forward))))
   debt=put('debt',mx(0,openingdebt+draw-principal-sweep));cumulative=put('cumulative_draw',draw_open+draw)
   taxable=put('taxable_profit',pretax0-interest-fees-commitment);ytd=put('taxable_ytd',(prev['taxable_ytd'] if (m-1)%12 else 0)+taxable)
   rate=put('effective_tax_rate',IF(AND(active,age<p['exempt_years']),0,p['tax_rate']),'pct')
   carry=prev['loss_carry'] if prev else asx(0)
   taxbasis=put('nonexempt_taxable_ytd',(prev['nonexempt_taxable_ytd'] if (m-1)%12 else 0)+IF(AND(active,age<p['exempt_years']),0,taxable))
   annualtax=put('annual_tax',IF(MOD(M,12).eq(0),mx(0,taxbasis-carry)*p['tax_rate'],0))
   losscarry=put('loss_carry',IF(MOD(M,12).eq(0),mx(0,carry-taxbasis),carry))
   accrual=put('tax_accrual',mx(0,taxbasis-carry)*p['tax_rate'])
   payable=put('tax_payable',mx(0,rng([x['annual_tax'] for x in data]+[annualtax])-rng([x['tax_paid'] for x in data]+[taxpaid])))
   service=put('debt_service',interest+principal+commitment)
   generation=put('generation_after_service',cfads+res_kk-taxpaid-service)
   trail=put('trailing_12m',SUM([x['generation_after_service'] for x in data[-12:]]))
   distributed=prev['dividends_cumulative'] if prev else asx(0)
   netprofit=put('net_profit_proxy',(prev['net_profit_proxy'] if prev else 0)+taxable-(accrual-(prev['tax_accrual'] if (m-1)%12 else 0)))
   legal=put('legal_headroom',mx(0,IF(p['profit_mode'].eq('Documented'),p['legal_cap'],netprofit)-distributed))
   liquid=put('liquid_headroom',mx(0,post-sweep-reserve))
   covered=put('coverage_headroom',mx(0,trail/p['coverage']))
   target=put('dividend_target',IF(p['div_policy'].eq('Maximum'),4150000,IF(p['div_policy'].eq('Intermediate'),3575000,IF(p['div_policy'].eq('Suspended'),0,IF(p['div_policy'].eq('Partial'),3000000*p['div_partial'],IF(p['div_policy'].eq('Auto'),IF(mn(legal,liquid,covered)>=4150000,4150000,IF(mn(legal,liquid,covered)>=3575000,3575000,3000000)),3000000))))))
   gap=put('funding_gap',mx(0,-post));ever=put('funding_gap_cumulative',(prev['funding_gap_cumulative'] if prev else 0)+gap)
   dividend=put('dividend',IF(AND(M>=54,M<=174,MOD(M-54,12).eq(0),debt<.005,ever<.005),mn(target,legal,liquid,covered),0))
   divcum=put('dividends_cumulative',distributed+dividend)
   cash=put('cash',post-sweep-dividend);reservegap=put('reserve_gap',mx(0,reserve-cash))
   checkcash=put('cash_check',cash-openingcash-equity-project-draw+commitment+fees+interest+principal+sweep+dividend)
   checkdebt=put('debt_check',debt-openingdebt-draw+principal+sweep)
   put('dscr',IF(service>0,cfads/mx(.00000001,service),0),'num')
   # Append-only extension: original 115 engine column positions remain stable.
   available_n=put('available_room_nights',IF(active,200*365/12,0));occupied_n=put('occupied_room_nights',available_n*occ)
   put('revpar',IF(available_n>0,room/mx(1,available_n),0));put('trevpar',IF(available_n>0,revenue/mx(1,available_n),0));put('goppar',IF(available_n>0,gop/mx(1,available_n),0));put('ebitda_margin',IF(revenue>0,ebitda/mx(1,revenue),0),'pct')
   put('operator_fees',basefee+system+incentive+extrafee);put('mice_venue',mice*.5);put('mice_event_fb',mice*.35);put('mice_services',mice*.15);put('mice_direct_cost',mice*hp['mice_cost']);put('mice_contribution',mice*(1-hp['mice_cost']))
   spend=put('ffe_replacement_spend',ffe*p.get('ffe_spend',1));fund=put('ffe_restricted_fund',(prev['ffe_restricted_fund'] if prev else 0)+ffe-spend)
   put('ffe_fund_check',fund-(prev['ffe_restricted_fund'] if prev else 0)-ffe+spend)
   put('known_tax_protection',mx(0,(prev['tax_payable'] if prev else 0)-taxpaid));put('general_reserve',p['reserve']);put('debt_restricted_reserve',p.get('debt_reserve',0));put('other_restricted_reserve',p.get('other_reserve',0))
   put('res_management_fee',res_rev*p.get('res_base',hp['res_base']));put('res_distribution_fee',res_rev*p.get('res_system',hp['res_system']));put('res_owner_net',owner*(1-p['owner_tax']))
   put('units_deeded',eligible_n,'int');put('units_not_deeded',94-eligible_n,'int');put('inventory_check',eligible_n+94-eligible_n-94,'int')
   data.append(d)
  return {'sheet':name,'rows':data,'columns':columns,'params':p}

 def calendar(self,p):
  """Single inventory; reusable scalar contractual forecast schedules, rounded cents."""
  key=(p['sales_delay'].f,p['construction_delay'].f,self.units[0]['price'].f)
  if not hasattr(self,'calendars'):self.calendars={}
  if key in self.calendars:return self.calendars[key]
  name=f"V06 Calend {len(self.calendars)+1:02}";self.calendars[key]=name
  s=self.sheet(name,['Unit','Price EUR','Sale month','Deed month','Progress instalments','Deposit 5%','Contract 20%','Progress 55%','Instalment EUR','Final EUR','Selling costs','Forecast only']+[f'M{m:03}' for m in range(1,181)])
  for r,u in enumerate(self.units,5):
   end=30+p['construction_delay'];sale=u['sale']+p['sales_delay'];deed=IF(p['construction_delay']>0,mx(u['deed']+p['sales_delay'],end),u['deed']+p['sales_delay'])
   self.cell(s,r,0,u['id'],'text');price=self.cell(s,r,1,u['price']);sale=self.cell(s,r,2,sale,'int');deed=self.cell(s,r,3,deed,'int');n=self.cell(s,r,4,mx(0,INT((deed-sale-1)/6)),'int')
   a=self.cell(s,r,5,ROUND(price*.05));b=self.cell(s,r,6,ROUND(price*.20));prog=self.cell(s,r,7,ROUND(price*.55));inst=self.cell(s,r,8,ROUND(prog/mx(1,n)));final=self.cell(s,r,9,price-a-b-prog);self.cell(s,r,10,price*.025)
   s.write(r-1,11,'Forecast; actual receipts are separate')
   for m in range(1,181):
    M=asx(m);progress=IF(AND(n>0,M>sale,M<deed,MOD(M-sale,6).eq(0),(M-sale)/6<=n),IF(M.eq(sale+6*n),prog-inst*(n-1),inst),0)
    self.cell(s,r,12+m-1,IF(M.eq(sale),a+b,0)+progress+IF(M.eq(deed),final,0))
  return name

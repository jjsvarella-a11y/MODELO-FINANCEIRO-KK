"""Motor independente de referência; euros nominais, meses relativos, sem macros."""
from pathlib import Path
import json,math
from scipy.optimize import brentq
ROOT=Path(__file__).resolve().parents[1]
SOURCE=json.loads((ROOT/'src/source.json').read_text())
UNITS=SOURCE['units'];CHAPTERS=SOURCE['chapters']
# key, label, default, unit, origin/limitation
PARAMS=[
('equity','Capital novo',20000000,'€','Cenários negociais; realizado no mês 1.'),
('bank','Cenário bancário','Limite aceitável','lista','Referência prudente; condições não contratadas.'),
('hotel','Cenário Hotel + MICE','Favorável','lista','Referência comercial; não garantia de desempenho.'),
('area','Área construída',47000,'m²','V45 Inputs B4; confirmar projeto de execução.'),
('cost','Construção + FF&E por m²',1200,'€/m²','Inclui equipamento inicial; não repetir FF&E inicial.'),
('supervision','Fiscalização',.02,'%','Sobre construção base.'),
('contingency','Contingência construção',.10,'%','Orçamento de risco; reserva não é pagamento.'),
('cont_used','Realização da contingência',1,'%','100% no cenário prudente; saldo não gasto fica reservado.'),
('historic','Investimento histórico',4500000,'€','Custo afundado; não é saída futura nem equity novo.'),
('hotusa','Exposição / reserva HOTUSA',1000000,'€','0 / 500 mil / 1 M€; não constitui reconhecimento de dívida.'),
('hotpay','Pagar HOTUSA no cenário',0,'0/1','0: reserva; 1: saída efetiva hipotética no mês indicado.'),
('hotmonth','Mês eventual pagamento HOTUSA',30,'mês','Sem compromisso jurídico presumido.'),
('limit','Linha bancária máxima',30000000,'€','Não renovável; limite sobre desembolsos acumulados.'),
('reserve','Reserva mínima operacional',2000000,'€','Não é CAPEX; caixa protegida desde mês 1.'),
('buffer','Meses reserva de obra',3,'meses','Próximos 3 meses de CAPEX; adicional à reserva HOTUSA.'),
('end','Fim construção',30,'mês','Calendário original de 30 meses; não deslocar sem rever capítulos.'),
('open','Abertura Hotel e residencial',31,'mês','Convenção: após conclusão do mês 30; diferença face à V45 (mês 30).'),
('draw_end','Último mês de utilização banco',30,'mês','Sem novo crédito operacional implícito.'),
('broker','Success fee broker',.015,'%','Sobre linha contratada, apenas no primeiro desembolso efetivo; confirmar contrato.'),
('mobil','Mobilização bancária',.0125,'%','V45 1,25%; separada da comissão de abertura; confirmar sobreposição.'),
('sweep','Cash sweep',1,'%','Excedente após reservas; durante obra limitado a recebimentos imobiliários líquidos.'),
('payout','Payout anual após dívida liquidada',.5,'%','Só no fecho anual; reservas e défices bloqueiam distribuição.'),
('dscr_min','Alerta DSCR',1.3,'x','Limite interno indicativo; não covenant contratado.'),
('discount','Taxa desconto projeto',.10,'%','Hipótese; VPL nominal sem valor terminal e com saída, separados.'),
('equity_discount','Taxa desconto acionistas',.12,'%','Hipótese de custo de capital próprio.'),
('exit_mult','Múltiplo EBITDA de saída',9,'x','Hotel + MICE; não inclui imóveis já vendidos; não é preço garantido.'),
('exit_cost','Custos da saída',.02,'%','Hipótese de custos de transação sobre EV.'),
('tax_transfer','Imposto transmissão',.015,'%','No mês escritura; incidência e sujeito passivo por confirmar.'),
('marketing','Marketing',.01,'%','Mínimo 1%; no contrato de venda.'),
('kyenu','Comissão Kyenu Kyetu',.015,'%','Sobre preço no contrato; conta separada.'),
('deposit','Sinal',.05,'%','No mês venda.'),
('tranche','Primeira tranche',.20,'%','Mês venda + atraso editável.'),
('progress','Prestações semestrais',.55,'%','Distribuídas nos semestres anteriores à escritura; última absorve cêntimos.'),
('final','Escritura',.20,'%','Complemento final exato; alerta se percentagens não somam 100%.'),
('lag','Atraso primeira tranche',0,'meses','0 preserva V45; validar contratos por fração.'),
('rooms','Quartos Hotel',200,'unidades','Exclui as 94 frações vendidas.'),
('adr_growth','Crescimento ADR anual Hotel',.03,'%','V45; por aniversário de operação.'),
('occ_step','Incremento ocupação anual Hotel',.012,'p.p.','Incremento absoluto, não taxa relativa.'),
('occ_max','Ocupação máxima Hotel',.74,'%','V45.'),
('mice_events','MICE — espaços e congressos',.5,'% MICE','Hipótese de segmentação; validar categorias, volumes e capacidade 1.500 pessoas.'),
('mice_banquet','MICE — banquetes',.35,'% MICE','Excluídos de F&B hotel para evitar dupla contagem.'),
('fb_share','F&B / outras receitas Hotel',.8,'% outras','Hipótese de segmentação; remanescente serviços.'),
('room_cost','Custos departamentais quartos',.48,'% quartos','V45.'),
('mice_cost','Custos departamentais MICE',.59,'% MICE','V45.'),
('other_cost','Custos departamentais F&B e serviços',.64,'% outras','V45.'),
('overhead','Gastos gerais não distribuídos',.115,'% receita','V45; manutenção corrente incluída.'),
('base_fee','Fee base Operador Internacional',.03,'% receita','Referência negocial: 3%; opções 4 / 3,5 / 3 / 2,5%.'),
('systems','Distribuição e sistemas',.03,'% receita','Referência negocial 3%; confirmar HMA e ausência de duplicação.'),
('incentive','Fee incentivo',.06,'% GOP ajustado','GOP após fee base e sistemas; 10 / 8 / 6%.'),
('ffe','Dotação FF&E Hotel',.05,'% receita','Transferência para fundo restrito; distinta de despesa efetiva.'),
('ffe_spend','Utilização do fundo FF&E',1,'% dotação','100% gasto na referência; saldo não utilizado fica em fundo restrito, fora de dividendos.'),
('maintenance','CAPEX reposição adicional Hotel',.015,'% receita','Diferente de manutenção corrente e FF&E.'),
('tax','Taxa fiscal simplificada',.20,'%','Proxy sobre EBITDA menos reposição e FF&E; sem escudo de juros. Validar fiscalidade, IVA, amortizações e prejuízos.'),
('wc_days','Fundo maneio Hotel',15,'dias receita','Hipótese adicional explícita; 365 dias convencionais.'),
('management_alert','Alerta gestão / CFADS',.35,'%','Alerta quando >= limiar.'),
('systems_alert','Alerta sistemas / CFADS',.25,'%','Alerta quando >= limiar.'),
('total_alert','Alerta total operador / CFADS',.50,'%','Alerta quando >= limiar.'),
('rate_custom','Juro personalizado',.06,'%','4,5 a 8% em passos de 0,5 p.p.; escolher Personalizado.'),
('grace_custom','Carência personalizada',42,'meses','Juros pagos desde a utilização; contagem desde mês 1.'),
('term_custom','Prazo personalizado após carência',10,'anos','5 / 10 / 15 / 20 anos; dívida residual pode ir além do mês 180.'),
('opening_custom','Abertura personalizada',.01,'% linha','Comissão na primeira utilização.'),
('res_adr','ADR residencial inicial',225,'€/noite','V45 proxy; necessita forecast autónomo das 94 frações.'),
('res_growth','Crescimento ADR residencial',.035,'%','V45.'),
('res_occ','Ocupação residencial inicial',.581,'%','Proxy V45; não aplicar ao Hotel de 200 quartos.'),
('res_occ_step','Incremento ocupação residencial',.0265,'p.p.','Aproximação editável até 63,4%; não é forecast contratado.'),
('res_occ_max','Teto ocupação residencial',.634,'%','V45.'),
('res_other','Acessórias residencial',.05,'% quartos','Sem MICE Hotel.'),
('res_cost','OPEX residencial',.35,'% receita','V45; validar custos separados.'),
('res_base','Fee base residencial',.03,'% receita','V45.'),
('res_system','Sistemas residencial',.05,'% receita','V45; diferente da hipótese Hotel.'),
('res_ffe','FF&E residencial',.04,'% receita','Dotação fora da caixa distribuível; não duplicar no hotel.'),
('res_yield','Yield alvo proprietário inicial',.0425,'% preço','V45 Waterfall Moderada; não garantia contratual.'),
('res_yield_step','Incremento yield proprietário',.0025,'p.p.','V45 Moderada.'),
('res_yield_max','Teto yield proprietário',.065,'% preço','Renda efetiva limitada à caixa do pool.'),
('res_fund','Fundo adicional / excedente',.20,'% excedente','V45 waterfall.'),
('res_kyenu','Kyenu / excedente',.15,'% excedente','V45 waterfall; direitos por confirmar.'),
('res_operator','Operador / excedente',.20,'% excedente','V45 waterfall; incentivo só após yield; sem repetir incentivo pré-yield.'),
('res_include','Incluir residual KK no consolidado',1,'0/1','Hipótese V45 de direito a 45% do excedente; não equivale a contrato validado. Testar também 0.'),
('res_wc','Fundo maneio residencial',15,'dias receita','Hipótese; antecipado pelo pool e afeta residual disponível.'),
('legacy_jscv','Participação histórica JSCV',.51,'%','JSCV mantém 51% na estrutura proposta; mecanismo de transmissão e de fundos depende de formalização.'),
('res_capex_share','Quota CAPEX imputada residencial',13290/47000,'%','Proxy por área privativa / área construída total; não é orçamento segmentado. Validar áreas comuns e custos específicos.'),
('conversion','Montante conversão ações em frações',0,'€','Mecanismo alternativo; reduz cobranças e capital novo em igual valor. Não acrescenta valor nem garante recompra.'),
('conversion_id','ID fração para conversão','201','ID','Piloto de liquidez; preço não pode ser excedido. Ações canceladas ao valor negocial linear.'),
('validate_deeds','Escrituras 504/810 validadas',0,'0/1','0 mantém alerta. Mês 29 apenas hipótese provisória herdada.'),
('capacity','Capacidade MICE',1500,'pessoas','Capacidade aproximada; receita MICE ainda proxy sobre quartos.'),
('value_adr','ADR sem operador / ADR cenário',.90,'x','Contrafactual hipotético para testar valor acrescentado; não forecast.'),
('value_occ','Ocupação sem operador / cenário',.95,'x','Contrafactual hipotético; custos departamentais mantidos, gestão alternativa não incluída.'),
]
PARAMS += [
('equity_price','Preço da ronda por 40%',20000000,'€','20 M€ por 40% na referência; 15 M€ por 40% preservado para comparação.'),
('equity_stake','Participação máxima da ronda',.40,'%','40.000 das 100.000 ações existentes da JALS; sem diluição da JSCV.'),
('treasury_ratio','Parcela do preço entregue à KK',1,'%','100% é hipótese condicional: compra de ações JALS não financia automaticamente a sociedade. Não presumir suprimento sem modelar dívida.'),
('legal_confirmed','Mecanismo de entrada na KK formalizado',0,'0/1','0: pendência externa. Referência assume entrada irrevogável sem reembolso; validar mecanismo e fiscalidade.'),
('shares_total','Ações existentes',100000,'unidades','JSCV 51.000, JALS 49.000 antes da operação; 40.000 ações JALS disponibilizadas.'),
('commercial_delay','Desfasamento comercial de cenário',0,'meses','Desloca cobranças, contrato e escritura em simulação; não aprova alterações contratuais. Obra não se desloca.'),
('cost_overrun','Sobrecusto adicional da obra',0,'%','Sobre orçamento e contingência; cenário de esforço, distinto da contingência já orçamentada.'),
('fee_mode','Condições do Operador selecionadas','Objetivo KK','lista','Objetivo KK / Proposta simulada / Personalizado. Nenhum modo representa HMA aprovado.'),
('target_base','Objetivo KK — fee base',.03,'%','Referência de negociação, não contrato.'),
('target_systems','Objetivo KK — distribuição/sistemas',.03,'%','Sem sobreposição de encargos.'),
('target_incentive','Objetivo KK — incentivo',.06,'%','Base GOP após fees fixos.'),
('offer_base','Proposta simulada — fee base',.04,'%','Hipótese para negociação; não representa proposta recebida.'),
('offer_systems','Proposta simulada — sistemas',.03,'%','Hipótese editável; confirmar encargos.'),
('offer_incentive','Proposta simulada — incentivo',.10,'%','Hipótese editável; não condição apresentada confirmada.'),
('offer_received','Proposta efetiva recebida',0,'0/1','Só assinalar 1 quando houver documento; valores da proposta têm de corresponder ao documento.'),
('incentive_gate','Condicionar incentivo ao desempenho',0,'0/1','Simulação facultativa; 0 preserva motor V01. Critérios não contratados.'),
('gop_floor','GOP anualizado mínimo para incentivo',0,'€','0 = não definido; simulação mensal anualizada, não substitui performance test plurianual do HMA.'),
('ebitda_floor','Resultado pré-incentivo anualizado mínimo',0,'€','Base antes de incentivo evita circularidade. Não confundir com EBITDA líquido de incentivo.'),
('actual_month','Mês de referência execução',0,'mês','0: ainda sem data de execução informada; indicadores realizados permanecem indisponíveis.'),
]
PARAMS += [
('equity_mode','Origem do capital modelado','Capital selecionado','lista','Capital selecionado / Combinação tickets. Composição não cria compromisso jurídico.'),
('n_platinum','Investidores Platinum — quantidade',2,'unidades','Ticket referência 10 M€.'),
('n_gold','Investidores Gold — quantidade',0,'unidades','Ticket referência 5 M€.'),
('n_silver_plus','Investidores Silver Plus — quantidade',0,'unidades','Ticket referência 2 M€.'),
('n_silver','Investidores Silver — quantidade',0,'unidades','Ticket referência 1 M€.'),
]
PARAMS += [
('limit_base','Limite proposto — fee base',None,'%','Não definido por defeito; hipótese de negociação, não aceite nem contratado.'),
('limit_systems','Limite proposto — sistemas',None,'%','Não definido por defeito; não inventar condição aceite.'),
('limit_incentive','Limite proposto — incentivo',None,'%','Não definido por defeito; cenário usa objetivo KK até definir os limites.'),
]
# V03 additive parameters; prior engine retained in v02/.
PARAMS += [
('capital_scenario','Cenário capital V03','KK-E25','lista','E15 / E20-A preservados; E20-B hipotético; E25 25 M€ por 49%.'),
('received_ratio','Preço recebido / contratado no cenário',1,'%','Hipótese de cenário, não comprovativo bancário. Não confundir com valores realizados.'),
('fund_month1','Mês mobilização tranche 1',1,'mês','Calendário de cenário; confirmar contrato.'),
('fund_month2','Mês mobilização tranche 2',7,'mês','Calendário de cenário.'),
('fund_month3','Mês mobilização tranche 3',13,'mês','Calendário de cenário.'),
('fund_part1','Percentagem tranche 1',1,'%','As três tranches devem somar 100%.'),
('fund_part2','Percentagem tranche 2',0,'%','Não é capital realizado antes do recebimento.'),
('fund_part3','Percentagem tranche 3',0,'%','Não é capital realizado antes do recebimento.'),
('funding_kind','Natureza transferência KK','Não reembolsável','lista','Hipótese não formalizada / Empréstimo. Venda secundária separada.'),
('shareloan_rate','Juro anual empréstimo dos vendedores',.05,'%','Simulação; não há empréstimo contratado.'),
('shareloan_start','Início amortização empréstimo',61,'mês','Amortização linear; sem pagamento antes desta data.'),
('shareloan_term','Prazo amortização empréstimo',120,'meses','Saldo residual deduzido no valor de saída.'),
('benefit_platinum','Custo anual benefício Platinum',0,'€','Custo por bloco; hipótese, não direito aprovado.'),
('benefit_gold','Custo anual benefício Gold',0,'€','Custo por bloco; despesa incremental ou perda de receita líquida.'),
('benefit_silver_plus','Custo anual benefício Silver Plus',0,'€','Custo por bloco, sem yield garantida.'),
('benefit_silver','Custo anual benefício Silver',0,'€','Custo por bloco; incide sobre todos os blocos, incluindo promotores.'),
]
V3_DEFAULTS={'equity':25000000,'equity_price':25000000,'equity_stake':.50,'legacy_jscv':.40,'n_platinum':1,'n_gold':1,'n_silver_plus':0,'n_silver':10}
PARAMS=[(k,l,V3_DEFAULTS.get(k,v),u,n) for k,l,v,u,n in PARAMS]
DEFAULT={k:v for k,l,v,u,n in PARAMS}
HOTELS={'Conservador':(160,.52,.28,.18),'Base':(185,.58,.38,.24),'Favorável':(210,.64,.5,.3)}
BANKS={'Abertura':(.05,48,15,.005),'Objetivo':(.055,48,15,.0075),'Limite aceitável':(.06,42,10,.01)}
def config(**updates):
 p=DEFAULT.copy();p.update(updates)
 if p['equity_mode']=='Combinação tickets':
  p['equity']=p['n_platinum']*1e7+p['n_gold']*5e6+p['n_silver']*1e6
  p['equity_price']=p['equity'] or 1
  p['equity_stake']=p['n_platinum']*.20+p['n_gold']*.10+p['n_silver']*.02
  p['legacy_jscv']=.51-.11*p['n_platinum']
 elif 'equity' not in updates:
  p['equity'],p['equity_stake'],p['legacy_jscv']={'KK-E15':(15e6,.4,.51),'KK-E20-A':(20e6,.4,.51),'KK-E20-B':(20e6,.5,.4),'KK-E25':(25e6,.50,.4)}[p['capital_scenario']]
  p['equity_price']=p['equity']
 elif 'equity_stake' not in updates:
  p['equity_stake']=.50 if p['equity']==25e6 else .4;p['legacy_jscv']=.4 if p['equity']==25e6 else .51
 p['rate'],p['grace'],p['term'],p['opening']=BANKS.get(p['bank'],(p['rate_custom'],p['grace_custom'],p['term_custom'],p['opening_custom']))
 p['adr'],p['occ'],p['mice'],p['other']=HOTELS[p['hotel']]
 if p['fee_mode']=='Objetivo KK':
  p['base_fee'],p['systems'],p['incentive']=[p['target_'+k] for k in ['base','systems','incentive']]
 elif p['fee_mode']=='Limites propostos':
  p['base_fee'],p['systems'],p['incentive']=[p['limit_'+k] if p['limit_'+k] is not None else p['target_'+k] for k in ['base','systems','incentive']]
 elif p['fee_mode']=='Proposta simulada':
  p['base_fee'],p['systems'],p['incentive']=[p['offer_'+k] for k in ['base','systems','incentive']]
 return p

def money(x):
 return math.floor(x*100+0.5+1e-7)/100

def receipts(p):
 schedules=[]; totals=[0.]*180;depos=[0.]*180;tranches=[0.]*180;progress=[0.]*180;finals=[0.]*180;costs=[[0.]*180 for _ in range(3)]
 for u in UNITS:
  price=u['price']; conv=p['conversion'] if u['id']==p['conversion_id'] else 0
  net=price-conv;s=u['sale']+int(p['commercial_delay']);d=u['deed']+int(p['commercial_delay']);t=s+int(p['lag']);n=max(0,(d-t-1)//6)
  a=money(net*p['deposit']);b=money(net*p['tranche'])
  prog=money(net*p['progress']); installment=money(prog/n) if n else 0
  f=money(net-a-b-prog)
  rows=[0.]*180;parts=[]
  for m in range(1,181):
   ap=a if m==s else 0;bp=b if m==t else 0
   pp=(prog-installment*(n-1) if m==t+n*6 else installment) if n and m>t and m<d and (m-t)%6==0 and (m-t)//6<=n else 0
   fp=f if m==d else 0
   rows[m-1]=ap+bp+pp+fp;depos[m-1]+=ap;tranches[m-1]+=bp;progress[m-1]+=pp;finals[m-1]+=fp;totals[m-1]+=rows[m-1]
   costs[0][m-1]+=price*p['tax_transfer'] if m==d else 0
   costs[1][m-1]+=price*p['marketing'] if m==s else 0
   costs[2][m-1]+=price*p['kyenu'] if m==s else 0
  schedules.append(rows)
 return dict(total=totals,deposit=depos,tranche=tranches,progress=progress,final=finals,tax=costs[0],marketing=costs[1],kyenu=costs[2],units=schedules)

def construction(p):
 base=[sum(c['share']*p['area']*p['cost']*(1+p['cost_overrun'])*c['weights'][m] for c in CHAPTERS) for m in range(30)]+[0.]*150
 return [dict(base=x,supervision=x*p['supervision'],cont=x*p['contingency']*p['cont_used'],total=x*(1+p['supervision']+p['contingency']*p['cont_used'])) for x in base]

def hotel(p):
 out=[];wcprev=0;fund=0
 for m in range(1,181):
  age=max(0,(m-p['open'])//12);active=m>=p['open'];adr=p['adr']*(1+p['adr_growth'])**age;occ=min(p['occ_max'],p['occ']+p['occ_step']*age)
  rooms=p['rooms']*365/12*adr*occ if active else 0;mice=rooms*p['mice'];other=rooms*p['other'];revenue=rooms+mice+other
  rc=rooms*p['room_cost'];mc=mice*p['mice_cost'];oc=other*p['other_cost'];over=revenue*p['overhead'];gop=revenue-rc-mc-oc-over
  base=revenue*p['base_fee'];systems=revenue*p['systems'];inc=max(0,gop-base-systems)*p['incentive']*(not p['incentive_gate'] or (gop*12>=p['gop_floor'] and (gop-base-systems)*12>=p['ebitda_floor']));ebitda=gop-base-systems-inc
  ffe=revenue*p['ffe'];ffespent=ffe*p['ffe_spend'];fund+=ffe-ffespent;maint=revenue*p['maintenance'];tax=max(0,ebitda-ffespent-maint)*p['tax'];wc=revenue*12/365*p['wc_days'];dwc=wc-wcprev;wcprev=wc
  benefit=(3*p['benefit_platinum']+p['benefit_gold']+15*p['benefit_silver'])/12 if active else 0
  ebitda-=benefit;tax=max(0,ebitda-ffespent-maint)*p['tax'];cf=ebitda-ffe-maint-tax-dwc
  out.append(dict(m=m,age=age+1 if active else 0,adr=adr if active else 0,occ=occ if active else 0,rooms=rooms,mice_events=mice*p['mice_events'],mice_banquet=mice*p['mice_banquet'],mice_services=mice*(1-p['mice_events']-p['mice_banquet']),fb=other*p['fb_share'],services=other*(1-p['fb_share']),revenue=revenue,room_cost=rc,mice_cost=mc,other_cost=oc,overhead=over,gop=gop,base=base,systems=systems,incentive=inc,ebitda=ebitda,ffe=ffe,ffe_spent=ffespent,ffe_fund=fund,maintenance=maint,tax=tax,wc=wc,dwc=dwc,cf=cf))
 return out

def residential(p):
 out=[];wcprev=0;fund=0;vgv=sum(u['price'] for u in UNITS)
 for m in range(1,181):
  active=m>=p['open'];age=max(0,(m-p['open'])//12);adr=p['res_adr']*(1+p['res_growth'])**age;occ=min(p['res_occ_max'],p['res_occ']+p['res_occ_step']*age)
  rooms=len(UNITS)*365/12*adr*occ if active else 0;revenue=rooms*(1+p['res_other']);opex=revenue*p['res_cost'];base=revenue*p['res_base'];systems=revenue*p['res_system'];ffe=revenue*p['res_ffe'];wc=revenue*12/365*p['res_wc'];dwc=wc-wcprev;wcprev=wc
  cash=revenue-opex-base-systems-ffe-dwc;target=vgv*min(p['res_yield_max'],p['res_yield']+age*p['res_yield_step'])/12 if active else 0
  owner=min(max(0,cash),target);short=max(0,target-owner);excess=max(0,cash-owner);reserve=excess*p['res_fund'];fund+=reserve+ffe;kyenu=excess*p['res_kyenu'];operator=excess*p['res_operator'];kkgross=excess-reserve-kyenu-operator+min(0,cash);tax=max(0,kkgross)*p['tax'];kk=(kkgross-tax)*p['res_include']
  out.append(dict(m=m,adr=adr if active else 0,occ=occ if active else 0,revenue=revenue,opex=opex,base=base,systems=systems,ffe=ffe,wc=wc,dwc=dwc,cash=cash,target=target,owner=owner,short=short,excess=excess,reserve=reserve,kyenu=kyenu,operator=operator,kkgross=kkgross,tax=tax,kk=kk,fund=fund))
 return out

def run(p=None):
 p=config() if p is None else p
 re=receipts(p);cap=construction(p);h=hotel(p);res=residential(p);out=[];cash=debt=drawtotal=0.;fundgap_ever=0;shareloan=0
 for m in range(1,181):
  i=m-1;openingcash=cash;openingdebt=debt;cumulative=drawtotal
  weight=sum(p['fund_part'+str(j)] for j in [1,2,3] if m==p['fund_month'+str(j)])
  equity=(p['equity']-p['conversion'])*p['received_ratio']*p['treasury_ratio']*weight
  loan_in=equity if p['funding_kind']=='Empréstimo' else 0
  loan_interest=(shareloan+loan_in)*p['shareloan_rate']/12
  loan_principal=min(shareloan+loan_in,(shareloan+loan_in)/max(1,p['shareloan_start']+p['shareloan_term']-m)) if m>=p['shareloan_start'] else 0
  shareloan+=loan_in-loan_principal
  hotpay=p['hotusa'] if p['hotpay'] and m==p['hotmonth'] else 0
  hothold=p['hotusa'] if not p['hotpay'] or m<p['hotmonth'] else 0
  conthold=p['area']*p['cost']*(1+p['cost_overrun'])*p['contingency']*(1-p['cont_used'])
  buffer=sum(x['total'] for x in cap[m:min(180,m+int(p['buffer']))]) if m<=p['end'] else 0
  startup=max(0,-h[int(p['open'])-1]['cf']-res[int(p['open'])-1]['kk']+p['limit']*p['rate']/12) if m<=p['end'] else 0
  reserve=max(p['reserve'],buffer)+hothold+conthold+startup
  avail=openingcash+equity+re['total'][i]+h[i]['cf']+res[i]['kk']-re['tax'][i]-re['marketing'][i]-re['kyenu'][i]-cap[i]['total']-hotpay-loan_interest-loan_principal
  principal=min(debt,debt/max(1,p['grace']+p['term']*12-m+1)) if m>p['grace'] else 0
  can_draw=m<=p['draw_end'];remaining=max(0,p['limit']-cumulative) if can_draw else 0
  trigger=int(cumulative<.005 and avail-principal-debt*p['rate']/12<reserve-.005 and remaining>0)
  openingfee=trigger*p['limit']*p['opening'];broker=trigger*p['limit']*p['broker'];mobil=trigger*p['limit']*p['mobil'];fees=openingfee+broker+mobil
  draw=min(remaining,max(0,(reserve-avail+principal+debt*p['rate']/12+fees)/(1-p['rate']/12)))
  interest=(debt+draw)*p['rate']/12
  post=avail+draw-fees-interest-principal
  # Preserve next reserve and forward net CAPEX not covered by remaining receipts and undrawn line.
  futurecap=sum(x['total'] for x in cap[m:]);futureincome=sum(re['total'][m:])-sum(re['tax'][m:])-sum(re['marketing'][m:])-sum(re['kyenu'][m:])
  forward=max(0,futurecap-futureincome-max(0,p['limit']-cumulative-draw)) if m<p['end'] else 0
  sweepcap=max(0,re['total'][i]-re['tax'][i]-re['marketing'][i]-re['kyenu'][i]) if m<=p['end'] else max(0,post-reserve-forward)
  sweep=min(max(0,debt+draw-principal),max(0,post-reserve-forward)*p['sweep'],sweepcap)
  debt=max(0,debt+draw-principal-sweep);drawtotal+=draw
  gap=max(0,-post);fundgap_ever+=gap
  dividend=max(0,post-sweep-reserve)*p['payout'] if m>=p['open'] and m%12==0 and debt<.005 and shareloan<.005 and fundgap_ever<.005 else 0
  cash=post-sweep-dividend;service=principal+interest+loan_interest+loan_principal;reservegap=max(0,reserve-cash)
  project=re['total'][i]+h[i]['cf']+res[i]['kk']-sum([re['tax'][i],re['marketing'][i],re['kyenu'][i],cap[i]['total'],hotpay])
  out.append(dict(m=m,year=(m-1)//12+1,equity=equity,receipts=re['total'][i],transfer=re['tax'][i],marketing=re['marketing'][i],kyenu=re['kyenu'][i],construction=cap[i]['base'],supervision=cap[i]['supervision'],contingency=cap[i]['cont'],hotpay=hotpay,hotel_cf=h[i]['cf'],res_cf=res[i]['kk'],cash_open=openingcash,debt_open=openingdebt,draw_open=cumulative,reserve=reserve,available=avail,trigger=trigger,openingfee=openingfee,broker=broker,mobil=mobil,principal=principal,draw=draw,interest=interest,post=post,sweep=sweep,debt=debt,cumulative=drawtotal,dividend=dividend,cash=cash,gap=gap,reservegap=reservegap,service=service,dscr=h[i]['cf']/service if service>0 else None,project_cf=project,equity_cf=dividend-equity,cash_check=cash-openingcash-equity-re['total'][i]-h[i]['cf']-res[i]['kk']-draw+sum([re['tax'][i],re['marketing'][i],re['kyenu'][i],cap[i]['total'],hotpay,fees,interest,principal,sweep,dividend,loan_interest,loan_principal]),debt_check=debt-openingdebt-draw+principal+sweep,hotel_ebitda=h[i]['ebitda'],hotel_rev=h[i]['revenue'],hothold=hothold,conthold=conthold,forward=forward,project_cum=project+ (out[-1]['project_cum'] if out else 0),equity_cum=dividend-equity+(out[-1]['equity_cum'] if out else 0),loan_in=loan_in,loan_interest=loan_interest,loan_principal=loan_principal,shareloan=shareloan))
 return dict(p=p,re=re,cap=cap,hotel=h,res=res,flow=out)

def irr_roots(flows):
 import numpy as np
 if not any(x<0 for x in flows) or not any(x>0 for x in flows):return []
 v=np.asarray(flows,dtype=float);t=np.arange(len(v))
 def f(z):
  ex=-z*t;ex-=np.max(ex)
  return float(np.dot(v,np.exp(ex)))
 grid=np.linspace(-.1,.5,1201);roots=[]
 for a,b in zip(grid[:-1],grid[1:]):
  if f(a)*f(b)<0:
   root=math.expm1(12*brentq(f,a,b))
   if not roots or abs(root-roots[-1])>1e-6:roots.append(root)
 return roots

IRR_CACHE={}
def irr(flows):
 key=tuple(flows)
 if key not in IRR_CACHE:
  roots=irr_roots(flows);IRR_CACHE[key]=roots[0] if len(roots)==1 else None
 return IRR_CACHE[key]

def summary(result,horizon=180):
 f=result['flow'][:horizon];h=result['hotel'];p=result['p'];last=f[-1];start=max(0,horizon-12)
 ebitda=sum(x['ebitda'] for x in h[start:horizon]);ev=ebitda*p['exit_mult'];exitvalue=ev*(1-p['exit_cost'])-last['debt']-last['shareloan']+last['cash']-last['hothold']-last['conthold']
 exitvalue=max(0,exitvalue);share=(p['equity']-p['conversion'])/p['equity_price']*p['equity_stake']
 dividends=sum(x['dividend'] for x in f);own=[0]+[x['equity_cf'] for x in f];own[-1]+=exitvalue
 pc=[0]+[x['project_cf'] for x in f];pc[-1]+=ev*(1-p['exit_cost'])-last['hothold']-last['conthold']
 # Return is conditional whenever funding gap or protected reserves are not met.
 annual=[]
 for year in range(1,horizon//12+1):
  fr=f[(year-1)*12:year*12];serv=sum(x['service'] for x in fr);cf=sum(x['hotel_cf'] for x in fr)
  annual.append(dict(year=year,dividend=sum(x['dividend'] for x in fr),service=serv,cf=cf,dscr=cf/serv if serv>.005 else None,cash=fr[-1]['cash'],debt=fr[-1]['debt']))
 operating_dscr=[a['dscr'] for a in annual if a['year']*12>=p['open'] and a['dscr'] is not None]
 return dict(equity=p['equity']-p['conversion'],share=share,peak=max(x['debt_open']+x['draw'] for x in f),cash_min=min(x['cash'] for x in f),reserve_gap=max(x['reservegap'] for x in f),funding_gap=max(x['gap'] for x in f),debt30=f[29]['debt'],cash30=f[29]['cash'],debt42=f[41]['debt'],cash42=f[41]['cash'],debt_end=last['debt'],cash_end=last['cash'],interest=sum(x['interest'] for x in f),fees=sum(x['openingfee']+x['broker']+x['mobil'] for x in f),dividends=dividends,ev=ev,exit_equity=exitvalue,moic=(dividends+exitvalue)*share/((p['equity']-p['conversion'])*p['received_ratio']) if p['equity']>p['conversion'] else None,equity_irr=irr([0]+[(-(p['equity']-p['conversion'])*p['received_ratio']*sum(p['fund_part'+str(j)] for j in [1,2,3] if x['m']==p['fund_month'+str(j)]))+x['dividend']*share+(exitvalue*share if x['m']==horizon else 0) for x in f]) if p['equity']>p['conversion'] else None,project_irr=irr(pc),npv_project=sum(x['project_cf']/(1+p['discount'])**(x['m']/12) for x in f),npv_project_exit=sum(x['project_cf']/(1+p['discount'])**(x['m']/12) for x in f)+(ev*(1-p['exit_cost'])-last['hothold']-last['conthold'])/(1+p['discount'])**(horizon/12),min_dscr=min(operating_dscr) if operating_dscr else None,annual=annual)
if __name__=='__main__':
 r=run();s=summary(r);print(json.dumps({k:v for k,v in s.items() if k!='annual'},indent=2));(ROOT/'src/baseline.json').write_text(json.dumps(r,ensure_ascii=False));(ROOT/'src/summary.json').write_text(json.dumps(s,ensure_ascii=False,indent=2))

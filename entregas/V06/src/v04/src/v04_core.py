# Executed after the inherited core definitions, before financing blocks are written.
_old_res=res_formulas
def res_formulas(r,extra_delay=0,extra_deed=0):
 d=_old_res(r);m='A'+str(r);sale=f'(\'Unidades\'!$I$5:$I$98+{pv("commercial_delay")}+{extra_delay})';deed=f'(\'Unidades\'!$L$5:$L$98+{pv("commercial_delay")}+{pv("deed_delay")}+{extra_delay+extra_deed})';deed=f'IF({pv("construction_delay")}>0,IF({deed}<{pv("end")},{pv("end")},{deed}),{deed})'
 eligible=f'SUMPRODUCT(--(({deed})<={m}),--({sale}<={pv("sales_stop")}))';vgv=f'SUMPRODUCT(\'Unidades\'!$H$5:$H$98,--(({deed})<={m}),--({sale}<={pv("sales_stop")}))'
 d['revenue']=f'{eligible}*365/12*B{r}*C{r}*(1+{pv("res_other")})';d['target']=d['target'].replace("SUM('Unidades'!$H$5:$H$98)",vgv);return d

for i,x in enumerate(BASE['res']):
 r=i+5;fs=res_formulas(r)
 for k in RK[1:]:fw(S['Residencial'],r,RIDX[k],fs[k],x[k],'pct' if k=='occ' else 'eur')
_old_flow=flow_formulas
def flow_formulas(r,first=5,last=184,equity_expr=None,hotel_sn='Hotel',hotel_offset=0,bank_override=None,capex_sn='CAPEX',re_sn='Imobiliario',cost_factor='1',residential_sn='Residencial',funding_delay=0,capital_sn='V04 Capital180'):
 d=_old_flow(r,first,last,equity_expr,hotel_sn,hotel_offset,bank_override,capex_sn,re_sn,cost_factor)
 if residential_sn!='Residencial':d={k:v.replace("'Residencial'!",f"'{residential_sn}'!") for k,v in d.items()}
 a=lambda k:col(FI[k])+str(r);prev=lambda k:'0' if r==first else col(FI[k])+str(r-1);m=a('m');sr=r-first+5
 eq=equity_expr or pv('equity');weightscol='B' if equity_expr else 'C';fundweights=f'IF(AND({m}-{funding_delay}>=1,{m}-{funding_delay}<=180),INDEX(\'{capital_sn}\'!${weightscol}$5:${weightscol}$184,{m}-{funding_delay}),0)'
 oldweight='+'.join(f'IF({m}={pv("fund_month"+str(j))},{pv("fund_part"+str(j))},0)' for j in [1,2,3])
 d['equity']=f'({eq}-{pv("conversion")})*{rawpv("treasury_ratio")}*IF({pv("capital_schedule")}="V03",{rawpv("received_ratio")}*({oldweight}),{fundweights})'
 d['commitment']=f'IF(AND({m}>={pv("bank_start")},{m}<={pv("draw_end")}),MAX(0,{pv("limit")}-{a("draw_open")})*{pv("commit_fee")}/12,0)'
 d['available']+='-'+a('commitment');d['service']+='+'+a('commitment');d['cash_check']+='+'+a('commitment')
 mobile=f'{pv("limit")}*{pv("bank_available")}';capPaid=f'SUM(\'{capex_sn}\'!$E$5:E{sr})';eqPaid=f'SUM(C${first}:C{r})';pp=f'IF({pv("pari_passu")}=1,{capPaid}*{pv("bank_capex_share")},{mobile})';can=f'AND({m}>={pv("bank_start")},{m}<={pv("draw_end")},{pv("bank_cp")}=1,{eqPaid}*IF({pv("funding_kind")} ="Empréstimo",{pv("subordinated_eligible")},1)+{pv("eligible_history")}>={capPaid}*{pv("min_equity_ratio")})';remain=f'IF({can},MAX(0,MIN({mobile},{pp})-{a("draw_open")}),0)'
 # Replace the previous remaining-line expression in the single bank engine.
 oldremain=f'IF({m}<={pv("draw_end")},MAX(0,{pv("limit")}-{a("draw_open")}),0)'
 for k in ['trigger','draw']:d[k]=d[k].replace(oldremain,remain)
 d['forward']=d['forward'].replace(f'MAX(0,{pv("limit")}-{a("draw_open")}-{a("draw")})',f'IF(AND({m}<{pv("draw_end")},{pv("bank_cp")}=1),MAX(0,{mobile}-{a("draw_open")}-{a("draw")}),0)')
 d['mobile_line']=f'IF({can},MAX(0,MIN({mobile},{pp})-{a("cumulative")}),0)'
 # Conservative earnings proxy, separate from statutory approval.
 sales=f'SUMIFS(\'Unidades\'!$H$5:$H$98,\'Unidades\'!$Z$5:$Z$98,{m},\'Unidades\'!$X$5:$X$98,"<="&{pv("sales_stop")} )'
 if re_sn!='Imobiliario':sales=f'INDEX(\'{re_sn}\'!$M$5:$M$184,{m})'
 profit=f'{sales}*(1-{pv("res_capex_share")}*SUM(\'{capex_sn}\'!$E$5:$E$184)/SUM(\'Unidades\'!$H$5:$H$98)-{pv("marketing")}-{pv("kyenu")}-{pv("tax_transfer")})+{a("hotel_cf")}+\'{hotel_sn}\'!AA{sr+hotel_offset}+{a("res_cf")}-{a("interest")}-SUM({a("openingfee")}:{a("mobil")})-{a("loan_interest")}-{a("commitment")}'
 d['profits']=f'{prev("profits")}+({profit})';d['legal_headroom']=f'IF({pv("payout_basis")}="Documentado",MAX(0,N({pv("legal_dividend_cap")})-{prev("dividends_cum")}),MAX(0,{a("profits")}-{prev("dividends_cum")}))';d['dividends_cum']=f'{prev("dividends_cum")}+{a("dividend")}'
 d['dividend']=d['dividend'].replace(f'MOD({m},12)=0',f'MOD({m},{pv("dividend_frequency")})=0').replace(f'MAX(0,{a("post")}-{a("sweep")}-{a("reserve")})*{pv("payout")}',f'MIN({a("legal_headroom")},MAX(0,{a("post")}-{a("sweep")}-{a("reserve")})*{pv("payout")})')
 return d

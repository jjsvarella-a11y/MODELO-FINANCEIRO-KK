def projected_bank_formula(fs,cap,pv,open_expr=None,limit_expr=None,available_expr=None):
 end=open_expr or pv('open');lim=limit_expr or pv('limit');av=available_expr or pv('bank_available');window=f'MAX(1,MIN(180,{pv("draw_end")},{end}))';cap_done=f'SUM(\'{cap}\'!E5:INDEX(\'{cap}\'!E5:E184,{window}))';eq=f'SUM(\'{fs}\'!C5:INDEX(\'{fs}\'!C5:C184,{window}))*IF({pv("funding_kind")}="Empréstimo",{pv("subordinated_eligible")},1)+{pv("eligible_history")}'
 ceiling=f'MIN(({lim})*({av}),IF({pv("pari_passu")}=1,({cap_done})*{pv("bank_capex_share")},({lim})*({av})))';cond=f'AND({pv("bank_cp")}=1,24<{window},{pv("bank_start")}<={window},({eq})>=({cap_done})*{pv("min_equity_ratio")})'
 return f'IF({cond},MAX(0,({ceiling})-\'{fs}\'!AC28),0)'

def terminal_assets_formula(fs,hs,pv):return f'MAX(0,SUM(\'{hs}\'!T173:T184)*{pv("exit_mult")}*(1-{pv("exit_cost")})-\'{fs}\'!AB184+\'{fs}\'!AE184-\'{fs}\'!AP184-\'{fs}\'!AQ184)'

def note_method(key):
 if key in ['equityirr','projectirr','projectirrexit','conventional']:return '(1+TIR mensal)^12−1; fluxos positivos e negativos em meses coerentes.'
 if key in ['npv','npvexit','equitynpv']:return 'Desconto mensal pela taxa anual efetiva convertida; pagamentos e recebimentos identificados na fonte.'
 if key in ['equitymoic']:return 'Total recebido / capital pago; sem efeito do prazo de recebimento.'
 if key in ['dscrstub','dscrfull']:return 'CFADS / serviço obrigatório; amortização voluntária excluída; sem serviço o rácio é indefinido.'
 if key in ['cfc_projected','cfc_secured']:return 'Fontes elegíveis / custos e reservas remanescentes; documentos e calendário de mobilização separados.'
 return 'Cálculo descrito na definição; fórmula integral e dependências no DicionárioKPI,colunaH.'

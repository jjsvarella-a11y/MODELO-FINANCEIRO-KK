"""10-year returns use the selected15-year policy, not a re-optimized10-year target."""
def returns10_updates(pv):
 u={};put=lambda a,f,fmt='eur':u.__setitem__(a,('formula',f,fmt));text=lambda a,v:u.__setitem__(a,('text',v,'text'))
 legacy=f'AND({pv("capital_scenario")}<>"KK-E25",{pv("equity_mode")}<>"Combinação tickets")';mode=f'IF({legacy},"Proporcional",{pv("waterfall_mode")})'
 put('M130',f'MAX(0,SUM(\'Hotel\'!T113:T124)*{pv("exit_mult")}*(1-{pv("exit_cost")})-\'Fluxo 180M\'!AB124+\'Fluxo 180M\'!AE124-\'Fluxo 180M\'!AP124-\'Fluxo 180M\'!AQ124)');put('M131',"MAX(0,M130-'Fluxo 180M'!AX124)");put('M132',"MIN('Fluxo 180M'!AX124,M130)")
 for cell,f in [('M133',"MIN(M131,MAX(0,'V04 Waterfall180'!M124))"),('M134',"MIN(MAX(0,M131-M133),MAX(0,'V04 Waterfall180'!N124))"),('M135',"MIN(MAX(0,M131-M133-M134),MAX(0,'V04 Waterfall180'!O124))"),('M136','MAX(0,M131-SUM(M133:M135))')]:put(cell,f)
 for k,share in enumerate([.4,.3,.3]):put('M'+str(137+k),f'IF(AND({mode}="Otimizado",\'V04 Politica\'!B30),M131*\'V04 Politica\'!K{k+5},IF(OR({mode}="Saída",{mode}="Combinado",{mode}="Otimizado"),M{133+k}+M136*{share},M131*{share}))')
 put('M140',f'IF({legacy},M131*\'V04 Capital180\'!G124,M137/2*\'V04 Waterfall180\'!Q124+M138/3*\'V04 Waterfall180\'!R124+M139/15*\'V04 Waterfall180\'!S124)');put('M141','M131-M140+M132');put('M142',"SUM('V04 Waterfall180'!U5:U124)");put('M143',"SUM('V04 Waterfall180'!T5:T124)+M140");put('M144',"SUM('V04 Waterfall180'!Y5:Y124)+M141")
 for m in range(121):
  r=m+5;fr=m+4;put('M'+str(r),str(m),'int')
  for k,(ticket,count,letter) in enumerate([(1e7,2,'N'),(5e6,3,'O'),(1e6,15,'P')]):
   f='0' if m==0 else f"'V04 Waterfall180'!{chr(71+k)}{fr}/{count}"+(f'-{ticket}' if m==1 else '')+(f'+M{137+k}/{count}' if m==120 else '');put(letter+str(r),f)
  put('Q'+str(r),'0' if m==0 else f"'V04 Waterfall180'!V{fr}"+('+M140' if m==120 else ''));put('R'+str(r),f'-{pv("historic")}' if m==0 else f"'V04 Waterfall180'!Y{fr}"+('+M141' if m==120 else ''))
 text('A45','Horizonte10anos — política15Aselecionada, saída económica hipotética')
 for c,label in zip('ABCDEFG',['Categoria','Base convencional','Dividendos10A','Saída10A','MOIC10A','TIR10A','VPL10A']):text(c+'46',label)
 for k,(label,ticket,count,letter) in enumerate([('Platinum',1e7,2,'N'),('Gold',5e6,3,'O'),('Silver',1e6,15,'P')]):
  r=47+k;text('A'+str(r),label);put('B'+str(r),str(ticket));put('C'+str(r),f'SUM(\'V04 Waterfall180\'!{chr(71+k)}5:{chr(71+k)}124)/{count}');put('D'+str(r),f'M{137+k}/{count}');put('E'+str(r),f'(C{r}+D{r})/B{r}','mult');put('F'+str(r),f'IFERROR((1+IRR({letter}5:{letter}125,0.005))^12-1,"TIR não definida")','pct');put('G'+str(r),f'NPV((1+{pv("equity_discount")})^(1/12)-1,{letter}6:{letter}125)')
 for r,label,f,fmt in [(52,'Valor saída equity10A','M131','eur'),(53,'Crédito recuperável10A','M132','eur'),(55,'TIR novos efetiva10A','IFERROR((1+IRR(Q5:Q125,0.005))^12-1,"TIR não definida")','pct'),(56,'MOIC novos10A','IF(M142>0,M143/M142,"")','mult'),(57,'VPL novos10A',f'NPV((1+{pv("equity_discount")})^(1/12)-1,Q6:Q125)','eur'),(58,'Promotores recebido10A','M144','eur'),(59,'TIR Grupo proxy10A','IFERROR((1+IRR(R5:R125,0.005))^12-1,"TIR não definida")','pct'),(60,'Novos dividendos10A','M143-M140','eur')]:text('A'+str(r),label);put('B'+str(r),f,fmt)
 return u

def payback_formula():
 return 'IF(SUM(\'V04 Waterfall180\'!U5:U184)<=0,"",IF(SUM(\'V04 Waterfall180\'!T5:T184)-\'V04 Retornos\'!B36<SUM(\'V04 Waterfall180\'!U5:U184),"Não recuperado",IFERROR(LOOKUP(2,1/((\'V04 Waterfall180\'!Z5:Z184-IF(\'V04 Waterfall180\'!A5:A184=180,\'V04 Retornos\'!B36,0))<0),\'V04 Waterfall180\'!A5:A184)+1,0)))'

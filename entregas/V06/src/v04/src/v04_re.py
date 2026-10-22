# Shared per-fraction receipt engine for stress cases and inherited delayed scenarios.
def write_case_re(code,rr,agg_existing=None):
 usn=f'V04 {code}Unid';us=sh(usn,['ID','Venda','Tranche','Escritura efetiva','Preço','Sinal','Tranche€','55%','Semestres','Final','Transmissão','Marketing','Kyenu'],'Preserva contrato original; calendário efetivo de cenário depende de construção. Renegociação jurídica pendente.');us.hide();recsn=f'V04 {code}Receb';receb=sh(recsn,['ID']+list(range(1,181)),'Mesmo cálculo5/20/55/20por fração.');receb.hide()
 for i,u in enumerate(UNITS):
  r=i+5;fw(us,r,0,vr('Unidades',r,0),u['id'],'text');fw(us,r,1,f'\'Unidades\'!I{r}+{pv("commercial_delay")}',None);fw(us,r,2,f'\'Unidades\'!K{r}+{pv("commercial_delay")}',None)
  deed=f'\'Unidades\'!L{r}+{pv("commercial_delay")}+{pv("deed_delay")}';fw(us,r,3,f'IF({pv("construction_delay")}>0,MAX({deed},{pv("end")}),{deed})',deed_month(rr['p'],u),'int')
  for c,uc in [(4,7),(5,13),(6,14),(7,15),(9,16),(10,18),(11,19),(12,20)]:fw(us,r,c,vr('Unidades',r,uc),None)
  fw(us,r,8,f'MAX(0,INT((D{r}-C{r}-1)/6))',None,'int');fw(receb,r,0,f"'{usn}'!A{r}",u['id'],'text')
  for m in range(1,181):
   R=lambda L:f"'{usn}'!{L}{r}";n=R('I');t=R('C');d=R('D');prog=f'IF(AND({n}>0,{m}>{t},{m}<{d},MOD({m}-{t},6)=0,INT(({m}-{t})/6)<={n}),IF({m}={t}+{n}*6,{R("H")}-ROUND({R("H")}/{n},2)*({n}-1),ROUND({R("H")}/{n},2)),0)';formula=f'IF({R("B")}<={pv("sales_stop")},IF({m}={R("B")},{R("F")},0)+IF({m}={t},{R("G")},0)+{prog}+IF({m}={d},{R("J")},0),0)';fw(receb,r,m,formula,rr['re']['units'][i][m-1])
 agg=agg_existing or f'V04 {code}RE';rs=S[agg] if agg_existing else sh(agg,['Mês','Sinal','Tranche','Semestrais','Final','Recebimentos','Transmissão','Marketing','Kyenu','Líquido','Acumulado','VGV vendido','VGV escriturado'],'Agregação do mapa individual do cenário.');rs.hide()
 for m in range(1,181):
  r=m+4;text(rs,r,0,m,'int');fw(rs,r,5,f'SUM(\'{recsn}\'!{col(m)}5:{col(m)}98)',rr['re']['total'][m-1])
  for c,datec,valc,key in [(1,'B','F','deposit'),(2,'C','G','tranche'),(4,'D','J','final'),(6,'D','K','tax'),(7,'B','L','marketing'),(8,'B','M','kyenu'),(11,'B','E',None),(12,'D','E',None)]:fw(rs,r,c,f'SUMIFS(\'{usn}\'!${valc}$5:${valc}$98,\'{usn}\'!${datec}$5:${datec}$98,A{r},\'{usn}\'!$B$5:$B$98,"<="&{pv("sales_stop")})',rr['re'][key][m-1] if key else None)
  fw(rs,r,3,f'F{r}-B{r}-C{r}-E{r}',rr['re']['progress'][m-1]);fw(rs,r,9,f'F{r}-SUM(G{r}:I{r})',None);fw(rs,r,10,f'SUM(J$5:J{r})',None)
 return agg

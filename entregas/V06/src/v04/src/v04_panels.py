exec((ROOT/'src/v04_re.py').read_text(),globals())
exec((ROOT/'src/v04_finance.py').read_text(),globals())
exec((ROOT/'src/v04_execution.py').read_text(),globals())
# Physical completion bounds effective deed scenarios; source contract dates remain unchanged.
for ur in range(5,99):
 d=f"'Unidades'!L{ur}+"+pv('commercial_delay')+'+'+pv('deed_delay');fw(S['Unidades'],ur,25,f'IF({pv("construction_delay")}>0,MAX({d},{pv("end")}),{d})',deed_month(BASE['p'],UNITS[ur-5]),'int')
# Risk and executive interface are appended after financial validation.
if (ROOT/'src/v04_risks.py').exists():exec((ROOT/'src/v04_risks.py').read_text(),globals())
if (ROOT/'src/v04_ui.py').exists():exec((ROOT/'src/v04_ui.py').read_text(),globals())
# Correct the benefit inventory for the20-block V04structure, preserving all Hotel fee scenarios.
oldbenefit='(3*'+pv('benefit_platinum')+'+'+pv('benefit_gold')+'+15*'+pv('benefit_silver')+')/12';newbenefit='(2*'+pv('benefit_platinum')+'+3*'+pv('benefit_gold')+'+15*'+pv('benefit_silver')+')/12'
for name,hs in S.items():
 if name=='Hotel' or name.startswith('CalcH'):
  for r in range(5,185):
   cell=hs.table[r-1].get(HIDX['ebitda'])
   if cell and hasattr(cell,'formula'):fw(hs,r,HIDX['ebitda'],cell.formula.lstrip('=').replace(oldbenefit,newbenefit),None)
(ROOT/'src/v04_metadata.json').write_text(json.dumps({'param_rows':PR,'flow_columns':FI,'workflow':WORKFLOW,'metas_total':METAS_TOTAL,'phase1':'14tests passed; rights and secured funding not approved'},ensure_ascii=False,indent=2))

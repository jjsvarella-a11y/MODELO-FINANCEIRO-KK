from engine import *
from waterfall import constant_policy,priority,solve
from completion import completion
import unittest,copy
class ExecutionTests(unittest.TestCase):
 def test_01_secondary_sale_not_capital(self):
  p=config(treasury_ratio=0);r=run(p);self.assertEqual(sum(investor_outlay(p,m) for m in range(1,181)),25e6);self.assertEqual(sum(x['equity'] for x in r['flow']),0);self.assertGreater(summary(r)['funding_gap'],0)
 def test_02_advanced_commitments(self):
  cal={'Base':[1,2]+[30]*10};p=config(capital_calendar=cal,capital_stop=24);self.assertEqual(sum(x['planned_received'] for x in pipeline(p)),15e6);self.assertAlmostEqual(economic_share(p,180),.3)
 def test_03_paid_not_promised(self):
  p=config(capital_delay=24);self.assertEqual(economic_share(p,24),0);self.assertAlmostEqual(economic_share(p,25),.2);self.assertAlmostEqual(economic_share(p,180),.5)
 def test_04_collections_not_new_sales(self):
  p=config(commercial_delay=12,sales_stop=24);r=run(p);expected=sum(u['price'] for u in UNITS if u['sale']+12<=24);self.assertAlmostEqual(sum(r['re']['total']),expected,2)
  self.assertTrue(any(r['re']['total'][m-1]>0 for m in range(25,50)))
 def test_05_pricing_source_preserved(self):
  originals=copy.deepcopy(UNITS);p=config(pricing_mode='Motor',unit_factors={UNITS[0]['id']:{'view':1.2}});u=effective_units(p)[0];self.assertEqual(u['price'],money(u['area']*p['price_m2']*1.2));self.assertEqual(UNITS,originals)
 def test_06_no_false_sequential_delay(self):
  p=config(construction_delay=6,commercial_delay=6);u=next(x for x in UNITS if x['deed']==30);self.assertEqual(deed_month(p,u),36)
 def test_07_physical_delivery(self):
  p=config(construction_delay=6);self.assertTrue(all(deed_month(p,u)>=36 for u in UNITS));r=run(p);self.assertAlmostEqual(sum(r['re']['final'][:35]),0,2)
 def test_08_residential_available(self):
  r=run(config(deed_delay=6));self.assertEqual(r['res'][30]['revenue'],0);self.assertGreater(r['hotel'][30]['revenue'],0);self.assertGreater(r['res'][35]['revenue'],0)
 def test_09_opening_not_physical(self):
  r=run(config(construction_delay=6,opening_delay=3));self.assertEqual(r['p']['end'],36);self.assertEqual(r['p']['open'],40);self.assertEqual(r['hotel'][38]['revenue'],0);self.assertGreater(r['hotel'][39]['revenue'],0)
 def test_10_eac_future_only(self):
  a=run();b=run(config(status_month=12,etc_override=1,etc_amount=50e6));self.assertAlmostEqual(sum(x['total'] for x in b['cap'][12:]),50e6,2)
  self.assertEqual(a['cap'][:12],b['cap'][:12])
 def test_11_delay_cost(self):
  r=run(config(construction_delay=6,delay_cost=1e5));self.assertAlmostEqual(sum(x['total'] for x in r['cap']),63168000+600000,2)
 def test_12_bankCP(self):
  r=run(config(bank_cp=0));self.assertEqual(sum(x['draw'] for x in r['flow']),0);self.assertEqual(sum(x['openingfee']+x['broker']+x['mobil'] for x in r['flow']),0)
 def test_13_pari_passu(self):
  r=run(config(pari_passu=1,bank_capex_share=.5));spent=0
  for x in r['flow']:
   spent+=x['construction']+x['supervision']+x['contingency'];self.assertLessEqual(x['cumulative'],spent*.5+.01)
 def test_14_subordinated_not_automatic_equity(self):
  r=run(config(funding_kind='Empréstimo',min_equity_ratio=.3));self.assertEqual(sum(x['draw'] for x in r['flow']),0)
 def test_15_monthly_cash(self):
  r=run(config(dividend_frequency=1));p=priority(r,'Combinado');self.assertLess(p['check'],.001);self.assertTrue(solve(r,'Combinado')['feasible_targets'])
 def test_16_waterfall_infeasible(self):
  r=run(config(target_platinum=.2,target_gold=.2,target_silver=.2));self.assertFalse(constant_policy(r)['feasible']);self.assertFalse(solve(r,'Combinado')['feasible_targets'])
 def test_17_with_operator(self):
  for extra in [{'fee_mode':'Proposta simulada'},{'hotel':'Conservador'},{'cost_overrun':.1}]:
   r=run(config(**extra));w=constant_policy(r);result=w if w['feasible'] else priority(r,'Combinado');self.assertLess(result['check'],.001)
 def test_18_documented_completion_timing(self):
  c=completion(run(config(completion_cash=30e6,completion_bank=0,completion_receipts=0,completion_equity=0,completion_cost_validated=1)));self.assertIn('Pendente',c['status'])
 def test_19_completion_reserves_once(self):
  r=run();c=completion(r);self.assertAlmostEqual(c['total_uses'],c['cost_to_complete']+c['charges_to_open']+c['opening_operating_need']+c['opening_reserve'],5)
 def test_20_all_cash_checks(self):
  for extra in [{'pricing_mode':'Motor'},{'status_month':12,'etc_override':1,'etc_amount':50e6},{'construction_delay':6,'commercial_delay':6,'bank_available':.75}]:
   r=run(config(**extra))
   for x in r['flow']:self.assertAlmostEqual(x['cash_check'],0,5);self.assertAlmostEqual(x['debt_check'],0,5)
if __name__=='__main__':
 r=unittest.TextTestRunner(verbosity=2).run(unittest.defaultTestLoader.loadTestsFromTestCase(ExecutionTests));(ROOT/'entrega/Testes_Fases23.json').write_text(json.dumps({'tests':r.testsRun,'failures':len(r.failures),'errors':len(r.errors)},indent=2));raise SystemExit(not r.wasSuccessful())

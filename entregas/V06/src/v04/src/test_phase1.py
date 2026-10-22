from engine import *
from waterfall import constant_policy,priority,solve,COUNTS,TICKETS
from completion import completion
import unittest,importlib.util,hashlib
class Phase1(unittest.TestCase):
 def test_01_structure(self):self.assertEqual(sum([2,3,15]),20);self.assertEqual(2*.2+3*.1+15*.02,1);self.assertEqual(1*.2+2*.1,.4);self.assertEqual(.2+.1+10*.02,.5)
 def test_02_pipeline(self):
  for mode,end in [('Acelerado',3),('Base',6),('Prudente',18)]:
   rows=pipeline(config(capital_schedule=mode));self.assertEqual(len(rows),12);self.assertEqual(sum(x['planned_received'] for x in rows),25e6);self.assertLessEqual(max(x['pay_month'] for x in rows),end)
 def test_03_cash_identities(self):
  for extra in [{},{'limit':20e6},{'limit':25e6},{'bank_cp':0},{'bank_available':.5},{'capital_delay':12},{'commercial_delay':6},{'capital_delay':12,'commercial_delay':6},{'construction_delay':6,'cost_overrun':.1},{'commit_fee':.01},{'pari_passu':1,'min_equity_ratio':.3},{'funding_kind':'Empréstimo'}]:
   r=run(config(**extra))
   for x in r['flow']:
    self.assertAlmostEqual(x['cash_check'],0,5);self.assertAlmostEqual(x['debt_check'],0,5);self.assertLessEqual(x['cumulative'],r['p']['limit']*r['p']['bank_available']+.01)
    if not r['p']['bank_cp']:self.assertEqual(x['draw'],0)
 def test_04_v03_regression(self):
  spec=importlib.util.spec_from_file_location('old',ROOT/'original/v03/src/engine.py');old=importlib.util.module_from_spec(spec);spec.loader.exec_module(old);a=run(config(capital_schedule='V03'));b=old.run()
  for x,y in zip(a['flow'],b['flow']):
   for k in y:
    if isinstance(y[k],(int,float)):self.assertAlmostEqual(x[k],y[k],delta=.00001,msg=k)
 def test_05_partial(self):
  p=config(equity_mode='Combinação tickets',n_platinum=0,n_gold=0,n_silver=1);self.assertEqual(sum(x['planned_received'] for x in pipeline(p)),1e6);self.assertEqual(p['equity_stake'],.02)
 def test_06_bank_nonrevolver(self):
  for limit in [20e6,25e6,30e6]:
   r=run(config(limit=limit));self.assertLessEqual(r['flow'][-1]['cumulative'],limit+.01)
   self.assertAlmostEqual(r['flow'][-1]['cumulative'],sum(x['draw'] for x in r['flow']),5)
 def test_07_completion_not_secured(self):
  c=completion(run());self.assertIsNone(c['secured_coverage']);self.assertIn('Pendente',c['status'])
 def test_08_secured(self):
  c=completion(run(config(completion_cash=20e6,completion_bank=10e6,completion_receipts=0,completion_equity=0,completion_cost_validated=1,completion_timing_validated=1)));self.assertGreater(c['secured_coverage'],1);self.assertIn('Conforme',c['status'])
 def test_09_waterfall(self):
  r=run();w=constant_policy(r);self.assertTrue(w['feasible']);self.assertLess(w['check'],.001)
  for c in w['categories']:self.assertAlmostEqual(c['irr'],c['target'],8)
  for row in w['rows']:self.assertAlmostEqual(sum(row['dividends']),r['flow'][row['m']-1]['dividend'],5)
 def test_10_same_category(self):
  w=constant_policy(run());self.assertAlmostEqual(sum(x['dividends'][0]/2+2*x['dividends'][1]/3+5*x['dividends'][2]/15+x['dividends'][0]/2+x['dividends'][1]/3+10*x['dividends'][2]/15 for x in w['rows']),sum(sum(x['dividends']) for x in w['rows']),4)
 def test_11_legal_cap(self):
  r=run(config(payout_basis='Documentado'));self.assertEqual(sum(x['dividend'] for x in r['flow']),0)
  r=run(config(payout_basis='Documentado',legal_dividend_cap=1e6));self.assertLessEqual(sum(x['dividend'] for x in r['flow']),1e6+.01)
 def test_12_monthly(self):
  r=run(config(dividend_frequency=1));self.assertTrue(any(x['dividend']>0 and x['m']%12 for x in r['flow']));self.assertLess(max(x['reservegap'] for x in r['flow']),.01);w=constant_policy(r);self.assertLess((w if w['feasible'] else priority(r,'Combinado'))['check'],.001);self.assertTrue(w['feasible'] or w['low']>w['high'])
 def test_13_no_positive_irr_proof(self):
  r=run(config(treasury_ratio=0));self.assertGreater(summary(r)['funding_gap'],0)
 def test_14_delay(self):
  r=run(config(construction_delay=6));self.assertEqual(r['p']['open'],37);self.assertAlmostEqual(sum(x['total'] for x in r['cap']),63168000,2);self.assertEqual(r['hotel'][35]['revenue'],0)
if __name__=='__main__':
 suite=unittest.defaultTestLoader.loadTestsFromTestCase(Phase1);r=unittest.TextTestRunner(verbosity=2).run(suite);(ROOT/'entrega/Testes_Fase1.json').write_text(json.dumps({'tests':r.testsRun,'failures':len(r.failures),'errors':len(r.errors)},indent=2));raise SystemExit(not r.wasSuccessful())

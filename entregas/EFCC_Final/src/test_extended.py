from test_links import *
CASES=[('Discount_8',[('V06 Matriz ADR','F3','8%')]),('Discount_Custom10',[('V06 Matriz ADR','F3','Custom'),('V06 Controlos','B65',.10)]),('Construction_6',[('V06 Controlos','B38',6)]),('VAT_18',[('V06 Controlos','B8',18)])]
if __name__=='__main__':
 results=[]
 with concurrent.futures.ThreadPoolExecutor(max_workers=2) as pool:
  for result in pool.map(lambda case:run_case(*case),CASES):results.append(result);print(result,flush=True)
 (ROOT/'Testes_Seletores_Adicionais.json').write_text(json.dumps({'platform':'LibreOffice, actual mutations, all formula caches invalidated','excel_native':'NOT TESTED','discount_test':'All 30 matrix NPVs checked from verified unchanged cash flows. Independent scenario flows need no recalculation when only the discount rate changes.','tests':results},ensure_ascii=False,indent=2));assert all(r['status']=='PASS' for r in results)

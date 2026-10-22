from test_selectors import *
if __name__=='__main__':
 cases=[(name+'_Compat',p,e) for name,p,e in CASES if name in ['V04_Final','V04_Pari_Passu','V04_ETC_Comissao']]
 with concurrent.futures.ThreadPoolExecutor(max_workers=2) as pool:
  result=[]
  for value in pool.map(worker,cases):result.append(value);print(json.dumps(value,ensure_ascii=False),flush=True)
 (ROOT/'entrega/Testes_Correcao_Compat.json').write_text(json.dumps(result,ensure_ascii=False,indent=2))
 raise SystemExit(not all(x['ok'] for x in result))

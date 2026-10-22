"""Refresh current proof metadata; preserve historic separation evidence explicitly."""
import json,hashlib
from zipfile import ZipFile
from premium_ui import ROOT,Source,DASHES

def main():
 master=ROOT/'KK01_V06_Executive_Master.xlsx';s=Source(master);sha=hashlib.sha256(master.read_bytes()).hexdigest();sheets=list(s.w.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheets'));names=[n.get('name') for n in sheets];hidden=[n.get('name') for n in sheets if n.get('state','visible')!='visible'];visible=[n for n in names if n not in hidden]
 def load(n):return json.loads((ROOT/n).read_text())
 def save(n,v):(ROOT/n).write_text(json.dumps(v,indent=2,ensure_ascii=False))
 org=load('Organizacao_Folhas.json');org.pop('presentation_followup',None);org.update(current_total_worksheets=len(names),sha256=sha,verification='Integridade_Premium.json',current_master_worksheets=len(names),current_master_sha256=sha,hidden_current_technical_count=len(hidden),hidden_current_technical_sheets=hidden,visible_current_count=len(visible),visible_current_sheets=visible,first_ten=[x[0] for x in DASHES]);save('Organizacao_Folhas.json',org)
 preserve=load('Preservacao.json');preserve.pop('presentation_followup',None);preserve['historic_retained_parts_byte_identical_at_separation']=preserve.pop('retained_worksheet_parts_byte_for_byte_unchanged',179);preserve.update(result_sha256=sha,current_master_worksheets=len(names),premium_preservation_proof='Integridade_Premium.json',original_input_transfer_proof='Parametros_Interface.json',current_financial_recalculation='Validacao_Financeira.json');save('Preservacao.json',preserve)
 identity=load('Prova_Final_Identidade.json');identity.pop('presentation_followup',None);identity['historical_retained_worksheets_unchanged_at_separation']=identity.pop('retained_worksheets_unchanged_after_separation',179);identity.update(current_master_worksheets=len(names),sha256=sha,latest_revision='Integridade_Premium.json',original_logo_embedded=False,logo_area='V06 Identidade!B8:H20');save('Prova_Final_Identidade.json',identity)
 reduction=load('Reducao_Ficheiro.json');audit=load('Auditoria_OOXML.json');reduction.update(current_master_bytes=master.stat().st_size,current_worksheets=len(names),size_reduction_percent=round(100*(1-master.stat().st_size/reduction['prior_master_bytes']),2),current_formula_count=load('Validacao_Financeira.json')['formula_count']);save('Reducao_Ficheiro.json',reduction)
 for n in ['Separacao_V04.json','Preservacao_Inputs.json']:
  proof=load(n);proof.update(scope='Historical base proof; current premium follow-up is documented separately.',current_premium_integrity='Integridade_Premium.json',current_input_locations='Parametros_Interface.json',current_sha256=sha);save(n,proof)
 for lang in ['PT','EN']:
  n=f'PDF_Validacao_{lang}.json'
  if (ROOT/n).exists():
   proof=load(n)
   save(n,{'scope':'Superseded historical PDF test','historical':proof.get('historical',proof),'current_validation':'PDF_Premium.json','native_excel365':'NOT TESTED'})
 import xlsxwriter
 book=xlsxwriter.Workbook(ROOT/'Parametros_Ativos_EFCC.xlsx');ws=book.add_worksheet('Parametros UI');ws.freeze_panes(1,0);ws.set_column('A:B',22);ws.set_column('C:C',36);ws.set_column('D:E',22);ws.write_row(0,0,['Dashboard técnico','Célula editável','Parâmetro','Valor inicial','Origem anterior']);inputs=load('Parametros_Interface.json')
 for r,item in enumerate(inputs,1):ws.write_row(r,0,[item['sheet'],item['cell'],item['label_pt'],item['value'],item.get('source_sheet','')+'!'+item.get('source_cell','')])
 active=book.add_worksheet('Ativos e Custom');active.set_column('A:D',28);active.write_row(0,0,['Dashboard','Leitura ativa','Entrada guardada','Quando se aplica'])
 for r,record in enumerate([['DASH 05 V06','H24 ADR / H26 ocupação','H25 ADR / H27 ocupação','Cenário Custom'],['DASH 05 V06','H28 base / N28 sistemas / N30 incentivo','H29 / N32 / H30','Custom Terms'],['DASH 04B V06','H24 gestão / N24 distribuição','H25 / N25','Custom Terms'],['DASH 07 V06','N26 desconto efetivo','N25 desconto Custom','Modo Custom']],1):active.write_row(r,0,record)
 book.close()
 readme=ROOT.parents[1]/'README.md'
 readme.write_text(f'''# KK-01 — Modelo Financeiro

A revisão atual está em [EFCC Final](entregas/EFCC_Final/README.md).

- [Modelo V06](entregas/EFCC_Final/KK01_V06_Executive_Master.xlsx) — {master.stat().st_size/1e6:.1f} MB, {len(names)} folhas, dez dashboards com navegação e filtros no painel.
- [Entrega completa](entregas/EFCC_Final/KK01_EFCC_Entrega_Candidata.zip).
- [Testes por filtro e escolha](entregas/EFCC_Final/Matriz_Testes_Filtros.csv).
- [Matriz de navegação](entregas/EFCC_Final/Matriz_Navegacao.csv).
- [Relatório integral PT](entregas/EFCC_Final/PDF/PT/DASH_10_Integral_PT.pdf).

Valores ativos e entradas Custom guardadas são identificados separadamente no Hotel, no residencial e no desconto. Os filtros continuam ligados ao motor central. Na revisão V06.3-SYNC, Base Case utiliza o perfil Base (185 EUR / 58%) e Conservative utiliza Conservador. As escolhas propagam-se à composição F&B/MICE e ao motor de 180 meses. [Comparação antes/depois](entregas/EFCC_Final/Relatorio_Sincronizacao_Cenarios.md).

As 313 folhas V04 foram fisicamente retiradas do master; a [V04 original](entregas/Arquivo/V04/README.md) permanece preservada em arquivo separado. O logótipo original pode ser inserido em **V06 Identidade**.

**Microsoft Excel 365: NOT TESTED nesta revisão.** O recálculo e as mutações são verificados com Python/LibreOffice. A observação do utilizador de que o erro anterior parecia resolvido não substitui a abertura, edição, recálculo e guardar/reabrir desta revisão no Excel. [Procedimento Windows](entregas/EFCC_Final/Excel365/LOGOTIPO.md).
''')
 print('Current proof metadata',len(names),len(visible),len(hidden),sha)
if __name__=='__main__':main()

"""Consolidated, reversible V08 build from the preserved V07 workbook.
Usage: python build_v08.py --source old.xlsx --target new.xlsx --workdir /tmp/kk08
Independent recalculation is a separate required release check; this command
never represents structural validation as native Excel Desktop acceptance.
"""
from pathlib import Path
import argparse,json
import robust_filters_v08 as robust
import navigation_v08 as navigation
import impact_layout_v08 as impact
import kpi_layout_v08 as kpi
import finalize_v08 as finalize
import charts_v08 as charts
import investor_v08 as investor
import direct_inputs_v08 as direct
import polish_v08 as polish
import release_v08 as release
import profile_controls_v08 as profiles
import wiring_v08 as wiring
import presentation_review_v08 as review
import hotel_helpers_v08 as helpers
import view_consistency_v08 as views
import budget_comparison_v08 as budget
import heatmap_render_v08 as heatmap

def build(source,target,work):
 source=Path(source);target=Path(target);work=Path(work);work.mkdir(parents=True,exist_ok=True)
 if source.resolve()==target.resolve():raise ValueError('Preserve source: source and target must differ')
 steps=[('01_inputs',robust.harden),('02_navigation',navigation.apply),('03_impact',impact.apply_impact),('04_kpis',kpi.apply_cards),('05_headers',finalize.finish_ui),('06_charts',charts.chart_refine),('07_investor',investor.investor_fix)]
 current=source
 for name,fn in steps:
  nxt=work/(name+'.xlsx');result=fn(current,nxt)
  if result is not None:(work/(name+'.json')).write_text(json.dumps(result,ensure_ascii=False,indent=2))
  current=nxt
 nxt=work/'08_direct.xlsx';mapping=work/'Mapa_Cenarios.json';proof=direct.direct_inputs(current,nxt,mapping);(work/'08_direct.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2));current=nxt
 nxt=work/'09_polished.xlsx';polish.polish(current,nxt);current=nxt
 nxt=work/'10_release.xlsx';release.release_metadata(current,nxt);current=nxt
 nxt=work/'11_profiles.xlsx';profiles.profiles(current,nxt,mapping,work/'11_profiles.json');current=nxt
 nxt=work/'12_wired.xlsx';result=wiring.wire(current,nxt,source);(work/'12_wiring.json').write_text(json.dumps(result,ensure_ascii=False,indent=2));current=nxt
 nxt=work/'13_reviewed.xlsx';result=review.review(current,nxt,source,work/'06_charts.json');(work/'13_review.json').write_text(json.dumps(result,ensure_ascii=False,indent=2));current=nxt
 nxt=work/'14_helpers.xlsx';result=helpers.protect_helpers(current,nxt,source);(work/'14_helpers.json').write_text(json.dumps(result,ensure_ascii=False,indent=2));current=nxt
 nxt=work/'15_consistency.xlsx';views.consistent_views(current,nxt,work/'06_charts.json');current=nxt
 nxt=work/'16_budget.xlsx';budget.budget_compare(current,nxt);current=nxt
 heatmap.legible(current,target)
 return target
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('--source',required=True,type=Path);p.add_argument('--target',required=True,type=Path);p.add_argument('--workdir',required=True,type=Path);a=p.parse_args();print(build(a.source,a.target,a.workdir))

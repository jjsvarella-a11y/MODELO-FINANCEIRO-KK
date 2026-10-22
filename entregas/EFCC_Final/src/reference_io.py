from pathlib import Path
import gzip,json
def load_reference(root):
 p=Path(root)/'Referencia_Python_V06.json'
 if p.exists():return json.loads(p.read_text())
 with gzip.open(p.with_suffix('.json.gz'),'rt',encoding='utf-8') as f:return json.load(f)

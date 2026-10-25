"""Reproduce an issued report from its frozen JSON, without reading today's XLSX."""
from report_v07 import *
def reproduce(snapshot_file,output=None):
 snapshot_file=Path(snapshot_file);snapshot=json.loads(snapshot_file.read_text(encoding='utf-8-sig'));output=Path(output or snapshot_file.with_name(snapshot_file.stem+'_REPRODUCED.pdf'))
 if output.exists():raise FileExistsError('Refusing to replace an existing report: '+str(output))
 if 'report_tables' not in snapshot:raise ValueError('Snapshot does not contain its frozen report tables')
 data={'contexts':{'Base':snapshot['base']},'calculated_at':snapshot['calculated_at'],'report_tables':snapshot['report_tables']};pdf_report(output,data,'Base',snapshot['comparisons'],snapshot['format'],snapshot['language'],snapshot['settings'],snapshot['id']);return output
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('snapshot',type=Path);p.add_argument('--output',type=Path);a=p.parse_args();print(reproduce(a.snapshot,a.output))

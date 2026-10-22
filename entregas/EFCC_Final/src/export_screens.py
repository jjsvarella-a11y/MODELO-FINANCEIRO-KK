from export_all import *
def screens(source,language):
 pdf=export(source,DASHES,'Dashboards_01_09',WORK/('screens_final_'+language));doc=fitz.open(pdf);assert len(doc)==9
 dest=ROOT/'PDF'/language
 for i in range(9):
  out=fitz.open();out.insert_pdf(doc,from_page=i,to_page=i);target=dest/f'DASH_{i+1:02}_{language}.pdf'
  if target.exists():target.unlink()
  out.save(target);out.close()
 print('Nine final screens exported',language)
if __name__=='__main__':screens(Path(sys.argv[1]),sys.argv[2])

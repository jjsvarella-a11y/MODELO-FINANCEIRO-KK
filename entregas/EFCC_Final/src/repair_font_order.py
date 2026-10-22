"""Repair only bold-font element order. Does not recalculate or rewrite worksheets.
Usage: python repair_font_order.py INPUT.xlsx OUTPUT.xlsx
Validate output with src/OpenXmlValidation (Open XML SDK 3.3.0).
"""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import sys,copy,json,hashlib
T=lambda n:'{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'+n

def repair(source,target):
 source,target=Path(source),Path(target)
 if target.exists():raise FileExistsError(target)
 with ZipFile(source) as old,ZipFile(target,'w',ZIP_DEFLATED,compresslevel=6) as new:
  styles=E.fromstring(old.read('xl/styles.xml'));changed=[]
  for i,font in enumerate(styles.find(T('fonts'))):
   bold=font.find(T('b'))
   if bold is not None and font.index(bold)>0:font.remove(bold);font.insert(0,bold);changed.append(i)
  for item in old.infolist():new.writestr(copy.copy(item),E.tostring(styles,encoding='UTF-8',xml_declaration=True,standalone=True) if item.filename=='xl/styles.xml' else old.read(item.filename))
 with ZipFile(source) as a,ZipFile(target) as b:
  assert a.namelist()==b.namelist()
  parts=[n for n in a.namelist() if a.read(n)!=b.read(n)];assert parts in [[],['xl/styles.xml']]
 return {'changed_parts':parts,'reordered_font_indices':changed,'sha256':hashlib.sha256(target.read_bytes()).hexdigest()}
if __name__=='__main__':print(json.dumps(repair(*sys.argv[1:3]),indent=2))

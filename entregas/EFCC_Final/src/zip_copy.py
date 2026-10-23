"""Rewrite selected ZIP members, copy unchanged compressed streams verbatim.
Retains each ZipInfo and safely rebuilds central directory offsets with zipfile.
No changes to OOXML content or compression of unaffected parts.
"""
from zipfile import ZipFile,ZIP_DEFLATED
import copy,struct

def rewrite(source,target,changes):
 with ZipFile(source) as src,ZipFile(target,'w',ZIP_DEFLATED,compresslevel=6) as dst:
  for entry in src.infolist():
   if entry.filename in changes:dst.writestr(copy.copy(entry),changes[entry.filename]);continue
   if entry.flag_bits & 8:
    dst.writestr(copy.copy(entry),src.read(entry.filename));continue
   src.fp.seek(entry.header_offset);header=src.fp.read(30)
   if header[:4]!=b'PK\x03\x04':raise ValueError('Invalid local ZIP header')
   fn,ex=struct.unpack_from('<HH',header,26);payload=src.fp.read(fn+ex+entry.compress_size);new=copy.copy(entry);new.header_offset=dst.fp.tell();dst.fp.write(header+payload);dst.filelist.append(new);dst.NameToInfo[new.filename]=new;dst.start_dir=dst.fp.tell();dst._didModify=True
  for name,data in changes.items():
   if name not in src.namelist():dst.writestr(name,data)
 with ZipFile(target) as result:
  if result.testzip() is not None:raise ValueError('ZIP integrity failure')

"""Produce a rendering-only DOCX while retaining the delivery file's live fields.

The bundled LibreOffice PDF exporter exhausted memory on the input university
template's live fields. This copy removes field instructions and controls only;
displayed results, hyperlinks, bookmarks, styles, and geometry are retained.
"""
import argparse,copy
from pathlib import Path
from zipfile import ZipFile
from lxml import etree as E

p=argparse.ArgumentParser();p.add_argument('source');p.add_argument('output');a=p.parse_args()
ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
with ZipFile(a.source) as z:
    root=E.fromstring(z.read('word/document.xml'))
    for x in root.xpath('//w:instrText|//w:fldChar',namespaces=ns):x.getparent().remove(x)
    with ZipFile(a.output,'w') as out:
        for info in z.infolist():out.writestr(copy.copy(info),E.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True) if info.filename=='word/document.xml' else z.read(info.filename))

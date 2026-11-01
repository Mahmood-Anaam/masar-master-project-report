"""Validate the delivered chapter's field structure, navigation, and page budget."""
import argparse,collections,json,re
from pathlib import Path
from zipfile import ZipFile
from lxml import etree as E
import fitz

ap=argparse.ArgumentParser();ap.add_argument('--pdf',required=True);ap.add_argument('--report',required=True);args=ap.parse_args()
root=Path(__file__).resolve().parents[1];ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'};W='{'+ns['w']+'}'
with ZipFile(root/'masar-master-project-report.docx') as z:
    assert z.testzip() is None
    tree=E.fromstring(z.read('word/document.xml'))
    for f in z.namelist():
        if f.endswith('.xml') or f.endswith('.rels'):E.fromstring(z.read(f))
codes=tree.xpath('//w:instrText/text()',namespaces=ns)
citations=[json.loads(t[t.index('{'):]) for t in codes if 'CSL_CITATION' in t]
assert len(citations)==len(set(c['citationID'] for c in citations))==53
for c in citations:
    assert c['citationItems']
    for item in c['citationItems']:assert item.get('uris') and item.get('itemData',{}).get('title')
stack=[]
for element in tree.iter(W+'fldChar'):
    typ=element.get(W+'fldCharType')
    if typ=='begin':stack.append(False)
    elif typ=='separate':assert stack and not stack[-1];stack[-1]=True
    elif typ=='end':assert stack and stack.pop()
assert not stack
names=tree.xpath('//w:bookmarkStart/@w:name',namespaces=ns);ids=tree.xpath('//w:bookmarkStart/@w:id',namespaces=ns)
assert len(ids)==len(set(ids))
missing=set(tree.xpath('//w:hyperlink/@w:anchor',namespaces=ns))-set(names);assert not missing,missing
assert len([n for n in names if n.startswith('CITE_MASAR_')])==32
assert sum('ZOTERO_BIBL' in s for s in codes)==1
visible=' '.join(tree.xpath('//w:t/text()',namespaces=ns));assert not re.search(r'\[@[^\]]+\]',visible)
register=json.loads((root/'chapter-2-source-register.json').read_text());assert sorted(x['ieee_number'] for x in register.values())==list(range(1,33))
keys=re.findall(r'@\w+\{([^,]+),',(root/'references.bib').read_text());assert len(keys)==len(set(keys))==32
pdf=fitz.open(args.pdf);pages=[p.get_text() for p in pdf]
ch2=next(i+1 for i,s in enumerate(pages) if i>7 and 'Chapter 2:' in s)
ch3=next(i+1 for i,s in enumerate(pages) if i>7 and 'Chapter 3:' in s)
assert 20<=ch3-ch2<=30
assert (len(pdf),ch2,ch3)==(50,15,42)
layout=json.loads((root/'chapter-2-layout-map.json').read_text())
assert next(x['page'] for x in layout['toc'] if x['title']=='References')==46
assert len(layout['captions'])==6
# Locate out-of-page text independently of document extraction.
outside=[]
for i,p in enumerate(pdf):
    for block in p.get_text('dict')['blocks']:
        if 'lines' not in block:continue
        for line in block['lines']:
            for span in line['spans']:
                x0,y0,x1,y1=span['bbox']
                if x0 < -1 or y0 < -1 or x1>p.rect.width+1 or y1>p.rect.height+1:outside.append(i+1)
assert not outside,outside
report={'date':'2026-11-01','zip_and_xml_integrity':'pass','balanced_field_boundaries':'pass','unique_citation_ids':len(citations),'bibliography_entries':32,'unresolved_internal_hyperlinks':0,'unresolved_markdown_citation_tokens':0,'report_pages':len(pdf),'chapter_start':ch2,'chapter_end':ch3-1,'chapter_pages':ch3-ch2,'table_count':5,'figure_count':1,'pdf_hyperlinks':sum(len(p.get_links()) for p in pdf),'out_of_page_text':0,'native_word_zotero_refresh':'not tested','rendering_basis':'Visible field results; field-preserving final DOCX validated separately'}
Path(args.report).write_text(json.dumps(report,indent=2)+'\n');print(json.dumps(report))

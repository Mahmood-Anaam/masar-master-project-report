"""Fill navigation page caches from a verified layout PDF without unlinking fields."""
import argparse,copy,json,re
from pathlib import Path
from zipfile import ZipFile
from lxml import etree as E
import fitz

ap=argparse.ArgumentParser();ap.add_argument('docx');ap.add_argument('pdf');args=ap.parse_args()
root=Path(__file__).resolve().parents[1]
layout=json.loads((root/'chapter-2-layout-map.json').read_text())
pdf=fitz.open(args.pdf)
def norm(s):return re.sub(r'[^\w]','',s).casefold()
texts=[p.get_text() for p in pdf]
def locate(title,label='',kind=''):
    # Arabic heading extraction order varies with PDF shaping. The unchanged
    # Arabic abstract leaf is verified visually at page 3 of this template.
    if title=='الخلاصة':return 3
    if title=='Table of Contents':return 4
    target=norm((label+' ' if label else '')+title)
    for i,t in enumerate(texts):
        if i==3:continue # The displayed contents list repeats every target.
        if label and i<8:continue
        if kind and i<14:continue
        scope=t if label or kind else '\n'.join(t.splitlines()[:8])
        found=target in norm(scope) if label or kind else any(norm(line)==target for line in scope.splitlines())
        if found:return i+1
    raise ValueError('No layout match for '+label+' '+title)
ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
with ZipFile(args.docx) as z:
    tree=E.fromstring(z.read('word/document.xml'))
    for ent in layout['toc']:
        page=locate(ent['title'],ent['label']);ent['page']=page
        p=tree.xpath('//w:p[w:bookmarkStart[@w:name=$name]]',namespaces=ns,name=ent['cache'])[0]
        p.xpath('.//w:t',namespaces=ns)[-1].text=str(page)
    for kind in ['Figure','Table']:
        captions=[c for c in layout['captions'] if c['kind']==kind]
        p=next(p for p in tree.xpath('//w:p',namespaces=ns) if any('\\c "'+kind+'"' in s for s in p.xpath('.//w:instrText/text()',namespaces=ns)))
        zero=[t for t in p.xpath('./w:r/w:t',namespaces=ns) if (t.text or '').isdigit()]
        assert len(zero)==len(captions)
        for c,t in zip(captions,zero):
            c['page']=locate(c['title'],c['label'],kind);t.text=str(c['page'])
    replacement=E.tostring(tree,xml_declaration=True,encoding='UTF-8',standalone=True)
    archive=[(copy.copy(f),replacement if f.filename=='word/document.xml' else z.read(f.filename)) for f in z.infolist()]
with ZipFile(args.docx,'w') as z:
    for f,data in archive:z.writestr(f,data)
(root/'chapter-2-layout-map.json').write_text(json.dumps(layout,indent=2)+'\n')
print(json.dumps({'pages':len(pdf),'chapter_2_start':next(e['page'] for e in layout['toc'] if e['title']=='Literature Review'),'chapter_3_start':next(e['page'] for e in layout['toc'] if e['title'].startswith('Methodology/')),'captions':layout['captions']}))

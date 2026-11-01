"""Rebuild Chapter 2 from its reviewed Markdown and source register.

Use the unmodified develop DOCX as --source. Existing Zotero items retain their
identities; additional sources use embedded CSL data and canonical source URIs.
This script does not claim to run Microsoft Word, Zotero, or its VBA add-in.
"""
import argparse
import copy
import hashlib
import json
import re
from pathlib import Path
from zipfile import ZipFile

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from lxml import etree as E

W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS={'w':W}
ROOT=Path(__file__).resolve().parents[2]

def node(tag,**attrs):
    e=OxmlElement('w:'+tag)
    for k,v in attrs.items():e.set(qn('w:'+k),str(v))
    return e

def field_run(kind):
    r=node('r');r.append(node('fldChar',fldCharType=kind));return r

def instruction(text):
    r=node('r');e=node('instrText');e.set(qn('xml:space'),'preserve');e.text=text;r.append(e);return r

def bookmark(p,name,bid):
    start=node('bookmarkStart',id=bid,name=name);stop=node('bookmarkEnd',id=bid)
    p.insert(1 if p.find(qn('w:pPr')) is not None else 0,start);p.append(stop)

def plain(element):return ''.join(element.xpath('.//w:t/text()'))

def initials(given):
    return ' '.join((x[0]+'.') for x in re.findall(r'[\wÀ-ž]+',given))

def author_text(c):
    authors=[a.get('literal') or (initials(a.get('given',''))+' '+a['family']).strip() for a in c.get('author',[])]
    if len(authors)>6:return authors[0]+' et al.'
    if len(authors)>1:return ', '.join(authors[:-1])+', and '+authors[-1] if len(authors)>2 else ' and '.join(authors)
    return authors[0] if authors else ''

def ieee(c):
    year=str(c.get('issued',{}).get('date-parts',[['']])[0][0])
    text=author_text(c)+', “'+c['title']+'”'
    if c.get('container-title'):
        text+=', '+('in ' if c['type']=='paper-conference' else '')+c['container-title'].replace('Computers \\&','Computers &')
    if c.get('publisher'):text+=', '+c['publisher']
    if c.get('number'):text+=', '+c['number']
    if c.get('volume'):text+=', vol. '+str(c['volume'])
    if c.get('issue'):text+=', no. '+str(c['issue'])
    pp=c.get('page','')
    if pp:text+=(', pp. '+pp if '–' in pp or '-' in pp else ', Art. no. '+pp)
    if year:text+=', '+year
    if c.get('DOI'):text+=', doi: '+c['DOI']+'.'
    elif c.get('URL'):text+='. Accessed: Nov. 1, 2026. [Online]. Available: '+c['URL']
    else:text+='.'
    return re.sub('”([,.])',r'\1”',text)

def bibtex(key,c):
    typ={'article-journal':'article','paper-conference':'inproceedings','report':'techreport'}.get(c['type'],'misc')
    vals={'author':' and '.join('{'+a['literal']+'}' if 'literal' in a else a['family']+', '+a.get('given','') for a in c.get('author',[])), 'title':c['title']}
    for ck,bk in [('container-title','booktitle' if typ=='inproceedings' else 'journal'),('volume','volume'),('issue','number'),('page','pages'),('DOI','doi'),('URL','url')]:
        if c.get(ck):vals[bk]=str(c[ck]).replace('–','--').replace('&','\\&')
    if c.get('issued'):vals['year']=str(c['issued']['date-parts'][0][0])
    if c.get('number'):vals['number']=c['number']
    return '@'+typ+'{'+key+',\n'+''.join('  '+k+' = {'+v+'},\n' for k,v in vals.items())+'}\n'

def main():
    ap=argparse.ArgumentParser();ap.add_argument('--source',required=True);ap.add_argument('--output',default=str(ROOT/'docs/masar-master-project-report.docx'));args=ap.parse_args()
    doc=Document(args.source);body=doc._element.body
    records=json.loads((ROOT/'docs/chapter-2-source-register.json').read_text())
    # Read original Zotero identities from the user's source document at build
    # time rather than publishing account-specific metadata in the register.
    original_items=[]
    for inst in doc._element.xpath('//w:instrText'):
        if 'CSL_CITATION' in (inst.text or ''):
            original_items.extend(json.loads(inst.text[inst.text.index('{'):])['citationItems'])
    for record in records.values():
        csl=record['csl']
        for item in original_items:
            old=item['itemData']
            same_doi=csl.get('DOI') and csl.get('DOI','').lower()==old.get('DOI','').lower()
            if same_doi or csl['title']==old.get('title'):
                record['citation_item']=copy.deepcopy(item);break
    text=(ROOT/'docs/chapter-2-literature-review.md').read_text()
    order=['nrcnationalframework','dga2025bcm','hevner2004design','peffers2007dsrm','torabi2014framework','sahebjamnia2015integrated']
    for match in re.finditer(r'\[@([^\]]+)\]',text):
        for key in match.group(1).split(';'):
            if key not in order:order.append(key)
    assert set(order)==set(records),(set(records)-set(order))
    numbers={k:i+1 for i,k in enumerate(order)}
    for key,n in numbers.items():records[key]['ieee_number']=n
    bid=max(int(x.get(qn('w:id'))) for x in doc._element.xpath('//w:bookmarkStart'))+1
    counter=0

    def rich(p,s):
        nonlocal counter
        pos=0
        for m in re.finditer(r'\[@([^\]]+)\]',s):
            p.add_run(s[pos:m.start()]);keys=m.group(1).split(';');counter+=1
            formatted=', '.join('['+str(numbers[k])+']' for k in keys)
            items=[]
            for k in keys:
                if 'citation_item' in records[k]:items.append(copy.deepcopy(records[k]['citation_item']))
                else:
                    c=copy.deepcopy(records[k]['csl']);uri=c.get('URL') or 'https://doi.org/'+c['DOI']
                    items.append({'id':k,'uris':[uri],'itemData':c})
            data={'citationID':hashlib.sha256((str(counter)+s).encode()).hexdigest()[:10],'properties':{'formattedCitation':formatted,'plainCitation':formatted,'noteIndex':0},'citationItems':items,'schema':'https://github.com/citation-style-language/schema/raw/master/csl-citation.json'}
            p._p.append(field_run('begin'));p._p.append(instruction(' ADDIN ZOTERO_ITEM CSL_CITATION '+json.dumps(data,ensure_ascii=False,separators=(',',':'))+' '));p._p.append(field_run('separate'))
            for i,k in enumerate(keys):
                if i:p.add_run(', ')
                h=node('hyperlink',anchor='CITE_MASAR_'+str(numbers[k]),history='1');r=node('r');rp=node('rPr');rp.append(node('color',val='0000FF'));rp.append(node('u',val='none'));r.append(rp);t=node('t');t.text='['+str(numbers[k])+']';r.append(t);h.append(r);p._p.append(h)
            p._p.append(field_run('end'));pos=m.end()
        p.add_run(s[pos:])

    # Preserve the original introductory content and the university style ladder.
    ch2=next(p for p in doc.paragraphs if p.text=='Literature Review')
    ch3=next(p for p in doc.paragraphs if p.text=='Methodology/ Research Design')
    a=list(body).index(ch2._p);b=list(body).index(ch3._p)
    for el in list(body)[a+1:b]:body.remove(el)
    # Blank template padding is replaced by a real page boundary.
    prev=ch2._p.getprevious()
    while prev is not None and prev.tag==qn('w:p') and not plain(prev) and not prev.xpath('.//w:sectPr'):
        prior=prev.getprevious();body.remove(prev);prev=prior
    ch2.paragraph_format.page_break_before=True;ch3.paragraph_format.page_break_before=True
    # Uniformly use the existing university font rather than introducing a new design.
    for name in ['Normal _ CoE_KSU','Heading 2','Heading 3']:
        doc.styles[name].font.color.rgb=RGBColor(0,0,0)
    sections=[];captions=[];tables=[];sec=0;sub=0;table_no=0
    lines=text.splitlines();i=0

    def insert_p(style='NormalCoEKSU'):
        after_table=ch3._p.getprevious().tag==qn('w:tbl')
        p=doc.add_paragraph(style=style);ch3._p.addprevious(p._p)
        if after_table:p.paragraph_format.space_before=Pt(8)
        return p

    def caption(p,kind,number,title):
        p.add_run(kind+' ')
        for code,value in [(' STYLEREF 1 \\s ','2'),(' SEQ '+kind+' \\* ARABIC \\s 1 ',str(number))]:
            p._p.append(field_run('begin'));p._p.append(instruction(code));p._p.append(field_run('separate'));p.add_run(value);p._p.append(field_run('end'))
            if value=='2' and 'STYLEREF' in code:p.add_run('.')
        p.add_run(' '+title)

    while i<len(lines):
        line=lines[i].strip();i+=1
        if not line or line.startswith('# '):continue
        if line.startswith('### '):
            sub+=1;p=insert_p('Heading 3');p.add_run(line[4:]);bookmark(p._p,'MASAR_CH2_'+str(sec)+'_'+str(sub),bid);bid+=1
        elif line.startswith('## '):
            sec+=1;sub=0;p=insert_p('Heading 2');p.add_run(line[3:]);name='MASAR_CH2_'+str(sec);bookmark(p._p,name,bid);bid+=1;sections.append({'title':line[3:],'label':'2.'+str(sec),'anchor':name})
        elif line.startswith('|'):
            rows=[line]
            while i<len(lines) and lines[i].strip().startswith('|'):rows.append(lines[i].strip());i+=1
            rows=[r for r in rows if not re.match(r'^\|[\s:|\-]+\|$',r)]
            cells=[[x.strip() for x in r.strip('|').split('|')] for r in rows]
            table_no+=1
            titles=['Recovery parameters and their interpretation','Comparison of methodological commitments','Research prototypes and their decision boundaries','Operational software and its relationship to Masar','Literature based evaluation questions']
            cp=insert_p('Caption');cp.paragraph_format.keep_with_next=True;caption(cp,'Table',table_no,titles[table_no-1]);anchor='MASAR_TABLE_'+str(table_no);bookmark(cp._p,anchor,bid);bid+=1
            captions.append({'label':'Table 2.'+str(table_no),'title':titles[table_no-1],'anchor':anchor,'kind':'Table'})
            table=doc.add_table(rows=0,cols=len(cells[0]));ch3._p.addprevious(table._tbl);table.autofit=False
            widths=([1.03,2.12,2.75] if len(cells[0])==3 else [1.03,1.62,1.57,1.68])
            for col,w in zip(table.columns,widths):col.width=Inches(w)
            pr=table._tbl.tblPr; borders=node('tblBorders')
            for side in ['top','left','bottom','right','insideH','insideV']:borders.append(node(side,val='single',sz='4',color='D9D9D9'))
            pr.append(borders);margins=node('tblCellMar')
            for side in ['top','left','bottom','right']:margins.append(node(side,w='95',type='dxa'))
            pr.append(margins)
            for ri,row in enumerate(cells):
                rr=table.add_row();rr._tr.get_or_add_trPr().append(node('cantSplit'))
                if ri==0:rr._tr.get_or_add_trPr().append(node('tblHeader'))
                for ci,(cell,s) in enumerate(zip(rr.cells,row)):
                    cell.width=Inches(widths[ci]);p=cell.paragraphs[0];p.style=doc.styles['Normal _ CoE_KSU'];p.paragraph_format.first_line_indent=Inches(0);p.paragraph_format.space_after=Pt(2);p.paragraph_format.line_spacing=1.05;p.alignment=WD_ALIGN_PARAGRAPH.LEFT
                    rich(p,s)
                    for run in p.runs:run.font.size=Pt(10);run.font.name='Calibri'
                    tcpr=cell._tc.get_or_add_tcPr();tcpr.append(node('vAlign',val='center'));tcpr.append(node('shd',fill='DDEBF7' if ri==0 else ('F5F7FA' if ri%2==0 else 'FFFFFF')))
                    if ri==0:
                        for run in p.runs:run.bold=True
            tables.append(table)
        elif line.startswith('!['):
            p=insert_p();p.paragraph_format.first_line_indent=Inches(0);p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.paragraph_format.keep_with_next=True
            p.paragraph_format.line_spacing=1.0
            p.add_run().add_picture(str(ROOT/'docs/figures/literature-concept-map.png'),width=Inches(5.8))
            cp=insert_p('Caption');caption(cp,'Figure',1,'Conceptual relationships linking temporal assessment service priorities and supporting dependencies');cp.paragraph_format.space_after=Pt(10);bookmark(cp._p,'MASAR_FIGURE_1',bid);bid+=1
            captions.append({'label':'Figure 2.1','title':'Conceptual relationships linking temporal assessment service priorities and supporting dependencies','anchor':'MASAR_FIGURE_1','kind':'Figure'})
        else:
            p=insert_p();rich(p,line)

    # Append new entries to the cached result of the existing Zotero bibliography.
    bibs=[p for p in doc.paragraphs if p.style.style_id=='Bibliography']
    last=bibs[-1]._p
    for j,p in enumerate(bibs,1):bookmark(p._p,'CITE_MASAR_'+str(j),bid);bid+=1
    for k in order[6:]:
        p=doc.add_paragraph(style='Bibliography');last.addnext(p._p);last=p._p;p.paragraph_format.first_line_indent=Inches(-.35);p.paragraph_format.left_indent=Inches(.35);p.paragraph_format.space_after=Pt(7);p.paragraph_format.line_spacing=1.05
        p.add_run('['+str(numbers[k])+']\t'+ieee(records[k]['csl']));bookmark(p._p,'CITE_MASAR_'+str(numbers[k]),bid);bid+=1
        for run in p.runs:run.font.size=Pt(11)
    # Replace template guidance in abbreviation list with the terms actually used.
    ab=next(p for p in doc.paragraphs if p.text.startswith('If appropriate, list any abbreviations'))
    ab.text=''
    terms={'BCM':'Business Continuity Management','BCMS':'Business Continuity Management System','BIA':'Business Impact Analysis','BWM':'Best-Worst Method','DGA':'Digital Government Authority','DRAMM':'Disaster Recovery Analysis and Management Method','DSS':'Decision Support System','FURIA':'Fuzzy Relational Incident Analysis','IT':'Information Technology','MADM':'Multi-Attribute Decision Making','MBCO':'Minimum Business Continuity Objective','MCDA':'Multicriteria Decision Analysis','MILP':'Mixed-Integer Linear Programming','MTD':'Maximum Tolerable Downtime','MTPD':'Maximum Tolerable Period of Disruption','NIST':'National Institute of Standards and Technology','PARAM':'Physical Asset Risk Assessment Matrix','ReaR':'Relax-and-Recover','ROPE':'Risk-Oriented Process Evaluation','RPO':'Recovery Point Objective','RTO':'Recovery Time Objective','TDRP':'Time-Dependent Recovery Prioritization','TOPSIS':'Technique for Order Preference by Similarity to Ideal Solution'}
    ab.style=doc.styles['Normal _ CoE_KSU'];ab.paragraph_format.first_line_indent=Inches(0);ab.paragraph_format.line_spacing=1.05;ab.alignment=WD_ALIGN_PARAGRAPH.LEFT
    for j,(k,v) in enumerate(terms.items()):
        if j:ab.add_run('\n')
        ab.add_run(k+'  ').bold=True;ab.add_run(v)
    # Correct the existing DGA sentence's reference target, without rewriting it.
    for p in doc.paragraphs:
        if p.text.startswith('Assessing the consequences of an interruption'):
            instrs=p._p.xpath('.//w:instrText')
            for instr in instrs:
                if 'CSL_CITATION' in (instr.text or ''):
                    data=json.loads(instr.text[instr.text.index('{'):]);data['citationItems']=[copy.deepcopy(records['dga2025bcm']['citation_item'])];data['properties']['formattedCitation']='[2]';data['properties']['plainCitation']='[2]';instr.text=' ADDIN ZOTERO_ITEM CSL_CITATION '+json.dumps(data,ensure_ascii=False)+' '
            for t in p._p.xpath('.//w:t'):
                if t.text=='[1]':t.text='[2]'
            for h in p._p.xpath('.//w:hyperlink'):
                if '[2]' in plain(h):h.set(qn('w:anchor'),'CITE_MASAR_2')

    # Retain live TOC fields but replace their displayed caches with complete entries.
    # Page values are materialized after layout using the companion QA workflow.
    tocpars=[p for p in doc.paragraphs if p.style.style_id in ['TOC1','TOC2']]
    entries=[]
    for p in doc.paragraphs:
        if p.style.style_id in ['PreliminaryHeadingCoEKSU','Heading1','Heading2','AppendicesBibliographyCoEKSU']:
            if not p.text.strip():continue
            if p.text=='Literature Review':continue
            # Only original headings; Chapter 2 entries are inserted below.
            if p._p.xpath('./w:bookmarkStart[starts-with(@w:name,"MASAR_CH2_")]'):continue
            names=p._p.xpath('./w:bookmarkStart/@w:name');anchor=names[0] if names else 'MASAR_NAV_'+str(len(entries))
            if not names:bookmark(p._p,anchor,bid);bid+=1
            entries.append({'title':p.text,'anchor':anchor,'style':p.style.style_id})
    # Derive labels for the original chapter/section hierarchy.
    result=[];chap=0;subno=0
    for ent in entries:
        if ent['style']=='Heading1':chap+=1;subno=0
        if ent['style']=='Heading2':subno+=1;ent['label']=str(chap)+'.'+str(subno)
        elif ent['style']=='Heading1':ent['label']='Chapter '+str(chap)+':'
        else:ent['label']=''
        # Chapter 2 was omitted from entries, so explicitly insert and advance here.
        if ent['title']=='Methodology/ Research Design':
            chap=3;ent['label']='Chapter 3:'
        if ent['title']=='Results and Discussions':chap=4;ent['label']='Chapter 4:'
        if ent['title']=='Analysis/ Discussions':chap=5;ent['label']='Chapter 5:'
        if ent['title']=='Conclusions':chap=6;ent['label']='Chapter 6:'
        result.append(ent)
        if ent['title']=='Objectives and Proposed Solution':
            bn=ch2._p.xpath('./w:bookmarkStart/@w:name');anchor=bn[0] if bn else 'MASAR_CH2'
            if not bn:bookmark(ch2._p,anchor,bid);bid+=1
            result.append({'title':'Literature Review','anchor':anchor,'style':'Heading1','label':'Chapter 2:'})
            result.extend({**s,'style':'Heading2'} for s in sections)
    insert_before=node('p');tocpars[0]._p.addprevious(insert_before)
    for p in tocpars:body.remove(p._p)
    for j,ent in enumerate(result):
        p=doc.add_paragraph(style='TOC2' if ent['style']=='Heading2' or ent['style']=='PreliminaryHeadingCoEKSU' else 'TOC1');insert_before.addprevious(p._p)
        p.paragraph_format.first_line_indent=Inches(0);p.paragraph_format.space_after=Pt(1);p.paragraph_format.line_spacing=1.0
        if j==0:p._p.append(field_run('begin'));p._p.append(instruction(' TOC \\o "1-2" \\h \\z '));p._p.append(field_run('separate'))
        h=node('hyperlink',anchor=ent['anchor'],history='1');r=node('r');t=node('t');t.text=(ent['label']+' ' if ent['label'] else '')+ent['title'];r.append(t);h.append(r);p._p.append(h);p.add_run('\t0')
        bookmark(p._p,'MASAR_TOC_CACHE_'+str(j),bid);bid+=1;ent['cache']='MASAR_TOC_CACHE_'+str(j)
    insert_before.append(field_run('end'))
    for kind in ['Figure','Table']:
        p=next(p for p in doc.paragraphs if any('\\c "'+kind+'"' in x.text for x in p._p.xpath('.//w:instrText')))
        for el in list(p._p):
            if el.tag!=qn('w:pPr'):p._p.remove(el)
        p._p.append(field_run('begin'));p._p.append(instruction(' TOC \\h \\z \\c "'+kind+'" '));p._p.append(field_run('separate'))
        for j,c in enumerate(x for x in captions if x['kind']==kind):
            if j:p.add_run('\n')
            h=node('hyperlink',anchor=c['anchor'],history='1');r=node('r');t=node('t');t.text=c['label']+' '+c['title'];r.append(t);h.append(r);p._p.append(h);p.add_run('\t0')
        p._p.append(field_run('end'));p.paragraph_format.line_spacing=1.15;p.alignment=WD_ALIGN_PARAGRAPH.LEFT;p.paragraph_format.first_line_indent=Inches(0);p.paragraph_format.left_indent=Inches(0)
    doc.save(args.output)
    public_records={k:{field:value for field,value in record.items() if field!='citation_item'} for k,record in records.items()}
    (ROOT/'docs/chapter-2-source-register.json').write_text(json.dumps(public_records,ensure_ascii=False,indent=2)+'\n')
    # Original six entries remain byte-for-byte intact in the BibTeX source.
    old=(ROOT/'docs/references.bib').read_text();old=old.split('% Chapter 2 sources')[0].rstrip()
    (ROOT/'docs/references.bib').write_text(old+'\n\n% Chapter 2 sources — metadata reviewed 2026-11-01\n\n'+'\n'.join(bibtex(k,records[k]['csl']) for k in order[6:]))
    (ROOT/'docs/chapter-2-layout-map.json').write_text(json.dumps({'toc':result,'captions':captions},indent=2)+'\n')
    print(json.dumps({'chapter_words':len(text.split()),'references':len(order),'new_citation_fields':counter,'main_sections':sec,'tables':table_no}))

if __name__=='__main__':main()

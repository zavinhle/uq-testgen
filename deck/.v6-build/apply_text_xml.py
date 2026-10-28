"""Apply tagged Markdown to existing OOXML without rebuilding the deck."""
from pathlib import Path
import copy, re, zipfile, json
from lxml import etree as E

ROOT=Path(__file__).resolve().parent.parent
NS={'a':'http://schemas.openxmlformats.org/drawingml/2006/main','p':'http://schemas.openxmlformats.org/presentationml/2006/main'}
def q(t):
    prefix,name=t.split(':'); return '{'+NS[prefix]+'}'+name
blocks={}; slide=None; tag=None
for line in (ROOT/'v5.txt.md').read_text(encoding='utf-8').splitlines():
    m=re.match(r'^=== SLIDE (\d+)',line)
    if m: slide=int(m[1]); tag=None; continue
    m=re.match(r'^\[(#\d+(?: r\d+c\d+)?|notes)\]\s?(.*)$',line)
    if m and slide: tag=m[1]; blocks[slide,tag]=[m[2]]
    elif tag and line.strip(): blocks[slide,tag].append(line[6:] if line.startswith('      ') else line.strip())

def norm(a):
    a=[x.rstrip() for x in a]
    while a and not a[-1]: a.pop()
    return a
def para_md(p):
    groups=[]
    for r in p.findall(q('a:r')):
        t=''.join(r.itertext()).replace('\n',' ').replace('\x0b',' ')
        rp=r.find(q('a:rPr')); b=rp.get('b') if rp is not None else None
        weight=(b=='1') if b is not None else (groups[-1][0] if groups else False)
        weight=weight and bool(t.strip())
        if groups and (groups[-1][0]==weight or not t.strip()): groups[-1][1]+=t
        else: groups.append([weight,t])
    return ''.join('**'+t+'**' if b else t for b,t in groups)
changed=[]; found=set()
def update(body, number, tag):
    key=(number,tag)
    if key not in blocks: return
    found.add(key); new=norm(blocks[key]); old=norm([para_md(p) for p in body.findall(q('a:p'))])
    if new==old: return
    paras=body.findall(q('a:p')); proto=paras[-1] if paras else E.Element(q('a:p'))
    runs=body.findall('.//'+q('a:r'))
    def bold(r):
        rp=r.find(q('a:rPr')); return rp is not None and rp.get('b')=='1'
    pb=next((r for r in runs if bold(r)),None); pp=next((r for r in runs if not bold(r)),None)
    for i,line in enumerate(new or ['']):
        p=paras[i] if i<len(paras) else copy.deepcopy(proto)
        if i>=len(paras): body.append(p)
        for child in list(p):
            if child.tag not in [q('a:pPr'),q('a:endParaRPr')]: p.remove(child)
        for j,text in enumerate(re.split(r'\*\*',line)):
            if not text: continue
            weight=j%2==1; template=pb if weight else pp
            if template is None: template=pp if pp is not None else pb
            r=E.Element(q('a:r'))
            if template is not None and template.find(q('a:rPr')) is not None: r.append(copy.deepcopy(template.find(q('a:rPr'))))
            rp=r.find(q('a:rPr'))
            if rp is None: rp=E.SubElement(r,q('a:rPr'))
            rp.set('b','1' if weight else '0'); E.SubElement(r,q('a:t')).text=text
            end=p.find(q('a:endParaRPr'))
            if end is not None: end.addprevious(r)
            else: p.append(r)
    for p in paras[max(1,len(new)):]: body.remove(p)
    changed.append([number,tag])

src=ROOT/'PhD_Proposal_ZaVinhLe_v5.pptx'; dst=ROOT/'PhD_Proposal_ZaVinhLe_v6.pptx'
with zipfile.ZipFile(src) as zi,zipfile.ZipFile(dst,'w',zipfile.ZIP_DEFLATED) as zo:
    for item in zi.infolist():
        data=zi.read(item.filename)
        m=re.fullmatch(r'ppt/slides/slide(\d+)\.xml',item.filename)
        n=re.fullmatch(r'ppt/notesSlides/notesSlide(\d+)\.xml',item.filename)
        before=len(changed)
        if m or n:
            number=int((m or n)[1]); root=E.fromstring(data)
            if m:
                for sp in root.findall('.//p:sp',NS):
                    cid=sp.find('p:nvSpPr/p:cNvPr',NS); body=sp.find('p:txBody',NS)
                    if cid is not None and body is not None: update(body,number,'#'+cid.get('id'))
                for frame in root.findall('.//p:graphicFrame',NS):
                    cid=frame.find('p:nvGraphicFramePr/p:cNvPr',NS)
                    if cid is None: continue
                    for ri,row in enumerate(frame.findall('.//a:tbl/a:tr',NS),1):
                        for ci,cell in enumerate(row.findall('a:tc',NS),1): update(cell.find('a:txBody',NS),number,f'#{cid.get("id")} r{ri}c{ci}')
            else:
                for sp in root.findall('.//p:sp',NS):
                    ph=sp.find('p:nvSpPr/p:nvPr/p:ph',NS)
                    if ph is not None and ph.get('type')=='body': update(sp.find('p:txBody',NS),number,'notes')
            if len(changed)>before: data=E.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True)
        zo.writestr(item,data)
assert set(blocks)<=found, f'Unknown blocks: {set(blocks)-found}'
(ROOT/'.v6-build/changes.json').write_text(json.dumps(changed,indent=2))
print('Changed blocks:',changed)
print('Written:',dst)

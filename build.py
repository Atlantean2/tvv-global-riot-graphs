"""Builds the TechVestVentures Global Riot Graphs site from Grok Bot's Oct 22 2026 Political Violence Focus Study (docx + G1-G10 PNGs). Text is taken verbatim from the study."""
import os,re,sys,glob,shutil,docx
from docx.oxml.ns import qn
sys.path.insert(0,'/workspace/risk-watcher/graphs-site');os.chdir(os.path.dirname(os.path.abspath(__file__)))
import common
from common import *
T="TechVestVentures Global Riot Graphs";F="/workspace/risk-watcher/issues/2026-10-22"
D=docx.Document(F+"/TVV Focus Study 2026-10-22 Political Violence.docx")
def rich(p):
    out=[]
    for el in p._p:
        if el.tag==qn('w:r'): out.append(E(''.join(t.text or '' for t in el.iter(qn('w:t')))))
        elif el.tag==qn('w:hyperlink'):
            rid=el.get(qn('r:id'));u=D.part.rels[rid].target_ref if rid else ''
            tx=E(''.join(t.text or '' for t in el.iter(qn('w:t'))))
            out.append(f'<a href="{E(u)}" target="_blank" rel="noopener">{tx}</a>' if u.startswith('https://') else tx)
    return ''.join(out)
ps=[p for p in D.paragraphs if p.text.strip()]
cap={};extra={}
for i,p in enumerate(ps):
    m=re.match(r"Graph (\d+)\. ",p.text)
    if m: cap[int(m[1])]=rich(p)
    m=re.match(r"(?:Reading )?Graph (\d+)(?: makes|\.)",p.text)
    if m and not re.match(r"Graph \d+\. ",p.text): extra[int(m[1])]=rich(p)
pngs={int(re.search(r" G(\d+) ",f)[1]):f for f in glob.glob(F+"/TVV Focus 2026-10-22 G*.png")}
assert sorted(cap)==sorted(pngs)==list(range(1,11)),(sorted(cap),sorted(pngs))
SEC=[("Q1","Protest Waves","The Data Picture: Protest Waves Are Trigger-Driven, Not Calendar-Driven",[1,2,3]),
     ("Q2","Election Cycle","Demonstrations Around National Elections",[4]),
     ("Q3","France October","France: Anatomy of the October Wave",[5,6,7]),
     ("Q4","Violence Share","Long-Run Context: France Is the Outlier on Protest Violence",[8,9]),
     ("Q5","Threat Counts","Threats Against Congress and German Extremist Violence",[10])]
common.CSS+="""
figure.g{display:grid;grid-template-columns:minmax(0,7fr) minmax(260px,3fr);gap:22px;align-items:start;margin:0 0 40px}
figure.g img{width:100%;height:auto;max-width:none}
figure.g figcaption{font-size:18px;line-height:1.5;margin:0}figure.g figcaption p{margin:0 0 12px}
figure.g figcaption .t{font-weight:700;display:block;margin-bottom:8px}
@media(max-width:1100px){figure.g{grid-template-columns:1fr}}"""
shutil.rmtree("issues",ignore_errors=True);os.makedirs("issues/img")
tabs=[]
def w(k,lab,body): open(f"issues/{k}.html","w").write(page(T,k,body));tabs.append((k,lab))
for q,lab,head,gs in SEC:
    figs=[]
    for g in gs:
        fn=f"G{g:02d}.webp";save_img(pngs[g],"issues/img/"+fn)
        x=f"<p>{extra[g]}</p>" if g in extra else ""
        figs.append(f'<figure class="g"><img src="img/{fn}" alt="Graph {g}" loading="lazy"><figcaption><p>{cap[g]}</p>{x}</figcaption></figure>')
    w(q,f"{q} {lab}",f'<h2 class="desk">{q}. {E(head)}</h2>'+"\n".join(figs))
src=[];seen=set()
for h in D.element.body.iter(qn('w:hyperlink')):
    rid=h.get(qn('r:id'));u=D.part.rels[rid].target_ref if rid else ''
    if u.startswith('https://') and u not in seen:
        seen.add(u);src.append((''.join(t.text or '' for t in h.iter(qn('w:t'))),u))
w("sources","Sources",f'<h2 class="desk">Sources ({len(src)} links)</h2><p>From the TVV Focus Study of October 22, 2026: Political Violence, Organized Protest and Threat Profiles.</p><ul class="src">'+"".join(f'<li>{E(t)}: <a href="{E(u)}">{E(u)}</a></li>' for t,u in src)+"</ul>")
index(T,tabs,"Sections",cur="riot");print(len(src))

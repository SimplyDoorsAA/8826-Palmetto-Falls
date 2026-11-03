"""Extract wood + fiberglass catalogs and all adder tables into one JSON payload."""
import openpyxl, pypdf, re, json

WOOD='/root/.claude/uploads/4a23149a-1401-565c-9803-04d52a1258a0/d8adb2fb-Hoelscher_Wood_Doors__Searchable_Pricing__Quoting_Master_2026.xlsx'
FG='/root/.claude/uploads/4a23149a-1401-565c-9803-04d52a1258a0/b71719ff-Hoelsher_Fiberglass_price.pdf'

def size_of(desc, part=''):
    m=re.match(r'^\s*(\d{4,5})\b', desc or '')
    if m: return m.group(1)
    m=re.search(r'(\d{4,5})$', part or '')
    return m.group(1) if m else ''

# ---------- WOOD ----------
wb=openpyxl.load_workbook(WOOD)
ws=wb['Doors Master Pricing']
rows=[r for r in ws.iter_rows(values_only=True) if any(v not in (None,'') for v in r)][1:]
wood=[]
for r in rows:
    wood.append(dict(
        species=r[0], category=r[1], page=r[2], part=r[3], desc=r[4], glazing=r[5],
        size=size_of(r[4], r[3]),
        uf_slab=r[6], uf_sph=r[8], uf_dph=r[10],
        pf_slab=r[12], pf_sph=r[14], pf_dph=r[16],
        notes=r[18] or ''))
print('wood rows:', len(wood))

# ---------- FIBERGLASS DOORS ----------
r_=pypdf.PdfReader(FG)
layout=[p.extract_text(extraction_mode='layout') for p in r_.pages]
plain=[p.extract_text() for p in r_.pages]

def num(tok):
    tok=tok.strip()
    if tok in ('N/A','','NA'): return None
    m=re.match(r'^\+?\$?([\d,]+)(\(ea\))?$', tok)
    return int(m.group(1).replace(',','')) if m else None

SKINS=[('White Brushed Smooth','White Brushed Smooth'),('Black Brushed Smooth','Black Brushed Smooth'),
       ('Mahogany Grain','Mahogany Grain'),('Oak Grain','Oak Grain'),('Fir Skin','Fir'),('Smooth Skin','Smooth')]
def skin_of(d):
    for needle,label in SKINS:
        if needle in d: return label
    return 'Other'
def cat_of(d):
    dl=d.lower()
    if 'sidelite' in dl: return 'Sidelite'
    if 'contemporary' in dl: return 'Contemporary'
    if 'tdl' in dl: return 'True Divided Lite (TDL)'
    if 'craftsman' in dl: return 'Craftsman'
    if 'mini blind' in dl: return 'Full Lite w/ Mini Blind'
    if 'full lite' in dl: return 'Full Lite'
    if '3/4 lite' in dl: return '3/4 Lite'
    if 'half lite' in dl: return 'Half Lite'
    if 'flush' in dl: return 'Flush'
    if 'panel' in dl: return 'Panel Doors'
    return 'Other'

fg=[]; seen=set()
for pi,t in enumerate(layout):
    for ln in t.split('\n'):
        s=ln.strip()
        if not re.match(r'^(FG|NL)[A-Z0-9]', s): continue
        f=[x for x in re.split(r'\s{3,}', s) if x]
        if len(f)!=9: continue
        part,desc,glaz=f[0],f[1],f[2]
        if part in seen: continue
        seen.add(part)
        p=f[3:]
        fg.append(dict(part=part, desc=desc, glazing=glaz, skin=skin_of(desc),
                       category=cat_of(desc), size=size_of(desc, part),
                       is_sidelite='Sidelite' in desc,
                       uf_slab=num(p[0]), uf_sph=num(p[1]), uf_dph=num(p[2]),
                       pf_slab=num(p[3]), pf_sph=num(p[4]), pf_dph=num(p[5])))
# configured items
for pi,t in enumerate(plain):
    lines=t.split('\n')
    for i,ln in enumerate(lines):
        if not ln.startswith('CONFIGURED ITEM'): continue
        blob=ln; j=i
        while len([x for x in blob.split() if re.match(r'^\d{3,4}$',x)])<6 and j+1<len(lines):
            j+=1; blob+=' '+lines[j].strip()
        toks=blob.replace('CONFIGURED ITEM','').strip().split()
        pr=[]
        while toks and re.match(r'^\d{3,4}$', toks[-1]) and len(pr)<6:
            pr.insert(0, int(toks.pop()))
        if len(pr)!=6: continue
        d=' '.join(toks).rstrip('*').strip()
        fg.append(dict(part='CONFIG-'+re.sub(r'[^A-Z0-9]','',d.upper())[:18], desc=d, glazing='See catalog',
                       skin=skin_of(d), category='Configured Item', size=size_of(d),
                       is_sidelite=False,
                       uf_slab=pr[0], uf_sph=pr[1], uf_dph=pr[2],
                       pf_slab=pr[3], pf_sph=pr[4], pf_dph=pr[5]))
print('fiberglass rows:', len(fg), '| sidelites:', sum(1 for x in fg if x['is_sidelite']))

# ---------- ADDERS ----------
ws2=wb['Hardware, Jambs & Adders']
a=[r for r in ws2.iter_rows(values_only=True) if any(v not in (None,'') for v in r)]
adders={'raw_wood':[list(r) for r in a]}

fgcomp=[]; sdl=[]; prehang=[]
t4=layout[3]
for ln in t4.split('\n'):
    s=ln.rstrip()
    m=re.match(r'^\s+(\d{2,3}")\s{3,}(.+?)\s{3,}(\d+)\s*$', s)
    if m: fgcomp.append([m.group(1), m.group(2).strip(), int(m.group(3))]); continue
    if 'Dentil Shelf' in s:
        m2=re.search(r'(Dentil Shelf.+?)\s{3,}(\d+)\s*$', s)
        if m2: fgcomp.append(['-', m2.group(1).strip(), int(m2.group(2))]); continue
    m=re.match(r'^(SDL8BAR\S+)\s{3,}(.+?)\s{3,}(.+?)\s{3,}\$(\d+)\s*$', s)
    if m: sdl.append([m.group(1), m.group(2).strip(), m.group(3).strip(), int(m.group(4))]); continue
    m=re.match(r'^\s+(\d{4}\s+(?:Single|Double)[^\s].*?|\d{4}\s+(?:Single|Double))\s{3,}(\d+)\s{3,}(\d+)\s*$', s)
    if m: prehang.append([m.group(1).strip(), int(m.group(2)), int(m.group(3))])
adders['fg_components']=fgcomp; adders['sdl_bars']=sdl; adders['fg_prehang']=prehang
print('fg components:', len(fgcomp), '| sdl bars:', len(sdl), '| prehang adders:', len(prehang))

ws3=wb['Glass Codes & Reference Guide']
ref=[list(r) for r in ws3.iter_rows(values_only=True) if any(v not in (None,'') for v in r)]
json.dump(dict(wood=wood, fg=fg, adders=adders, reference=ref), open('data.json','w'), indent=1, default=str)
print('OK -> data.json')

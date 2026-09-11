"""Reconcile every price in the built workbook against the original sources."""
import openpyxl, pypdf, re, json

built=openpyxl.load_workbook('wip.xlsx')
src=openpyxl.load_workbook('/root/.claude/uploads/4a23149a-1401-565c-9803-04d52a1258a0/d8adb2fb-Hoelscher_Wood_Doors__Searchable_Pricing__Quoting_Master_2026.xlsx')

# ---- WOOD: built catalog vs original workbook ----
s=src['Doors Master Pricing']
orig={}
for r in s.iter_rows(min_row=2, values_only=True):
    if not r[3]: continue
    orig[r[3]]=tuple((None if (isinstance(v,str) and v.strip().upper() in ('N/A','NA','-')) else v) for v in (r[6],r[8],r[10],r[12],r[14],r[16]))
wc=built['Wood Catalog']
bad=[]; n=0
for rr in range(5, wc.max_row+1):
    part=wc.cell(row=rr,column=3).value
    if not part: continue
    got=tuple(wc.cell(row=rr,column=c).value for c in range(8,14))
    exp=orig.get(part)
    n+=1
    if exp is None: bad.append((part,'not in source',None,None)); continue
    for i,(a,b) in enumerate(zip(got,exp)):
        if (a is None)!=(b is None) or (a is not None and abs(float(a)-float(b))>0.005):
            bad.append((part,f'price col {i+1}',a,b))
print(f'WOOD  : {n} rows checked, {len(bad)} mismatches')
for x in bad[:8]: print('   ',x)

# ---- FIBERGLASS: built catalog vs a fresh independent parse of the PDF ----
r_=pypdf.PdfReader('/root/.claude/uploads/4a23149a-1401-565c-9803-04d52a1258a0/b71719ff-Hoelsher_Fiberglass_price.pdf')
def num(t):
    t=t.strip()
    if t in ('N/A','','NA'): return None
    m=re.match(r'^\+?\$?([\d,]+)(\(ea\))?$', t)
    return float(m.group(1).replace(',','')) if m else None
pdf={}
for p in r_.pages:
    for ln in p.extract_text(extraction_mode='layout').split('\n'):
        s2=ln.strip()
        if not re.match(r'^(FG|NL)[A-Z0-9]', s2): continue
        f=[x for x in re.split(r'\s{3,}', s2) if x]
        if len(f)!=9: continue
        pdf.setdefault(f[0], tuple(num(x) for x in f[3:]))
fc=built['Fiberglass Catalog']
bad2=[]; n2=0; conf=0
for rr in range(5, fc.max_row+1):
    part=fc.cell(row=rr,column=3).value
    if not part: continue
    if str(part).startswith('CONFIG-'): conf+=1; continue
    got=tuple(fc.cell(row=rr,column=c).value for c in range(8,14))
    exp=pdf.get(part); n2+=1
    if exp is None: bad2.append((part,'not in PDF',None,None)); continue
    for i,(a,b) in enumerate(zip(got,exp)):
        if (a is None)!=(b is None) or (a is not None and abs(float(a)-float(b))>0.005):
            bad2.append((part,f'price col {i+1}',a,b))
print(f'FIBERGLASS: {n2} rows checked ({conf} configured items skipped - no part number in PDF), {len(bad2)} mismatches')
for x in bad2[:8]: print('   ',x)

# ---- ADDERS: spot-check against the original workbook rows ----
raw=[list(r) for r in src['Hardware, Jambs & Adders'].iter_rows(values_only=True) if any(v not in (None,'') for v in r)]
ad=built['Adders']
idx={}
for rr in range(5, ad.max_row+1):
    idx[(str(ad.cell(row=rr,column=3).value), str(ad.cell(row=rr,column=4).value))]=ad.cell(row=rr,column=5).value
checks=[('Top Mount Strap Kit (Matte Black)','Matte Black',561),
        ('Top Mount Strap Kit (Stainless)','Stainless',665),
        ('Iron Mask - Windsor (IMW)','Iron / Black',242),
        ('100"-6-9/16" Exterior Jamb Leg','Knotty Alder',121),
        ('100"-6-9/16" Exterior Jamb Leg','Frame-Saver 8/0',88),
        ('3680 Construction Door','Unfinished',349),
        ('Square Top / 3080 Double','1x4 Trim',197),
        ('PVC Jamb 6-9/16" x 8-0 (100")','100"',94),
        ('SDL 8\' Lineal Bar Textured Black SR 1/4" X 5"','8 pcs/box',171),
        ('3080 Double','1x4 Trim',197)]
bad3=[]
for desc,opt,exp in checks:
    got=idx.get((desc,opt))
    if got is None or abs(float(got)-exp)>0.005: bad3.append((desc,opt,got,exp))
print(f'ADDERS: {len(checks)} spot-checks, {len(bad3)} mismatches')
for x in bad3: print('   ',x)
print()
print('RECONCILIATION', 'CLEAN' if not (bad or bad2 or bad3) else 'HAS ERRORS')

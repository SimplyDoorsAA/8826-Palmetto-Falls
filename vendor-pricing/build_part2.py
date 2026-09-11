"""Adders (flattened), Freight Rules, Lists."""
import json, openpyxl, re
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

D=json.load(open('data.json')); adders=D['adders']; raw=adders['raw_wood']
wb=openpyxl.load_workbook('wip.xlsx')

STEEL='2E5C8A'; LIGHT='D9E2F3'; NAVY='1F3864'; CREAM='FFF2CC'
H_FONT=Font(bold=True, color='FFFFFF', size=10); H_FILL=PatternFill('solid', fgColor=STEEL)
TITLE=Font(bold=True, size=14, color=NAVY); SUB=Font(bold=True, size=11, color=NAVY)
thin=Side(style='thin', color='BFBFBF'); BOX=Border(left=thin,right=thin,top=thin,bottom=thin)
MONEY='#,##0.00'; MONEY0='#,##0'
def hdr(ws,row,labels,widths=None):
    for i,l in enumerate(labels):
        c=ws.cell(row=row,column=1+i,value=l); c.font=H_FONT; c.fill=H_FILL; c.border=BOX
        c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    if widths:
        for i,w in enumerate(widths): ws.column_dimensions[get_column_letter(1+i)].width=w
    ws.row_dimensions[row].height=30

def nn(v):
    if v in (None,'','N/A','#ERROR!'): return None
    try: return float(v)
    except: return None

# ---- flatten every adder into one quotable table ----
A=[]  # code, group, description, option, list, frt_units_each, notes
def add(code, group, desc, opt, lst, frt, notes=''):
    if lst is None: return
    A.append([code, group, desc, opt, float(lst), frt, notes])

for r in raw[1:23]:                                   # barn door hardware
    mb, ss = nn(r[3]), nn(r[5])
    if mb is not None: add(r[0], 'Barn Door Hardware', r[1], 'Matte Black', mb, 0, r[7] or '')
    if ss is not None: add(r[0], 'Barn Door Hardware', r[1], 'Stainless',  ss, 0, r[7] or '')
for r in raw[24:34]:                                  # speakeasy / iron / SDL
    add(r[0], 'Speakeasy & Iron', r[1], r[2], nn(r[3]), 0,
        '' if str(r[5]).startswith('#ERROR') else (r[5] or ''))
JAMB=[('Mahogany',1),('Knotty Alder',3),('Frame-Saver 6/8',5),('Frame-Saver 8/0',7)]
for r in raw[35:49]:                                  # jambs & components
    comp=r[0]
    is_leg='Jamb Leg' in comp
    frt=1/25 if is_leg else 1/50
    for label,ci in JAMB:
        v=nn(r[ci])
        if v is not None:
            code=re.sub(r'[^A-Za-z0-9]','', comp)[:14].upper()+'-'+label.split()[0][:3].upper()
            add(code, 'Jambs & Components', comp, label, v, frt, r[9] or '')
for r in raw[50:55]:                                  # construction doors
    add(re.sub(r'\D','',str(r[0]))+'-CONSTR', 'Construction Doors', r[0], 'Unfinished', nn(r[1]), 1, r[3] or '')
for r in raw[56:75]:                                  # wood prehang adders
    base=f'{r[0]} / {r[1]}'
    add('PH-'+re.sub(r'[^A-Za-z0-9]','',str(r[1]))[:16].upper()+'-IC', 'Prehang Adders (Wood)', base, 'Interior Casing', nn(r[2]), 0, r[6] or '')
    add('PH-'+re.sub(r'[^A-Za-z0-9]','',str(r[1]))[:16].upper()+'-1X4','Prehang Adders (Wood)', base, '1x4 Trim',        nn(r[4]), 0, r[6] or '')
for size,desc,price in adders['fg_components']:       # fiberglass PVC
    frt=1/25 if 'Jamb' in desc else 1/50
    add('FGC-'+re.sub(r'[^A-Za-z0-9]','',desc)[:16].upper(), 'Fiberglass Components', desc, size, price, frt, 'Fiberglass PDF p.4')
for part,desc,box,price in adders['sdl_bars']:        # SDL bars
    add(part, 'SDL Bars (Fiberglass)', desc, box, price, 1/50, 'Fiberglass PDF p.4')
for label,ic,x4 in adders['fg_prehang']:              # fiberglass prehang adders
    add('FGPH-'+re.sub(r'[^A-Za-z0-9]','',label)[:16].upper()+'-IC','Prehang Adders (Fiberglass)', label,'Interior Casing', ic, 0,'Fiberglass PDF p.4')
    add('FGPH-'+re.sub(r'[^A-Za-z0-9]','',label)[:16].upper()+'-1X4','Prehang Adders (Fiberglass)', label,'1x4 Trim',       x4, 0,'Fiberglass PDF p.4')

seen={}
for row in A:                                         # guarantee unique codes
    c=row[0]
    if c in seen:
        seen[c]+=1; row[0]=f'{c}-{seen[c]}'
    else: seen[c]=1

ad=wb.create_sheet('Adders')
ad['A1']='Adders, Hardware, Jambs & Components  -  LIST PRICES'; ad['A1'].font=TITLE
ad['A2']=('Every adder from both price sheets, flattened into one quotable list. Frt Units/Ea encodes the vendor freight rule: '
          '25 jamb legs = 1 unit, 50 other components = 1 unit, hardware ships free.')
ad['A2'].font=Font(italic=True, size=9, color='595959')
hdr(ad,4,['Code','Group','Description','Option / Finish','List Price','Net Cost','Frt Units/Ea','Notes'],
    [24,26,44,20,12,12,12,34])
for i,row in enumerate(A):
    rr=5+i
    for j,v in enumerate([row[0],row[1],row[2],row[3],row[4]]):
        c=ad.cell(row=rr,column=1+j,value=v); c.font=Font(size=9)
        if j==4: c.number_format=MONEY0
    c=ad.cell(row=rr,column=6,value=f'=E{rr}*Multiplier'); c.number_format=MONEY; c.font=Font(size=9)
    c=ad.cell(row=rr,column=7,value=round(row[5],4)); c.number_format='0.0000'; c.font=Font(size=9)
    ad.cell(row=rr,column=8,value=row[6]).font=Font(size=9)
    if i%2:
        for j in range(1,9): ad.cell(row=rr,column=j).fill=PatternFill('solid',fgColor='FAFAFA')
ad.freeze_panes='A5'; ad.auto_filter.ref=f'A4:H{4+len(A)}'
ADD_LAST=4+len(A)

# ---- FREIGHT RULES ----
fr=wb.create_sheet('Freight Rules')
fr['A1']='Freight & Crating Rules  -  Hoelscher, FOB Run (San Antonio / Austin)'; fr['A1'].font=TITLE
fr['A2']=('Freight is INBOUND to your warehouse, so it is landed cost, not a customer line. '
          'The Quote tab folds it into door cost before margin.')
fr['A2'].font=Font(italic=True,size=9,color='595959')
for col,w in zip('ABCDEF',[34,14,14,14,16,46]): fr.column_dimensions[col].width=w

fr['A4']='Freight Unit Equivalents'; fr['A4'].font=SUB; fr['A4'].fill=PatternFill('solid',fgColor=LIGHT)
hdr(fr,5,['Item','Units Each','','','',''])
UNITS=[('1 door slab',1),('1 single door (prehung)',1),('1 double door (prehung)',2),
       ('1 sidelite',0.5),('1 transom',1),('25 jamb legs (4-9/16" or 6-9/16")',1),
       ('50 pcs all other components (brickmould, casing, 1x4, T-astragal)',1)]
for i,(k,v) in enumerate(UNITS):
    fr.cell(row=6+i,column=1,value=k).font=Font(size=9)
    c=fr.cell(row=6+i,column=2,value=v); c.number_format='0.0'; c.font=Font(size=9); c.border=BOX

fr['A15']='Shipping Tier  (charged on total freight units per shipment)'; fr['A15'].font=SUB
fr['A15'].fill=PatternFill('solid',fgColor=LIGHT)
hdr(fr,16,['Tier','Min Units','Max Units','Charge','',''])
TIERS=[('1-3 units',1,3,120),('4-6 units',4,6,165),('7-9 units',7,9,200),('10+ units (prepaid)',10,9999,0)]
for i,(lab,lo,hi,ch) in enumerate(TIERS):
    rr=17+i
    fr.cell(row=rr,column=1,value=lab).font=Font(size=9)
    for col,v,fmt in ((2,lo,'0'),(3,hi,'0'),(4,ch,MONEY0)):
        c=fr.cell(row=rr,column=col,value=v); c.number_format=fmt; c.font=Font(size=9); c.border=BOX
fr['A22']='Crate Charge'; fr['A22'].font=SUB; fr['A22'].fill=PatternFill('solid',fgColor=LIGHT)
fr['A23']='Per prehung unit (opening), up to 6080'
c=fr['B23']; c.value=100; c.number_format=MONEY0; c.fill=PatternFill('solid',fgColor=CREAM); c.border=BOX
fr['C23']='Slabs do not incur a crate charge. Counted per opening, not per door.'
fr['C23'].font=Font(size=9,italic=True,color='595959')
wb.save('wip.xlsx')
print(f'adders: {len(A)} rows (5-{ADD_LAST}); freight rules built')
json.dump({'ADD_LAST':ADD_LAST,'N_ADD':len(A)}, open('meta.json','w'))

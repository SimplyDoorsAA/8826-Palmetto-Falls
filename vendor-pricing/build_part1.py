"""Hoelscher Quoting Workbook - catalogs, config, adders, lists, freight."""
import json, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

D=json.load(open('data.json'))
wood, fg, adders, ref = D['wood'], D['fg'], D['adders'], D['reference']

NAVY='1F3864'; STEEL='2E5C8A'; LIGHT='D9E2F3'; CREAM='FFF2CC'; GREEN='E2EFDA'; GREY='F2F2F2'
H_FONT=Font(bold=True, color='FFFFFF', size=10)
H_FILL=PatternFill('solid', fgColor=STEEL)
TITLE=Font(bold=True, size=14, color=NAVY)
SUB=Font(bold=True, size=11, color=NAVY)
SECT_FILL=PatternFill('solid', fgColor=LIGHT)
INPUT_FILL=PatternFill('solid', fgColor=CREAM)
CALC_FILL=PatternFill('solid', fgColor=GREY)
thin=Side(style='thin', color='BFBFBF')
BOX=Border(left=thin,right=thin,top=thin,bottom=thin)
MONEY='#,##0.00'; MONEY0='#,##0'; PCT='0.0%'

wb=openpyxl.Workbook()
wb.remove(wb.active)

def header_row(ws, row, labels, widths=None, start=1):
    for i,l in enumerate(labels):
        c=ws.cell(row=row, column=start+i, value=l)
        c.font=H_FONT; c.fill=H_FILL; c.border=BOX
        c.alignment=Alignment(horizontal='center', vertical='center', wrap_text=True)
    if widths:
        for i,w in enumerate(widths):
            ws.column_dimensions[get_column_letter(start+i)].width=w
    ws.row_dimensions[row].height=30

def title(ws, text, sub_text=''):
    ws['A1']=text; ws['A1'].font=TITLE
    if sub_text:
        ws['A2']=sub_text; ws['A2'].font=Font(italic=True, size=9, color='595959')

# ============ CONFIG ============
cfg=wb.create_sheet('Config')
title(cfg,'Configuration','Every constant in this workbook lives here. Change it once, everything downstream follows.')
cfg.column_dimensions['A'].width=34; cfg.column_dimensions['B'].width=20; cfg.column_dimensions['C'].width=66
rows=[
 ('SECTION','Vendor',''),
 ('Vendor','Hoelscher Weatherstrip Mfg. Co.','Hoelscher Doors'),
 ('Facility','2400 S. Persimmon St., Tomball, TX 77375',''),
 ('Contact','713-869-6466','hoelscherdoors.com'),
 ('SECTION','Pricing',''),
 ('Cost Multiplier','MULT','Dealer NET = List x this. Applies to wood, fiberglass, hardware and jambs.'),
 ('Default Target Margin','MARGIN','Starting margin on new quotes. Override per quote on the Quote tab.'),
 ('SECTION','Price Sheet Version',''),
 ('Wood price list effective','2026-07-01','Revised 2026-06-19'),
 ('Fiberglass price list effective','2026-07-01','Revised 2026-05-20'),
 ('SECTION','Open Items to Confirm with Hoelscher',''),
 ('Freight 10-door clause','OPEN','Shipping sheet reads "$100 NET PER UNIT + ABOVE SHIPPING CHARGES IF 10 DOOR MINIMUM", which contradicts their own Examples 1 and 2. Encoded per the worked examples. Confirm.'),
 ('Missing descriptions','OPEN','NLSMFL3080NB and FGSHCMLEF3680NP have no description on the vendor PDF. Inferred from part number - verify.'),
 ('Sidelite pricing','OPEN','Sidelite prehung prices are per-each ADDERS, not standalone unit prices. Confirm a sidelite is never ordered prehung on its own.'),
]
r=4
for a,b,c in rows:
    if a=='SECTION':
        cfg.cell(row=r, column=1, value=b).font=SUB
        for col in range(1,4): cfg.cell(row=r, column=col).fill=SECT_FILL
        r+=1; continue
    cfg.cell(row=r, column=1, value=a).font=Font(bold=True, size=10)
    if b=='MULT':
        cell=cfg.cell(row=r, column=2, value=0.48); cell.number_format='0.00'; cell.fill=INPUT_FILL; cell.border=BOX
        MULT_CELL=f'Config!$B${r}'
    elif b=='MARGIN':
        cell=cfg.cell(row=r, column=2, value=0.30); cell.number_format=PCT; cell.fill=INPUT_FILL; cell.border=BOX
        MARGIN_CELL=f'Config!$B${r}'
    elif b=='OPEN':
        cell=cfg.cell(row=r, column=2, value='CONFIRM'); cell.fill=PatternFill('solid', fgColor='FFC7CE')
        cell.font=Font(bold=True, color='9C0006', size=9); cell.border=BOX
    else:
        cfg.cell(row=r, column=2, value=b)
    cfg.cell(row=r, column=3, value=c).font=Font(size=9, color='404040')
    cfg.cell(row=r, column=3).alignment=Alignment(wrap_text=True, vertical='top')
    r+=1

# ============ WOOD CATALOG ============
wc=wb.create_sheet('Wood Catalog')
title(wc,'Wood Door Catalog  -  LIST PRICES','LIST price only. Dealer net is calculated on the Quote tab from the Config multiplier. To update: paste new list prices over columns H-M only.')
WCOLS=['Species','Category','Part Number','Description','Glazing','Size','Cat. Page',
       'UF Slab','UF Single PH','UF Double PH','PF Slab','PF Single PH','PF Double PH','Notes']
header_row(wc, 4, WCOLS, [14,26,18,46,30,8,9,11,13,13,11,13,13,30])
for i,d in enumerate(wood):
    rr=5+i
    vals=[d['species'],d['category'],d['part'],d['desc'],d['glazing'],d['size'],d['page'],
          d['uf_slab'],d['uf_sph'],d['uf_dph'],d['pf_slab'],d['pf_sph'],d['pf_dph'],d['notes']]
    for j,v in enumerate(vals):
        c=wc.cell(row=rr, column=1+j, value=v)
        if 8<=j+1<=13: c.number_format=MONEY0
        c.font=Font(size=9)
        if i%2: c.fill=PatternFill('solid', fgColor='FAFAFA')
wc.freeze_panes='D5'
wc.auto_filter.ref=f'A4:N{4+len(wood)}'
WOOD_LAST=4+len(wood)

# ============ FIBERGLASS CATALOG ============
fc=wb.create_sheet('Fiberglass Catalog')
title(fc,'Fiberglass Door Catalog  -  LIST PRICES','LIST price only. Sidelite prehung columns are per-each ADDERS, not standalone prices - see the Type column. To update: paste new list prices over columns H-M only.')
FCOLS=['Skin','Category','Part Number','Description','Glazing','Size','Type',
       'UF Slab','UF Single PH','UF Double PH','PF Slab','PF Single PH','PF Double PH']
header_row(fc, 4, FCOLS, [20,24,24,50,26,8,10,11,13,13,11,13,13])
for i,d in enumerate(fg):
    rr=5+i
    vals=[d['skin'],d['category'],d['part'],d['desc'],d['glazing'],d['size'],
          'Sidelite (adder)' if d['is_sidelite'] else 'Door',
          d['uf_slab'],d['uf_sph'],d['uf_dph'],d['pf_slab'],d['pf_sph'],d['pf_dph']]
    for j,v in enumerate(vals):
        c=fc.cell(row=rr, column=1+j, value=v)
        if 8<=j+1<=13: c.number_format=MONEY0
        c.font=Font(size=9)
        if d['is_sidelite']: c.fill=PatternFill('solid', fgColor='FFF7E6')
        elif i%2: c.fill=PatternFill('solid', fgColor='FAFAFA')
fc.freeze_panes='D5'
fc.auto_filter.ref=f'A4:M{4+len(fg)}'
FG_LAST=4+len(fg)

wb.save('wip.xlsx')
print(f'catalogs built: wood rows 5-{WOOD_LAST}, fg rows 5-{FG_LAST}')
print('MULT cell', MULT_CELL, '| MARGIN cell', MARGIN_CELL)

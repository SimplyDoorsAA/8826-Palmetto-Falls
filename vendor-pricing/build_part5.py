"""Finder, Export Block, Reference tabs."""
import json, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter

D=json.load(open('data.json')); wood, fg, ref = D['wood'], D['fg'], D['reference']
wb=openpyxl.load_workbook('wip.xlsx')
wc, fc, ls = wb['Wood Catalog'], wb['Fiberglass Catalog'], wb['Lists']
WL, FL = wc.max_row, fc.max_row
NAVY='1F3864'; STEEL='2E5C8A'; LIGHT='D9E2F3'; CREAM='FFF2CC'; GREY='F2F2F2'
H_FONT=Font(bold=True,color='FFFFFF',size=9); H_FILL=PatternFill('solid',fgColor=STEEL)
SUB=Font(bold=True,size=11,color=NAVY); TITLE=Font(bold=True,size=14,color=NAVY)
thin=Side(style='thin',color='BFBFBF'); BOX=Border(left=thin,right=thin,top=thin,bottom=thin)
IN_FILL=PatternFill('solid',fgColor=CREAM); MONEY='#,##0.00'

def uniq(vals): return sorted({str(v) for v in vals if v not in (None,'')})
lists={
 'H':('Species / Skin', uniq([d['species'] for d in wood])+uniq([d['skin'] for d in fg])),
 'I':('Category',       uniq([d['category'] for d in wood])+uniq([d['category'] for d in fg])),
 'J':('Size',           uniq([d['size'] for d in wood]+[d['size'] for d in fg])),
}
for col,(lab,vals) in lists.items():
    c=ls[f'{col}4']; c.value=lab; c.font=H_FONT; c.fill=H_FILL; c.border=BOX
    ls.column_dimensions[col].width=26
    for i,v in enumerate(vals): ls.cell(row=5+i,column=ord(col)-64,value=v).font=Font(size=9)
for col,(lab,vals),nm in zip('HIJ', lists.values(), ['SpeciesList','CategoryList','SizeList']):
    wb.defined_names.add(DefinedName(nm, attr_text=f'Lists!${col}$5:${col}${4+len(vals)}'))

# ---------- FINDER ----------
fd=wb.create_sheet('Finder')
wb.move_sheet('Finder', offset=-(wb.sheetnames.index('Finder')-1))
fd['A1']='Door Finder  -  browse the catalog by attribute'; fd['A1'].font=TITLE
fd['A2']=('Set any combination of filters below; leave one blank to mean "any". Matching doors list underneath with all six price points. '
          'Copy the part number you want into the Quote tab.')
fd['A2'].font=Font(italic=True,size=9,color='595959')
fd['A3']='NOTE: this tab uses FILTER and comes alive once the workbook is opened in Google Sheets.'
fd['A3'].font=Font(italic=True,size=9,color='9C5700')
picks=[('Material','B5','=MaterialList'),('Species / Skin','B6','=SpeciesList'),
       ('Category','B7','=CategoryList'),('Size','B8','=SizeList')]
for lab,cell,src in picks:
    r=int(cell[1:]); fd.cell(row=r,column=1,value=lab).font=Font(bold=True,size=10)
    c=fd[cell]; c.fill=IN_FILL; c.border=BOX
    dv=DataValidation(type='list', formula1=src, allow_blank=True, showDropDown=False)
    fd.add_data_validation(dv); dv.add(cell)
fd['B5']='Wood'
fd.column_dimensions['A'].width=18
for col,w in zip('BCDEFGHIJKLM',[22,26,44,26,9,11,11,13,13,11,13,13]): fd.column_dimensions[col].width=w

def cond(sheet, last, sp_col, cat_col):
    return (f"((LEN($B$6)=0)+('{sheet}'!${sp_col}$5:${sp_col}${last}=$B$6)>0)*"
            f"((LEN($B$7)=0)+('{sheet}'!${cat_col}$5:${cat_col}${last}=$B$7)>0)*"
            f"((LEN($B$8)=0)+('{sheet}'!$F$5:$F${last}=$B$8)>0)")
fd['A10']=(f'=IF($B$5="Wood",'
           f'IFERROR(FILTER(\'Wood Catalog\'!$A$5:$M${WL},{cond("Wood Catalog",WL,"A","B")}),"No matches - widen your filters"),'
           f'IFERROR(FILTER(\'Fiberglass Catalog\'!$A$5:$M${FL},{cond("Fiberglass Catalog",FL,"A","B")}),"No matches - widen your filters"))')
hdr=['Species/Skin','Category','Part Number','Description','Glazing','Size','Page/Type',
     'UF Slab','UF Single PH','UF Double PH','PF Slab','PF Single PH','PF Double PH']
for i,l in enumerate(hdr):
    c=fd.cell(row=9,column=1+i,value=l); c.font=H_FONT; c.fill=H_FILL; c.border=BOX
    c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
fd.row_dimensions[9].height=30

# ---------- EXPORT BLOCK ----------
ex=wb.create_sheet('Export Block')
ex['A1']='Export Block  -  paste straight into your proposal template'; ex['A1'].font=TITLE
ex['A2']='Sell prices only, no cost or margin. Blank rows are harmless - select the whole block and paste.'
ex['A2'].font=Font(italic=True,size=9,color='595959')
for i,(l,w) in enumerate([('Qty',8),('Description',74),('Unit Price',14),('Extended',14)]):
    c=ex.cell(row=4,column=1+i,value=l); c.font=H_FONT; c.fill=H_FILL; c.border=BOX
    ex.column_dimensions[get_column_letter(1+i)].width=w
r=5
for qr in range(12,27):
    ex.cell(row=r,column=1,value=f'=IF(Quote!$C{qr}="","",Quote!$I{qr})')
    ex.cell(row=r,column=2,value=(f'=IF(Quote!$C{qr}="","",Quote!$D{qr}'
        f'&IF(Quote!$F{qr}="",""," - "&Quote!$F{qr})'
        f'&" ("&Quote!$G{qr}&", "&Quote!$H{qr}&")"'
        f'&IF(N(Quote!$K{qr})=0,""," + "&Quote!$K{qr}&" sidelite(s)")'
        f'&IF(Quote!$L{qr}="Yes"," + transom",""))'))
    ex.cell(row=r,column=3,value=f'=IF(OR(Quote!$C{qr}="",N(Quote!$I{qr})=0),"",Quote!$W{qr}/Quote!$I{qr})')
    ex.cell(row=r,column=4,value=f'=IF(Quote!$C{qr}="","",Quote!$W{qr})')
    r+=1
for qr in range(30,40):
    ex.cell(row=r,column=1,value=f'=IF(Quote!$B{qr}="","",Quote!$E{qr})')
    ex.cell(row=r,column=2,value=f'=IF(Quote!$B{qr}="","",Quote!$C{qr}&IF(Quote!$D{qr}="",""," - "&Quote!$D{qr}))')
    ex.cell(row=r,column=3,value=f'=IF(OR(Quote!$B{qr}="",N(Quote!$E{qr})=0),"",Quote!$K{qr}/Quote!$E{qr})')
    ex.cell(row=r,column=4,value=f'=IF(Quote!$B{qr}="","",Quote!$K{qr})')
    r+=1
for rr in range(5,r):
    for col in range(1,5):
        c=ex.cell(row=rr,column=col); c.border=BOX; c.font=Font(size=9)
        if col in (3,4): c.number_format=MONEY
ex.cell(row=r+1,column=2,value='TOTAL').font=Font(bold=True,size=11)
c=ex.cell(row=r+1,column=4,value=f'=SUM(D5:D{r-1})'); c.number_format=MONEY
c.font=Font(bold=True,size=11,color=NAVY); c.fill=PatternFill('solid',fgColor='E2EFDA'); c.border=BOX

# ---------- REFERENCE ----------
rf=wb.create_sheet('Reference')
rf['A1']='Glass Codes, Families & Vendor Policy'; rf['A1'].font=TITLE
rf['A2']='Reference only - nothing here feeds the quote math.'; rf['A2'].font=Font(italic=True,size=9,color='595959')
for i,(l,w) in enumerate([('Code',9),('Glass Description',26),('Application / Notes',52),
                          ('Glass Family',15),('Family Style',44),('Common Door Pairings',34)]):
    c=rf.cell(row=4,column=1+i,value=l); c.font=H_FONT; c.fill=H_FILL; c.border=BOX
    c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    rf.column_dimensions[get_column_letter(1+i)].width=w
rf.row_dimensions[4].height=28
for i,row in enumerate(ref[1:]):
    for j in range(6):
        v=row[j] if j<len(row) else None
        if v in ('None',): v=None
        c=rf.cell(row=5+i,column=1+j,value=v); c.font=Font(size=9); c.border=BOX
        c.alignment=Alignment(wrap_text=True,vertical='top')
pr=5+len(ref[1:])+2
rf.cell(row=pr,column=1,value='Vendor Policy').font=SUB
rf.cell(row=pr,column=1).fill=PatternFill('solid',fgColor=LIGHT)
pol=[(row[6],row[7]) for row in ref[1:] if len(row)>7 and row[6] not in (None,'','None')]
for i,(k,v) in enumerate(pol):
    rf.cell(row=pr+1+i,column=1,value=k).font=Font(bold=True,size=9)
    c=rf.cell(row=pr+1+i,column=2,value=v); c.font=Font(size=9)
    c.alignment=Alignment(wrap_text=True,vertical='top')
    rf.merge_cells(start_row=pr+1+i,start_column=2,end_row=pr+1+i,end_column=6)
wb.save('wip.xlsx')
print('tabs:', wb.sheetnames)

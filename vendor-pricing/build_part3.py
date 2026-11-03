"""Defined names, freight self-test block, Lists tab."""
import json, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

M=json.load(open('meta.json')); ADD_LAST=M['ADD_LAST']
wb=openpyxl.load_workbook('wip.xlsx')
wc, fc, ad, fr = wb['Wood Catalog'], wb['Fiberglass Catalog'], wb['Adders'], wb['Freight Rules']
WOOD_LAST=wc.max_row; FG_LAST=fc.max_row
STEEL='2E5C8A'; LIGHT='D9E2F3'; NAVY='1F3864'; CREAM='FFF2CC'
H_FONT=Font(bold=True,color='FFFFFF',size=10); H_FILL=PatternFill('solid',fgColor=STEEL)
SUB=Font(bold=True,size=11,color=NAVY)
thin=Side(style='thin',color='BFBFBF'); BOX=Border(left=thin,right=thin,top=thin,bottom=thin)
MONEY='#,##0.00'; MONEY0='#,##0'

NAMES={
 'Multiplier':   'Config!$B$9',
 'DefaultMargin':'Config!$B$10',
 'WoodParts':   f"'Wood Catalog'!$C$5:$C${WOOD_LAST}",
 'WoodDesc':    f"'Wood Catalog'!$D$5:$D${WOOD_LAST}",
 'WoodGlaz':    f"'Wood Catalog'!$E$5:$E${WOOD_LAST}",
 'WoodPrices':  f"'Wood Catalog'!$H$5:$M${WOOD_LAST}",
 'FGParts':     f"'Fiberglass Catalog'!$C$5:$C${FG_LAST}",
 'FGDesc':      f"'Fiberglass Catalog'!$D$5:$D${FG_LAST}",
 'FGGlaz':      f"'Fiberglass Catalog'!$E$5:$E${FG_LAST}",
 'FGPrices':    f"'Fiberglass Catalog'!$H$5:$M${FG_LAST}",
 'AdderCodes':  f"Adders!$A$5:$A${ADD_LAST}",
 'AdderDesc':   f"Adders!$C$5:$C${ADD_LAST}",
 'AdderOpt':    f"Adders!$D$5:$D${ADD_LAST}",
 'AdderList':   f"Adders!$E$5:$E${ADD_LAST}",
 'AdderFrt':    f"Adders!$G$5:$G${ADD_LAST}",
 'TierMin':      "'Freight Rules'!$B$17:$B$20",
 'TierCharge':   "'Freight Rules'!$D$17:$D$20",
 'CrateCharge':  "'Freight Rules'!$B$23",
}
for n,ref in NAMES.items():
    wb.defined_names.add(DefinedName(n, attr_text=ref))

# ---- freight self-test: rebuild the vendor's five worked examples ----
fr['A26']='Self-Test  -  the vendor\'s own five worked examples, computed by these rules'
fr['A26'].font=SUB; fr['A26'].fill=PatternFill('solid',fgColor=LIGHT)
fr['A27']=('If any row below reads FAIL, the freight engine does not match how Hoelscher bills and the quote numbers '
           'cannot be trusted. These recompute live from the tables above.')
fr['A27'].font=Font(italic=True,size=9,color='595959')
labels=['Example','Scenario','Frt Units','Prehung Openings','Tier Charge','Crate','Computed','Vendor Says','Check']
for i,l in enumerate(labels):
    c=fr.cell(row=28,column=1+i,value=l); c.font=H_FONT; c.fill=H_FILL; c.border=BOX
    c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
fr.row_dimensions[28].height=30
for col,w in zip('GHI',[13,13,10]): fr.column_dimensions[col].width=w
EX=[('#1','1 single prehung unit',1,1,220),
    ('#2','1 single prehung + 1 double prehung',3,2,320),
    ('#3','5 double prehung units',10,5,500),
    ('#4','1 single w/ 2 sidelites + transom, plus 1 double prehung',5,2,365),
    ('#5','8 door slabs + 1 double prehung',10,1,100)]
for i,(lab,scen,units,openings,expect) in enumerate(EX):
    r=29+i
    fr.cell(row=r,column=1,value=lab).font=Font(size=9,bold=True)
    fr.cell(row=r,column=2,value=scen).font=Font(size=9)
    for col,v in ((3,units),(4,openings)):
        c=fr.cell(row=r,column=col,value=v); c.number_format='0.0'; c.font=Font(size=9); c.border=BOX
    c=fr.cell(row=r,column=5,value=f'=IFERROR(INDEX(TierCharge,MATCH(C{r},TierMin,1)),0)')
    c.number_format=MONEY0; c.font=Font(size=9)
    c=fr.cell(row=r,column=6,value=f'=D{r}*CrateCharge'); c.number_format=MONEY0; c.font=Font(size=9)
    c=fr.cell(row=r,column=7,value=f'=E{r}+F{r}'); c.number_format=MONEY0; c.font=Font(size=9,bold=True)
    c=fr.cell(row=r,column=8,value=expect); c.number_format=MONEY0; c.font=Font(size=9)
    c=fr.cell(row=r,column=9,value=f'=IF(G{r}=H{r},"PASS","FAIL")')
    c.font=Font(size=9,bold=True); c.alignment=Alignment(horizontal='center')
fr.conditional_formatting.add('I29:I33',
    CellIsRule(operator='equal', formula=['"PASS"'], fill=PatternFill('solid',fgColor='C6EFCE'), font=Font(color='006100',bold=True)))
fr.conditional_formatting.add('I29:I33',
    CellIsRule(operator='equal', formula=['"FAIL"'], fill=PatternFill('solid',fgColor='FFC7CE'), font=Font(color='9C0006',bold=True)))

# ---- LISTS (dropdown sources) ----
ls=wb.create_sheet('Lists')
ls['A1']='Dropdown Sources'; ls['A1'].font=Font(bold=True,size=12,color=NAVY)
ls['A2']='Feeds the Quote tab dropdowns. No need to edit by hand - it rebuilds from the catalogs.'
ls['A2'].font=Font(italic=True,size=9,color='595959')
cols=[('A','All Part Numbers',28),('B','Config',14),('C','Finish',14),('D','Yes/No',10),
      ('E','Material',14),('F','Sidelite Parts',24),('G','Adder Codes',24)]
for col,lab,w in cols:
    c=ls[f'{col}4']; c.value=lab; c.font=H_FONT; c.fill=H_FILL; c.border=BOX
    ls.column_dimensions[col].width=w
allparts=[wc.cell(row=r,column=3).value for r in range(5,WOOD_LAST+1)]
allparts+=[fc.cell(row=r,column=3).value for r in range(5,FG_LAST+1)]
allparts=[p for p in allparts if p]
for i,p in enumerate(allparts): ls.cell(row=5+i,column=1,value=p).font=Font(size=9)
for i,v in enumerate(['Slab','Single PH','Double PH']): ls.cell(row=5+i,column=2,value=v).font=Font(size=9)
for i,v in enumerate(['Unfinished','Prefinished']):     ls.cell(row=5+i,column=3,value=v).font=Font(size=9)
for i,v in enumerate(['No','Yes']):                     ls.cell(row=5+i,column=4,value=v).font=Font(size=9)
for i,v in enumerate(['Wood','Fiberglass']):            ls.cell(row=5+i,column=5,value=v).font=Font(size=9)
sl=[fc.cell(row=r,column=3).value for r in range(5,FG_LAST+1) if fc.cell(row=r,column=7).value=='Sidelite (adder)']
sl+=[wc.cell(row=r,column=3).value for r in range(5,WOOD_LAST+1) if 'Sidelite' in str(wc.cell(row=r,column=14).value)]
sl=[s for s in sl if s]
for i,p in enumerate(sl): ls.cell(row=5+i,column=6,value=p).font=Font(size=9)
codes=[ad.cell(row=r,column=1).value for r in range(5,ADD_LAST+1)]
for i,p in enumerate(codes): ls.cell(row=5+i,column=7,value=p).font=Font(size=9)

for n,ref in {'AllParts':f'Lists!$A$5:$A${4+len(allparts)}',
              'ConfigList':'Lists!$B$5:$B$7','FinishList':'Lists!$C$5:$C$6',
              'YesNo':'Lists!$D$5:$D$6','MaterialList':'Lists!$E$5:$E$6',
              'SideliteList':f'Lists!$F$5:$F${4+len(sl)}',
              'AdderCodeList':f'Lists!$G$5:$G${4+len(codes)}'}.items():
    wb.defined_names.add(DefinedName(n, attr_text=ref))
wb.save('wip.xlsx')
print(f'names={len(wb.defined_names)}  parts={len(allparts)}  sidelites={len(sl)}  adders={len(codes)}')

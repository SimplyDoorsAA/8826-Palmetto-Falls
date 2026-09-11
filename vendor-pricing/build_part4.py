"""The Quote tab."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule, FormulaRule

wb=openpyxl.load_workbook('wip.xlsx')
q=wb.create_sheet('Quote')
wb.move_sheet('Quote', offset=-(len(wb.sheetnames)-1))   # make it the first tab

NAVY='1F3864'; STEEL='2E5C8A'; LIGHT='D9E2F3'; CREAM='FFF2CC'; GREY='F2F2F2'; GREEN='E2EFDA'
H_FONT=Font(bold=True,color='FFFFFF',size=9); H_FILL=PatternFill('solid',fgColor=STEEL)
SUB=Font(bold=True,size=11,color=NAVY)
thin=Side(style='thin',color='BFBFBF'); BOX=Border(left=thin,right=thin,top=thin,bottom=thin)
IN_FILL=PatternFill('solid',fgColor=CREAM); CALC_FILL=PatternFill('solid',fgColor=GREY)
MONEY='#,##0.00'; MONEY0='#,##0'; PCT='0.0%'
F1=12; FN=26          # door lines
A1=30; AN=39          # adder lines

q['A1']='HOELSCHER  -  QUOTE BUILDER'; q['A1'].font=Font(bold=True,size=16,color=NAVY)
q['A2']='Shaded gold = you type it.  Grey = calculated, do not overwrite.  Freight is inbound landed cost and is folded into door cost before margin.'
q['A2'].font=Font(italic=True,size=9,color='595959')

# ---- quote header ----
q['A4']='QUOTE HEADER'; q['A4'].font=SUB; q['A4'].fill=PatternFill('solid',fgColor=LIGHT)
hdrs=[('A5','Job / Project',''),('A6','Customer',''),('A7','Quote Date',''),('A8','Target Margin','MARGIN')]
for cell,label,kind in hdrs:
    r=int(cell[1:]); q.cell(row=r,column=1,value=label).font=Font(bold=True,size=10)
    c=q.cell(row=r,column=2); c.fill=IN_FILL; c.border=BOX
    if kind=='MARGIN':
        c.value='=DefaultMargin'; c.number_format=PCT
        q.cell(row=r,column=3,value='Override per quote. Defaults to the Config value.').font=Font(size=9,italic=True,color='595959')

# ---- door lines ----
q['A10']='DOOR & UNIT LINES'; q['A10'].font=SUB; q['A10'].fill=PatternFill('solid',fgColor=LIGHT)
DCOLS=[('#',5),('Material',11),('Part Number',22),('Description',40),('Glazing Options',26),('Glass Choice',16),
       ('Config',12),('Finish',12),('Qty',6),('Sidelite Part',20),('SL Qty',7),('Transom',9),('Prehang Adder',22),
       ('Unit List',11),('SL Adder Ea',11),('Prehang $',10),('Ext List',12),('Net Cost',12),
       ('Frt Units',9),('Crate $',10),('Alloc Frt',10),('Landed Cost',13),('Sell Price',13)]
for i,(lab,w) in enumerate(DCOLS):
    c=q.cell(row=11,column=1+i,value=lab); c.font=H_FONT; c.fill=H_FILL; c.border=BOX
    c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    q.column_dimensions[get_column_letter(1+i)].width=w
q.row_dimensions[11].height=32

CFGIDX='IF($H{r}="Prefinished",3,0)+MATCH($G{r},{{"Slab","Single PH","Double PH"}},0)'
for r in range(F1,FN+1):
    idx=CFGIDX.format(r=r)
    q.cell(row=r,column=1,value=r-F1+1).font=Font(size=9,color='808080')
    q.cell(row=r,column=2,value=f'=IF($C{r}="","",IF(ISNUMBER(MATCH($C{r},WoodParts,0)),"Wood",'
                               f'IF(ISNUMBER(MATCH($C{r},FGParts,0)),"Fiberglass","NOT FOUND")))')
    q.cell(row=r,column=4,value=f'=IFERROR(INDEX(WoodDesc,MATCH($C{r},WoodParts,0)),IFERROR(INDEX(FGDesc,MATCH($C{r},FGParts,0)),""))')
    q.cell(row=r,column=5,value=f'=IFERROR(INDEX(WoodGlaz,MATCH($C{r},WoodParts,0)),IFERROR(INDEX(FGGlaz,MATCH($C{r},FGParts,0)),""))')
    q.cell(row=r,column=14,value=f'=IFERROR(INDEX(WoodPrices,MATCH($C{r},WoodParts,0),{idx}),IFERROR(INDEX(FGPrices,MATCH($C{r},FGParts,0),{idx}),0))')
    q.cell(row=r,column=15,value=f'=IF($J{r}="",0,IFERROR(INDEX(WoodPrices,MATCH($J{r},WoodParts,0),{idx}),IFERROR(INDEX(FGPrices,MATCH($J{r},FGParts,0),{idx}),0)))')
    q.cell(row=r,column=16,value=f'=IF($M{r}="",0,IFERROR(INDEX(AdderList,MATCH($M{r},AdderCodes,0)),0))')
    q.cell(row=r,column=17,value=f'=IF($C{r}="",0,($N{r}+$O{r}*N($K{r})+$P{r})*N($I{r}))')
    q.cell(row=r,column=18,value=f'=$Q{r}*Multiplier')
    q.cell(row=r,column=19,value=f'=IF($C{r}="",0,(IF($G{r}="Double PH",2,1)+0.5*N($K{r})+IF($L{r}="Yes",1,0))*N($I{r}))')
    q.cell(row=r,column=20,value=f'=IF(OR($C{r}="",$G{r}="Slab",$G{r}=""),0,CrateCharge*N($I{r}))')
    q.cell(row=r,column=21,value=f'=IF($B$43=0,0,$B$45*$S{r}/$B$43)')
    q.cell(row=r,column=22,value=f'=$R{r}+$T{r}+$U{r}')
    q.cell(row=r,column=23,value=f'=IF($V{r}=0,"",$V{r}/(1-$B$8))')
    for col in range(1,24):
        c=q.cell(row=r,column=col); c.border=BOX; c.font=Font(size=9)
        if col in (3,6,7,8,9,10,11,12,13): c.fill=IN_FILL
        elif col in (1,): pass
        else: c.fill=CALC_FILL
        if col in (14,15,16,17,18,20,21,22,23): c.number_format=MONEY0 if col in (14,15,16) else MONEY
        if col==19: c.number_format='0.0'
        if col==23: c.font=Font(size=9,bold=True,color='1F3864')

# ---- adder lines ----
q['A28']='ADDERS, HARDWARE, JAMBS & COMPONENTS'; q['A28'].font=SUB; q['A28'].fill=PatternFill('solid',fgColor=LIGHT)
ACOLS=['#','Adder Code','Description','Option / Finish','Qty','List Ea','Ext List','Frt Units','Net Cost','Landed Cost','Sell Price']
for i,lab in enumerate(ACOLS):
    c=q.cell(row=29,column=1+i,value=lab); c.font=H_FONT; c.fill=H_FILL; c.border=BOX
    c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
q.row_dimensions[29].height=28
for r in range(A1,AN+1):
    q.cell(row=r,column=1,value=r-A1+1).font=Font(size=9,color='808080')
    q.cell(row=r,column=3,value=f'=IFERROR(INDEX(AdderDesc,MATCH($B{r},AdderCodes,0)),"")')
    q.cell(row=r,column=4,value=f'=IFERROR(INDEX(AdderOpt,MATCH($B{r},AdderCodes,0)),"")')
    q.cell(row=r,column=6,value=f'=IF($B{r}="",0,IFERROR(INDEX(AdderList,MATCH($B{r},AdderCodes,0)),0))')
    q.cell(row=r,column=7,value=f'=IF($B{r}="",0,$F{r}*N($E{r}))')
    q.cell(row=r,column=8,value=f'=IF($B{r}="",0,IFERROR(INDEX(AdderFrt,MATCH($B{r},AdderCodes,0)),0)*N($E{r}))')
    q.cell(row=r,column=9,value=f'=$G{r}*Multiplier')
    q.cell(row=r,column=10,value=f'=$I{r}+IF($B$43=0,0,$B$45*$H{r}/$B$43)')
    q.cell(row=r,column=11,value=f'=IF($J{r}=0,"",$J{r}/(1-$B$8))')
    for col in range(1,12):
        c=q.cell(row=r,column=col); c.border=BOX; c.font=Font(size=9)
        if col in (2,5): c.fill=IN_FILL
        elif col!=1: c.fill=CALC_FILL
        if col in (6,7,9,10,11): c.number_format=MONEY
        if col==8: c.number_format='0.0000'
        if col==11: c.font=Font(size=9,bold=True,color='1F3864')

# ---- freight block ----
q['A41']='FREIGHT  -  inbound to your warehouse (landed cost, not a customer line)'
q['A41'].font=SUB; q['A41'].fill=PatternFill('solid',fgColor=LIGHT)
FRB=[('Total freight units on this shipment', f'=SUM($S${F1}:$S${FN})+SUM($H${A1}:$H${AN})','0.0'),
     ('Prehung openings (crate charge applies)', f'=SUMPRODUCT(($G${F1}:$G${FN}<>"Slab")*($G${F1}:$G${FN}<>"")*N($I${F1}:$I${FN}))','0'),
     ('Tier charge', '=IFERROR(INDEX(TierCharge,MATCH($B$43,TierMin,1)),0)',MONEY),
     ('Crate charge total', '=$B$44*CrateCharge',MONEY),
     ('TOTAL FREIGHT', '=$B$45+$B$46',MONEY)]
for i,(lab,f,fmt) in enumerate(FRB):
    r=43+i
    q.cell(row=r,column=1,value=lab).font=Font(bold=(i==4),size=10)
    c=q.cell(row=r,column=2,value=f); c.number_format=fmt; c.border=BOX; c.fill=CALC_FILL
    c.font=Font(bold=(i==4),size=10)
q.cell(row=43,column=3,value='Doors + sidelites + transoms + (jamb legs/25) + (components/50)').font=Font(size=9,italic=True,color='595959')
q.cell(row=45,column=3,value='1-3 units $120  |  4-6 $165  |  7-9 $200  |  10+ prepaid').font=Font(size=9,italic=True,color='595959')
q.cell(row=46,column=3,value='$100 per prehung opening. Slabs do not crate.').font=Font(size=9,italic=True,color='595959')

# ---- totals ----
q['A49']='QUOTE TOTALS'; q['A49'].font=SUB; q['A49'].fill=PatternFill('solid',fgColor=LIGHT)
TOT=[('Total List Price', f'=SUM($Q${F1}:$Q${FN})+SUM($G${A1}:$G${AN})',MONEY,False),
     ('Total Dealer NET (List x multiplier)', f'=SUM($R${F1}:$R${FN})+SUM($I${A1}:$I${AN})',MONEY,False),
     ('Inbound Freight', '=$B$47',MONEY,False),
     ('TOTAL LANDED COST', '=$B$51+$B$52',MONEY,True),
     ('Target Margin', '=$B$8',PCT,False),
     ('CUSTOMER SELL PRICE', '=IF($B$53=0,0,$B$53/(1-$B$54))',MONEY,True),
     ('Gross Profit $', '=$B$55-$B$53',MONEY,False),
     ('Effective Margin %', '=IF($B$55=0,0,$B$56/$B$55)',PCT,False)]
for i,(lab,f,fmt,bold) in enumerate(TOT):
    r=51+i
    q.cell(row=r,column=1,value=lab).font=Font(bold=bold,size=11 if bold else 10)
    c=q.cell(row=r,column=2,value=f); c.number_format=fmt; c.border=BOX
    c.font=Font(bold=bold,size=12 if bold else 10,color=NAVY if bold else '000000')
    c.fill=PatternFill('solid',fgColor=GREEN) if bold else CALC_FILL
q.column_dimensions['A'].width=38

# ---- validation ----
DVS=[('=AllParts',      f'C{F1}:C{FN}'), ('=ConfigList', f'G{F1}:G{FN}'),
     ('=FinishList',    f'H{F1}:H{FN}'), ('=SideliteList', f'J{F1}:J{FN}'),
     ('=YesNo',         f'L{F1}:L{FN}'), ('=AdderCodeList', f'M{F1}:M{FN}'),
     ('=AdderCodeList', f'B{A1}:B{AN}')]
for formula, rng in DVS:
    dv=DataValidation(type='list', formula1=formula, allow_blank=True, showDropDown=False)
    q.add_data_validation(dv); dv.add(rng)

q.conditional_formatting.add(f'B{F1}:B{FN}',
    CellIsRule(operator='equal', formula=['"NOT FOUND"'],
               fill=PatternFill('solid',fgColor='FFC7CE'), font=Font(color='9C0006',bold=True)))
q.conditional_formatting.add(f'G{F1}:G{FN}',
    FormulaRule(formula=[f'AND($C{F1}<>"",$G{F1}="")'], fill=PatternFill('solid',fgColor='FFC7CE')))
q.conditional_formatting.add(f'H{F1}:H{FN}',
    FormulaRule(formula=[f'AND($C{F1}<>"",$H{F1}="")'], fill=PatternFill('solid',fgColor='FFC7CE')))
q.conditional_formatting.add(f'I{F1}:I{FN}',
    FormulaRule(formula=[f'AND($C{F1}<>"",N($I{F1})=0)'], fill=PatternFill('solid',fgColor='FFC7CE')))
q.freeze_panes='D12'
wb.save('wip.xlsx')
print('Quote tab built. Door lines %d-%d, adder lines %d-%d' % (F1,FN,A1,AN))

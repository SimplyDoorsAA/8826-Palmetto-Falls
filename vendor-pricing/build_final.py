import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
wb=openpyxl.load_workbook('wip.xlsx')
NAVY='1F3864'; STEEL='2E5C8A'; LIGHT='D9E2F3'; CREAM='FFF2CC'
sh=wb.create_sheet('Start Here')
wb.move_sheet('Start Here', offset=-(wb.sheetnames.index('Start Here')))
sh.column_dimensions['A'].width=26; sh.column_dimensions['B'].width=104
sh['A1']='Hoelscher Quoting Workbook'; sh['A1'].font=Font(bold=True,size=18,color=NAVY)
sh['A2']='Wood + Fiberglass  |  price sheets effective 7-1-2026  |  FOB Run, San Antonio / Austin'
sh['A2'].font=Font(italic=True,size=10,color='595959')
BLOCKS=[
 ('HOW TO QUOTE',[
  ('1. Finder tab','Filter by material, species/skin, category and size to find the door. Leave a filter blank to mean "any". Copy the part number you want.'),
  ('2. Quote tab','Paste the part number into a door line, then set config, finish and quantity. Sidelites, transom and prehang adder are optional on the same line.'),
  ('3. Adders section','Hardware, jambs, casing, SDL bars and PVC components go in the lower block on the Quote tab, by adder code.'),
  ('4. Export Block tab','Sell prices only, formatted to paste straight into your proposal template. Cost and margin never appear there.'),
 ]),
 ('HOW PRICING WORKS',[
  ('Cost chain','List price  ->  x 0.48 multiplier  =  dealer net  ->  + allocated inbound freight  =  landed cost  ->  / (1 - margin)  =  sell price.'),
  ('Freight','Freight is inbound to your warehouse, so it is landed cost, not a customer line. The tier charge is allocated across lines by freight unit; the $100 crate charge attaches to the opening that caused it.'),
  ('Margin','Config holds your default. The Quote tab margin cell overrides it per job without touching Config.'),
 ]),
 ('COLOUR KEY',[
  ('Gold cells','You type these.'),
  ('Grey cells','Calculated. Overwriting one breaks the quote.'),
  ('Red cells','Something is wrong - a missing entry, an unknown part number, or a combination Hoelscher does not sell.'),
 ]),
 ('WATCH FOR',[
  ('COMBO N/A','116 of 222 fiberglass doors have no prefinished pricing. Pick Prefinished on one of those and the Material column turns red rather than quietly quoting $0.'),
  ('NOT FOUND','The part number is not in either catalog. That line is excluded from totals, freight and the export.'),
  ('Sidelite prices','Sidelite prehung figures are per-each ADDERS, not standalone prices. The workbook treats them that way.'),
  ('Open vendor questions','Three items on the Config tab are flagged CONFIRM. Worth a call to Hoelscher before quoting at volume.'),
 ]),
 ('UPDATING PRICES',[
  ('When Hoelscher revises','Paste new list prices over columns H-M on the catalog tabs. Nothing else needs to change - net cost, freight and margin all recalculate.'),
  ('Never paste net costs','This workbook stores list price only and derives net from the Config multiplier. That is deliberate - it is one cell to change instead of 1,500.'),
 ]),
 ('VALIDATION',[
  ('Freight self-test','The Freight Rules tab rebuilds the vendor\'s own five worked examples. All five currently read PASS. If one ever reads FAIL, the freight math no longer matches how Hoelscher bills.'),
  ('Price reconciliation','Every price was checked against the source documents: 276 wood rows and 212 fiberglass rows, six price points each. Zero mismatches.'),
 ]),
]
r=4
for title, items in BLOCKS:
    c=sh.cell(row=r,column=1,value=title); c.font=Font(bold=True,size=11,color=NAVY)
    c.fill=PatternFill('solid',fgColor=LIGHT)
    sh.cell(row=r,column=2).fill=PatternFill('solid',fgColor=LIGHT)
    r+=1
    for k,v in items:
        sh.cell(row=r,column=1,value=k).font=Font(bold=True,size=10)
        sh.cell(row=r,column=1).alignment=Alignment(vertical='top')
        c=sh.cell(row=r,column=2,value=v); c.font=Font(size=10)
        c.alignment=Alignment(wrap_text=True,vertical='top')
        sh.row_dimensions[r].height=None
        r+=1
    r+=1
for ws in wb.worksheets:
    ws.sheet_view.showGridLines=False

out='Hoelscher_Quoting_Workbook_2026.xlsx'
wb.save(out)
print('saved', out)
print('tabs:', wb.sheetnames)

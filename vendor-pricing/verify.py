import formulas, warnings
warnings.filterwarnings('ignore')
xl = formulas.ExcelModel().loads('test_nofinder.xlsx').finish()
sol = xl.calculate()
LOOKUP = {}
for k, v in sol.items():
    ku = k.upper()
    if ']' in ku and "'!" in ku:
        sheet = ku.split(']', 1)[1].split("'!")[0]
        cell  = ku.split("'!", 1)[1]
        try: val = v.value[0, 0]
        except Exception:
            try: val = v.value
            except Exception: val = v
        LOOKUP[(sheet, cell)] = val
def g(sheet, cell):
    return LOOKUP.get((sheet.upper(), cell.upper()), '<missing>')
def show(sheet, cells):
    for c, lab in cells:
        print('  %-5s %-26s = %s' % (c, lab, g(sheet, c)))

print('=== FREIGHT SELF-TEST : the vendor\'s own five worked examples ===')
ok = True
for r in range(29, 34):
    comp, vend, chk = g('FREIGHT RULES','G%d'%r), g('FREIGHT RULES','H%d'%r), g('FREIGHT RULES','I%d'%r)
    print('  Example %s  units=%-5s openings=%-4s computed=%-8s vendor=%-8s  %s' % (
        g('FREIGHT RULES','A%d'%r), g('FREIGHT RULES','C%d'%r), g('FREIGHT RULES','D%d'%r), comp, vend, chk))
    if str(chk) != 'PASS': ok = False
print('  ALL FIVE PASS' if ok else '  *** FAILURE ***')
print()
print('=== QUOTE LINE 1 : M1LV-3080, Single PH, Unfinished, +2 sidelites, +transom ===')
show('QUOTE', [('B12','material (derived)'),('D12','description'),('N12','unit list'),
               ('O12','sidelite adder ea'),('Q12','ext list'),('R12','net cost'),
               ('S12','freight units'),('T12','crate $'),('U12','alloc freight'),
               ('V12','landed cost'),('W12','sell price')])
print('=== QUOTE LINE 2 : M3GPW-3080, Double PH, Unfinished ===')
show('QUOTE', [('B13','material'),('N13','unit list'),('Q13','ext list'),('R13','net'),
               ('S13','freight units'),('T13','crate $'),('V13','landed')])
print('=== ADDER LINE : BDH001-MB x2 ===')
show('QUOTE', [('C30','description'),('D30','option'),('F30','list ea'),('G30','ext list'),
               ('H30','freight units'),('I30','net'),('K30','sell')])
print('=== FREIGHT BLOCK ===')
show('QUOTE', [('B43','total freight units'),('B44','prehung openings'),('B45','tier charge'),
               ('B46','crate total'),('B47','TOTAL FREIGHT')])
print('=== TOTALS ===')
show('QUOTE', [('B51','total list'),('B52','total dealer net'),('B53','inbound freight'),
               ('B54','TOTAL LANDED'),('B55','target margin'),('B56','CUSTOMER SELL'),
               ('B57','gross profit'),('B58','effective margin')])
print('=== EXPORT BLOCK (first rows) ===')
for r in range(5, 9):
    print('  qty=%-6s %-58s unit=%-10s ext=%s' % (
        g('EXPORT BLOCK','A%d'%r), str(g('EXPORT BLOCK','B%d'%r))[:58],
        g('EXPORT BLOCK','C%d'%r), g('EXPORT BLOCK','D%d'%r)))

print('=== GUARD CASES ===')
for r, what in [(14,'smooth-skin door + Prefinished (vendor offers none)'), (15,'bogus part number')]:
    print('  row %d  %-52s -> B=%-12s N=%s Q=%s' % (r, what, g('QUOTE','B%d'%r), g('QUOTE','N%d'%r), g('QUOTE','Q%d'%r)))

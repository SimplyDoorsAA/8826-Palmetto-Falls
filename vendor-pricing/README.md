# Hoelscher Quoting Workbook

Interactive quoting workbook for Hoelscher Weatherstrip Mfg. Co. (Hoelscher Doors),
covering both the wood and fiberglass lines. Built for Google Sheets.

**Deliverable:** `Hoelscher_Quoting_Workbook_2026.xlsx` — upload to Google Drive and
open as a Sheet.

## Tabs

| Tab | Purpose |
|---|---|
| Start Here | How to quote, colour key, what to watch for |
| Quote | 15 configured door lines + 10 adder lines, live totals |
| Finder | Browse the catalog by material / species / category / size |
| Config | Multiplier, default margin, price sheet dates, open vendor questions |
| Wood Catalog | 276 SKUs, list price only |
| Fiberglass Catalog | 222 SKUs, list price only |
| Adders | 146 rows — hardware, jambs, casings, SDL bars, PVC, prehang adders |
| Freight Rules | Unit equivalents, tier table, and a live self-test |
| Lists | Dropdown sources |
| Export Block | Sell prices only, formatted to paste into a proposal template |
| Reference | Glass codes and vendor policy |

## Pricing model

```
List price
  x 0.48 multiplier        -> dealer net
  + allocated inbound freight -> landed cost
  / (1 - margin)           -> sell price
```

Freight is inbound to the warehouse, so it is landed cost rather than a customer
line. The shipment tier charge is allocated across lines by freight unit; the $100
crate charge attaches directly to the prehung opening that caused it.

### Freight rules (FOB Run — San Antonio / Austin)

Units: slab 1 · single prehung 1 · double prehung 2 · sidelite 0.5 · transom 1 ·
25 jamb legs 1 · 50 other components 1.

Tiers: 1–3 units $120 · 4–6 $165 · 7–9 $200 · 10+ prepaid. Plus $100 per prehung
opening (slabs do not crate).

## Design decisions

- **List price only.** The source workbook carried 1,500 hardcoded net-cost cells at
  List x 0.48. Net is now derived from a single Config cell.
- **No hand-counted column offsets.** Price selection uses `MATCH` on the config and
  finish rather than arithmetic on column indices, so inserting a column cannot
  silently reprice a quote.
- **Invalid lines contribute nothing.** A bad part number or a finish the vendor does
  not offer is flagged red and excluded from totals, freight and the export, rather
  than quoting at $0.
- **Sidelites are adders.** Sidelite prehung figures are per-each adders, not
  standalone unit prices, and are handled as such.

## Validation

- `reconcile.py` — checks every price against the source documents.
  276 wood rows and 212 fiberglass rows, six price points each. Zero mismatches.
- `verify.py` — evaluates the workbook formulas and rebuilds the vendor's own five
  worked freight examples. All five pass ($220, $320, $500, $365, $100).
- The Freight Rules tab carries the same self-test live, so it re-runs whenever the
  tier table is edited.

## Open questions for the vendor

Flagged `CONFIRM` on the Config tab:

1. The shipping sheet's "$100 NET PER UNIT + ABOVE SHIPPING CHARGES IF 10 DOOR
   MINIMUM" contradicts its own Examples 1 and 2. Encoded per the worked examples.
2. `NLSMFL3080NB` and `FGSHCMLEF3680NP` have no description on the vendor PDF.
3. Confirm a sidelite is never ordered prehung on its own.

## Rebuilding

```
python3 extract_all.py      # parse both source documents -> data.json
python3 build_part1.py      # config + catalogs
python3 build_part2.py      # adders + freight rules
python3 build_part3.py      # defined names, freight self-test, lists
python3 build_part4.py      # quote engine
python3 build_part5.py      # finder, export block, reference
python3 build_final.py      # start-here tab -> final workbook
python3 reconcile.py        # price reconciliation
python3 verify.py           # formula evaluation + freight self-test
```

## Not yet built

Per-row cascading dropdowns, an "add line item" button and one-click quote export
need Apps Script. The pricing and freight math is complete and independent of that.

# Borpathar Dashboard — Project State & Handoff

Read this first when picking up the project in a new chat. It captures the
architecture, the hard rules, where things stand, and the analysis in flight.
Operator: **Manu (Jaibir Singh)** — Borpathar Tea Company LLP, a bought-leaf tea
factory (BLF) in Assam. Single operator, MacBook Pro 14" M1.

---

## What this is

A master-dashboard system for the factory. A Python pipeline parses dropped
report files and writes a clean `data.xlsx`; Manu's own Excel master
(`Borpathar_Master File.xlsx`, lives on his Mac, NOT in the repo) pulls from it
via Power Query and holds all his charts/layout.

**The split (sacred):** code owns `data.xlsx` (data only, no charts). Manu owns
the master (presentation). The pipeline never touches his master.

---

## Architecture

```
feed folders  →  scripts/build_data.py  →  data.xlsx (tidy Excel Tables)
                 (run via RUN.command)        ↓  Power Query (refresh on open)
                                          Manu's Excel master  →  charts/tiles
```

`data.xlsx` tabs (each a named Excel Table): `KPIs`, `Mandays`, `Electricity`,
`GreenLeaf`, `Auction`, `Production`, `Expenses`, `Economics`, `YoY_Expenses`,
`Timeline`.

Parsers in `scripts/`: `mfg_parser`, `advice_parser`, `green_leaf_parser`,
`tally_parser`, `mandays_parser`, `electricity_parser`, `private_advice_parser`.

### Feed folders (contents gitignored; `.gitkeep` keeps the structure)
| folder | holds | cadence |
|---|---|---|
| `mfg_reports/` | daily MFG `.xls` (one tab: MFG) | daily |
| `sale_advices/` | Parcon auction advice PDFs | weekly |
| `green_leaf_prices/` | green-leaf rate file(s) | fortnightly |
| `tally_reports/` | Tally expense ledger (PDF/CSV) | monthly |
| `mandays/` | wages workbook (Mandays sheet) | fortnightly |
| `electricity_bills/` | APDCL bill PDFs | monthly |
| `private_sale_advices/` | per-invoice GST `.xlsx` (HUL/Tata/direct), read recursively | occasional |

---

## Core formulas & constants (`build_data.py`)

- `RECOVERY = 0.22` (made tea ÷ green leaf; "out-turn")
- `MULT = 4.545`, `CONV_COST = 45.0` (fallback; real conversion from Tally)
- `made_cost = avg_leaf_rate × 4.545 + conversion_cost`
- `marginal_floor = avg_leaf_rate × 4.545 + marginal_conversion` (variable + ½ semi)
- `full_cost_spread = auction_realization − made_cost` (the crush spread)
- **Total sold = auction (advices) + private (invoices)** — they add, no overlap.
- `closing_stock = tea_made_todate − total_sold` (live, every run)
- Primary vs secondary: grade ending in `1` (BP1, D1, BOPSM1) = secondary; else
  primary. Realization blends auction lots **and** private invoice lines →
  all-channel.

---

## CONVENTIONS — read before editing anything

1. **Append-only KPIs/Economics. NEVER insert or reorder rows.** Manu wires his
   dashboard tiles with direct `=` cell references (position-based, e.g.
   `=KPIs!B10`). Any row shift silently breaks his tiles. New metrics go at the
   **end** (before `generated` in KPIs). `avg_primary_rate`/`avg_secondary_rate`
   are frozen at KPIs rows 10–11 for this reason — do not move them.
2. **Confirm before every push.** Always confirm the approach with Manu before
   pushing to GitHub. Never push unprompted. (He runs `git pull` to receive.)
3. **Business data stays local.** All feed-folder contents + `data.xlsx` are
   gitignored. Never commit his data.
4. **Code owns data; Manu owns the master.** Never edit or overwrite his Excel
   master. Header/label cosmetics live in his master, not in the data tabs
   (renaming source headers breaks his Power Query refresh).
5. **Git attribution:** end commits with the Co-Authored-By + Claude-Session
   lines the current session provides.

---

## Operator workflow

New data: drop file(s) in the right folder → close `data.xlsx` → double-click
`RUN.command` → open master → it refreshes on open (or Data → Refresh All).
Code change: Claude pushes → Manu `git pull` → RUN → refresh.
(macOS Gatekeeper on first RUN: `xattr -dr com.apple.quarantine .` in the folder.)

---

## Current state (shipped — latest commit before this doc: `9b93769`)

- Private-sales feed live; stock computes across all channels every run.
- `Timeline` tab: cost vs realization on one weekly auction grid (crush-spread chart).
- Primary/secondary realization now all-channel (auction + private blended).
- KPI row order frozen.

---

## IN PROGRESS: Premium-over-market analysis ("is there alpha in how I buy leaf?")

**Question:** when Manu pays up for higher-fine% green leaf, does he actually
earn it back at auction? And with volume as a second lever, where's the
quality-vs-throughput sweet spot?

**Method:**
1. **Premium over market**, per sale, per grade = *Manu's grade price − the BLF
   market average for that grade, that sale.* Strips out market movement AND
   estate-tier bias. Read the **movement** (does premium rise when he buys better
   leaf), not the absolute level.
2. **Benchmark source:** Parcon broker website → REPORTS → **GRADE GAP**, filtered
   **GUW + BLF** (peer bought-leaf factories, not estates), per sale, all grades.
   Manu is pulling sales **1–42** and will Export.
3. Correlate premium (lagged by the leaf→auction production cycle) against leaf
   fine%/rate from the green-leaf file.
4. **Volume layer:** plot each fortnight — leaf rate (x) vs crush spread (y),
   dot size = volume made. Look for the efficient frontier → "could I have
   maximised it, and how."

**Green-leaf data note:** `GREEN_LEAF_RATE_2026.xlsx` has two parts — the
`RATE -2026` sheet holds fine%-band **rate cards** (base rate per band, Made Tea =
base×4.55, ₹50 mfg, COP); the per-fortnight **period sheets** hold **actuals**
(KG + amount + weighted-avg rate per band). Leaf is priced by fine% band and the
weighted-avg rate swings hard fortnight to fortnight (~₹21.9–₹32.2). If this file
becomes the source of truth, `green_leaf_parser` needs rework to read it.

---

## Open items

- **Grade-name mapping:** Manu's grades (BP, BOP, BOPSM, PF, PD, DUST, BP1, D1…)
  ↔ broker's labels (e.g. broker **BOP** = Manu's **BP**?). Confirm before matching.
- **Confirm grade-gap export granularity:** must be per-sale (a row per sale),
  not one aggregate for the 1–42 range.
- **Leaf→auction lag:** need Manu's estimate (weeks from leaf-in to lot selling).
- **50-invoice parse safeguard (not built):** add a per-run report
  (`private: N/N invoices parsed, X kg`, loudly list any skipped) + dedup by
  invoice number, so a silently-skipped invoice can't under-count sold.
- If premium analysis carries signal: wire the GRADE GAP report in as a weekly feed.

---

## Resume in a new chat

1. `add_repo jaibirthesingh/borpathar_dashboard`, then read this file.
2. Data lives on Manu's Mac (gitignored) — he attaches files or works via the
   desktop bridge. The repo has only code + empty feed folders.
3. Next concrete step: Manu sends the Parcon GRADE GAP export (sales 1–42,
   GUW+BLF, all grades) → build premium-per-sale + correlation vs leaf.

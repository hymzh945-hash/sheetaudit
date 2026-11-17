# SheetAudit

**Find what breaks a spreadsheet before someone else does.**

A small, offline audit for Excel workbooks: it reads formulas, not values, and reports the
structural problems that quietly wreck models — circular references, error values, external
workbook links, volatile functions, hardcoded constants and formula drift inside a column.

```
sheetaudit audit -i budget.xlsx --report audit.html
```

```
SheetAudit 0.1.0 · budget.xlsx
  score      46/100  Risky — fix the high findings before relying on it
  findings   17 (2 high, 6 medium, 9 low)
  sheets     3 · formulas 812
  - Error values                       1  [high]
  - Circular references                1  [high]
  - External workbook links            1  [medium]
  - Formula drift in a column          1  [medium]
  - Hardcoded constants in formulas    6  [low]
```

## What it checks

| Check | Severity | Why it matters |
|---|---|---|
| Error values (`#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`, `#N/A`) | high | The model already broke; the cell just hasn't been looked at |
| Circular references | high | Excel stops calculating, or silently iterates to a wrong number |
| External workbook links | medium | The file is only correct on one machine, with one other file present |
| Division by zero | medium | Passes review today, explodes the first time a driver goes to zero |
| Volatile functions (`INDIRECT`, `OFFSET`, `TODAY`, `RAND`, `CELL`…) | medium | Recalc storms in real workbooks; often a sign of a hack |
| Formula drift in a column | medium | One row divides by 1.08 while its siblings use 1.05 — the classic copy-paste bug |
| Hidden sheets with live formulas | medium | Calculations nobody reviewing the model can see |
| Hardcoded constants in formulas | low | `=B2*1.05` looks like a model, behaves like a snapshot |
| Numbers stored as text | low | Sums come back short and nobody notices |
| Merged cells / hidden rows / no structure lock | low | Breaks sorting, filtering and copy-down |

Each finding carries sheet, cell address and a one-line reason — the list you can hand to
whoever owns the model.

## Score and verdict

0–100, weighted by severity, then capped when high-severity findings exist:

- `85+` Audit-ready — no structural problems found
- `70–84` Usable, but review the flagged cells
- `50–69` Risky — fix the high findings before relying on it
- `<50` Not audit-ready — high-severity errors present

The score scales gently with model size: a 5,000-formula workbook is allowed a few cosmetic
findings without being called broken.

## Free vs Pro

| | Free | Pro |
|---|---|---|
| Cell-level findings in the terminal | counts only | full list |
| HTML report | — | single self-contained file |
| CSV export for a review tracker | — | yes |
| Everything runs offline | yes | yes |

Pro is a signed licence file, verified locally. No server, no telemetry, no per-seat phone home.
See `docs/quickstart.md`.

## Install

```
pip install sheetaudit          # once published
pip install -e ".[dev]"         # from source
```

Requires Python 3.9+. One runtime dependency: `openpyxl`.

## Other commands

```
sheetaudit demo                    # build a broken demo model and audit it
sheetaudit license                 # licence status
sheetaudit audit -i book.xlsx --csv findings.csv
```

## Scope

SheetAudit is **static analysis**. It does not recalculate your formulas — it reads the
formulas and the cached values Excel last saved. That means:

- Circular references are found by walking the dependency graph, so they're reported even
  when Excel has iteration enabled and shows a number.
- Error values are found both as cached error strings and as error tokens inside formulas.
- A formula whose result Excel never cached shows up as "no cached value", not as an error.

It supports `.xlsx` and `.xlsm`. Legacy `.xls` is not supported (openpyxl cannot read it).

## Built on

`openpyxl` (MIT) for workbook parsing. SheetAudit adds the checks, scoring, report and licence
layer. See `NOTICE`.

## Licence

Business source: free to use, including commercially, on your own workbooks. Redistribution or
resale of the source requires a commercial licence. `LICENSE` holds the full text.

# SheetAudit Pro — quickstart

## 1. Install

```bash
pip install sheetaudit
```

Python 3.9+, one dependency (`openpyxl`). Windows, macOS and Linux are all fine.

## 2. Store your licence

You received a `sheetaudit-license.json`. Put it here:

```bash
mkdir -p ~/.sheetaudit
mv ~/Downloads/sheetaudit-license.json ~/.sheetaudit/license.json
```

Or point an environment variable at it:

```bash
export SHEETAUDIT_LICENSE=/path/to/sheetaudit-license.json
```

Check it:

```bash
sheetaudit license
```

`Pro license · you@example.com · expires 2027-11-18`

## 3. Audit a workbook

```bash
sheetaudit audit -i budget.xlsx
```

You get a score, a verdict and one line per check with counts. Add the report and the
CSV:

```bash
sheetaudit audit -i budget.xlsx --report audit.html --csv findings.csv
```

- `audit.html` is one self-contained file — mail it, drop it in a ticket, no assets to lose.
- `findings.csv` has severity, check, sheet, cell and detail — paste it straight into a
  review tracker.

## 4. Read the output

| Column | Meaning |
|---|---|
| Severity | `high` = the model is wrong today; `medium` = fragile; `low` = hygiene |
| Sheet / Cell | Exactly where to look; `-` means workbook-level |
| Detail | What is wrong and, where possible, the offending value |

Score bands: 85+ audit-ready · 70–84 review flagged cells · 50–69 risky · below 50 not
audit-ready.

## 5. See it work before you trust it

```bash
sheetaudit demo -o /tmp/demo
```

Builds a small P&L workbook with planted errors (circular reference, `#REF!`, external
link, a column that drifts to 1.08, a divide-by-zero) and audits it. Open
`/tmp/demo/sheetaudit-report.html`.

## Common questions

**Does it recalculate my formulas?**
No. It reads formulas and the values Excel cached at last save. That is deliberate:
recalculation needs a full formula engine and gives different answers to Excel in edge
cases. SheetAudit answers "is this file structurally sound", not "what does it compute".

**Does it upload anything?**
No. No network calls at all. You can run it on a machine with no internet.

**`.xlsm` macros?**
Supported as a workbook. Macros are never executed; only their host cells are read.

**Legacy `.xls`?**
Not supported — export to `.xlsx` first.

**A clean score means my model is right?**
No. It means no structural problems were found. Logic errors, wrong assumptions and bad
inputs are outside what static analysis can see.

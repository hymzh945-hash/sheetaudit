"""`sheetaudit audit | demo | license`."""

from __future__ import annotations

import argparse
import csv
import sys
from pathlib import Path

from . import audit as audit_mod, licensing, report
from .version import __version__

URL = "https://sheetaudit.gumroad.com"


def _require_pro(feature: str) -> None:
    print(f"{feature} is part of SheetAudit Pro — {URL}")
    raise SystemExit(1)


def demo_workbook(path: Path) -> Path:
    """A small model with the mistakes people actually ship."""
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "P&L"

    ws["A1"], ws["B1"], ws["C1"] = "Item", "Amount", "Growth"
    items = [("Revenue", 1200), ("COGS", 480), ("Gross profit", 720), ("Opex", 300), ("EBIT", 420)]
    for i, (label, amount) in enumerate(items, start=2):
        ws[f"A{i}"] = label
        ws[f"B{i}"] = amount
        ws[f"C{i}"] = f"=B{i}*1.05"          # hardcoded 1.05 on every row
    ws["C6"] = "=B6*1.08"                     # one row drifts from the pattern
    ws["D2"] = '=INDIRECT("B2")'              # volatile
    ws["E2"] = "='[Budget2025.xlsx]Sheet1'!B4"  # external workbook link
    ws["B9"] = "#REF!"                        # error value pasted as text
    ws["A10"] = "1234"                        # number stored as text
    ws.merge_cells("A12:B12")
    ws["A12"] = "Notes merged across two columns"
    ws["F2"] = "=SUM(B2:B6)/0"                # will blow up on recalculation

    hidden = wb.create_sheet("Scratch")
    hidden["A1"] = "=A1+1"                    # circular
    hidden["A2"] = "=OFFSET(B1,1,0)"          # volatile + hidden
    hidden.sheet_state = "hidden"

    wb.create_sheet("Assumptions")["A1"] = 0.05
    wb.save(path)
    return path


def cmd_audit(args: argparse.Namespace) -> int:
    pro = licensing.is_pro()
    result = audit_mod.audit(args.input)

    print(f"\nSheetAudit {__version__} · {result['book']['name']}")
    print(f"  score      {result['score']:.0f}/100  {result['verdict']}")
    print(f"  findings   {len(result['findings'])} "
          f"({result['severity'].get('high', 0)} high, "
          f"{result['severity'].get('medium', 0)} medium, "
          f"{result['severity'].get('low', 0)} low)")
    print(f"  sheets     {result['sheets']} · formulas {result['formula_cells']}")

    for row in audit_mod.summary_rows(result):
        print(f"  - {row['check']:<38} {row['count']:>3}  [{row['severity']}]")

    if pro and args.csv:
        out = Path(args.csv)
        out.parent.mkdir(parents=True, exist_ok=True)
        with out.open("w", newline="", encoding="utf-8") as handle:
            writer = csv.DictWriter(handle, fieldnames=["severity", "check", "sheet", "cell", "detail"])
            writer.writeheader()
            for f in result["findings"]:
                writer.writerow({k: f[k] for k in writer.fieldnames})
        print(f"\nfindings -> {out}")

    if args.report:
        if not pro:
            _require_pro("HTML report")
        out = Path(args.report)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(report.build(result, pro=pro, version=__version__, top=args.top), "utf-8")
        print(f"report -> {out}")
    return 0


def cmd_demo(args: argparse.Namespace) -> int:
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    book = demo_workbook(out / "demo_model.xlsx")
    result = audit_mod.audit(book)

    print(f"{len(result['findings'])} findings · score {result['score']:.0f} · {result['verdict']}")
    for row in audit_mod.summary_rows(result):
        print(f"  - {row['check']}: {row['count']}")

    html = report.build(result, pro=licensing.is_pro(), version=__version__)
    (out / "sheetaudit-report.html").write_text(html, "utf-8")
    print(f"report -> {out / 'sheetaudit-report.html'} "
          f"({'Pro' if licensing.is_pro() else 'free tier'})")
    return 0


def cmd_license(args: argparse.Namespace) -> int:
    st = licensing.status()
    if st["valid"]:
        print(f"Pro license · {st['email']} · expires {st['expires'] or 'lifetime'}")
        print(f"source: {licensing.license_path()}")
    else:
        print(f"Free tier · {st['info']}")
        print(f"Pro: {URL}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="sheetaudit",
                                description="Find what breaks a spreadsheet before someone else does.")
    p.add_argument("--version", action="version", version=f"sheetaudit {__version__}")
    sub = p.add_subparsers(dest="command", required=True)

    a = sub.add_parser("audit", help="audit an .xlsx workbook")
    a.add_argument("-i", "--input", required=True)
    a.add_argument("--report", help="write an HTML report (Pro)")
    a.add_argument("--csv", help="write findings to CSV (Pro)")
    a.add_argument("--top", type=int, default=40)
    a.set_defaults(func=cmd_audit)

    d = sub.add_parser("demo", help="build a broken demo model and audit it")
    d.add_argument("-o", "--out", default="sheetaudit-demo")
    d.set_defaults(func=cmd_demo)

    l = sub.add_parser("license", help="license status")
    l.set_defaults(func=cmd_license)
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        return args.func(args)
    except KeyboardInterrupt:
        return 130
    except Exception as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

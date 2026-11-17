"""Read an .xlsx workbook and report the things that actually break models."""

from __future__ import annotations

import re
from collections import Counter, defaultdict
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Set, Tuple

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

CELL_REF = re.compile(r"(?<![A-Z0-9_])(\$?)([A-Z]{1,3})(\$?)(\d{1,7})(?![A-Z0-9_(])")
RANGE_REF = re.compile(r"\$?[A-Z]{1,3}\$?\d{1,7}:\$?[A-Z]{1,3}\$?\d{1,7}")
NUMBER = re.compile(r"(?<![A-Z0-9_.])\d+(?:\.\d+)?")
VOLATILE = {"INDIRECT", "OFFSET", "NOW", "TODAY", "RAND", "RANDBETWEEN", "INFO", "CELL"}
ERRORS = ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NULL!", "#NUM!")

SEVERITY_WEIGHT = {"high": 20, "medium": 8, "low": 2}


def _is_formula(value) -> bool:
    return isinstance(value, str) and value.startswith("=")


def _normalise(formula: str) -> str:
    """Collapse refs but KEEP numbers: a row that divides by 1.05 while its siblings
    divide by 1.08 is exactly the drift this check exists to catch."""
    text = RANGE_REF.sub("<R>", formula.upper())
    text = CELL_REF.sub("<R>", text)
    return re.sub(r"\s+", "", text)


def _refs(formula: str, sheet_max_row: int = 10_000) -> Iterable[Tuple[int, int]]:
    for match in CELL_REF.finditer(formula):
        col = 0
        for ch in match.group(2):
            col = col * 26 + (ord(ch) - 64)
        row = int(match.group(4))
        if row <= sheet_max_row:
            yield (row, col)


def load(path: str | Path) -> Dict:
    path = Path(path)
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError("only .xlsx / .xlsm workbooks are supported")
    formulas = load_workbook(path, data_only=False, read_only=False, keep_links=True)
    values = load_workbook(path, data_only=True, read_only=False, keep_links=True)

    sheets = []
    for name in formulas.sheetnames:
        ws, wv = formulas[name], values[name]
        cells = {}
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                cells[(cell.row, cell.column)] = {
                    "ref": f"{get_column_letter(cell.column)}{cell.row}",
                    "value": cell.value,
                    "cached": wv[cell.coordinate].value,
                    "formula": cell.value if _is_formula(cell.value) else None,
                    "number_format": cell.number_format,
                }
        sheets.append({
            "name": name,
            "index": formulas.sheetnames.index(name),
            "hidden": ws.sheet_state != "visible",
            "state": ws.sheet_state,
            "max_row": ws.max_row,
            "max_column": ws.max_column,
            "cells": cells,
            "merged": [str(rng) for rng in ws.merged_cells.ranges],
            "hidden_rows": sum(1 for i, d in ws.row_dimensions.items() if d.hidden),
            "hidden_cols": sum(1 for i, d in ws.column_dimensions.items() if d.hidden),
            "protected": ws.protection.sheet,
        })

    return {
        "path": str(path),
        "name": path.name,
        "sheets": sheets,
        "defined_names": [str(n) for n in formulas.defined_names],
        "external_links": [link.file_link.Target if hasattr(link, "file_link") else str(link)
                           for link in getattr(formulas, "_external_links", [])],
        "structure_protected": bool(formulas.security and formulas.security.lockStructure),
        "creator": formulas.properties.creator or "",
    }


def _finding(check: str, severity: str, sheet: str, ref: str, detail: str) -> Dict:
    return {"check": check, "severity": severity, "sheet": sheet, "cell": ref,
            "detail": detail, "weight": SEVERITY_WEIGHT[severity]}


def check_errors(book: Dict) -> List[Dict]:
    out = []
    for sheet in book["sheets"]:
        for cell in sheet["cells"].values():
            formula = cell["formula"] or ""
            for err in ERRORS:
                if err in formula or (isinstance(cell["value"], str) and cell["value"].startswith(err)):
                    out.append(_finding("error_value", "high", sheet["name"], cell["ref"],
                                        f"{err} in {cell['ref']}"))
                    break
    return out


def check_volatile(book: Dict) -> List[Dict]:
    out = []
    for sheet in book["sheets"]:
        for cell in sheet["cells"].values():
            formula = cell["formula"]
            if not formula:
                continue
            upper = formula.upper()
            hits = [fn for fn in VOLATILE if re.search(rf"\b{fn}\(", upper)]
            if hits:
                out.append(_finding("volatile_function", "medium", sheet["name"], cell["ref"],
                                    f"{', '.join(sorted(hits))} forces recalculation"))
    return out


def check_hardcoded(book: Dict, threshold: int = 1) -> List[Dict]:
    out = []
    for sheet in book["sheets"]:
        for cell in sheet["cells"].values():
            formula = cell["formula"]
            if not formula:
                continue
            stripped = CELL_REF.sub("", RANGE_REF.sub("", formula))
            numbers = [n for n in NUMBER.findall(stripped)]
            meaningful = [n for n in numbers if n not in {"0", "1", "2"}]
            if len(meaningful) >= threshold:
                out.append(_finding("hardcoded_constant", "low", sheet["name"], cell["ref"],
                                    f"formula carries {len(meaningful)} constants: "
                                    + ", ".join(meaningful[:5])))
    return out


def check_division_by_zero(book: Dict) -> List[Dict]:
    out = []
    for sheet in book["sheets"]:
        for cell in sheet["cells"].values():
            formula = cell["formula"]
            if formula and re.search(r"/\s*\(?\s*0\s*\)?(?![0-9.])", formula):
                out.append(_finding("division_by_zero", "medium", sheet["name"], cell["ref"],
                                    "formula divides by a literal zero"))
    return out


def check_external(book: Dict) -> List[Dict]:
    out = []
    for sheet in book["sheets"]:
        for cell in sheet["cells"].values():
            formula = cell["formula"]
            if formula and "[" in formula and "]" in formula:
                out.append(_finding("external_link", "medium", sheet["name"], cell["ref"],
                                    "formula points at another workbook"))
    for link in book["external_links"]:
        out.append(_finding("external_link", "medium", "<workbook>", "-",
                            f"workbook link: {link}"))
    return out


def check_circular(book: Dict, max_visited: int = 4000) -> List[Dict]:
    out = []
    for sheet in book["sheets"]:
        graph: Dict[Tuple[int, int], Set[Tuple[int, int]]] = {}
        for key, cell in sheet["cells"].items():
            formula = cell["formula"]
            if formula:
                graph[key] = set(_refs(formula, sheet["max_row"]))

        seen_cycles: Set[Tuple] = set()
        visited = 0

        def walk(node, stack, path):
            nonlocal visited
            visited += 1
            if visited > max_visited:
                return
            for nxt in sorted(graph.get(node, ())):
                if nxt in stack:
                    cycle = path[path.index(nxt):] if nxt in path else path
                    key = tuple(sorted(cycle))
                    if key not in seen_cycles:
                        seen_cycles.add(key)
                        refs = [f"{get_column_letter(c)}{r}" for r, c in cycle][:6]
                        out.append(_finding("circular_reference", "high", sheet["name"],
                                            refs[0] if refs else "-",
                                            "cycle: " + " → ".join(refs)))
                    continue
                if nxt in graph:
                    walk(nxt, stack | {nxt}, path + [nxt])

        for node in graph:
            walk(node, {node}, [node])
    return out


def check_inconsistent(book: Dict, min_block: int = 3) -> List[Dict]:
    """A column of identical formulas with one odd cell out is how models drift."""
    out = []
    for sheet in book["sheets"]:
        by_col: Dict[int, List[Tuple[int, str, str]]] = defaultdict(list)
        for key, cell in sheet["cells"].items():
            if cell["formula"]:
                by_col[key[1]].append((key[0], _normalise(cell["formula"]), cell["ref"]))

        for column, items in by_col.items():
            items.sort()
            block, start = [], None
            for row, shape, ref in items:
                if block and row != block[-1][0] + 1:
                    out.extend(_judge_block(sheet["name"], block, min_block))
                    block = []
                block.append((row, shape, ref))
            out.extend(_judge_block(sheet["name"], block, min_block))
    return out


def _judge_block(sheet_name: str, block: List[Tuple[int, str, str]], min_block: int) -> List[Dict]:
    if len(block) < min_block:
        return []
    shapes = Counter(shape for _, shape, _ in block)
    if len(shapes) == 1:
        return []
    dominant, _ = shapes.most_common(1)[0]
    odd = [(ref, shape) for _, shape, ref in block if shape != dominant]
    if not odd:
        return []
    return [_finding("inconsistent_formula", "medium", sheet_name, ref,
                     f"{len(odd)} of {len(block)} cells in this column deviate from the pattern")
            for ref, _ in odd[:8]]


def check_hygiene(book: Dict) -> List[Dict]:
    out = []
    for sheet in book["sheets"]:
        if sheet["hidden"]:
            formulas = sum(1 for c in sheet["cells"].values() if c["formula"])
            out.append(_finding("hidden_sheet", "medium" if formulas else "low", sheet["name"], "-",
                                f"sheet is {sheet['state']} with {formulas} live formulas"))
        if sheet["hidden_rows"]:
            out.append(_finding("hidden_rows", "low", sheet["name"], "-",
                                f"{sheet['hidden_rows']} hidden rows"))
        for rng in sheet["merged"]:
            out.append(_finding("merged_cells", "low", sheet["name"], rng,
                                "merged range can break sorting and copy-down"))
        for cell in sheet["cells"].values():
            if isinstance(cell["value"], str) and cell["value"].strip().lstrip("-").replace(".", "", 1).isdigit():
                out.append(_finding("number_as_text", "low", sheet["name"], cell["ref"],
                                    f"numeric text {cell['value']!r} will not sum"))
            if cell["formula"] and len(cell["formula"]) > 300:
                out.append(_finding("long_formula", "low", sheet["name"], cell["ref"],
                                    f"{len(cell['formula'])} characters in one cell"))
    if not book["structure_protected"]:
        out.append(_finding("no_structure_protection", "low", "<workbook>", "-",
                            "sheets can be added, deleted or reordered without a password"))
    return out


CHECKS = [
    ("error_value", "Error values", check_errors, "high"),
    ("circular_reference", "Circular references", check_circular, "high"),
    ("external_link", "External workbook links", check_external, "medium"),
    ("division_by_zero", "Division by zero", check_division_by_zero, "medium"),
    ("volatile_function", "Volatile functions", check_volatile, "medium"),
    ("inconsistent_formula", "Formula drift in a column", check_inconsistent, "medium"),
    ("hidden_sheet", "Hidden sheets with live formulas", check_hygiene, "medium"),
    ("hardcoded_constant", "Hardcoded constants in formulas", check_hardcoded, "low"),
    ("number_as_text", "Numbers stored as text", check_hygiene, "low"),
    ("merged_cells", "Merged cells", check_hygiene, "low"),
    ("no_structure_protection", "Workbook structure not protected", check_hygiene, "low"),
]


def audit(path: str | Path) -> Dict:
    book = load(path)
    findings: List[Dict] = []
    for _, _, fn, _ in CHECKS:
        findings.extend(fn(book))
    unique = {(f["check"], f["sheet"], f["cell"], f["detail"]): f for f in findings}
    findings = sorted(unique.values(), key=lambda f: (-f["weight"], f["sheet"], f["cell"]))

    counts = Counter(f["check"] for f in findings)
    severity = Counter(f["severity"] for f in findings)
    formula_cells = sum(1 for s in book["sheets"] for c in s["cells"].values() if c["formula"])
    total_cells = sum(len(s["cells"]) for s in book["sheets"])

    score, verdict, colour = _score(findings, severity, formula_cells)
    return {
        "book": book,
        "findings": findings,
        "counts": counts,
        "severity": severity,
        "score": score,
        "verdict": verdict,
        "verdict_color": colour,
        "formula_cells": formula_cells,
        "total_cells": total_cells,
        "sheets": len(book["sheets"]),
    }


def _score(findings, severity, formula_cells) -> Tuple[float, str, str]:
    penalty = 0.0
    for f in findings:
        penalty += f["weight"]
    # a big model can legitimately carry a few low findings; scale gently, then cap
    scale = max(1.0, (formula_cells / 400) ** 0.5)
    # a floor of 12: a real workbook is never "literally zero" good, and 0/100
    # reads as a broken tool rather than a broken model
    value = max(12.0, min(100.0, 100 - penalty / scale))

    highs = severity.get("high", 0)
    if highs >= 3:
        value = min(value, 50.0)
    elif highs == 2:
        value = min(value, 62.0)
    elif highs == 1:
        value = min(value, 72.0)
    if value >= 85:
        return round(value, 1), "Audit-ready — no structural problems found", "#34d399"
    if value >= 70:
        return round(value, 1), "Usable, but review the flagged cells", "#fbbf24"
    if value >= 50:
        return round(value, 1), "Risky — fix the high findings before relying on it", "#fb923c"
    return round(value, 1), "Not audit-ready — high-severity errors present", "#f87171"


def summary_rows(result: Dict) -> List[Dict]:
    rows = []
    for key, label, _, _ in CHECKS:
        n = result["counts"].get(key, 0)
        if n:
            worst = "high" if any(f["check"] == key and f["severity"] == "high" for f in result["findings"]) \
                else "medium" if any(f["check"] == key and f["severity"] == "medium" for f in result["findings"]) \
                else "low"
            rows.append({"check": label, "count": n, "severity": worst})
    return rows

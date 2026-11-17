"""SheetAudit tests — every check gets a workbook that should trip it, and one that shouldn't."""

from __future__ import annotations

import base64
import json

import pytest
from openpyxl import Workbook

from sheetaudit import audit, licensing, report


def book(tmp_path, cells, name="model.xlsx", hidden=None, merge=None):
    wb = Workbook()
    ws = wb.active
    for ref, value in cells.items():
        ws[ref] = value
    if merge:
        ws.merge_cells(merge)
    if hidden:
        sh = wb.create_sheet("Scratch")
        for ref, value in hidden.items():
            sh[ref] = value
        sh.sheet_state = "hidden"
    path = tmp_path / name
    wb.save(path)
    return path


def checks(result) -> set[str]:
    return {f["check"] for f in result["findings"]}


def test_error_value_is_high_severity(tmp_path):
    path = book(tmp_path, {"A1": 1, "A2": "#REF!", "A3": "=B1/0"})
    result = audit.audit(path)
    assert "error_value" in checks(result)
    assert result["severity"]["high"] >= 1
    assert result["score"] < 80


def test_circular_self_reference(tmp_path):
    path = book(tmp_path, {"A1": "=A1+1", "A2": 5})
    result = audit.audit(path)
    assert "circular_reference" in checks(result)


def test_circular_two_cell_loop(tmp_path):
    path = book(tmp_path, {"A1": "=B1+1", "B1": "=A1+1"})
    result = audit.audit(path)
    circular = [f for f in result["findings"] if f["check"] == "circular_reference"]
    assert circular
    assert "→" in circular[0]["detail"]


def test_healthy_formula_chain_is_not_flagged(tmp_path):
    path = book(tmp_path, {"A1": "=B1+B2", "B1": 3, "B2": 4, "A2": "=A1*2"})
    result = audit.audit(path)
    assert "circular_reference" not in checks(result)


def test_external_workbook_link(tmp_path):
    path = book(tmp_path, {"A1": "='[Budget.xlsx]Sheet1'!B4"})
    result = audit.audit(path)
    assert "external_link" in checks(result)


def test_volatile_functions(tmp_path):
    path = book(tmp_path, {"A1": '=INDIRECT("B1")', "A2": "=OFFSET(B1,1,0)", "A3": "=TODAY()"})
    result = audit.audit(path)
    volatile = [f for f in result["findings"] if f["check"] == "volatile_function"]
    assert len(volatile) == 3


def test_hardcoded_constant_in_formula(tmp_path):
    path = book(tmp_path, {"A1": "=B1*1.05", "B1": 100})
    result = audit.audit(path)
    assert "hardcoded_constant" in checks(result)
    assert "1.05" in next(f["detail"] for f in result["findings"] if f["check"] == "hardcoded_constant")


def test_column_drift_with_different_constant(tmp_path):
    path = book(tmp_path, {"A1": "=B1*1.05", "A2": "=B2*1.05", "A3": "=B3*1.05",
                           "A4": "=B4*1.08"})
    result = audit.audit(path)
    drift = [f for f in result["findings"] if f["check"] == "inconsistent_formula"]
    assert drift
    assert any(f["cell"] == "A4" for f in drift)


def test_consistent_column_has_no_drift(tmp_path):
    path = book(tmp_path, {"A1": "=B1*2", "A2": "=B2*2", "A3": "=B3*2"})
    result = audit.audit(path)
    assert "inconsistent_formula" not in checks(result)


def test_number_stored_as_text(tmp_path):
    path = book(tmp_path, {"A1": "1234", "A2": "56.7"})
    result = audit.audit(path)
    assert "number_as_text" in checks(result)


def test_division_by_zero_literal(tmp_path):
    path = book(tmp_path, {"A1": "=SUM(B1:B4)/0", "B1": 1})
    result = audit.audit(path)
    assert "division_by_zero" in checks(result)


def test_hidden_sheet_with_formulas(tmp_path):
    path = book(tmp_path, {"A1": 1}, hidden={"A1": "=B1*2"})
    result = audit.audit(path)
    assert "hidden_sheet" in checks(result)


def test_clean_workbook_scores_well(tmp_path):
    path = book(tmp_path, {"A1": "Item", "B1": "Value", "A2": "x", "B2": 10, "B3": 20,
                           "B4": "=SUM(B2:B3)"})
    result = audit.audit(path)
    assert result["score"] >= 85
    assert "Audit-ready" in result["verdict"]


def test_broken_workbook_scores_poorly(tmp_path):
    path = book(tmp_path, {"A1": "=A1+1", "A2": "#DIV/0!", "A3": '=INDIRECT("A1")'})
    result = audit.audit(path)
    assert result["score"] < 65
    assert result["verdict_color"] in {"#fb923c", "#f87171"}
    assert "Risky" in result["verdict"] or "Not audit-ready" in result["verdict"]


def test_summary_rows_only_lists_hits(tmp_path):
    path = book(tmp_path, {"A1": "=B1*1.05", "B1": 100})
    rows = audit.summary_rows(audit.audit(path))
    assert any(r["check"] == "Hardcoded constants in formulas" for r in rows)
    assert all(r["count"] > 0 for r in rows)


def test_report_gates_pro_findings(tmp_path):
    path = book(tmp_path, {"A1": "=A1+1"})
    result = audit.audit(path)
    free = report.build(result, pro=False)
    assert "Cell-level findings are Pro" in free
    pro = report.build(result, pro=True)
    assert "Cell-level findings are Pro" not in pro
    assert "Circular references" in pro


def test_load_rejects_non_xlsx(tmp_path):
    path = tmp_path / "data.csv"
    path.write_text("a,b\n1,2\n")
    with pytest.raises(ValueError):
        audit.load(path)


def test_bad_license_falls_back_to_free(tmp_path, monkeypatch):
    doc = {"payload": {"product": "sheetaudit"}, "sig": base64.b64encode(b"no").decode()}
    path = tmp_path / "license.json"
    path.write_text(json.dumps(doc))
    monkeypatch.setenv("SHEETAUDIT_LICENSE", str(path))
    assert licensing.status()["valid"] is False

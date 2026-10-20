from pathlib import Path

import pytest

from randomoffice.contracts import ExcelCell, OfficeDocument
from randomoffice.web_office import export_office_file, open_office_file


def test_office_document_round_trips_existing_editor_payload() -> None:
    payload = {
        "app": "excel",
        "format": ".xlsx",
        "title": "Budget",
        "sheets": [{"name": "Data", "rows": [[2, 3, "=SUM(A1:B1)"]]}],
        "loss_report": [
            {"feature": "charts", "message": "Chart styling may be rebuilt", "severity": "warning"}
        ],
    }

    document = OfficeDocument.from_editor_data(payload)

    assert document.schema_version == "office-document/1"
    assert document.app == "excel"
    assert document.loss_report[0].feature == "charts"
    assert document.to_editor_data()["sheets"] == payload["sheets"]


def test_office_document_rejects_mismatched_format() -> None:
    with pytest.raises(ValueError, match="must use"):
        OfficeDocument.from_editor_data({"app": "excel", "format": ".docx"})


def test_office_document_rejects_unknown_schema_and_metadata_types() -> None:
    with pytest.raises(ValueError, match="schema version"):
        OfficeDocument.from_editor_data({"app": "word", "schema_version": "office-document/99"})
    with pytest.raises(TypeError, match="source_name"):
        OfficeDocument.from_editor_data({"app": "word", "source_name": 123})


def test_office_document_preserves_safe_source_metadata() -> None:
    document = OfficeDocument.from_editor_data(
        {
            "app": "word",
            "source_name": "report.docx",
            "source_fingerprint": "sha256:abc",
        },
        source_path=Path("report.docx"),
    )

    assert document.source_name == "report.docx"
    assert document.source_path == Path("report.docx")
    assert document.to_editor_data()["source_fingerprint"] == "sha256:abc"


def test_excel_cell_keeps_formula_separate_from_value() -> None:
    cell = ExcelCell(value=5, formula="=SUM(A1:B1)", value_type="formula")

    assert cell.to_dict() == {
        "value": 5,
        "formula": "=SUM(A1:B1)",
        "value_type": "formula",
        "number_format": None,
        "style_ref": None,
    }


def test_open_office_file_exposes_canonical_schema_version(tmp_path: Path) -> None:
    from openpyxl import Workbook

    source = tmp_path / "sheet.xlsx"
    workbook = Workbook()
    workbook.active["A1"] = "=SUM(1,2)"
    workbook.save(source)

    data = open_office_file(source)

    assert data["schema_version"] == "office-document/1"
    assert data["sheets"][0]["rows"][0][0] == "=SUM(1,2)"
    assert export_office_file("xlsx", OfficeDocument.from_editor_data(data))

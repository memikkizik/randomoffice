from pathlib import Path
from docx import Document
from openpyxl import Workbook
from pptx import Presentation

from randomoffice import default_registry


def test_registry_parses_docx_with_tables_and_metadata(tmp_path: Path) -> None:
    doc_path = tmp_path / "test.docx"
    doc = Document()
    doc.core_properties.title = "Document Core Title"
    doc.add_paragraph("Introductory text")
    tbl = doc.add_table(rows=2, cols=2)
    tbl.rows[0].cells[0].text = "H1"
    tbl.rows[0].cells[1].text = "H2"
    tbl.rows[1].cells[0].text = "V1"
    tbl.rows[1].cells[1].text = "V2"
    doc.add_paragraph("Concluding text")
    doc.save(str(doc_path))

    registry = default_registry()
    parsed = registry.parse(doc_path)

    assert parsed.title == "Document Core Title"
    assert "Introductory text" in parsed.text
    assert "H1 | H2" in parsed.text
    assert "Concluding text" in parsed.text


def test_registry_parses_xlsx_with_metadata(tmp_path: Path) -> None:
    xlsx_path = tmp_path / "test.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws.append(["Metric", "Score"])
    ws.append(["Accuracy", 99])
    wb.properties.title = "XLSX Core Title"
    wb.save(str(xlsx_path))

    registry = default_registry()
    parsed = registry.parse(xlsx_path)

    assert parsed.title == "XLSX Core Title"
    assert "[Summary]" in parsed.text
    assert "Accuracy\t99" in parsed.text


def test_registry_parses_pptx_with_tables_and_notes(tmp_path: Path) -> None:
    pptx_path = tmp_path / "test.pptx"
    prs = Presentation()
    prs.core_properties.title = "Deck Presentation Title"
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    tb = slide.shapes.add_textbox(0, 0, 100, 100)
    tb.text_frame.text = "Slide 1 Content"

    # Presenter note
    slide.notes_slide.notes_text_frame.text = "Key speaker point"
    prs.save(str(pptx_path))

    registry = default_registry()
    parsed = registry.parse(pptx_path)

    assert parsed.title == "Deck Presentation Title"
    assert "Slide 1 Content" in parsed.text
    assert "[Note] Key speaker point" in parsed.text

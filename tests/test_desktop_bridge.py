import base64
import io

from docx import Document

from randomoffice.desktop import NativeDesktopBridge
from randomoffice.ui import render_office_ui


def test_native_drop_opens_office_data_without_http():
    document = Document()
    document.add_paragraph("Dropped document")
    buffer = io.BytesIO()
    document.save(buffer)
    data = NativeDesktopBridge().open_file_data("report.docx", base64.b64encode(buffer.getvalue()).decode("ascii"))
    assert data["filename"] == "report.docx"
    assert data["items"][0]["text"] == "Dropped document"


def test_native_drop_rejects_invalid_file():
    bridge = NativeDesktopBridge()
    assert "error" in bridge.open_file_data("report.exe", "YWJj")
    assert "error" in bridge.open_file_data("report.docx", "not valid base64!")


def test_shared_ui_replaces_legacy_script(tmp_path):
    source = tmp_path / "index.html"
    source.write_text('<html><head></head><body><script>legacy()</script></body></html>', encoding="utf-8")
    source.with_name("editor.js").write_text("function workingEditor() {}", encoding="utf-8")
    source.with_name("formulas.js").write_text("function workingFormula() {}", encoding="utf-8")
    rendered = render_office_ui(source)
    assert "legacy()" not in rendered
    assert "workingEditor()" in rendered
    assert "workingFormula()" in rendered
    assert rendered.count("<script") == 1

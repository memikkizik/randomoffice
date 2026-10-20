import io
from docx import Document
from fastapi.testclient import TestClient
import pytest

from randomoffice.api.app import app


@pytest.fixture
def client():
    with TestClient(app) as test_client:
        yield test_client


def test_health_endpoint(client: TestClient) -> None:
    response = client.get("/health")
    assert response.status_code == 200
    data = response.json()
    assert data["status"] == "healthy"
    assert "uptime_seconds" in data
    assert ".docx" in data["supported_formats"]


def test_api_parse_docx(client: TestClient) -> None:
    doc = Document()
    doc.core_properties.title = "API Test Title"
    doc.add_paragraph("Testing API extraction.")
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)

    response = client.post(
        "/api/v1/parse",
        files={"file": ("test.docx", buf, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
    )
    assert response.status_code == 200
    data = response.json()
    assert data["title"] == "API Test Title"
    assert "Testing API extraction." in data["text"]


def test_api_convert_docx_to_markdown(client: TestClient) -> None:
    doc = Document()
    doc.add_heading("API Heading", level=1)
    doc.add_paragraph("Paragraph in API.")
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)

    response = client.post(
        "/api/v1/convert",
        files={"file": ("test.docx", buf, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
        data={"target": "markdown"},
    )
    assert response.status_code == 200
    data = response.json()
    assert "# API Heading" in data["content"]
    assert "Paragraph in API." in data["content"]


def test_api_rejects_unsupported_file(client: TestClient) -> None:
    response = client.post(
        "/api/v1/parse",
        files={"file": ("test.pdf", b"%PDF-1.4...", "application/pdf")},
    )
    assert response.status_code == 400
    assert "Unsupported file extension" in response.json()["detail"]

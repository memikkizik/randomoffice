"""Small standard-library client for the RandomOffice REST API."""

import argparse
import json
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen
from uuid import uuid4


def request_json(url: str, method: str = "GET", body: bytes | None = None,
                 headers: dict[str, str] | None = None) -> dict:
    """Send a JSON HTTP request and decode its object response."""
    request = Request(url, data=body, headers=headers or {}, method=method)
    with urlopen(request, timeout=30) as response:
        payload = response.read()
    decoded = json.loads(payload)
    if not isinstance(decoded, dict):
        raise ValueError(f"Expected a JSON object from {url}")
    return decoded


def request_raw(url: str, method: str = "GET", body: bytes | None = None,
                headers: dict[str, str] | None = None) -> bytes:
    """Send an HTTP request and return the raw byte payload."""
    request = Request(url, data=body, headers=headers or {}, method=method)
    with urlopen(request, timeout=30) as response:
        return response.read()


def multipart_body(field_name: str, filename: str, content: bytes,
                   fields: dict[str, str] | None = None) -> tuple[bytes, str]:
    """Build a multipart/form-data body for one uploaded file."""
    boundary = f"----RandomOffice{uuid4().hex}"
    chunks: list[bytes] = []
    for name, value in (fields or {}).items():
        chunks.extend([
            f"--{boundary}\r\n".encode(),
            f'Content-Disposition: form-data; name="{name}"\r\n\r\n'.encode(),
            value.encode(),
            b"\r\n",
        ])
    chunks.extend([
        f"--{boundary}\r\n".encode(),
        f'Content-Disposition: form-data; name="{field_name}"; filename="{filename}"\r\n'.encode(),
        b"Content-Type: application/octet-stream\r\n\r\n",
        content,
        b"\r\n",
        f"--{boundary}--\r\n".encode(),
    ])
    return b"".join(chunks), f"multipart/form-data; boundary={boundary}"


def upload(url: str, path: Path, fields: dict[str, str] | None = None) -> dict:
    """Upload a local document to a RandomOffice endpoint."""
    body, content_type = multipart_body("file", path.name, path.read_bytes(), fields)
    return request_json(url, method="POST", body=body, headers={"Content-Type": content_type})


def roundtrip_document(base_url: str, source_path: Path, output_path: Path) -> Path:
    """Open a document via /api/v1/office/open and export it back via /api/v1/office/export."""
    if source_path.resolve() == output_path.resolve():
        raise ValueError(f"Kaynak ve çıktı dosyası aynı olamaz: {source_path}")

    # 1. Open document via API
    open_url = f"{base_url.rstrip('/')}/api/v1/office/open"
    doc_data = upload(open_url, source_path)
    if "error" in doc_data:
        raise RuntimeError(f"Belge açılamadı: {doc_data['error']}")

    doc_type = doc_data.get("app")
    if not doc_type:
        raise RuntimeError(f"Belge türü (app) yanıtta bulunamadı: {doc_data}")

    # 2. Export back via /api/v1/office/export
    query = urlencode({"doc_type": doc_type, "filename": source_path.name})
    export_url = f"{base_url.rstrip('/')}/api/v1/office/export?{query}"
    payload_bytes = json.dumps(doc_data).encode("utf-8")

    exported_bytes = request_raw(
        export_url,
        method="POST",
        body=payload_bytes,
        headers={"Content-Type": "application/json", "Accept": "application/octet-stream"}
    )

    if not exported_bytes:
        raise RuntimeError("API boş dosya çıktısı döndürdü.")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_bytes(exported_bytes)
    return output_path


def main() -> int:
    """Call health, parse, convert endpoints and optionally roundtrip Office documents."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("document", type=Path, help="DOCX, XLSX, or PPTX file to upload")
    parser.add_argument("--base-url", default="http://localhost:8000")
    parser.add_argument(
        "--roundtrip-output",
        type=Path,
        default=None,
        help="Optional path to save binary roundtripped document from /office/open -> /office/export",
    )
    args = parser.parse_args()
    if not args.document.is_file():
        parser.error(f"document does not exist: {args.document}")

    base_url = args.base_url.rstrip("/")
    try:
        print("GET /health")
        print(json.dumps(request_json(f"{base_url}/health"), indent=2))
        print("POST /api/v1/parse")
        print(json.dumps(upload(f"{base_url}/api/v1/parse", args.document), indent=2))
        print("POST /api/v1/convert")
        print(json.dumps(
            upload(f"{base_url}/api/v1/convert", args.document, {"target": "markdown"}),
            indent=2,
        ))

        if args.roundtrip_output:
            print(f"POST /api/v1/office/open -> POST /api/v1/office/export (Roundtrip)")
            out = roundtrip_document(base_url, args.document, args.roundtrip_output)
            print(f"Başarıyla kaydedildi: {out} ({out.stat().st_size} bayt)")

    except (HTTPError, URLError, OSError, ValueError, RuntimeError, json.JSONDecodeError) as exc:
        print(f"API request failed: {exc}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

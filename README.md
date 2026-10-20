# RandomOffice

RandomOffice is a typed foundation for parsing local Microsoft Office files.
The current scope is intentionally limited to read-only extraction:

- `.docx` via `python-docx`
- `.xlsx` via `openpyxl`
- `.pptx` via `python-pptx`

## Setup

```powershell
uv sync --extra test
```

## Contract

Every parser implements `DocumentParser.parse(Path) -> ParsedDocument`.
`ParsedDocument` contains the source path, normalized extension, optional title,
plain-text extraction, and filesystem modification time. Invalid paths,
unsupported extensions, missing dependencies, and malformed documents raise
explicit exceptions from `randomoffice.errors`.

## Example

```python
from pathlib import Path

from randomoffice import default_registry

document = default_registry().parse(Path(r"C:\docs\report.docx"))
print(document.title)
print(document.text)
```

The same parser can be launched from the command line:

```powershell
uv run randomoffice
uv run randomoffice "C:\docs\report.docx"
```

This layer does not modify files, call Microsoft Graph, or use mock data.
Those capabilities should be added as separate adapters after their
authentication and write contracts are defined.

## Tests

```powershell
uv run pytest
```

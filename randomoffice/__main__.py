"""Command-line entry point for the local document parser."""

import argparse
import sys
from pathlib import Path

from .errors import DocumentParseError
from .registry import default_registry


def build_parser() -> argparse.ArgumentParser:
    """Build the CLI argument parser.

    Input: no arguments.
    Output: configured argparse parser for status and document parsing commands.
    """
    parser = argparse.ArgumentParser(
        prog="randomoffice",
        description="Parse local DOCX, XLSX, and PPTX documents.",
    )
    parser.add_argument(
        "document",
        nargs="?",
        type=Path,
        help="local .docx, .xlsx, or .pptx file to parse",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    """Run the CLI and return a process exit code."""
    args = build_parser().parse_args(argv)
    registry = default_registry()

    if args.document is None:
        print("RandomOffice backend is ready.")
        print("Supported formats: .docx, .xlsx, .pptx")
        print("Mode: local read-only document extraction")
        print("Usage: uv run randomoffice <path-to-document>")
        return 0

    try:
        document = registry.parse(args.document)
    except DocumentParseError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 2

    print(f"File: {document.source_path}")
    print(f"Format: {document.format}")
    print(f"Title: {document.title or '<untitled>'}")
    print("Content:")
    print(document.text or "<empty>")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

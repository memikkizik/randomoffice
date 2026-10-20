"""Command-line entry point for the local document parser and server."""

import argparse
from pathlib import Path
import sys
import time
from typing import Sequence

from .converters.contracts import ConversionTarget
from .converters.pipeline import ConversionPipeline
from .errors import DocumentParseError
from .registry import default_registry
from .sync.watcher import DirectorySyncManager


def build_parser() -> argparse.ArgumentParser:
    """Build the CLI argument parser with subcommands and legacy compatibility."""
    parser = argparse.ArgumentParser(
        prog="randomoffice",
        description="Production CLI for Microsoft Office parsing, conversion, and server.",
    )
    subparsers = parser.add_subparsers(dest="subcommand", help="Available subcommands")

    # 'desktop' subcommand
    desktop_parser = subparsers.add_parser("desktop", help="Launch native desktop GUI application")
    desktop_parser.add_argument("document", nargs="?", type=Path, help="Office document to open at startup")

    # 'serve' subcommand
    serve_parser = subparsers.add_parser("serve", help="Start the production FastAPI server")
    serve_parser.add_argument(
        "--host",
        default="127.0.0.1",
        help="Bind host (default: 127.0.0.1; use 0.0.0.0 only behind a secured proxy)",
    )
    serve_parser.add_argument("--port", type=int, default=8000, help="Bind port (default: 8000)")
    serve_parser.add_argument("--workers", type=int, default=1, help="Uvicorn workers count")

    # 'parse' subcommand
    parse_parser = subparsers.add_parser("parse", help="Parse an Office document")
    parse_parser.add_argument("document", type=Path, help="Path to .docx, .xlsx, or .pptx file")

    # 'convert' subcommand
    convert_parser = subparsers.add_parser("convert", help="Convert an Office document")
    convert_parser.add_argument("document", type=Path, help="Path to .docx, .xlsx, or .pptx file")
    convert_parser.add_argument(
        "--target",
        choices=["markdown", "json", "csv", "text"],
        default="markdown",
        help="Target format (default: markdown)",
    )
    convert_parser.add_argument("--output", "-o", type=Path, help="Output destination file")

    # 'watch' subcommand
    watch_parser = subparsers.add_parser("watch", help="Watch a directory for Office documents")
    watch_parser.add_argument("directory", type=Path, help="Directory to monitor")
    watch_parser.add_argument(
        "--target",
        choices=["markdown", "json", "csv", "text"],
        default="markdown",
        help="Target conversion format for changed files",
    )

    # Legacy fallback argument
    parser.add_argument(
        "legacy_document",
        nargs="?",
        type=Path,
        help=argparse.SUPPRESS,
    )

    return parser


def main(argv: Sequence[str] | None = None) -> int:
    """Run the CLI and return process exit code."""
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.subcommand == "desktop":
        from .desktop import run_desktop_app
        if args.document is None:
            run_desktop_app()
        else:
            run_desktop_app(startup_document=args.document)
        return 0

    if args.subcommand == "serve":
        try:
            import uvicorn
            from .api.app import app
        except ImportError as exc:
            print(f"Error: API dependencies missing: {exc}", file=sys.stderr)
            return 1
        uvicorn.run(app, host=args.host, port=args.port)
        return 0

    if args.subcommand == "convert":
        target = ConversionTarget(args.target)
        pipeline = ConversionPipeline()
        try:
            result = pipeline.convert(args.document, target=target)
            if args.output:
                args.output.write_text(result.content, encoding="utf-8")
                print(f"Converted {args.document} -> {args.output}")
            else:
                print(result.content)
            return 0
        except DocumentParseError as exc:
            print(f"Error: {exc}", file=sys.stderr)
            return 2

    if args.subcommand == "watch":
        pipeline = ConversionPipeline()
        target = ConversionTarget(args.target)

        def handle_change(change):
            print(f"[{change.event_type.upper()}] {change.path}")
            if change.event_type != "deleted":
                try:
                    res = pipeline.convert(change.path, target=target)
                    out_path = change.path.with_suffix(f".{target.value}")
                    out_path.write_text(res.content, encoding="utf-8")
                    print(f"  -> Converted to {out_path.name}")
                except Exception as exc:
                    print(f"  -> Error converting {change.path.name}: {exc}")

        watcher = DirectorySyncManager(args.directory, on_change=handle_change)
        print(f"Watching directory {args.directory} for Office documents... (Ctrl+C to stop)")
        watcher.start()
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            watcher.stop()
            print("Stopped directory watcher.")
            return 0

    # Parse command or legacy document argument
    doc_path = args.document if args.subcommand == "parse" else args.legacy_document

    if doc_path is None:
        print("RandomOffice backend is ready.")
        print("Supported formats: .docx, .xlsx, .pptx")
        print("Mode: local read-only document extraction & conversion pipeline")
        print("Usage: uv run randomoffice [parse|convert|serve|watch] <path>")
        return 0

    registry = default_registry()
    try:
        document = registry.parse(doc_path)
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

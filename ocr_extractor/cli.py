"""Command-line interface for the OCR table extractor."""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from .extractor import extract


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="ocr-extract",
        description=(
            "Convert .jpg/.jpeg/.png/.tif/.tiff/.pdf files into a "
            "layout-preserving output file that reconstructs the table structure."
        ),
    )
    parser.add_argument(
        "input",
        metavar="INPUT",
        help="Path to the input file (.jpg/.jpeg/.png/.tif/.tiff/.pdf).",
    )
    parser.add_argument(
        "-o",
        "--output",
        metavar="OUTPUT",
        default=None,
        help=(
            "Path for the output file. "
            "Defaults to INPUT with the .html extension when --tenancy-mode is used "
            "(or --format html), otherwise .xlsx."
        ),
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        default=False,
        help=(
            "Enable debug mode: log detected columns/rows/merges and save an "
            "annotated preview image (INPUT.debug.png)."
        ),
    )
    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        default=False,
        help="Enable verbose (INFO-level) logging.",
    )
    parser.add_argument(
        "--tenancy-mode",
        action="store_true",
        default=False,
        help=(
            "Enable tenancy schedule parsing mode. "
            "Uses specialized parsing for real estate lease documents with fields like "
            "Property, Tenant, Suite, Lease dates, Area, Rent, etc. "
            "Outputs a clean HTML <table> (implies --format html)."
        ),
    )
    parser.add_argument(
        "--format",
        choices=["xlsx", "html"],
        default="xlsx",
        dest="output_format",
        help=(
            "Output format. 'xlsx' writes a .xlsx file. "
            "'html' writes an .html file containing a clean HTML <table> "
            "with normalized tenancy schedule data (implies --tenancy-mode). "
            "When --tenancy-mode is specified, defaults to 'html'."
        ),
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    """Entry point for the ``ocr-extract`` command."""
    parser = _build_parser()
    args = parser.parse_args(argv)

    log_level = logging.DEBUG if args.debug else (logging.INFO if args.verbose else logging.WARNING)
    logging.basicConfig(
        level=log_level,
        format="%(levelname)s %(name)s: %(message)s",
        stream=sys.stderr,
    )

    try:
        output = extract(
            input_path=args.input,
            output_path=args.output,
            debug=args.debug,
            tenancy_mode=args.tenancy_mode,
            output_format=args.output_format,
        )
        print(f"Output written to: {output}")
        return 0
    except FileNotFoundError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    except ValueError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:  # noqa: BLE001
        logging.getLogger(__name__).exception("Unexpected error: %s", exc)
        return 2


if __name__ == "__main__":
    sys.exit(main())

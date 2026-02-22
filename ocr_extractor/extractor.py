"""Main extraction pipeline.

Orchestrates the full conversion flow:
    Input file (.jpg/.png/.tif/.pdf)
    → Image(s)
    → Pre-processing
    → Table detection
    → OCR of each cell
    → Merge detection
    → HTML table writing
"""

from __future__ import annotations

import logging
import os
import tempfile
from pathlib import Path
from typing import List, Optional

from PIL import Image

from .ocr_engine import ocr_cell
from .preprocessor import preprocess
from .table_detector import TableGrid, detect_merges, detect_table

logger = logging.getLogger(__name__)

# Supported image extensions
_IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".tif", ".tiff"}
_PDF_EXT = ".pdf"


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def extract(
    input_path: str | Path,
    output_path: Optional[str | Path] = None,
    debug: bool = False,
) -> Path:
    """Convert *input_path* to a layout-preserving HTML table file.

    Parameters
    ----------
    input_path:
        Path to a .jpg/.jpeg/.png/.tif/.tiff or .pdf file.
    output_path:
        Destination path.  Defaults to the input file name with ``.html``
        in the same directory.
    debug:
        When *True* extra diagnostic information is logged and (for image
        inputs) an annotated preview PNG is saved alongside the output.

    Returns
    -------
    Path
        Resolved path to the written file.
    """
    input_path = Path(input_path)
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    if output_path is None:
        output_path = input_path.with_suffix(".html")
    output_path = Path(output_path)

    # Validate the output directory exists and is writable
    out_dir = output_path.parent
    if not out_dir.exists():
        raise FileNotFoundError(f"Output directory does not exist: {out_dir}")
    if not os.access(out_dir, os.W_OK):
        raise PermissionError(f"Output directory is not writable: {out_dir}")

    ext = input_path.suffix.lower()
    if ext == _PDF_EXT:
        images = _load_pdf(input_path, debug=debug)
    elif ext in _IMAGE_EXTS:
        images = [Image.open(input_path).convert("RGB")]
    else:
        raise ValueError(
            f"Unsupported file type: {ext!r}. "
            f"Supported: {sorted(_IMAGE_EXTS | {_PDF_EXT})}"
        )

    if not images:
        raise ValueError(f"No pages/images could be loaded from {input_path}")

    # Process each page and accumulate into one combined grid
    combined_grid = _process_pages(images, debug=debug)

    if debug:
        _save_debug_preview(images, combined_grid, output_path)
        _save_debug_artifacts(input_path, combined_grid, output_path)

    # Export as HTML table
    logger.info("Using tenancy schedule parser for HTML table output")
    from .tenancy_parser import export_tenancy_to_html, parse_grid_to_rows

    tenancy_rows = parse_grid_to_rows(combined_grid)
    export_tenancy_to_html(tenancy_rows, output_path)

    return output_path.resolve()


# ---------------------------------------------------------------------------
# Page loading
# ---------------------------------------------------------------------------


def _load_pdf(pdf_path: Path, debug: bool = False) -> List[Image.Image]:
    """Load all pages of a PDF as PIL Images.

    Uses pdfplumber first to check whether the PDF has native text.
    For each page, if native text is found it renders a high-resolution
    raster (so the same table-detection pipeline is used uniformly).
    """
    try:
        from pdf2image import convert_from_path  # noqa: PLC0415

        dpi = 200
        images = convert_from_path(str(pdf_path), dpi=dpi)
        if debug:
            logger.debug("Loaded %d PDF page(s) at %d DPI", len(images), dpi)
        return images
    except Exception as exc:  # noqa: BLE001
        logger.warning("pdf2image failed (%s); trying pdfplumber fallback", exc)
        return _load_pdf_pdfplumber(pdf_path, debug=debug)


def _load_pdf_pdfplumber(pdf_path: Path, debug: bool = False) -> List[Image.Image]:
    """Fallback: render PDF pages via pdfplumber (uses pypdfium2 internally)."""
    import pdfplumber  # noqa: PLC0415

    images: list[Image.Image] = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            img = page.to_image(resolution=200).original
            images.append(img.convert("RGB"))
    if debug:
        logger.debug("pdfplumber loaded %d page(s)", len(images))
    return images


# ---------------------------------------------------------------------------
# Per-page processing
# ---------------------------------------------------------------------------


def _process_pages(
    images: List[Image.Image],
    debug: bool = False,
) -> TableGrid:
    """Process all pages and return a single combined :class:`TableGrid`.

    For multi-page tables the grids are concatenated vertically: subsequent
    pages are appended below the first page's rows.  The header is taken
    from the first page only.
    """
    combined: Optional[TableGrid] = None
    row_offset = 0

    for page_num, image in enumerate(images):
        logger.info("Processing page %d / %d …", page_num + 1, len(images))

        # Pre-process
        clean = preprocess(image, debug=debug)

        # Detect table structure
        grid = detect_table(clean, debug=debug)

        # Detect merged header cells
        grid = detect_merges(clean, grid, debug=debug)

        # OCR each cell
        _ocr_grid(clean, grid, debug=debug)

        if combined is None:
            combined = grid
            row_offset = grid.num_rows
        else:
            # Append rows from subsequent pages (skip header repeat)
            skip_header = grid.header_rows if page_num > 0 else 0
            for cell in grid.cells:
                if cell.row < skip_header:
                    continue
                cell.row = cell.row - skip_header + row_offset
                combined.cells.append(cell)
            row_offset = combined.num_rows

    return combined if combined is not None else TableGrid()


# ---------------------------------------------------------------------------
# OCR pass
# ---------------------------------------------------------------------------


def _ocr_grid(
    image: Image.Image,
    grid: TableGrid,
    debug: bool = False,
) -> None:
    """OCR each cell in *grid* and populate ``cell.text`` in-place."""
    img_array = image  # PIL image; we crop below

    for cell in grid.cells:
        left, top, right, bottom = cell.bbox
        # Guard against zero-size or out-of-bounds crops
        width, height = image.size
        left = max(0, left)
        top = max(0, top)
        right = min(width, right)
        bottom = min(height, bottom)

        if right <= left or bottom <= top:
            cell.text = ""
            continue

        crop = img_array.crop((left, top, right, bottom))
        result = ocr_cell(crop)
        cell.text = result.text.strip()
        cell.low_confidence = result.low_confidence

        if debug and cell.text:
            logger.debug(
                "Cell (%d,%d) conf=%.1f text=%r",
                cell.row,
                cell.col,
                result.confidence,
                cell.text[:40],
            )


# ---------------------------------------------------------------------------
# Debug preview
# ---------------------------------------------------------------------------


def _save_debug_preview(
    images: List[Image.Image],
    grid: TableGrid,
    output_path: Path,
) -> None:
    """Save an annotated preview image showing detected cell bounding boxes."""
    try:
        import cv2  # noqa: PLC0415
        import numpy as np  # noqa: PLC0415

        # Use the first page only for the preview
        image = images[0]
        arr = cv2.cvtColor(np.array(image.convert("RGB")), cv2.COLOR_RGB2BGR)

        for cell in grid.cells:
            if cell.row >= (grid.header_rows * len(images)):
                break
            l, t, r, b = cell.bbox
            color = (0, 0, 200) if cell.row < grid.header_rows else (0, 180, 0)
            cv2.rectangle(arr, (l, t), (r, b), color, 2)
            label = f"({cell.row},{cell.col})"
            cv2.putText(arr, label, (l + 2, t + 12), cv2.FONT_HERSHEY_SIMPLEX, 0.35, color, 1)

        preview_path = output_path.with_suffix(".debug.png")
        cv2.imwrite(str(preview_path), arr)
        logger.info("Debug preview saved to %s", preview_path)
    except Exception as exc:  # noqa: BLE001
        logger.warning("Could not save debug preview: %s", exc)


# ---------------------------------------------------------------------------
# Debug text artifacts
# ---------------------------------------------------------------------------


def _col_letter(n: int) -> str:
    """Convert 1-based column index to spreadsheet-style column letter (A, B, …, Z, AA, …)."""
    result = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(ord("A") + rem) + result
    return result


def _save_debug_artifacts(
    input_path: Path,
    grid: TableGrid,
    output_path: Path,
) -> None:
    """Write ``pipeline_diagram.md`` and ``grid_preview.txt`` to a debug folder.

    The debug folder is created next to the output file as
    ``<output_stem>.debug/``.
    """
    try:
        debug_dir = output_path.parent / f"{output_path.stem}.debug"
        debug_dir.mkdir(parents=True, exist_ok=True)

        merged_cells = [c for c in grid.cells if c.is_merged]
        low_conf_cells = [c for c in grid.cells if c.low_confidence]

        # ------------------------------------------------------------------ #
        # pipeline_diagram.md
        # ------------------------------------------------------------------ #
        diagram_path = debug_dir / "pipeline_diagram.md"
        diagram_lines = [
            "# OCR Extraction Pipeline Diagram",
            "",
            "```",
            f"  INPUT  ─────────────────────────────────────────────────",
            f"  File   : {input_path}",
            f"  ┌──────────────────────────────────────────────────────┐",
            f"  │ 1. Preprocessing (grayscale, denoise, threshold)     │",
            f"  └───────────────────────┬──────────────────────────────┘",
            f"                          ▼",
            f"  ┌──────────────────────────────────────────────────────┐",
            f"  │ 2. Table Detection (ruling lines / projection)       │",
            f"  │    columns detected : {grid.num_cols:<6}                      │",
            f"  │    rows    detected : {grid.num_rows:<6}                      │",
            f"  └───────────────────────┬──────────────────────────────┘",
            f"                          ▼",
            f"  ┌──────────────────────────────────────────────────────┐",
            f"  │ 3. OCR (pytesseract per cell)                        │",
            f"  │    total cells  : {len(grid.cells):<6}                        │",
            f"  │    low-conf     : {len(low_conf_cells):<6}                        │",
            f"  └───────────────────────┬──────────────────────────────┘",
            f"                          ▼",
            f"  ┌──────────────────────────────────────────────────────┐",
            f"  │ 4. Header / Merge Detection                          │",
            f"  │    header rows  : {grid.header_rows:<6}                        │",
            f"  │    merged cells : {len(merged_cells):<6}                        │",
            f"  └───────────────────────┬──────────────────────────────┘",
            f"                          ▼",
            f"  ┌──────────────────────────────────────────────────────┐",
            f"  │ 5. HTML Write                                        │",
            f"  │    sheet name   : Table                              │",
            f"  │    output       : {output_path.name:<34}│",
            f"  └──────────────────────────────────────────────────────┘",
            "```",
            "",
            "## Metrics",
            "",
            f"| Metric | Value |",
            f"|--------|-------|",
            f"| Input file | `{input_path.name}` |",
            f"| Columns | {grid.num_cols} |",
            f"| Rows | {grid.num_rows} |",
            f"| Header rows | {grid.header_rows} |",
            f"| Total cells | {len(grid.cells)} |",
            f"| Merged cells | {len(merged_cells)} |",
            f"| Low-confidence cells | {len(low_conf_cells)} |",
            "",
        ]
        if merged_cells:
            diagram_lines += [
                "## Merged Cell Ranges",
                "",
            ]
            for mc in merged_cells:
                start_col = _col_letter(mc.col + 1)
                end_col = _col_letter(mc.col + mc.colspan)
                start_row = mc.row + 1
                end_row = mc.row + mc.rowspan
                addr = f"{start_col}{start_row}:{end_col}{end_row}"
                diagram_lines.append(f"- `{addr}` — row={mc.row}, col={mc.col}, text={mc.text!r}")
            diagram_lines.append("")

        diagram_path.write_text("\n".join(diagram_lines), encoding="utf-8")
        logger.info("Debug diagram saved to %s", diagram_path)

        # ------------------------------------------------------------------ #
        # grid_preview.txt
        # ------------------------------------------------------------------ #
        _write_grid_preview(debug_dir / "grid_preview.txt", grid)

    except Exception as exc:  # noqa: BLE001
        logger.warning("Could not save debug artifacts: %s", exc)


def _write_grid_preview(
    preview_path: Path,
    grid: TableGrid,
    max_rows: int = 15,
    max_cols: int = 12,
) -> None:
    """Write a monospaced grid preview showing first *max_rows* rows."""
    if not grid.cells:
        preview_path.write_text("(empty grid)\n", encoding="utf-8")
        return

    n_rows = min(grid.num_rows, max_rows)
    n_cols = min(grid.num_cols, max_cols)

    # Build a 2-D text grid
    text_grid: list[list[str]] = [[""] * n_cols for _ in range(n_rows)]
    for cell in grid.cells:
        if cell.row < n_rows and cell.col < n_cols:
            text_grid[cell.row][cell.col] = cell.text

    # Compute column widths
    col_widths = [8] * n_cols
    for r in range(n_rows):
        for c in range(n_cols):
            col_widths[c] = max(col_widths[c], len(text_grid[r][c]))
    col_widths = [min(w, 25) for w in col_widths]

    def _row_sep(char: str = "-") -> str:
        return "+" + "+".join(char * (w + 2) for w in col_widths) + "+"

    def _row_line(row_idx: int, label: str = "") -> str:
        cells = []
        for c in range(n_cols):
            val = text_grid[row_idx][c][:col_widths[c]]
            cells.append(f" {val:<{col_widths[c]}} ")
        prefix = f"[{label}]" if label else f"[{row_idx:>3}]"
        return f"{prefix} |" + "|".join(cells) + "|"

    lines = [
        f"Grid Preview  ({grid.num_rows} rows × {grid.num_cols} cols, "
        f"showing first {n_rows}×{n_cols})",
        f"Header rows: {grid.header_rows}",
        "",
        _row_sep("="),
    ]
    for r in range(n_rows):
        label = "HDR" if r < grid.header_rows else f"{r:>3}"
        lines.append(_row_line(r, label))
        sep_char = "=" if r == grid.header_rows - 1 else "-"
        lines.append(_row_sep(sep_char))

    lines.append("")
    preview_path.write_text("\n".join(lines), encoding="utf-8")
    logger.info("Grid preview saved to %s", preview_path)

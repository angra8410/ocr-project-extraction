"""Main extraction pipeline.

Orchestrates the full conversion flow:
    Input file (.jpg/.png/.tif/.pdf)
    → Image(s)
    → Pre-processing
    → Table detection
    → OCR of each cell
    → Merge detection
    → Excel writing
"""

from __future__ import annotations

import logging
import tempfile
from pathlib import Path
from typing import List, Optional

from PIL import Image

from .excel_writer import write_xlsx
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
    """Convert *input_path* to a layout-preserving .xlsx file.

    Parameters
    ----------
    input_path:
        Path to a .jpg/.jpeg/.png/.tif/.tiff or .pdf file.
    output_path:
        Destination .xlsx path.  Defaults to the input file name with
        ``.xlsx`` extension in the same directory.
    debug:
        When *True* extra diagnostic information is logged and (for image
        inputs) an annotated preview PNG is saved alongside the output.

    Returns
    -------
    Path
        Resolved path to the written .xlsx file.
    """
    input_path = Path(input_path)
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    if output_path is None:
        output_path = input_path.with_suffix(".xlsx")
    output_path = Path(output_path)

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

    write_xlsx(combined_grid, output_path)
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

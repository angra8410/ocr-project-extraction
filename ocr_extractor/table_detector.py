"""Table structure detection.

Analyses a pre-processed PIL image (or a pdfplumber page) to find:
- Column boundaries
- Row boundaries
- Merged cell regions (especially header merges)

The result is a :class:`TableGrid` – a 2-D list of :class:`CellRegion`
objects that the excel_writer can turn directly into an openpyxl worksheet.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from typing import List, Optional, Sequence, Tuple

import cv2
import numpy as np
from PIL import Image

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Data models
# ---------------------------------------------------------------------------


@dataclass
class CellRegion:
    """A single table cell, possibly spanning multiple rows/columns."""

    row: int  # 0-based top-left row
    col: int  # 0-based top-left column
    rowspan: int = 1
    colspan: int = 1
    # Pixel bounding box in the *source* image (left, top, right, bottom)
    bbox: Tuple[int, int, int, int] = field(default_factory=lambda: (0, 0, 0, 0))
    text: str = ""
    low_confidence: bool = False

    @property
    def is_merged(self) -> bool:
        return self.rowspan > 1 or self.colspan > 1


@dataclass
class TableGrid:
    """Complete table structure returned by the detector."""

    cells: List[CellRegion] = field(default_factory=list)
    header_rows: int = 1  # number of rows that form the header band

    @property
    def num_rows(self) -> int:
        if not self.cells:
            return 0
        return max(c.row + c.rowspan for c in self.cells)

    @property
    def num_cols(self) -> int:
        if not self.cells:
            return 0
        return max(c.col + c.colspan for c in self.cells)


# ---------------------------------------------------------------------------
# Main detector
# ---------------------------------------------------------------------------


def detect_table(
    image: Image.Image,
    debug: bool = False,
) -> TableGrid:
    """Detect table structure in a PIL image.

    Strategy:
    1. Convert to grayscale + threshold.
    2. Detect horizontal and vertical ruling lines with morphology.
    3. Intersect them to find cell bounding boxes.
    4. Fall back to whitespace-gap analysis when no lines are found.

    Returns
    -------
    TableGrid
        Detected grid with bbox info for each cell.  Text fields are
        empty at this stage – they are filled by the OCR pass in
        :mod:`extractor`.
    """
    arr = cv2.cvtColor(np.array(image.convert("RGB")), cv2.COLOR_RGB2BGR)
    gray = cv2.cvtColor(arr, cv2.COLOR_BGR2GRAY)
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)

    h_lines, v_lines = _detect_ruling_lines(binary, debug=debug)

    if h_lines is not None and v_lines is not None:
        grid = _grid_from_lines(h_lines, v_lines, gray.shape, debug=debug)
    else:
        if debug:
            logger.debug("No ruling lines found; falling back to whitespace analysis")
        grid = _grid_from_whitespace(gray, debug=debug)

    return grid


# ---------------------------------------------------------------------------
# Ruling-line detection
# ---------------------------------------------------------------------------


def _detect_ruling_lines(
    binary: np.ndarray,
    debug: bool = False,
) -> Tuple[Optional[np.ndarray], Optional[np.ndarray]]:
    """Return sorted arrays of (y-coord) and (x-coord) ruling lines.

    Returns (None, None) if insufficient lines are detected.
    """
    height, width = binary.shape

    # Horizontal lines: long thin rectangles
    h_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (max(width // 5, 40), 1))
    h_lines_img = cv2.morphologyEx(binary, cv2.MORPH_OPEN, h_kernel, iterations=2)

    # Vertical lines: tall thin rectangles
    v_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(height // 5, 40)))
    v_lines_img = cv2.morphologyEx(binary, cv2.MORPH_OPEN, v_kernel, iterations=2)

    h_coords = _line_positions(h_lines_img, axis=0)
    v_coords = _line_positions(v_lines_img, axis=1)

    if debug:
        logger.debug(
            "Detected %d horizontal lines, %d vertical lines",
            len(h_coords),
            len(v_coords),
        )

    if len(h_coords) < 2 or len(v_coords) < 2:
        return None, None

    return np.array(h_coords), np.array(v_coords)


def _line_positions(line_img: np.ndarray, axis: int) -> list[int]:
    """Collapse *line_img* along *axis* and return sorted peak positions.

    axis=0 → find y-positions of horizontal lines
    axis=1 → find x-positions of vertical lines
    """
    # axis=0 → find y-positions (sum each row = numpy axis=1)
    # axis=1 → find x-positions (sum each col = numpy axis=0)
    projection = line_img.sum(axis=1 - axis)
    threshold = projection.max() * 0.3
    above = projection > threshold

    positions: list[int] = []
    in_run = False
    run_start = 0
    for i, val in enumerate(above):
        if val and not in_run:
            in_run = True
            run_start = i
        elif not val and in_run:
            in_run = False
            positions.append((run_start + i) // 2)
    if in_run:
        positions.append((run_start + len(above)) // 2)

    return positions


# ---------------------------------------------------------------------------
# Grid building from ruling lines
# ---------------------------------------------------------------------------


def _grid_from_lines(
    h_lines: np.ndarray,
    v_lines: np.ndarray,
    img_shape: Tuple[int, int],
    debug: bool = False,
) -> TableGrid:
    """Build a TableGrid whose cells are the intersections of lines."""
    # Add image-edge sentinel lines if not already present
    height, width = img_shape
    if h_lines[0] > 5:
        h_lines = np.concatenate([[0], h_lines])
    if h_lines[-1] < height - 5:
        h_lines = np.concatenate([h_lines, [height]])
    if v_lines[0] > 5:
        v_lines = np.concatenate([[0], v_lines])
    if v_lines[-1] < width - 5:
        v_lines = np.concatenate([v_lines, [width]])

    cells: list[CellRegion] = []
    for r, (y_top, y_bot) in enumerate(zip(h_lines, h_lines[1:])):
        for c, (x_left, x_right) in enumerate(zip(v_lines, v_lines[1:])):
            cells.append(
                CellRegion(
                    row=r,
                    col=c,
                    bbox=(int(x_left), int(y_top), int(x_right), int(y_bot)),
                )
            )

    n_rows = len(h_lines) - 1
    header_rows = _estimate_header_rows(h_lines, n_rows)

    if debug:
        logger.debug(
            "Grid: %d rows × %d cols, header_rows=%d",
            n_rows,
            len(v_lines) - 1,
            header_rows,
        )

    return TableGrid(cells=cells, header_rows=header_rows)


def _estimate_header_rows(h_lines: np.ndarray, n_rows: int) -> int:
    """Guess how many rows form the header band.

    Heuristic: header rows are often shorter than data rows.  Look for
    the first inter-row gap that is notably larger than the previous ones.
    Default is 1.
    """
    if n_rows <= 1:
        return 1
    gaps = np.diff(h_lines).astype(float)
    if len(gaps) < 2:
        return 1
    # Find first gap that is >= 1.5× the first gap
    first_gap = gaps[0]
    for i, g in enumerate(gaps[1:], start=1):
        if g >= first_gap * 1.5:
            return i
    return 1


# ---------------------------------------------------------------------------
# Whitespace-based fallback
# ---------------------------------------------------------------------------


def _grid_from_whitespace(
    gray: np.ndarray,
    debug: bool = False,
) -> TableGrid:
    """Detect rows and columns from whitespace gaps when no lines exist."""
    h_positions = _whitespace_separators(gray, axis=1)  # horizontal separators
    v_positions = _whitespace_separators(gray, axis=0)  # vertical separators

    height, width = gray.shape

    # Ensure boundary sentinels
    if not h_positions or h_positions[0] > 5:
        h_positions = [0] + h_positions
    if not h_positions or h_positions[-1] < height - 5:
        h_positions = h_positions + [height]
    if not v_positions or v_positions[0] > 5:
        v_positions = [0] + v_positions
    if not v_positions or v_positions[-1] < width - 5:
        v_positions = v_positions + [width]

    cells: list[CellRegion] = []
    for r, (y_top, y_bot) in enumerate(zip(h_positions, h_positions[1:])):
        for c, (x_left, x_right) in enumerate(zip(v_positions, v_positions[1:])):
            cells.append(
                CellRegion(
                    row=r,
                    col=c,
                    bbox=(x_left, y_top, x_right, y_bot),
                )
            )

    n_rows = len(h_positions) - 1
    if debug:
        logger.debug(
            "Whitespace grid: %d rows × %d cols",
            n_rows,
            len(v_positions) - 1,
        )

    return TableGrid(cells=cells, header_rows=1)


def _whitespace_separators(gray: np.ndarray, axis: int) -> list[int]:
    """Find separator positions by looking for all-white stripes.

    axis=0 → scan columns (find vertical separators → x positions)
    axis=1 → scan rows    (find horizontal separators → y positions)
    """
    # Project to find dark-pixel counts per row/column
    projection = (gray < 200).sum(axis=axis)
    is_empty = projection == 0

    # Cluster consecutive empty rows/cols into single separator positions
    separators: list[int] = []
    in_gap = False
    gap_start = 0
    for i, empty in enumerate(is_empty):
        if empty and not in_gap:
            in_gap = True
            gap_start = i
        elif not empty and in_gap:
            in_gap = False
            if (i - gap_start) >= 3:  # require at least 3px gap
                separators.append((gap_start + i) // 2)
    return separators


# ---------------------------------------------------------------------------
# Merge detection (for header rows)
# ---------------------------------------------------------------------------


def detect_merges(
    image: Image.Image,
    grid: TableGrid,
    debug: bool = False,
) -> TableGrid:
    """Identify merged cells in the header area and update *grid* in-place.

    Strategy: for each pair of horizontally adjacent header cells, check
    whether the pixel column between them contains any dark vertical-line
    pixels.  If the dividing line is absent, the two cells are merged.

    Returns the same *grid* object with merged :class:`CellRegion` objects
    replacing the original pair.
    """
    arr = np.array(image.convert("L"))
    _, binary = cv2.threshold(arr, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)

    if not grid.cells:
        return grid

    num_cols = grid.num_cols

    # Build a lookup: (row, col) → CellRegion
    cell_map: dict[tuple[int, int], CellRegion] = {
        (c.row, c.col): c for c in grid.cells
    }

    for row_idx in range(grid.header_rows):
        col_idx = 0
        while col_idx < num_cols:
            cell = cell_map.get((row_idx, col_idx))
            if cell is None:
                col_idx += 1
                continue

            # Try to extend the merge rightward
            while cell.col + cell.colspan < num_cols:
                next_cell = cell_map.get((row_idx, cell.col + cell.colspan))
                if next_cell is None:
                    break
                # Check if the vertical divider between them is absent
                divider_x = cell.bbox[2]  # right edge of current cell
                y_top = cell.bbox[1]
                y_bot = cell.bbox[3]
                strip = binary[y_top:y_bot, max(0, divider_x - 2) : divider_x + 3]
                dark_pixels = strip.sum() // 255
                if dark_pixels < (y_bot - y_top) * 0.3:
                    # No divider → merge
                    if debug:
                        logger.debug(
                            "Merging header cell (%d,%d) with (%d,%d)",
                            row_idx,
                            cell.col,
                            row_idx,
                            next_cell.col,
                        )
                    cell.colspan += next_cell.colspan
                    cell.bbox = (
                        cell.bbox[0],
                        cell.bbox[1],
                        next_cell.bbox[2],
                        next_cell.bbox[3],
                    )
                    # Remove the absorbed cell
                    grid.cells.remove(next_cell)
                    del cell_map[(row_idx, next_cell.col)]
                else:
                    break

            col_idx = cell.col + cell.colspan

    return grid

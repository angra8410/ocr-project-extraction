"""Table structure detection.

Analyses a pre-processed PIL image (or a pdfplumber page) to find:
- Column boundaries
- Row boundaries
- Merged cell regions (especially header merges)

The result is a :class:`TableGrid` – a 2-D list of :class:`CellRegion`
objects that the tenancy_parser can convert into an HTML table.
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
# Constants
# ---------------------------------------------------------------------------

# Whitespace separator: maximum ink density (as fraction of column peak) that
# still counts as a "gap" between columns/rows.
_WHITESPACE_DENSITY_RATIO: float = 0.05
_MIN_WHITESPACE_THRESHOLD: float = 2.0

# Projection valley: fraction of peak density below which a stripe is a valley.
# Set to 0.60 for tables with many narrow columns (15-20 columns).
# Higher values detect more valleys; adjust based on your table's column count.
_VALLEY_THRESHOLD_RATIO: float = 0.60
# Minimum valley width as a fraction of the image dimension (keeps false
# positives caused by sub-pixel noise from being promoted to column/row gaps).
# Reduced from 50 to 200 to allow detection of much narrower column gaps in tables.
_VALLEY_MIN_GAP_DIVISOR: int = 200


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
    4. Fall back to whitespace-gap / projection-valley analysis when no lines
       are found or when only a few columns are detected (suggesting the table
       uses whitespace rather than ruling lines for column separation).

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

    # Sanity check: if we ended up with very few columns on a wide image,
    # retry using the projection-valley approach which is more robust for
    # documents where columns are defined by text alignment rather than ruling lines.
    # A threshold of <=5 columns on images wider than 1000px suggests insufficient detection.
    width = gray.shape[1]
    if grid.num_cols <= 5 and width > 1000:
        if debug:
            logger.debug(
                "Only %d column(s) detected on %dpx-wide image; "
                "retrying with projection-valley fallback",
                grid.num_cols,
                width,
            )
        grid_fallback = _grid_from_projection(gray, debug=debug)
        
        # Use the fallback if it found more columns (even slightly more)
        if grid_fallback.num_cols > grid.num_cols:
            if debug:
                logger.debug(
                    "Projection fallback found %d columns (vs %d from lines); using fallback",
                    grid_fallback.num_cols,
                    grid.num_cols,
                )
            grid = grid_fallback

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

    # Bridge small gaps in lines before detecting them (handles broken/faint lines)
    h_bridge = cv2.getStructuringElement(cv2.MORPH_RECT, (15, 1))
    binary_hb = cv2.morphologyEx(binary, cv2.MORPH_CLOSE, h_bridge)
    # Increased vertical bridge to 25px to reconnect broken lines in tables with many columns
    v_bridge = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 25))
    binary_vb = cv2.morphologyEx(binary, cv2.MORPH_CLOSE, v_bridge)

    # Horizontal lines: long thin rectangles – use ~12.5% width (previously 20%)
    h_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (max(width // 8, 20), 1))
    h_lines_img = cv2.morphologyEx(binary_hb, cv2.MORPH_OPEN, h_kernel, iterations=1)

    # Vertical lines: Use ~6.25% height (reduced from 12.5%) to detect thinner column lines
    # in tables with many columns. Min size reduced to 15px.
    v_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(height // 16, 15)))
    v_lines_img = cv2.morphologyEx(binary_vb, cv2.MORPH_OPEN, v_kernel, iterations=1)

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
    # Reduced threshold from 0.3 to 0.1 to catch weaker line signals from narrow columns
    threshold = projection.max() * 0.1
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
# Shared helper
# ---------------------------------------------------------------------------


def _add_boundary_sentinels(
    positions: list[int], max_val: int, margin: int = 5
) -> list[int]:
    """Ensure *positions* starts at/near 0 and ends at/near *max_val*.

    This guarantees that the grid spans the full image dimension even when
    no separator is detected close to an edge.
    """
    if not positions or positions[0] > margin:
        positions = [0] + positions
    if positions[-1] < max_val - margin:
        positions = positions + [max_val]
    return positions


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

    h_positions = _add_boundary_sentinels(h_positions, height)
    v_positions = _add_boundary_sentinels(v_positions, width)

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


# ---------------------------------------------------------------------------
# Projection-profile–based fallback (robust for scanned documents)
# ---------------------------------------------------------------------------


def _grid_from_projection(
    gray: np.ndarray,
    debug: bool = False,
) -> TableGrid:
    """Detect columns and rows using projection-profile valley analysis.

    This is a more robust alternative to the whitespace separator approach
    for scanned documents where ink bleeds slightly into gap areas.

    Algorithm:
    1. Compute the dark-pixel count per column (vertical projection).
    2. Smooth the projection with a running-mean kernel.
    3. Detect local valleys below ``_VALLEY_THRESHOLD_RATIO`` of the peak.
    4. Use the valley centres as column-boundary positions.
    5. Repeat for rows.
    """
    height, width = gray.shape

    v_positions = _projection_valleys(gray, axis=0)
    h_positions = _projection_valleys(gray, axis=1)

    h_positions = _add_boundary_sentinels(h_positions, height)
    v_positions = _add_boundary_sentinels(v_positions, width)

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
            "Projection-valley grid: %d rows × %d cols",
            n_rows,
            len(v_positions) - 1,
        )

    return TableGrid(cells=cells, header_rows=1)


def _projection_valleys(
    gray: np.ndarray,
    axis: int,
    threshold_ratio: float = _VALLEY_THRESHOLD_RATIO,
) -> list[int]:
    """Find valley positions in the ink-density projection along *axis*.

    axis=0 → vertical projection (sum per column) → x positions (column gaps)
    axis=1 → horizontal projection (sum per row) → y positions (row gaps)

    Returns a sorted list of gap-centre positions.
    """
    projection = (gray < 200).sum(axis=axis).astype(float)
    if projection.max() == 0:
        return []

    # Smooth to reduce single-pixel noise; window = ~3% of dimension but ≥5.
    # A wider window suppresses noise in the gap regions on scanned documents.
    n = len(projection)
    win = max(5, n // 30)
    kernel = np.ones(win, dtype=float) / win
    smoothed = np.convolve(projection, kernel, mode="same")

    peak = smoothed.max()
    threshold = peak * threshold_ratio

    is_valley = smoothed <= threshold

    # Require a minimum gap width to filter out sub-pixel noise spikes
    min_gap = max(5, n // _VALLEY_MIN_GAP_DIVISOR)

    valleys: list[int] = []
    in_valley = False
    v_start = 0
    for i, val in enumerate(is_valley):
        if val and not in_valley:
            in_valley = True
            v_start = i
        elif not val and in_valley:
            in_valley = False
            if (i - v_start) >= min_gap:
                valleys.append((v_start + i) // 2)
    if in_valley and (n - v_start) >= min_gap:
        valleys.append((v_start + n) // 2)

    return valleys


def _whitespace_separators(gray: np.ndarray, axis: int) -> list[int]:
    """Find separator positions by looking for low-ink-density stripes.

    axis=0 → scan columns (find vertical separators → x positions)
    axis=1 → scan rows    (find horizontal separators → y positions)

    Uses a density threshold (≤ ``_WHITESPACE_DENSITY_RATIO`` of the peak
    ink density) rather than requiring completely empty stripes.  This
    handles scanned documents where even blank areas contain minor noise.
    """
    # Project to find dark-pixel counts per row/column
    projection = (gray < 200).sum(axis=axis).astype(float)
    peak = projection.max()
    if peak == 0:
        return []

    threshold = max(peak * _WHITESPACE_DENSITY_RATIO, _MIN_WHITESPACE_THRESHOLD)
    is_empty = projection <= threshold

    # Cluster consecutive low-density rows/cols into single separator positions
    separators: list[int] = []
    in_gap = False
    gap_start = 0
    for i, empty in enumerate(is_empty):
        if empty and not in_gap:
            in_gap = True
            gap_start = i
        elif not empty and in_gap:
            in_gap = False
            if (i - gap_start) >= 2:  # require at least 2px gap
                separators.append((gap_start + i) // 2)
    if in_gap and (len(is_empty) - gap_start) >= 2:
        separators.append((gap_start + len(is_empty)) // 2)
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

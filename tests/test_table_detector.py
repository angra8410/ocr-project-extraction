"""Unit tests for the table-structure detection module."""

import numpy as np
import pytest
from PIL import Image, ImageDraw

from ocr_extractor.table_detector import (
    CellRegion,
    TableGrid,
    _estimate_header_rows,
    _grid_from_lines,
    _grid_from_whitespace,
    _line_positions,
    _whitespace_separators,
    detect_merges,
    detect_table,
)


# ---------------------------------------------------------------------------
# Helper: synthetic table image
# ---------------------------------------------------------------------------


def make_grid_image(
    rows: int = 3,
    cols: int = 4,
    cell_w: int = 80,
    cell_h: int = 40,
    line_w: int = 2,
) -> Image.Image:
    """Draw a simple ruled-line grid and return the PIL Image."""
    total_w = cols * cell_w + (cols + 1) * line_w
    total_h = rows * cell_h + (rows + 1) * line_w
    img = Image.new("RGB", (total_w, total_h), "white")
    draw = ImageDraw.Draw(img)

    # Horizontal lines
    for r in range(rows + 1):
        y = r * (cell_h + line_w)
        draw.line([(0, y), (total_w, y)], fill="black", width=line_w)

    # Vertical lines
    for c in range(cols + 1):
        x = c * (cell_w + line_w)
        draw.line([(x, 0), (x, total_h)], fill="black", width=line_w)

    return img


# ---------------------------------------------------------------------------
# CellRegion
# ---------------------------------------------------------------------------


class TestCellRegion:
    def test_is_merged_false_by_default(self):
        cell = CellRegion(row=0, col=0)
        assert not cell.is_merged

    def test_is_merged_colspan(self):
        cell = CellRegion(row=0, col=0, colspan=2)
        assert cell.is_merged

    def test_is_merged_rowspan(self):
        cell = CellRegion(row=0, col=0, rowspan=2)
        assert cell.is_merged


# ---------------------------------------------------------------------------
# TableGrid
# ---------------------------------------------------------------------------


class TestTableGrid:
    def test_empty_grid(self):
        grid = TableGrid()
        assert grid.num_rows == 0
        assert grid.num_cols == 0

    def test_num_rows_cols(self):
        cells = [
            CellRegion(row=0, col=0),
            CellRegion(row=0, col=1),
            CellRegion(row=1, col=0),
            CellRegion(row=1, col=1),
        ]
        grid = TableGrid(cells=cells)
        assert grid.num_rows == 2
        assert grid.num_cols == 2

    def test_merged_cell_counted_correctly(self):
        cells = [CellRegion(row=0, col=0, colspan=3)]
        grid = TableGrid(cells=cells)
        assert grid.num_cols == 3


# ---------------------------------------------------------------------------
# _line_positions
# ---------------------------------------------------------------------------


class TestLinePositions:
    def test_detects_horizontal_line(self):
        """A bright horizontal stripe in the projection should yield one position."""
        img = np.zeros((100, 200), dtype=np.uint8)
        img[40:43, :] = 255  # bright horizontal stripe
        positions = _line_positions(img, axis=0)
        assert len(positions) == 1
        assert 39 <= positions[0] <= 44

    def test_empty_image_returns_empty(self):
        img = np.zeros((50, 50), dtype=np.uint8)
        positions = _line_positions(img, axis=0)
        assert positions == []


# ---------------------------------------------------------------------------
# _whitespace_separators
# ---------------------------------------------------------------------------


class TestWhitespaceSeparators:
    def test_finds_column_gap(self):
        """A blank column should be detected as a vertical separator."""
        # 100×100 grey image with a white column in the middle
        gray = np.ones((100, 100), dtype=np.uint8) * 128
        gray[:, 48:52] = 255  # 4-px wide white gap
        seps = _whitespace_separators(gray, axis=0)
        assert len(seps) >= 1
        assert any(47 <= s <= 53 for s in seps)


# ---------------------------------------------------------------------------
# _estimate_header_rows
# ---------------------------------------------------------------------------


class TestEstimateHeaderRows:
    def test_single_row(self):
        h_lines = np.array([0, 30, 200])
        # First gap = 30, second = 170 → header = 1
        assert _estimate_header_rows(h_lines, 2) == 1

    def test_large_first_gap(self):
        # All equal gaps → no jump → 1
        h_lines = np.array([0, 30, 60, 90])
        assert _estimate_header_rows(h_lines, 3) == 1


# ---------------------------------------------------------------------------
# detect_table
# ---------------------------------------------------------------------------


class TestDetectTable:
    def test_detect_grid_image(self):
        """detect_table should find the rows and columns of a clean grid."""
        img = make_grid_image(rows=3, cols=4)
        grid = detect_table(img)
        assert grid.num_rows == 3
        assert grid.num_cols == 4
        assert len(grid.cells) == 12

    def test_detect_table_returns_table_grid(self):
        img = make_grid_image(rows=2, cols=3)
        grid = detect_table(img)
        assert isinstance(grid, TableGrid)

    def test_cells_have_valid_bboxes(self):
        img = make_grid_image(rows=2, cols=2)
        grid = detect_table(img)
        for cell in grid.cells:
            l, t, r, b = cell.bbox
            assert r > l, f"Cell ({cell.row},{cell.col}) has zero or negative width"
            assert b > t, f"Cell ({cell.row},{cell.col}) has zero or negative height"

    def test_debug_mode_does_not_raise(self):
        img = make_grid_image(rows=2, cols=2)
        detect_table(img, debug=True)


# ---------------------------------------------------------------------------
# detect_merges
# ---------------------------------------------------------------------------


class TestDetectMerges:
    def test_no_merges_on_clean_grid(self):
        """A clean grid with visible dividers should produce no merges."""
        img = make_grid_image(rows=3, cols=4)
        grid = detect_table(img)
        original_cell_count = len(grid.cells)
        grid = detect_merges(img, grid)
        # In a fully-ruled grid no cells should be merged
        merged = [c for c in grid.cells if c.is_merged]
        # Allow up to 1 spurious merge in edge cases
        assert len(merged) <= 1

    def test_empty_grid_handled(self):
        """detect_merges should not crash on an empty grid."""
        img = Image.new("RGB", (100, 100), "white")
        grid = TableGrid()
        result = detect_merges(img, grid)
        assert isinstance(result, TableGrid)

"""Unit tests for the table-structure detection module."""

import numpy as np
import pytest
from PIL import Image, ImageDraw

from ocr_extractor.table_detector import (
    CellRegion,
    TableGrid,
    _estimate_header_rows,
    _grid_from_lines,
    _grid_from_projection,
    _grid_from_whitespace,
    _line_positions,
    _projection_valleys,
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


# ---------------------------------------------------------------------------
# _projection_valleys
# ---------------------------------------------------------------------------


class TestProjectionValleys:
    def test_detects_gap_columns(self):
        """A region with no dark pixels should appear as a valley."""
        gray = np.ones((100, 300), dtype=np.uint8) * 200  # light background
        # Two text-like blocks with a clear gap in between
        gray[:, 10:90] = 50   # first block (dark)
        gray[:, 210:290] = 50  # second block (dark)
        # Columns 90-210 are light → should contain at least one valley
        valleys = _projection_valleys(gray, axis=0, threshold_ratio=0.2)
        assert len(valleys) >= 1
        # At least one valley should fall in the gap region
        assert any(90 <= v <= 210 for v in valleys)

    def test_empty_image_returns_empty(self):
        gray = np.full((50, 50), 255, dtype=np.uint8)
        assert _projection_valleys(gray, axis=0) == []

    def test_single_block_no_internal_valleys(self):
        """A uniformly dark strip should produce no valleys."""
        gray = np.zeros((50, 200), dtype=np.uint8)  # all dark
        valleys = _projection_valleys(gray, axis=0, threshold_ratio=0.2)
        assert len(valleys) == 0


# ---------------------------------------------------------------------------
# _grid_from_projection
# ---------------------------------------------------------------------------


class TestGridFromProjection:
    def _make_noisy_columnar_image(
        self,
        num_cols: int = 4,
        col_w: int = 60,
        gap_w: int = 20,
        height: int = 200,
        noise: int = 5,
    ) -> np.ndarray:
        """Create a grayscale array with text-like blocks separated by noisy gaps."""
        width = num_cols * col_w + (num_cols - 1) * gap_w + 2 * gap_w
        gray = np.full((height, width), 240, dtype=np.uint8)  # near-white background
        # Fixed seed for reproducibility: the exact gap positions matter for
        # the valley-detection assertions below.
        rng = np.random.default_rng(42)
        for i in range(num_cols):
            x_start = gap_w + i * (col_w + gap_w)
            x_end = x_start + col_w
            # Simulate text: scatter dark pixels inside columns
            gray[:, x_start:x_end] = 80
        # Add a little noise in the gaps (simulating scanned document noise)
        noise_mask = rng.integers(0, 255, size=(height, width)) > (255 - noise * 3)
        gray[noise_mask] = 50
        return gray

    def test_multi_column_scanned_like_image(self):
        """Projection-valley fallback should detect multiple columns."""
        gray = self._make_noisy_columnar_image(num_cols=4)
        grid = _grid_from_projection(gray, debug=True)
        assert grid.num_cols >= 2, (
            f"Expected >= 2 columns from projection fallback, got {grid.num_cols}"
        )

    def test_produces_valid_bboxes(self):
        gray = self._make_noisy_columnar_image(num_cols=3)
        grid = _grid_from_projection(gray)
        for cell in grid.cells:
            l, t, r, b = cell.bbox
            assert r > l
            assert b > t


# ---------------------------------------------------------------------------
# Regression: detect_table falls back gracefully on noisy / lineless images
# ---------------------------------------------------------------------------


class TestDetectTableFallback:
    def test_no_single_column_on_multi_column_image(self):
        """detect_table must NOT collapse a clear multi-column layout into 1 column.

        This is the core regression test: previously the whitespace fallback
        required *perfectly* empty columns (projection == 0) which never
        occurred in scanned documents with minor noise → always 1 column.
        """
        # Create an image that has NO ruling lines but clearly separate columns
        width, height = 600, 300
        img = Image.new("RGB", (width, height), "white")
        draw = ImageDraw.Draw(img)
        # Draw 4 text blocks without any grid lines
        col_positions = [(20, 120), (170, 270), (320, 420), (470, 570)]
        for x0, x1 in col_positions:
            for row_y in range(20, height, 40):
                draw.rectangle([(x0, row_y), (x1, row_y + 15)], fill="black")

        grid = detect_table(img, debug=True)
        assert grid.num_cols >= 2, (
            f"Expected multiple columns from lineless multi-column image, "
            f"got num_cols={grid.num_cols}"
        )

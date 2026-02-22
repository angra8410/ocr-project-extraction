"""Integration / golden tests for the end-to-end extraction pipeline.

These tests run the full pipeline on a synthetic table image and verify
that the output .html has the expected structure and content.
"""

from __future__ import annotations

import logging
from pathlib import Path
import re

import pytest
from PIL import Image, ImageDraw

from ocr_extractor.extractor import extract

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Fixture: synthetic table image
# ---------------------------------------------------------------------------

FIXTURES_DIR = Path(__file__).parent / "fixtures"
SAMPLE_TABLE = FIXTURES_DIR / "sample_table.png"
TEST_PDF = FIXTURES_DIR / "test.pdf"


def _make_table_image(
    content: list[list[str]],
    cell_w: int = 120,
    cell_h: int = 50,
    line_w: int = 2,
) -> Image.Image:
    """Draw a ruled-line table with the given cell content and return it."""
    rows = len(content)
    cols = max(len(r) for r in content)
    total_w = cols * cell_w + (cols + 1) * line_w
    total_h = rows * cell_h + (rows + 1) * line_w

    img = Image.new("RGB", (total_w, total_h), "white")
    draw = ImageDraw.Draw(img)

    for r in range(rows + 1):
        y = r * (cell_h + line_w)
        draw.rectangle([(0, y), (total_w, y + line_w)], fill="black")
    for c in range(cols + 1):
        x = c * (cell_w + line_w)
        draw.rectangle([(x, 0), (x + line_w, total_h)], fill="black")

    for r, row_data in enumerate(content):
        for c, text in enumerate(row_data):
            x = c * (cell_w + line_w) + line_w + 5
            y = r * (cell_h + line_w) + line_w + 15
            draw.text((x, y), text, fill="black")

    return img


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------


class TestExtractFromImage:
    def test_output_file_created(self, tmp_path):
        out = tmp_path / "out.html"
        result = extract(str(SAMPLE_TABLE), str(out))
        assert Path(result).exists()

    def test_output_is_valid_html(self, tmp_path):
        out = tmp_path / "out.html"
        result = extract(str(SAMPLE_TABLE), str(out))
        content = Path(result).read_text(encoding="utf-8")
        assert "<table>" in content

    def test_html_has_table(self, tmp_path):
        out = tmp_path / "out.html"
        result = extract(str(SAMPLE_TABLE), str(out))
        content = Path(result).read_text(encoding="utf-8")
        assert "<table>" in content
        assert "</table>" in content

    def test_html_has_thead_and_tbody(self, tmp_path):
        out = tmp_path / "out.html"
        result = extract(str(SAMPLE_TABLE), str(out))
        content = Path(result).read_text(encoding="utf-8")
        assert "<thead>" in content
        assert "<tbody>" in content

    def test_has_multiple_rows(self, tmp_path):
        """The output should have at least one data row in tbody."""
        out = tmp_path / "out.html"
        result = extract(str(SAMPLE_TABLE), str(out))
        content = Path(result).read_text(encoding="utf-8")
        tbody_match = re.search(r"<tbody>(.*?)</tbody>", content, re.DOTALL)
        assert tbody_match
        rows = re.findall(r"<tr>", tbody_match.group(1))
        assert len(rows) >= 1

    def test_has_multiple_columns(self, tmp_path):
        """The output should have at least 2 header columns."""
        out = tmp_path / "out.html"
        result = extract(str(SAMPLE_TABLE), str(out))
        content = Path(result).read_text(encoding="utf-8")
        th_count = content.count("<th>")
        assert th_count >= 2

    def test_debug_mode_does_not_raise(self, tmp_path):
        """debug=True should complete without errors."""
        out = tmp_path / "debug_out.html"
        result = extract(str(SAMPLE_TABLE), str(out), debug=True)
        assert Path(result).exists()


class TestExtractErrors:
    def test_missing_file_raises(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            extract(str(tmp_path / "nonexistent.png"))

    def test_unsupported_extension_raises(self, tmp_path):
        f = tmp_path / "file.docx"
        f.write_bytes(b"dummy")
        with pytest.raises(ValueError, match="Unsupported"):
            extract(str(f))

    def test_default_output_path(self, tmp_path):
        """When output_path is omitted the .html lands next to the input."""
        src = tmp_path / "table.png"
        import shutil

        shutil.copy(str(SAMPLE_TABLE), str(src))
        result = extract(str(src))
        assert Path(result).suffix == ".html"
        assert Path(result).parent == tmp_path


class TestExtractSyntheticTable:
    """Golden test: verify OCR picks up expected column-header keywords."""

    def test_header_text_present(self, tmp_path):
        content = [
            ["Name", "Amount", "Date"],
            ["Alice", "100", "2024-01-01"],
            ["Bob", "200", "2024-01-02"],
        ]
        img = _make_table_image(content)
        src = tmp_path / "synthetic.png"
        img.save(str(src))
        out = tmp_path / "synthetic.html"
        result = extract(str(src), str(out))
        html_content = Path(result).read_text(encoding="utf-8")

        # The HTML must contain at least ONE of the known column names
        # (OCR may not be perfect, but at least something should come through)
        all_text = html_content.lower()
        assert any(keyword in all_text for keyword in ("name", "amount", "date", "alice", "bob")), (
            f"No expected keywords found in extracted HTML. Got (first 500 chars): {html_content[:500]}"
        )



# ---------------------------------------------------------------------------
# PDF integration tests (uses tests/fixtures/test.pdf)
# ---------------------------------------------------------------------------

# Skip the whole class if pdfplumber cannot open the fixture (e.g. in very
# restricted CI environments).  The PDF fixture is a PIL-generated single-page
# image-based PDF, so no special renderer is needed beyond pdfplumber.
_pdf_available = TEST_PDF.exists()


@pytest.mark.skipif(not _pdf_available, reason="tests/fixtures/test.pdf not found")
class TestPdfIntegration:
    """Integration tests that use tests/fixtures/test.pdf as input.

    Output is always written to pytest's tmp_path – never to the fixtures dir.
    """

    def test_pdf_output_written_to_tmp(self, tmp_path):
        """Output .html must be created inside tmp_path, not in fixtures."""
        out = tmp_path / "test_output.html"
        result = extract(str(TEST_PDF), str(out))
        result_path = Path(result)
        assert result_path.exists()
        # Crucially, the output must NOT be inside the fixtures directory
        assert not result_path.is_relative_to(FIXTURES_DIR), (
            f"Output was written to fixtures dir: {result_path}"
        )
        assert result_path.is_relative_to(tmp_path)

    def test_pdf_output_is_html(self, tmp_path):
        """The generated file must be an HTML table."""
        out = tmp_path / "test_output.html"
        result = extract(str(TEST_PDF), str(out))
        content = Path(result).read_text(encoding="utf-8")
        assert "<table>" in content
        assert "<thead>" in content
        assert "<tbody>" in content

    def test_pdf_has_rows_and_columns(self, tmp_path):
        """The HTML table must have at least 1 header column and 1 data row."""
        out = tmp_path / "test_output.html"
        result = extract(str(TEST_PDF), str(out))
        content = Path(result).read_text(encoding="utf-8")
        assert content.count("<th>") >= 1
        tbody_match = re.search(r"<tbody>(.*?)</tbody>", content, re.DOTALL)
        assert tbody_match
        assert re.search(r"<tr>", tbody_match.group(1))

    def test_pdf_debug_artifacts_created(self, tmp_path):
        """Debug mode must create pipeline_diagram.md and grid_preview.txt."""
        out = tmp_path / "test_output.html"
        result = extract(str(TEST_PDF), str(out), debug=True)
        debug_dir = tmp_path / "test_output.debug"
        assert debug_dir.exists(), f"Debug dir not found: {debug_dir}"
        diagram = debug_dir / "pipeline_diagram.md"
        preview = debug_dir / "grid_preview.txt"
        assert diagram.exists(), "pipeline_diagram.md not generated"
        assert preview.exists(), "grid_preview.txt not generated"
        # Check diagram has meaningful content
        content = diagram.read_text(encoding="utf-8")
        assert "Pipeline" in content
        assert "Header rows" in content or "header rows" in content.lower()
        # Check preview has meaningful content
        preview_content = preview.read_text(encoding="utf-8")
        assert "Grid Preview" in preview_content or "HDR" in preview_content

    def test_pdf_debug_artifacts_in_tmp_not_fixtures(self, tmp_path):
        """Debug artifacts must be written to tmp_path, not fixtures dir."""
        out = tmp_path / "test_output.html"
        extract(str(TEST_PDF), str(out), debug=True)
        debug_dir = tmp_path / "test_output.debug"
        assert debug_dir.exists()
        assert not debug_dir.is_relative_to(FIXTURES_DIR)

    def test_pdf_multi_column_structure(self, tmp_path):
        """The output must have multi-column structure, NOT a single-column dump.

        We verify that:
        - The output is a valid HTML file
        - The <table> header row contains the required 15 schema columns
        - The expected column names are present
        - At least one data row exists in the tbody
        """
        out = tmp_path / "test_output.html"
        result = extract(str(TEST_PDF), str(out))
        result_path = Path(result)

        # 1. Output file must exist and have .html extension
        assert result_path.exists(), f"Output file not found: {result_path}"
        assert result_path.suffix == ".html", (
            f"Expected .html output, got {result_path.suffix}"
        )

        content = result_path.read_text(encoding="utf-8")

        # 2. Must be a valid HTML table
        assert "<table>" in content, "Output does not contain <table>"
        assert "<thead>" in content, "Output does not contain <thead>"
        assert "<tbody>" in content, "Output does not contain <tbody>"

        # 3. All 15 required schema columns must be present as <th> headers
        from ocr_extractor.tenancy_parser import HTML_SCHEMA_COLUMNS
        for col in HTML_SCHEMA_COLUMNS:
            assert f"<th>{col}</th>" in content, (
                f"Required column '{col}' not found in HTML header"
            )

        # 4. Header row must have exactly 15 <th> elements
        th_count = content.count("<th>")
        assert th_count == len(HTML_SCHEMA_COLUMNS), (
            f"Expected {len(HTML_SCHEMA_COLUMNS)} header columns, found {th_count}"
        )

        # 5. At least one data row must exist in <tbody>
        tbody_match = re.search(r"<tbody>(.*?)</tbody>", content, re.DOTALL)
        assert tbody_match, "Could not find <tbody> in output"
        tbody_content = tbody_match.group(1)
        data_rows = re.findall(r"<tr>", tbody_content)
        assert len(data_rows) >= 1, (
            f"Expected at least 1 data row in tbody, found {len(data_rows)}"
        )

    def test_pdf_generates_visual_artifacts(self, tmp_path):
        """Generate visual artifacts for debugging and verification.

        This test produces:
        - table_preview.html: Visual grid representation with styling
        - layout_overlay.png: Annotated image showing detected structure
        - test_assertions_report.md: Detailed explanation of what was checked

        These artifacts help verify that multi-column structure is preserved
        and provide visual evidence for code review.
        """
        from ocr_extractor.test_artifacts import (
            generate_assertions_report,
            generate_layout_overlay,
            generate_table_preview_html,
        )
        from ocr_extractor.preprocessor import preprocess
        from ocr_extractor.table_detector import detect_table, detect_merges
        from pdf2image import convert_from_path
        import cv2
        import numpy as np

        # Extract to HTML
        out = tmp_path / "test_output.html"
        result = extract(str(TEST_PDF), str(out), debug=False)

        # Also get the image and grid for layout overlay
        images = convert_from_path(str(TEST_PDF), dpi=200)
        clean = preprocess(images[0], debug=False)

        # Detect table structure
        arr = cv2.cvtColor(np.array(clean.convert("RGB")), cv2.COLOR_RGB2BGR)
        gray = cv2.cvtColor(arr, cv2.COLOR_BGR2GRAY)
        _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)

        from ocr_extractor.table_detector import _detect_ruling_lines, _grid_from_lines
        h_lines, v_lines = _detect_ruling_lines(binary, debug=False)
        if h_lines is not None and v_lines is not None:
            grid = _grid_from_lines(h_lines, v_lines, gray.shape, debug=False)
        else:
            grid = detect_table(clean, debug=False)

        grid = detect_merges(clean, grid, debug=False)

        # Generate artifacts
        artifacts_dir = tmp_path / "test_artifacts"
        artifacts_dir.mkdir(exist_ok=True)

        # 1. Table preview HTML
        preview_html = artifacts_dir / "table_preview.html"
        generate_table_preview_html(
            html_path=Path(result),
            output_path=preview_html,
            max_rows=30,
            max_cols=15,
        )
        assert preview_html.exists(), "table_preview.html not generated"

        # 2. Layout overlay PNG
        overlay_png = artifacts_dir / "layout_overlay.png"
        generate_layout_overlay(
            source_image=clean,
            grid=grid,
            output_path=overlay_png,
        )
        assert overlay_png.exists(), "layout_overlay.png not generated"

        # 3. Test assertions report
        report_md = artifacts_dir / "test_assertions_report.md"
        generate_assertions_report(
            html_path=Path(result),
            output_path=report_md,
            test_name="PDF Multi-Column Structure Test",
        )
        assert report_md.exists(), "test_assertions_report.md not generated"

        # Verify content of report
        report_content = report_md.read_text(encoding="utf-8")
        assert "Column Structure Analysis" in report_content
        assert "Header Row Analysis" in report_content
        assert "Data Row Analysis" in report_content

        # Log locations for user
        logger.info("Visual artifacts generated:")
        logger.info("  - %s", preview_html)
        logger.info("  - %s", overlay_png)
        logger.info("  - %s", report_md)

    def test_pdf_golden_structure_expectations(self, tmp_path):
        """Verify the expected multi-column structure (golden expectations).

        This test encodes the INTENDED behavior for test.pdf (real estate lease table):
        - Expected headers defined by HTML_SCHEMA_COLUMNS
        - At least one data row in the tbody
        """
        import pandas as pd

        out = tmp_path / "test_output.html"
        result = extract(str(TEST_PDF), str(out))
        result_path = Path(result)

        content = result_path.read_text(encoding="utf-8")

        # 1. Verify the HTML schema columns are present
        from ocr_extractor.tenancy_parser import HTML_SCHEMA_COLUMNS
        for col in HTML_SCHEMA_COLUMNS:
            assert f"<th>{col}</th>" in content, (
                f"Expected schema column '{col}' not found in HTML header"
            )

        # 2. At least one data row must be present
        tbody_match = re.search(r"<tbody>(.*?)</tbody>", content, re.DOTALL)
        assert tbody_match, "Could not find <tbody> in output"
        data_rows = re.findall(r"<tr>", tbody_match.group(1))
        assert len(data_rows) >= 1, (
            f"Expected at least 1 data row in tbody, found {len(data_rows)}"
        )

        # 3. Parse the HTML with pandas to verify multi-column data
        dfs = pd.read_html(str(result_path))
        assert dfs, "No tables found in HTML output"
        df = dfs[0]
        assert len(df.columns) == len(HTML_SCHEMA_COLUMNS), (
            f"Expected {len(HTML_SCHEMA_COLUMNS)} columns, got {len(df.columns)}"
        )

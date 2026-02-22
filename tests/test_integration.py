"""Integration / golden tests for the end-to-end extraction pipeline.

These tests run the full pipeline on a synthetic table image and verify
that the output .xlsx has the expected structure and content.
"""

from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import load_workbook
from PIL import Image, ImageDraw

from ocr_extractor.extractor import extract


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
        out = tmp_path / "out.xlsx"
        result = extract(str(SAMPLE_TABLE), str(out))
        assert Path(result).exists()

    def test_output_is_valid_xlsx(self, tmp_path):
        out = tmp_path / "out.xlsx"
        result = extract(str(SAMPLE_TABLE), str(out))
        wb = load_workbook(str(result))
        assert wb is not None

    def test_sheet_named_table(self, tmp_path):
        out = tmp_path / "out.xlsx"
        result = extract(str(SAMPLE_TABLE), str(out))
        wb = load_workbook(str(result))
        assert "Table" in wb.sheetnames

    def test_single_sheet_only(self, tmp_path):
        out = tmp_path / "out.xlsx"
        result = extract(str(SAMPLE_TABLE), str(out))
        wb = load_workbook(str(result))
        assert len(wb.sheetnames) == 1

    def test_has_multiple_rows(self, tmp_path):
        """The output should have at least the header + one data row."""
        out = tmp_path / "out.xlsx"
        result = extract(str(SAMPLE_TABLE), str(out))
        wb = load_workbook(str(result))
        ws = wb.active
        assert ws.max_row >= 2

    def test_has_multiple_columns(self, tmp_path):
        """The output should have at least 2 columns."""
        out = tmp_path / "out.xlsx"
        result = extract(str(SAMPLE_TABLE), str(out))
        wb = load_workbook(str(result))
        ws = wb.active
        assert ws.max_column >= 2

    def test_freeze_panes_set(self, tmp_path):
        """Panes should be frozen below the header row."""
        out = tmp_path / "out.xlsx"
        result = extract(str(SAMPLE_TABLE), str(out))
        wb = load_workbook(str(result))
        ws = wb.active
        assert ws.freeze_panes is not None

    def test_debug_mode_does_not_raise(self, tmp_path):
        """debug=True should complete without errors."""
        out = tmp_path / "debug_out.xlsx"
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
        """When output_path is omitted the .xlsx lands next to the input."""
        src = tmp_path / "table.png"
        import shutil

        shutil.copy(str(SAMPLE_TABLE), str(src))
        result = extract(str(src))
        assert Path(result).suffix == ".xlsx"
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
        out = tmp_path / "synthetic.xlsx"
        result = extract(str(src), str(out))
        wb = load_workbook(str(result))
        ws = wb.active

        # Collect all non-None cell values from the sheet
        all_values = []
        for row in ws.iter_rows(values_only=True):
            for v in row:
                if v is not None:
                    all_values.append(str(v))

        all_text = " ".join(all_values).lower()

        # The header row must contain at least ONE of the known column names
        # (OCR may not be perfect, but at least something should come through)
        assert any(keyword in all_text for keyword in ("name", "amount", "date", "alice", "bob")), (
            f"No expected keywords found in extracted text. Got: {all_values}"
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
        """Output .xlsx must be created inside tmp_path, not in fixtures."""
        out = tmp_path / "test_output.xlsx"
        result = extract(str(TEST_PDF), str(out))
        result_path = Path(result)
        assert result_path.exists()
        # Crucially, the output must NOT be inside the fixtures directory
        assert not result_path.is_relative_to(FIXTURES_DIR), (
            f"Output was written to fixtures dir: {result_path}"
        )
        assert result_path.is_relative_to(tmp_path)

    def test_pdf_sheet_named_table(self, tmp_path):
        """The generated workbook must have exactly one sheet named 'Table'."""
        out = tmp_path / "test_output.xlsx"
        result = extract(str(TEST_PDF), str(out))
        wb = load_workbook(str(result))
        assert wb.sheetnames == ["Table"]

    def test_pdf_has_rows_and_columns(self, tmp_path):
        """The workbook must have at least 1 row and 1 column."""
        out = tmp_path / "test_output.xlsx"
        result = extract(str(TEST_PDF), str(out))
        wb = load_workbook(str(result))
        ws = wb.active
        assert ws.max_row >= 1
        assert ws.max_column >= 1

    def test_pdf_debug_artifacts_created(self, tmp_path):
        """Debug mode must create pipeline_diagram.md and grid_preview.txt."""
        out = tmp_path / "test_output.xlsx"
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
        out = tmp_path / "test_output.xlsx"
        extract(str(TEST_PDF), str(out), debug=True)
        debug_dir = tmp_path / "test_output.debug"
        assert debug_dir.exists()
        assert not debug_dir.is_relative_to(FIXTURES_DIR)

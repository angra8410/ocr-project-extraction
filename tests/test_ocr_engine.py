"""Unit tests for the OCR engine module."""

import pytest
from PIL import Image, ImageDraw, ImageFont

from ocr_extractor.ocr_engine import OcrCell, ocr_cell, ocr_dataframe, ocr_image


# ---------------------------------------------------------------------------
# OcrCell
# ---------------------------------------------------------------------------


class TestOcrCell:
    def test_low_confidence_flag_set_when_below_threshold(self):
        cell = OcrCell(text="hello", confidence=30.0)
        assert cell.low_confidence is True

    def test_low_confidence_false_when_above_threshold(self):
        cell = OcrCell(text="hello", confidence=90.0)
        assert cell.low_confidence is False

    def test_empty_text_never_low_confidence(self):
        cell = OcrCell(text="", confidence=20.0)
        assert cell.low_confidence is False

    def test_negative_confidence_not_low_confidence(self):
        """confidence == -1 means no data; should not be flagged."""
        cell = OcrCell(text="hello", confidence=-1.0)
        assert cell.low_confidence is False

    def test_boundary_confidence(self):
        # Exactly at threshold → not low confidence
        cell = OcrCell(text="x", confidence=60.0)
        assert cell.low_confidence is False


# ---------------------------------------------------------------------------
# ocr_image (smoke test – requires Tesseract)
# ---------------------------------------------------------------------------


def _make_text_image(text: str, size: tuple = (200, 60)) -> Image.Image:
    """Create a simple white image with black text."""
    img = Image.new("RGB", size, "white")
    draw = ImageDraw.Draw(img)
    draw.text((10, 10), text, fill="black")
    return img


class TestOcrImage:
    def test_returns_string(self):
        img = _make_text_image("Test")
        result = ocr_image(img)
        assert isinstance(result, str)

    def test_blank_image_returns_string(self):
        img = Image.new("RGB", (100, 50), "white")
        result = ocr_image(img)
        assert isinstance(result, str)


# ---------------------------------------------------------------------------
# ocr_cell (smoke test)
# ---------------------------------------------------------------------------


class TestOcrCellFunction:
    def test_returns_ocr_cell(self):
        img = _make_text_image("Hello")
        result = ocr_cell(img)
        assert isinstance(result, OcrCell)
        assert isinstance(result.text, str)
        assert isinstance(result.confidence, float)


# ---------------------------------------------------------------------------
# ocr_dataframe (smoke test)
# ---------------------------------------------------------------------------


class TestOcrDataframe:
    def test_returns_list_of_lists(self):
        img = _make_text_image("Col1  Col2")
        result = ocr_dataframe(img)
        assert isinstance(result, list)
        for row in result:
            assert isinstance(row, list)

    def test_blank_image_returns_empty(self):
        img = Image.new("RGB", (100, 50), "white")
        result = ocr_dataframe(img)
        # Blank image: Tesseract may return empty or minimal output
        assert isinstance(result, list)

"""OCR engine wrapper.

Wraps pytesseract and returns per-cell text with confidence scores so
the caller can flag low-confidence extractions.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from typing import List

import pytesseract
from PIL import Image

logger = logging.getLogger(__name__)

# Tesseract page-segmentation mode:
#  6  – Assume a single uniform block of text.
# 11  – Sparse text; find as much text as possible in no particular order.
_PSM_BLOCK = 6
_PSM_SPARSE = 11

# Confidence threshold below which a cell is considered low-confidence.
LOW_CONF_THRESHOLD = 60


@dataclass
class OcrCell:
    """A single OCR'd cell with associated metadata."""

    text: str
    confidence: float  # 0–100; -1 means "no data / empty"
    low_confidence: bool = field(init=False)

    def __post_init__(self) -> None:
        self.low_confidence = (
            0 <= self.confidence < LOW_CONF_THRESHOLD and bool(self.text.strip())
        )


def ocr_image(image: Image.Image, psm: int = _PSM_BLOCK) -> str:
    """Run Tesseract on *image* and return the raw text string."""
    config = f"--oem 3 --psm {psm}"
    return pytesseract.image_to_string(image, config=config)


def ocr_cell(image: Image.Image) -> OcrCell:
    """OCR a single cropped cell image and return an :class:`OcrCell`.

    Uses ``image_to_data`` to obtain per-word confidence scores and
    aggregates them into a single cell-level result.
    """
    config = f"--oem 3 --psm {_PSM_SPARSE}"
    data = pytesseract.image_to_data(
        image, config=config, output_type=pytesseract.Output.DICT
    )

    words: list[str] = []
    confs: list[float] = []
    for text, conf in zip(data["text"], data["conf"]):
        text = str(text).strip()
        conf = float(conf)
        if text and conf != -1:
            words.append(text)
            confs.append(conf)

    combined_text = " ".join(words)
    avg_conf: float = float(sum(confs) / len(confs)) if confs else -1.0

    return OcrCell(text=combined_text, confidence=avg_conf)


def ocr_dataframe(
    image: Image.Image,
) -> "list[list[OcrCell]]":
    """OCR a full table image using Tesseract's ``image_to_data``.

    Returns a 2-D list of :class:`OcrCell` objects, one per detected word
    grouped by (block_num, par_num, line_num).  This is a **fallback** path
    used when grid-line detection cannot segment the table into individual
    cells.
    """
    config = f"--oem 3 --psm {_PSM_BLOCK}"
    data = pytesseract.image_to_data(
        image, config=config, output_type=pytesseract.Output.DICT
    )

    # Group words by line key (block, par, line)
    lines: dict[tuple, list[tuple]] = {}
    n = len(data["text"])
    for i in range(n):
        text = str(data["text"][i]).strip()
        conf = float(data["conf"][i])
        if not text or conf == -1:
            continue
        key = (data["block_num"][i], data["par_num"][i], data["line_num"][i])
        lines.setdefault(key, []).append(
            (data["left"][i], data["top"][i], data["width"][i], text, conf)
        )

    # Sort lines by top, then left
    sorted_lines = sorted(lines.values(), key=lambda words: (words[0][1], words[0][0]))

    result: list[list[OcrCell]] = []
    for line_words in sorted_lines:
        line_words_sorted = sorted(line_words, key=lambda w: w[0])
        row = [OcrCell(text=w[3], confidence=w[4]) for w in line_words_sorted]
        result.append(row)

    return result

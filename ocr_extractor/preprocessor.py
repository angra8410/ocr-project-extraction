"""Image pre-processing utilities.

Applies rotation correction, de-skewing, contrast enhancement, and
noise reduction so that the downstream OCR step gets the cleanest
possible input.
"""

from __future__ import annotations

import logging
from typing import Optional

import cv2
import numpy as np
from PIL import Image

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Public helpers
# ---------------------------------------------------------------------------


def pil_to_cv(image: Image.Image) -> np.ndarray:
    """Convert a PIL Image to an OpenCV BGR array."""
    return cv2.cvtColor(np.array(image.convert("RGB")), cv2.COLOR_RGB2BGR)


def cv_to_pil(arr: np.ndarray) -> Image.Image:
    """Convert an OpenCV BGR array to a PIL Image."""
    return Image.fromarray(cv2.cvtColor(arr, cv2.COLOR_BGR2RGB))


def preprocess(image: Image.Image, debug: bool = False) -> Image.Image:
    """Run the full pre-processing pipeline on *image*.

    Steps performed:
    1. Convert to grayscale
    2. Denoise
    3. Deskew (correct small rotations ≤ 45°)
    4. Binarise / enhance contrast (Otsu threshold)

    Parameters
    ----------
    image:
        Input PIL Image (any mode).
    debug:
        When *True* additional logging is emitted.

    Returns
    -------
    PIL.Image
        Pre-processed PIL Image in RGB mode.
    """
    arr = pil_to_cv(image)
    gray = cv2.cvtColor(arr, cv2.COLOR_BGR2GRAY)

    # --- denoise ---
    gray = cv2.fastNlMeansDenoising(gray, h=10, templateWindowSize=7, searchWindowSize=21)

    # --- deskew ---
    angle = _detect_skew_angle(gray)
    if debug:
        logger.debug("Detected skew angle: %.2f°", angle)
    if abs(angle) > 0.1:
        gray = _rotate_image(gray, angle)

    # --- binarise / enhance contrast ---
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

    # Return as RGB PIL image (pytesseract works with both greyscale and RGB)
    return Image.fromarray(binary).convert("RGB")


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------


def _detect_skew_angle(gray: np.ndarray) -> float:
    """Return the dominant skew angle (degrees) of *gray*.

    Uses the Hough line transform on a morphologically thinned edge map.
    Returns 0.0 when no reliable angle can be determined.
    """
    edges = cv2.Canny(gray, 50, 150, apertureSize=3)
    lines = cv2.HoughLinesP(
        edges,
        rho=1,
        theta=np.pi / 180,
        threshold=100,
        minLineLength=100,
        maxLineGap=10,
    )
    if lines is None:
        return 0.0

    angles: list[float] = []
    for line in lines:
        x1, y1, x2, y2 = line[0]
        if x2 != x1:
            angle = np.degrees(np.arctan2(y2 - y1, x2 - x1))
            # Only consider near-horizontal lines (-45° … 45°)
            if -45 < angle < 45:
                angles.append(angle)

    if not angles:
        return 0.0

    median_angle = float(np.median(angles))
    # Ignore very small rotations
    return median_angle if abs(median_angle) > 0.5 else 0.0


def _rotate_image(gray: np.ndarray, angle: float) -> np.ndarray:
    """Rotate *gray* by *angle* degrees around its centre."""
    h, w = gray.shape[:2]
    centre = (w / 2, h / 2)
    M = cv2.getRotationMatrix2D(centre, angle, 1.0)
    rotated = cv2.warpAffine(
        gray,
        M,
        (w, h),
        flags=cv2.INTER_CUBIC,
        borderMode=cv2.BORDER_REPLICATE,
    )
    return rotated

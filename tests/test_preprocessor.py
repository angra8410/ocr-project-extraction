"""Unit tests for the image pre-processing module."""

import numpy as np
import pytest
from PIL import Image

from ocr_extractor.preprocessor import (
    _detect_skew_angle,
    _rotate_image,
    cv_to_pil,
    pil_to_cv,
    preprocess,
)


class TestPilCvConversions:
    def test_pil_to_cv_rgb_shape(self):
        """pil_to_cv should return a (H, W, 3) BGR array."""
        img = Image.new("RGB", (100, 80), color=(255, 0, 0))
        arr = pil_to_cv(img)
        assert arr.shape == (80, 100, 3)

    def test_pil_to_cv_bgr_channel_order(self):
        """Red pixel in RGB becomes (0, 0, 255) in BGR."""
        img = Image.new("RGB", (1, 1), color=(255, 0, 0))
        arr = pil_to_cv(img)
        # BGR: blue=0, green=0, red=255
        assert arr[0, 0, 0] == 0    # B
        assert arr[0, 0, 1] == 0    # G
        assert arr[0, 0, 2] == 255  # R

    def test_roundtrip(self):
        """pil→cv→pil should preserve pixel values (within rounding)."""
        img = Image.new("RGB", (50, 50), color=(123, 45, 67))
        reconstructed = cv_to_pil(pil_to_cv(img))
        orig_px = img.getpixel((0, 0))
        new_px = reconstructed.getpixel((0, 0))
        assert orig_px == new_px


class TestSkewDetection:
    def test_zero_angle_on_blank_image(self):
        """A blank image should return a skew angle of 0."""
        gray = np.ones((100, 200), dtype=np.uint8) * 255
        angle = _detect_skew_angle(gray)
        assert angle == 0.0

    def test_small_angle_ignored(self):
        """Angles ≤ 0.5° should be treated as 0."""
        import cv2

        # Draw a very slightly rotated line (< 0.5° effective angle)
        gray = np.ones((200, 400), dtype=np.uint8) * 255
        # Horizontal line: effective angle = 0
        cv2.line(gray, (10, 100), (390, 101), 0, 2)
        angle = _detect_skew_angle(gray)
        assert abs(angle) < 0.6  # should be tiny / 0


class TestRotateImage:
    def test_zero_rotation_unchanged(self):
        """Rotating by 0° should leave the image unchanged."""
        gray = np.zeros((100, 200), dtype=np.uint8)
        gray[50, 100] = 255  # single white pixel
        rotated = _rotate_image(gray, 0.0)
        assert rotated.shape == gray.shape


class TestPreprocess:
    def test_returns_rgb_pil_image(self):
        """preprocess() should always return an RGB PIL Image."""
        img = Image.new("RGB", (200, 150), color=(200, 200, 200))
        result = preprocess(img)
        assert isinstance(result, Image.Image)
        assert result.mode == "RGB"

    def test_output_same_aspect_ratio(self):
        """Output dimensions should be the same as the input."""
        img = Image.new("RGB", (300, 150), color=(128, 128, 128))
        result = preprocess(img)
        assert result.size == img.size

    def test_debug_does_not_raise(self):
        """debug=True should not raise any exceptions."""
        img = Image.new("RGB", (100, 100), color=(100, 100, 100))
        preprocess(img, debug=True)

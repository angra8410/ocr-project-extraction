# ocr-project-extraction

**OCR + table-structure reconstruction tool** that converts scanned or native
`.jpg` / `.jpeg` / `.png` / `.tif` / `.tiff` / `.pdf` documents into a
layout-preserving `.xlsx` file — mimicking the output style of
[jpgtoexcel.com](https://jpgtoexcel.com/es).

---

## Features

| Capability | Details |
|---|---|
| **Input formats** | `.jpg`, `.jpeg`, `.png`, `.tif`, `.tiff`, `.pdf` (native or scanned) |
| **Output** | Single `.xlsx` file with one sheet named **Table** |
| **Grid reconstruction** | Detects ruling lines; falls back to whitespace gap analysis |
| **Merged header cells** | Multi-column/row header merges detected and written as Excel merges |
| **Freeze panes** | Frozen below the header band |
| **Data integrity** | IDs kept as text, leading zeros preserved, no scientific notation |
| **Low-confidence OCR** | Cells flagged with `[?]` suffix and Excel cell comments |
| **Multi-page PDFs** | All pages appended into one continuous table |
| **Debug mode** | Extra logging + annotated preview image (`.debug.png`) |

---

## Requirements

- Python ≥ 3.9
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) installed and on `PATH`
  - Ubuntu/Debian: `sudo apt-get install tesseract-ocr tesseract-ocr-eng`
  - macOS: `brew install tesseract`
- `poppler-utils` for PDF rendering (`pdftoppm`)
  - Ubuntu/Debian: `sudo apt-get install poppler-utils`
  - macOS: `brew install poppler`

Python package dependencies are listed in `requirements.txt`:

```
pytesseract>=0.3.10
pdfplumber>=0.9.0
pdf2image>=1.16.0
openpyxl>=3.1.0
Pillow>=10.0.0
opencv-python-headless>=4.7.0
numpy>=1.24.0
```

---

## Installation

```bash
# 1. Clone the repository
git clone https://github.com/angra8410/ocr-project-extraction.git
cd ocr-project-extraction

# 2. Install Python dependencies
pip install -r requirements.txt

# 3. (Optional) Install as a package so `ocr-extract` is on PATH
pip install -e .
```

---

## Usage

### Command-line interface

```bash
# Basic usage — output file is placed next to the input
ocr-extract invoice.png

# Specify output path
ocr-extract scan.pdf -o results/table.xlsx

# Enable verbose logging
ocr-extract photo.jpg --verbose

# Enable debug mode (extra logging + annotated preview image)
ocr-extract table.tif --debug
```

#### Options

```
positional arguments:
  INPUT              Path to the input file

options:
  -o OUTPUT          Path for the output .xlsx (default: INPUT.xlsx)
  --debug            Log detected columns/rows/merges; save INPUT.debug.png
  -v, --verbose      Enable INFO-level logging
  -h, --help         Show this help message and exit
```

### Python API

```python
from ocr_extractor import extract

# Returns the resolved Path of the written .xlsx file
output_path = extract("invoice.pdf")

# With explicit output path and debug mode
output_path = extract("scan.png", output_path="output/table.xlsx", debug=True)
```

---

## Output format

The output `.xlsx` contains exactly **one sheet** named **`Table`** that:

- Reconstructs the table grid as faithfully as possible
- Uses **merged cells** for header spans (multi-column or multi-row)
- **Freezes panes** below the detected header area
- Applies a **thin border** to every cell
- **Left-aligns text**; **right-aligns** clearly numeric values
- Keeps dates and IDs as **text** to avoid auto-formatting
- Marks uncertain OCR cells with a `[?]` suffix and an Excel cell comment

---

## Pipeline overview

```
Input file
    │
    ▼
Page loading
  ├── PDF  → render pages at 200 DPI via pdf2image (or pdfplumber fallback)
  └── Image → open directly with Pillow
    │
    ▼
Pre-processing  (preprocessor.py)
  • Grayscale conversion
  • Fast non-local means denoising
  • Skew detection & correction (Hough lines)
  • Otsu binarisation
    │
    ▼
Table detection  (table_detector.py)
  • Morphological ruling-line detection
  • Whitespace-gap fallback
  • Header-row count estimation
    │
    ▼
Merge detection  (table_detector.detect_merges)
  • For header rows: check if vertical dividers are absent
    │
    ▼
OCR  (ocr_engine.py)
  • Crop each cell → pytesseract image_to_data (PSM 11 sparse)
  • Aggregate word confidences → flag low-confidence cells
    │
    ▼
Excel writing  (excel_writer.py)
  • openpyxl: merged cells, borders, freeze panes, column widths
  • Low-confidence cells: append [?], add cell comment
    │
    ▼
output.xlsx
```

---

## Running tests

```bash
# Run all tests
python -m pytest tests/ -v

# Unit tests only (fast, no Tesseract required)
python -m pytest tests/test_preprocessor.py tests/test_table_detector.py tests/test_excel_writer.py -v

# Integration + OCR tests
python -m pytest tests/test_integration.py tests/test_ocr_engine.py -v
```

---

## Supported file types

| Extension | Notes |
|---|---|
| `.jpg`, `.jpeg` | Standard JPEG images |
| `.png` | PNG images |
| `.tif`, `.tiff` | TIFF images |
| `.pdf` | Native text or scanned; all pages processed and appended |

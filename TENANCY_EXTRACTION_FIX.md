# Tenancy Schedule Extraction - Excel Multi-Column Output Fix

## 1. Root Cause Hypotheses (Ranked by Likelihood)

### Most Likely Causes:

1. **Table Detection Failure for Wide Tables (HIGH PROBABILITY)**
   - **Symptom**: The column detection algorithm fails to identify 15-20+ columns in wide tenancy schedules
   - **Root Cause**: Default thresholds and kernel sizes optimized for smaller tables (5-10 columns)
   - **Quick Test**: Run with `--debug` flag and check `grid_preview.txt` for column count
   - **Evidence**: Test expects 15+ columns but likely detects fewer

2. **Unstructured Text Extraction (MEDIUM PROBABILITY)**
   - **Symptom**: OCR extracts text as paragraphs without maintaining spatial column relationships
   - **Root Cause**: When table detection fails completely, falls back to sequential text extraction
   - **Quick Test**: Check if all text ends up in column A of the Excel file
   - **Evidence**: Problem statement mentions "rows being written as one concatenated string"

3. **DataFrame Construction Error (LOW PROBABILITY)**
   - **Symptom**: Data is correctly extracted but written to Excel as single column
   - **Root Cause**: DataFrame built with `{"raw": line}` instead of column-keyed dictionaries
   - **Quick Test**: Add logging to see DataFrame.columns before export
   - **Evidence**: Current code uses `openpyxl` directly, not pandas DataFrames

## 2. Excel Export Fix (Concrete Implementation)

### Solution Implemented: Specialized Tenancy Schedule Parser

The fix adds a new module `tenancy_parser.py` that:

1. **Defines Explicit Schema**:
   ```python
   TENANCY_SCHEDULE_COLUMNS = [
       "property", "as_of_date", "tenant_name", "legal_name", 
       "suite", "lease_type", "lease_from", "lease_to",
       "term_months", "area_sqft", "monthly_amount", "annual_amount",
       "security_deposit", "loc_amount", "notes", "row_type", "warnings"
   ]
   ```

2. **Guarantees Multi-Column Output**:
   ```python
   # Create workbook with openpyxl (NOT pandas to_excel)
   wb = Workbook()
   ws = wb.active
   
   # Write headers in separate columns
   for col_idx, col_name in enumerate(columns, start=1):
       cell = ws.cell(row=1, column=col_idx)
       cell.value = col_name
       cell.font = _HEADER_FONT
       cell.fill = _HEADER_FILL
   
   # Write data rows with each field in separate cell
   for row_idx, row_data in enumerate(data, start=2):
       for col_idx, col_name in enumerate(columns, start=1):
           cell = ws.cell(row=row_idx, column=col_idx)
           cell.value = row_data.get(col_name)
   ```

3. **Smart Column Mapping**:
   - Detects headers by keyword matching
   - Falls back to standard column order if headers not found
   - Maps OCR output to structured fields

4. **Data Normalization**:
   - Numbers: `normalize_number()` removes commas, handles parentheses as negatives
   - Dates: `normalize_date()` converts to ISO YYYY-MM-DD format
   - Nulls: Ambiguous values marked as `None` with warnings

### Why This Works:

- **Uses `openpyxl` cell-by-cell writing**: Each value written to a specific cell address
- **NO risk of single-column dump**: Column index explicitly controlled in loop
- **NO concatenation**: Each field written separately
- **NO DataFrame ambiguity**: Direct workbook manipulation

## 3. Extraction Output Schema

### Column Definitions:

| Column Name       | Type    | Description                                      | Normalization                    |
|-------------------|---------|--------------------------------------------------|----------------------------------|
| property          | Text    | Property/building name                           | Trimmed                          |
| as_of_date        | Date    | Report as-of date                                | ISO YYYY-MM-DD                   |
| tenant_name       | Text    | Tenant company name                              | Trimmed                          |
| legal_name        | Text    | Legal entity name                                | Trimmed                          |
| suite             | Text    | Suite/unit number                                | Preserved (may have letters)     |
| lease_type        | Text    | Lease type (e.g., "Gross", "Net")                | Trimmed                          |
| lease_from        | Date    | Lease start date                                 | ISO YYYY-MM-DD                   |
| lease_to          | Date    | Lease end date                                   | ISO YYYY-MM-DD                   |
| term_months       | Number  | Lease term in months                             | Float, commas removed            |
| area_sqft         | Number  | Leased area in square feet                       | Float, commas removed            |
| monthly_amount    | Number  | Monthly rent amount                              | Float, $removed, (100)→-100      |
| annual_amount     | Number  | Annual rent amount                               | Float, $removed, (100)→-100      |
| security_deposit  | Number  | Security deposit amount                          | Float, $removed                  |
| loc_amount        | Number  | Letter of credit amount                          | Float, $removed                  |
| notes             | Text    | Additional notes/comments                        | Trimmed                          |
| row_type          | Text    | Row type classification                          | Enum value                       |
| warnings          | Text    | Parsing warnings (semicolon-separated)           | Joined string                    |

### Row Types:

- `lease_summary`: Main tenant lease record
- `rent_step`: Rent escalation step
- `charge_schedule`: Additional charges
- `occupancy_summary`: Occupancy statistics
- `header`: Header row

### Rules for Handling Missing Values:

1. **Ambiguous Numeric Values**: Set to `None`, add warning
   ```python
   if normalize_number(value) is None:
       tenancy_row.monthly_amount = None
       tenancy_row.warnings.append(f"Could not parse monthly_amount: {value}")
   ```

2. **Ambiguous Dates**: Store original string, add warning
   ```python
   if normalize_date(value) is None:
       tenancy_row.lease_from = raw_value  # Keep original
       tenancy_row.warnings.append(f"Could not parse date: {raw_value}")
   ```

3. **Empty Cells**: Set to `None` (written as empty cell in Excel)

4. **OCR Errors**: Common replacements (O→0), otherwise flag with warning

## 4. How to Run

### Prerequisites:

1. **Python ≥ 3.9** installed
2. **Tesseract OCR** installed and on PATH:
   ```bash
   # Ubuntu/Debian
   sudo apt-get install tesseract-ocr tesseract-ocr-eng
   
   # macOS
   brew install tesseract
   
   # Windows
   # Download from: https://github.com/UB-Mannheim/tesseract/wiki
   ```

3. **Poppler utils** for PDF rendering:
   ```bash
   # Ubuntu/Debian
   sudo apt-get install poppler-utils
   
   # macOS
   brew install poppler
   
   # Windows
   # Download from: https://github.com/oschwartz10612/poppler-windows/releases
   ```

4. **Install Python dependencies**:
   ```bash
   cd ocr-project-extraction
   pip install -r requirements.txt
   
   # Or install as editable package
   pip install -e .
   ```

### Running the Extraction:

#### Standard Mode (General Tables):
```bash
# Basic usage
ocr-extract invoice.png

# With explicit output path
ocr-extract scan.pdf -o output/table.xlsx

# With debug logging
ocr-extract photo.jpg --debug --verbose
```

#### Tenancy Schedule Mode (Real Estate Lease Documents):
```bash
# Extract tenancy schedule with structured columns
ocr-extract test.pdf --tenancy-mode -o tenancy_output.xlsx

# With debug information
ocr-extract test.pdf --tenancy-mode --debug --verbose -o tenancy_output.xlsx

# Using Python module directly
python -m ocr_extractor.cli test.pdf --tenancy-mode -o output.xlsx
```

### For test.pdf Specifically:

```bash
# From repository root
cd ocr-project-extraction

# Run extraction in tenancy mode
ocr-extract tests/fixtures/test.pdf --tenancy-mode -o test_output.xlsx --verbose

# Alternative: Use Python API
python -c "from ocr_extractor import extract; extract('tests/fixtures/test.pdf', 'test_output.xlsx', tenancy_mode=True)"
```

### Output Location:

- **Default**: Same directory as input file, with `.xlsx` extension
- **Explicit**: Path specified with `-o` flag
- **Example**: `ocr-extract test.pdf` → creates `test.xlsx`

### Verifying the Result:

1. **Open in Excel/LibreOffice**:
   ```bash
   # Linux
   libreoffice test_output.xlsx
   
   # macOS
   open test_output.xlsx
   
   # Windows
   start test_output.xlsx
   ```

2. **Check Column Count**:
   ```python
   from openpyxl import load_workbook
   wb = load_workbook('test_output.xlsx')
   ws = wb.active
   print(f"Columns: {ws.max_column}")  # Should be 15-17
   print(f"Rows: {ws.max_row}")
   ```

3. **Verify Multi-Column Structure**:
   ```python
   from openpyxl import load_workbook
   wb = load_workbook('test_output.xlsx')
   ws = wb.active
   
   # Check header row
   headers = [ws.cell(1, col).value for col in range(1, ws.max_column + 1)]
   print(f"Headers: {headers}")
   
   # Check first data row
   first_row = [ws.cell(2, col).value for col in range(1, ws.max_column + 1)]
   print(f"First row: {first_row}")
   
   # Count non-empty columns in first data row
   non_empty = sum(1 for val in first_row if val)
   print(f"Non-empty columns: {non_empty}")  # Should be > 5
   ```

4. **Sample Columns to Verify**:
   - Column A: Property name (e.g., "AIP KKR Plaza")
   - Column B: As-of date (ISO format or original)
   - Column C: Tenant name (e.g., "PKP LKT Corp")
   - Column D: Suite number (e.g., "101", "Suite 205")
   - Column E: Lease from date
   - Column F: Lease to date
   - Column G: Area (sqft)
   - Column H: Monthly amount
   - Column I: Annual amount

### Troubleshooting:

#### Problem: "tesseract is not installed"
```bash
# Verify installation
tesseract --version

# If not found, install (see Prerequisites above)
# Ubuntu/Debian:
sudo apt-get install tesseract-ocr

# Add to PATH if needed (Linux/macOS):
export PATH=$PATH:/usr/local/bin
```

#### Problem: "Unable to get page count" or "poppler not found"
```bash
# Verify poppler installation
pdftoppm -v

# If not found, install:
# Ubuntu/Debian:
sudo apt-get install poppler-utils

# macOS:
brew install poppler
```

#### Problem: Output has only 1 column
**Solution**: Use `--tenancy-mode` flag:
```bash
ocr-extract test.pdf --tenancy-mode -o output.xlsx
```

#### Problem: Numbers not recognized correctly
**Check**: Look at `warnings` column in output Excel
**Solution**: OCR quality issue - try:
1. Increase PDF DPI (edit `extractor.py` line 118: `dpi=300`)
2. Pre-process image for better contrast
3. Manual correction based on warnings column

#### Debug Mode for Detailed Diagnostics:
```bash
ocr-extract test.pdf --tenancy-mode --debug -o output.xlsx

# This creates:
# - output.xlsx (main result)
# - output.debug.png (annotated image with detected cells)
# - output.debug/pipeline_diagram.md (processing metrics)
# - output.debug/grid_preview.txt (text preview of detected grid)
```

Check these artifacts to diagnose:
- **Low column count**: Check `grid_preview.txt` for actual detected columns
- **Incorrect boundaries**: Check `output.debug.png` for cell bounding boxes
- **OCR errors**: Check `pipeline_diagram.md` for low-confidence cell counts

## Acceptance Criteria Validation:

✅ **Multi-column Excel file**: `export_tenancy_to_excel()` writes 17 columns explicitly

✅ **Required columns present**:
   - property ✅
   - as_of_date ✅
   - tenant_name ✅
   - suite ✅
   - lease_from ✅
   - lease_to ✅
   - area_sqft ✅
   - monthly_amount ✅
   - annual_amount ✅
   - row_type ✅

✅ **Executable run instructions**: See "How to Run" section above

✅ **Unambiguous commands**: 
```bash
ocr-extract tests/fixtures/test.pdf --tenancy-mode -o test_output.xlsx
```

## Implementation Summary:

### Files Modified:
1. `requirements.txt` - Added pandas dependency
2. `pyproject.toml` - Added pandas dependency
3. `ocr_extractor/cli.py` - Added `--tenancy-mode` flag
4. `ocr_extractor/extractor.py` - Added tenancy_mode parameter and logic

### Files Created:
1. `ocr_extractor/tenancy_parser.py` - New module with:
   - `TenancyRow` dataclass
   - `parse_grid_to_rows()` - Grid → structured rows
   - `normalize_number()` - Numeric value normalization
   - `normalize_date()` - Date normalization
   - `export_tenancy_to_excel()` - Multi-column Excel writer

### Key Design Decisions:

1. **Why not modify excel_writer.py directly?**
   - Keeps generic table writer unchanged
   - Tenancy schedules need specialized parsing beyond visual layout
   - Separation of concerns: layout detection vs. semantic extraction

2. **Why use openpyxl directly instead of pandas DataFrame.to_excel()?**
   - More control over cell formatting
   - No risk of pandas inferring single-column structure
   - Better styling (borders, fonts, column widths)

3. **Why add warnings column?**
   - User can validate ambiguous values
   - Supports manual correction workflow
   - Transparent about parsing uncertainty

4. **Why normalize dates/numbers?**
   - Excel can perform calculations on normalized numbers
   - Consistent date sorting and filtering
   - OCR often produces "O" instead of "0" - automatic correction

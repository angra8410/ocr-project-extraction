"""Unit tests for tenancy_parser module."""

from __future__ import annotations

import tempfile
from pathlib import Path

import pytest
from openpyxl import load_workbook

from ocr_extractor.table_detector import CellRegion, TableGrid
from ocr_extractor.tenancy_parser import (
    TenancyRow,
    export_tenancy_to_excel,
    normalize_date,
    normalize_number,
    parse_grid_to_rows,
)


class TestNormalizeNumber:
    """Test numeric value normalization."""

    def test_simple_integer(self):
        assert normalize_number("123") == 123.0

    def test_simple_float(self):
        assert normalize_number("123.45") == 123.45

    def test_with_commas(self):
        assert normalize_number("1,234.56") == 1234.56

    def test_with_dollar_sign(self):
        assert normalize_number("$1,234.56") == 1234.56

    def test_negative_parentheses(self):
        assert normalize_number("(100)") == -100.0

    def test_negative_parentheses_with_commas(self):
        assert normalize_number("($1,234.56)") == -1234.56

    def test_ocr_o_to_zero(self):
        # Common OCR error: O instead of 0
        assert normalize_number("1O0") == 100.0
        assert normalize_number("1o0") == 100.0

    def test_invalid_returns_none(self):
        assert normalize_number("abc") is None
        assert normalize_number("") is None
        assert normalize_number("   ") is None

    def test_none_returns_none(self):
        assert normalize_number(None) is None


class TestNormalizeDate:
    """Test date normalization to ISO format."""

    def test_mdy_slash(self):
        assert normalize_date("01/15/2024") == "2024-01-15"
        assert normalize_date("1/5/2024") == "2024-01-05"

    def test_mdy_slash_short_year(self):
        assert normalize_date("01/15/24") == "2024-01-15"

    def test_iso_format(self):
        assert normalize_date("2024-01-15") == "2024-01-15"

    def test_month_name(self):
        assert normalize_date("January 15, 2024") == "2024-01-15"
        assert normalize_date("Jan 15, 2024") == "2024-01-15"

    def test_invalid_returns_none(self):
        assert normalize_date("not-a-date") is None
        assert normalize_date("") is None
        assert normalize_date("   ") is None

    def test_none_returns_none(self):
        assert normalize_date(None) is None


class TestTenancyRow:
    """Test TenancyRow dataclass."""

    def test_to_dict_all_fields(self):
        row = TenancyRow(
            property="Test Property",
            as_of_date="2024-01-15",
            tenant_name="Test Tenant",
            legal_name="Test Tenant LLC",
            suite="101",
            lease_type="Gross",
            lease_from="2024-01-01",
            lease_to="2024-12-31",
            term_months=12.0,
            area_sqft=1500.0,
            monthly_amount=1000.0,
            annual_amount=12000.0,
            security_deposit=2000.0,
            loc_amount=5000.0,
            notes="Test notes",
            row_type="lease_summary",
        )
        data = row.to_dict()
        assert data["property"] == "Test Property"
        assert data["as_of_date"] == "2024-01-15"
        assert data["tenant_name"] == "Test Tenant"
        assert data["legal_name"] == "Test Tenant LLC"
        assert data["suite"] == "101"
        assert data["lease_type"] == "Gross"
        assert data["lease_from"] == "2024-01-01"
        assert data["lease_to"] == "2024-12-31"
        assert data["term_months"] == 12.0
        assert data["area_sqft"] == 1500.0
        assert data["monthly_amount"] == 1000.0
        assert data["annual_amount"] == 12000.0
        assert data["security_deposit"] == 2000.0
        assert data["loc_amount"] == 5000.0
        assert data["notes"] == "Test notes"
        assert data["row_type"] == "lease_summary"

    def test_to_dict_with_warnings(self):
        row = TenancyRow(
            property="Test",
            warnings=["Warning 1", "Warning 2"],
        )
        data = row.to_dict()
        assert data["warnings"] == "Warning 1; Warning 2"

    def test_to_dict_no_warnings(self):
        row = TenancyRow(property="Test")
        data = row.to_dict()
        assert data["warnings"] is None


class TestParseGridToRows:
    """Test grid-to-rows parsing."""

    def test_empty_grid(self):
        grid = TableGrid()
        rows = parse_grid_to_rows(grid)
        assert rows == []

    def test_single_row_header_only(self):
        grid = TableGrid(
            cells=[
                CellRegion(row=0, col=0, text="Property"),
                CellRegion(row=0, col=1, text="Tenant"),
                CellRegion(row=0, col=2, text="Suite"),
            ],
            header_rows=1,
        )
        rows = parse_grid_to_rows(grid)
        assert rows == []  # No data rows

    def test_header_and_data_row(self):
        grid = TableGrid(
            cells=[
                # Header
                CellRegion(row=0, col=0, text="Property"),
                CellRegion(row=0, col=1, text="Tenant"),
                CellRegion(row=0, col=2, text="Suite"),
                # Data
                CellRegion(row=1, col=0, text="Building A"),
                CellRegion(row=1, col=1, text="ABC Corp"),
                CellRegion(row=1, col=2, text="101"),
            ],
            header_rows=1,
        )
        rows = parse_grid_to_rows(grid)
        assert len(rows) == 1
        assert rows[0].property == "Building A"
        assert rows[0].tenant_name == "ABC Corp"
        assert rows[0].suite == "101"

    def test_numeric_field_normalization(self):
        grid = TableGrid(
            cells=[
                # Header
                CellRegion(row=0, col=0, text="Monthly"),
                CellRegion(row=0, col=1, text="Annual"),
                # Data with commas
                CellRegion(row=1, col=0, text="$1,000.50"),
                CellRegion(row=1, col=1, text="$12,006.00"),
            ],
            header_rows=1,
        )
        rows = parse_grid_to_rows(grid)
        assert len(rows) == 1
        assert rows[0].monthly_amount == 1000.50
        assert rows[0].annual_amount == 12006.00

    def test_date_field_normalization(self):
        grid = TableGrid(
            cells=[
                # Header
                CellRegion(row=0, col=0, text="Lease From"),
                CellRegion(row=0, col=1, text="Lease To"),
                # Data
                CellRegion(row=1, col=0, text="01/01/2024"),
                CellRegion(row=1, col=1, text="12/31/2024"),
            ],
            header_rows=1,
        )
        rows = parse_grid_to_rows(grid)
        assert len(rows) == 1
        assert rows[0].lease_from == "2024-01-01"
        assert rows[0].lease_to == "2024-12-31"

    def test_invalid_numeric_adds_warning(self):
        grid = TableGrid(
            cells=[
                # Header
                CellRegion(row=0, col=0, text="Property"),
                CellRegion(row=0, col=1, text="Monthly"),
                # Data with property (so it's not filtered) and invalid number
                CellRegion(row=1, col=0, text="Building A"),
                CellRegion(row=1, col=1, text="invalid"),
            ],
            header_rows=1,
        )
        rows = parse_grid_to_rows(grid)
        assert len(rows) == 1
        assert rows[0].property == "Building A"
        assert rows[0].monthly_amount is None
        assert len(rows[0].warnings) > 0
        assert "monthly_amount" in rows[0].warnings[0].lower()

    def test_invalid_date_adds_warning(self):
        grid = TableGrid(
            cells=[
                # Header
                CellRegion(row=0, col=0, text="Lease From"),
                # Data with invalid date
                CellRegion(row=1, col=0, text="invalid-date"),
            ],
            header_rows=1,
        )
        rows = parse_grid_to_rows(grid)
        assert len(rows) == 1
        assert rows[0].lease_from == "invalid-date"  # Keeps original
        assert len(rows[0].warnings) > 0
        assert "date" in rows[0].warnings[0].lower()


class TestExportTenancyToExcel:
    """Test Excel export with multi-column guarantee."""

    def test_export_creates_file(self, tmp_path):
        rows = [
            TenancyRow(
                property="Test Property",
                tenant_name="Test Tenant",
                suite="101",
            ),
        ]
        output_path = tmp_path / "test.xlsx"
        result = export_tenancy_to_excel(rows, output_path)
        assert Path(result).exists()

    def test_export_has_multiple_columns(self, tmp_path):
        rows = [
            TenancyRow(
                property="Test Property",
                tenant_name="Test Tenant",
                suite="101",
                monthly_amount=1000.0,
                annual_amount=12000.0,
            ),
        ]
        output_path = tmp_path / "test.xlsx"
        export_tenancy_to_excel(rows, output_path)

        wb = load_workbook(output_path)
        ws = wb.active

        # Should have at least 10 columns (schema defines 17)
        assert ws.max_column >= 10

    def test_export_header_row(self, tmp_path):
        rows = [
            TenancyRow(property="Test", tenant_name="Tenant"),
        ]
        output_path = tmp_path / "test.xlsx"
        export_tenancy_to_excel(rows, output_path)

        wb = load_workbook(output_path)
        ws = wb.active

        # Check header row contains expected column names
        headers = [ws.cell(1, col).value for col in range(1, ws.max_column + 1)]
        assert "property" in headers
        assert "tenant_name" in headers
        assert "suite" in headers
        assert "monthly_amount" in headers

    def test_export_data_in_correct_columns(self, tmp_path):
        rows = [
            TenancyRow(
                property="Building A",
                tenant_name="ABC Corp",
                suite="101",
                monthly_amount=1000.0,
            ),
        ]
        output_path = tmp_path / "test.xlsx"
        export_tenancy_to_excel(rows, output_path)

        wb = load_workbook(output_path)
        ws = wb.active

        # Get header row to find column indices
        headers = [ws.cell(1, col).value for col in range(1, ws.max_column + 1)]

        # Find column indices
        property_col = headers.index("property") + 1
        tenant_col = headers.index("tenant_name") + 1
        suite_col = headers.index("suite") + 1
        monthly_col = headers.index("monthly_amount") + 1

        # Check data row (row 2)
        assert ws.cell(2, property_col).value == "Building A"
        assert ws.cell(2, tenant_col).value == "ABC Corp"
        assert ws.cell(2, suite_col).value == "101"
        assert ws.cell(2, monthly_col).value == 1000.0

    def test_export_multiple_rows(self, tmp_path):
        rows = [
            TenancyRow(property="Building A", tenant_name="ABC Corp"),
            TenancyRow(property="Building B", tenant_name="XYZ Inc"),
            TenancyRow(property="Building C", tenant_name="Test LLC"),
        ]
        output_path = tmp_path / "test.xlsx"
        export_tenancy_to_excel(rows, output_path)

        wb = load_workbook(output_path)
        ws = wb.active

        # Should have 1 header + 3 data rows
        assert ws.max_row == 4

    def test_export_without_warnings_column(self, tmp_path):
        rows = [
            TenancyRow(property="Test", warnings=["Some warning"]),
        ]
        output_path = tmp_path / "test.xlsx"
        export_tenancy_to_excel(rows, output_path, include_warnings=False)

        wb = load_workbook(output_path)
        ws = wb.active

        # Check that warnings column is not present
        headers = [ws.cell(1, col).value for col in range(1, ws.max_column + 1)]
        assert "warnings" not in headers

    def test_export_freeze_panes(self, tmp_path):
        rows = [
            TenancyRow(property="Test"),
        ]
        output_path = tmp_path / "test.xlsx"
        export_tenancy_to_excel(rows, output_path)

        wb = load_workbook(output_path)
        ws = wb.active

        # Panes should be frozen below header row
        assert ws.freeze_panes is not None

    def test_export_column_widths_set(self, tmp_path):
        rows = [
            TenancyRow(property="Test"),
        ]
        output_path = tmp_path / "test.xlsx"
        export_tenancy_to_excel(rows, output_path)

        wb = load_workbook(output_path)
        ws = wb.active

        # Check that column widths are set (not default)
        # Default width is typically around 8-10
        col_a_width = ws.column_dimensions["A"].width
        assert col_a_width > 10  # Should be wider than default

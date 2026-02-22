"""Tenancy schedule parser for extracting structured data from real estate lease documents.

This module provides specialized parsing for tenancy schedules that contain
multiple tables and repeated sections like:
- Tenancy Schedule (main lease data)
- Rent Steps
- Charge Schedules
- Occupancy Summary

The parser handles:
- Column boundary detection for wide tables (15-20+ columns)
- Numeric normalization (remove commas, handle parentheses as negatives)
- Date normalization to ISO YYYY-MM-DD format
- Null handling for ambiguous data
- Warning tracking for uncertain extractions
"""

from __future__ import annotations

import html
import json
import logging
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Optional


from .table_detector import TableGrid

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Schema Definition
# ---------------------------------------------------------------------------

# Column definitions for tenancy schedule tables
TENANCY_SCHEDULE_COLUMNS = [
    "property",
    "as_of_date",
    "tenant_name",
    "legal_name",
    "suite",
    "lease_type",
    "lease_from",
    "lease_to",
    "term_months",
    "area_sqft",
    "charge_label",
    "period_from",
    "period_to",
    "monthly_amount",
    "annual_amount",
    "management_fee_rate",
    "security_deposit",
    "loc_amount",
    "notes",
    "row_type",
    "warnings",
]

# Ordered columns for the HTML table output (spec-mandated order)
HTML_SCHEMA_COLUMNS = [
    "property",
    "as_of_date",
    "row_type",
    "tenant_name",
    "suite",
    "lease_from",
    "lease_to",
    "area_sqft",
    "charge_label",
    "period_from",
    "period_to",
    "monthly_amount",
    "annual_amount",
    "management_fee_rate",
    "notes",
]

# Row type constants
ROW_TYPE_LEASE_SUMMARY = "lease_summary"
ROW_TYPE_RENT_STEP = "rent_step"
ROW_TYPE_CHARGE_SCHEDULE = "charge_schedule"
ROW_TYPE_OCCUPANCY_SUMMARY = "occupancy_summary"
ROW_TYPE_HEADER = "header"

# ---------------------------------------------------------------------------
# Data Models
# ---------------------------------------------------------------------------


@dataclass
class TenancyRow:
    """A single row in the tenancy schedule."""

    property: Optional[str] = None
    as_of_date: Optional[str] = None
    tenant_name: Optional[str] = None
    legal_name: Optional[str] = None
    suite: Optional[str] = None
    lease_type: Optional[str] = None
    lease_from: Optional[str] = None
    lease_to: Optional[str] = None
    term_months: Optional[float] = None
    area_sqft: Optional[float] = None
    charge_label: Optional[str] = None
    period_from: Optional[str] = None
    period_to: Optional[str] = None
    monthly_amount: Optional[float] = None
    annual_amount: Optional[float] = None
    management_fee_rate: Optional[float] = None
    security_deposit: Optional[float] = None
    loc_amount: Optional[float] = None
    notes: Optional[str] = None
    row_type: str = ROW_TYPE_LEASE_SUMMARY
    warnings: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        """Convert to dictionary for DataFrame construction."""
        data = {
            "property": self.property,
            "as_of_date": self.as_of_date,
            "tenant_name": self.tenant_name,
            "legal_name": self.legal_name,
            "suite": self.suite,
            "lease_type": self.lease_type,
            "lease_from": self.lease_from,
            "lease_to": self.lease_to,
            "term_months": self.term_months,
            "area_sqft": self.area_sqft,
            "charge_label": self.charge_label,
            "period_from": self.period_from,
            "period_to": self.period_to,
            "monthly_amount": self.monthly_amount,
            "annual_amount": self.annual_amount,
            "management_fee_rate": self.management_fee_rate,
            "security_deposit": self.security_deposit,
            "loc_amount": self.loc_amount,
            "notes": self.notes,
            "row_type": self.row_type,
            "warnings": "; ".join(self.warnings) if self.warnings else None,
        }
        return data


# ---------------------------------------------------------------------------
# Normalization Functions
# ---------------------------------------------------------------------------


def normalize_number(value: str) -> Optional[float]:
    """Normalize numeric values from OCR text.

    Handles:
    - Removes commas (1,234.56 → 1234.56)
    - Handles parentheses as negatives: (100) → -100
    - Handles OCR errors: O vs 0
    - Returns None for invalid/ambiguous values

    Args:
        value: String value to normalize

    Returns:
        Normalized float or None if invalid
    """
    if not value or not isinstance(value, str):
        return None

    # Clean the value
    value = value.strip()

    if not value:
        return None

    # Handle parentheses as negative
    is_negative = False
    if value.startswith("(") and value.endswith(")"):
        is_negative = True
        value = value[1:-1].strip()

    # Remove commas and dollar signs
    value = value.replace(",", "").replace("$", "").strip()

    # Common OCR errors: replace O with 0
    # Only do this if the value looks like a number (contains digits)
    if re.search(r"\d", value):
        value = value.replace("O", "0").replace("o", "0")

    # Try to convert to float
    try:
        result = float(value)
        return -result if is_negative else result
    except ValueError:
        logger.debug("Cannot normalize number: %r", value)
        return None


def normalize_date(value: str) -> Optional[str]:
    """Normalize date values to ISO YYYY-MM-DD format.

    Handles common formats:
    - MM/DD/YYYY, M/D/YYYY
    - DD-MM-YYYY, D-M-YYYY
    - Month DD, YYYY
    - DD Month YYYY

    Args:
        value: String value to normalize

    Returns:
        ISO date string (YYYY-MM-DD) or None if invalid
    """
    if not value or not isinstance(value, str):
        return None

    value = value.strip()

    if not value:
        return None

    # Common date formats to try
    formats = [
        "%m/%d/%Y",
        "%m/%d/%y",
        "%d/%m/%Y",
        "%d/%m/%y",
        "%Y-%m-%d",
        "%d-%m-%Y",
        "%B %d, %Y",
        "%b %d, %Y",
        "%d %B %Y",
        "%d %b %Y",
    ]

    for fmt in formats:
        try:
            dt = datetime.strptime(value, fmt)
            return dt.strftime("%Y-%m-%d")
        except ValueError:
            continue

    logger.debug("Cannot normalize date: %r", value)
    return None


# ---------------------------------------------------------------------------
# Grid-to-Row Parsing
# ---------------------------------------------------------------------------


def parse_grid_to_rows(grid: TableGrid) -> list[TenancyRow]:
    """Parse a TableGrid into structured tenancy rows.

    This function extracts data from the grid and converts it into
    structured TenancyRow objects with proper column mapping.

    Section header rows (e.g. "Rent Steps", "Charge Schedule",
    "Occupancy Summary") are detected and used to set the ``row_type``
    of subsequent data rows.  Property name and as-of date are extracted
    from document header lines when present.

    Args:
        grid: TableGrid from table detection

    Returns:
        List of TenancyRow objects
    """
    if not grid.cells:
        logger.warning("Empty grid provided to parse_grid_to_rows")
        return []

    rows: list[TenancyRow] = []

    # Build a map: (row, col) -> cell for fast lookup
    cell_map: dict[tuple[int, int], Any] = {}
    for cell in grid.cells:
        cell_map[(cell.row, cell.col)] = cell

    # Detect header row and column mapping
    header_map = _detect_header_mapping(grid, cell_map)

    if not header_map:
        logger.warning("Could not detect header mapping; using fallback column order")
        # Fallback: assume columns in standard order
        header_map = _create_fallback_header_mapping(grid.num_cols)

    # Extract property name and as_of_date from header/title rows if present
    property_name: Optional[str] = None
    as_of_date: Optional[str] = None
    for row_idx in range(min(grid.header_rows, grid.num_rows)):
        row_text = _get_row_full_text(row_idx, grid.num_cols, cell_map)
        prop, aod = _extract_property_as_of_date(row_text)
        if prop and not property_name:
            property_name = prop
        if aod and not as_of_date:
            as_of_date = aod

    # Track the current section type; starts as lease_summary
    current_row_type: str = ROW_TYPE_LEASE_SUMMARY

    # Process data rows (skip header rows)
    for row_idx in range(grid.header_rows, grid.num_rows):
        row_text = _get_row_full_text(row_idx, grid.num_cols, cell_map)

        # Check if this row is a section header that changes the row type
        detected_type = _detect_section_type(row_text)
        if detected_type is not None:
            current_row_type = detected_type
            logger.debug("Section header detected at row %d: %r → %s",
                         row_idx, row_text[:60], current_row_type)
            continue  # Section header rows are not data rows

        tenancy_row = _extract_row_data(row_idx, grid.num_cols, cell_map, header_map)
        tenancy_row.row_type = current_row_type

        # Propagate property/date from document header.
        # The document header is the authoritative source for these fields
        # (each tenancy schedule document covers a single property).
        if property_name:
            tenancy_row.property = property_name
        if as_of_date and not tenancy_row.as_of_date:
            tenancy_row.as_of_date = as_of_date

        # Only add rows that have at least some data
        if _has_meaningful_data(tenancy_row):
            rows.append(tenancy_row)

    logger.info("Parsed %d tenancy rows from grid (%d rows × %d cols)",
                len(rows), grid.num_rows, grid.num_cols)

    return rows


def _get_row_full_text(row_idx: int, num_cols: int, cell_map: dict) -> str:
    """Return the combined text of all cells in a row, space-joined."""
    parts = []
    for col_idx in range(num_cols):
        cell = cell_map.get((row_idx, col_idx))
        if cell and cell.text:
            parts.append(cell.text.strip())
    return " ".join(parts)


def _detect_section_type(row_text: str) -> Optional[str]:
    """Detect whether *row_text* is a section header and return the row type.

    Returns the matching :data:`ROW_TYPE_*` constant if the row is a section
    header, or ``None`` if it is a regular data row.

    Detection is case-insensitive and tolerant of minor OCR noise (e.g.
    "Rent  Steps", "RENT STEPS", "Charge  Schedule").
    """
    if not row_text:
        return None
    text_lower = row_text.lower()
    if re.search(r"rent\s*steps?", text_lower):
        return ROW_TYPE_RENT_STEP
    if re.search(r"charge\s*schedules?", text_lower):
        return ROW_TYPE_CHARGE_SCHEDULE
    if re.search(r"occupancy\s*summary", text_lower):
        return ROW_TYPE_OCCUPANCY_SUMMARY
    return None


def _extract_property_as_of_date(row_text: str) -> tuple[Optional[str], Optional[str]]:
    """Extract property name and as-of date from a document header line.

    Recognises patterns such as:
    - "Property: Cornet Axol Date: 09/30/2024"
    - "Tenancy Schedule | Property: Acme Corp | As of: 2024-09-30"

    Returns a ``(property_name, as_of_date)`` tuple.  Either element may be
    ``None`` if not found.
    """
    property_name: Optional[str] = None
    as_of_date: Optional[str] = None

    if not row_text:
        return property_name, as_of_date

    # Property name: "Property: <name>" up to next keyword / end of field
    prop_match = re.search(
        r"property\s*[:|]\s*([A-Za-z0-9 ,.\-'&]+?)(?=\s*(?:date|as\s*of|$|\|))",
        row_text,
        re.IGNORECASE,
    )
    if prop_match:
        property_name = prop_match.group(1).strip() or None

    # As-of date: "Date: MM/DD/YYYY" or "As of: YYYY-MM-DD" etc.
    date_match = re.search(
        r"(?:date|as\s*of)\s*[:|]\s*(\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}|\d{4}-\d{2}-\d{2})",
        row_text,
        re.IGNORECASE,
    )
    if date_match:
        raw_date = date_match.group(1).strip()
        as_of_date = normalize_date(raw_date) or raw_date

    return property_name, as_of_date


def _detect_header_mapping(grid: TableGrid, cell_map: dict) -> dict[str, int]:
    """Detect which columns correspond to which fields based on header text.

    Args:
        grid: TableGrid
        cell_map: Map of (row, col) -> cell

    Returns:
        Dictionary mapping field names to column indices
    """
    header_map: dict[str, int] = {}

    # Keywords to look for in headers
    keywords = {
        "property": ["property", "building", "site"],
        "tenant_name": ["tenant", "name", "company"],
        "legal_name": ["legal", "legal name"],
        "suite": ["suite", "unit", "space"],
        "lease_type": ["type", "lease type"],
        "lease_from": ["from", "start", "commencement", "commence"],
        "lease_to": ["to", "end", "expiration", "expire", "expiry"],
        "term_months": ["term", "months", "duration"],
        "area_sqft": ["area", "sqft", "sf", "square"],
        "charge_label": ["charge", "charge label", "charge type"],
        "period_from": ["period from", "period start"],
        "period_to": ["period to", "period end"],
        "monthly_amount": ["monthly", "month", "rent/month"],
        "annual_amount": ["annual", "year", "rent/year"],
        "management_fee_rate": ["management fee", "mgmt fee", "management"],
        "security_deposit": ["security", "deposit"],
        "loc_amount": ["loc", "letter", "credit"],
    }

    # Check first few rows for headers
    for row_idx in range(min(grid.header_rows, 3)):
        for col_idx in range(grid.num_cols):
            cell = cell_map.get((row_idx, col_idx))
            if not cell or not cell.text:
                continue

            cell_text = cell.text.lower().strip()

            # Match against keywords
            for field, kws in keywords.items():
                if field in header_map:
                    continue  # Already found

                for kw in kws:
                    if kw in cell_text:
                        header_map[field] = col_idx
                        logger.debug("Mapped %r to column %d (header: %r)",
                                     field, col_idx, cell.text)
                        break

    return header_map


def _create_fallback_header_mapping(num_cols: int) -> dict[str, int]:
    """Create a fallback header mapping when automatic detection fails.

    Assumes columns in this order:
    Property, Tenant, Suite, Lease From, Lease To, Area, Monthly, Annual, ...
    """
    fallback_order = [
        "property",
        "tenant_name",
        "suite",
        "lease_from",
        "lease_to",
        "area_sqft",
        "charge_label",
        "period_from",
        "period_to",
        "monthly_amount",
        "annual_amount",
        "management_fee_rate",
        "security_deposit",
        "loc_amount",
        "notes",
    ]

    header_map = {}
    for idx, field in enumerate(fallback_order):
        if idx < num_cols:
            header_map[field] = idx

    return header_map


def _extract_row_data(
    row_idx: int,
    num_cols: int,
    cell_map: dict,
    header_map: dict[str, int],
) -> TenancyRow:
    """Extract data from a single row into a TenancyRow object.

    Args:
        row_idx: Row index to extract
        num_cols: Total number of columns
        cell_map: Map of (row, col) -> cell
        header_map: Mapping of field names to column indices

    Returns:
        TenancyRow object
    """
    tenancy_row = TenancyRow()

    # Helper to get cell text
    def get_cell_text(col_idx: int) -> str:
        cell = cell_map.get((row_idx, col_idx))
        return cell.text.strip() if cell and cell.text else ""

    # Extract text fields
    if "property" in header_map:
        tenancy_row.property = get_cell_text(header_map["property"]) or None

    if "tenant_name" in header_map:
        tenancy_row.tenant_name = get_cell_text(header_map["tenant_name"]) or None

    if "legal_name" in header_map:
        tenancy_row.legal_name = get_cell_text(header_map["legal_name"]) or None

    if "suite" in header_map:
        tenancy_row.suite = get_cell_text(header_map["suite"]) or None

    if "lease_type" in header_map:
        tenancy_row.lease_type = get_cell_text(header_map["lease_type"]) or None

    if "charge_label" in header_map:
        tenancy_row.charge_label = get_cell_text(header_map["charge_label"]) or None

    if "notes" in header_map:
        tenancy_row.notes = get_cell_text(header_map["notes"]) or None

    # Extract and normalize dates
    if "lease_from" in header_map:
        raw_from = get_cell_text(header_map["lease_from"])
        if raw_from:
            normalized = normalize_date(raw_from)
            if normalized:
                tenancy_row.lease_from = normalized
            else:
                tenancy_row.lease_from = raw_from
                tenancy_row.warnings.append(f"Could not parse date: {raw_from}")

    if "lease_to" in header_map:
        raw_to = get_cell_text(header_map["lease_to"])
        if raw_to:
            normalized = normalize_date(raw_to)
            if normalized:
                tenancy_row.lease_to = normalized
            else:
                tenancy_row.lease_to = raw_to
                tenancy_row.warnings.append(f"Could not parse date: {raw_to}")

    if "period_from" in header_map:
        raw_period_from = get_cell_text(header_map["period_from"])
        if raw_period_from:
            normalized = normalize_date(raw_period_from)
            if normalized:
                tenancy_row.period_from = normalized
            else:
                tenancy_row.period_from = raw_period_from
                tenancy_row.warnings.append(f"Could not parse date: {raw_period_from}")

    if "period_to" in header_map:
        raw_period_to = get_cell_text(header_map["period_to"])
        if raw_period_to:
            normalized = normalize_date(raw_period_to)
            if normalized:
                tenancy_row.period_to = normalized
            else:
                tenancy_row.period_to = raw_period_to
                tenancy_row.warnings.append(f"Could not parse date: {raw_period_to}")

    # Extract and normalize numeric fields
    numeric_fields = [
        ("term_months", "term_months"),
        ("area_sqft", "area_sqft"),
        ("monthly_amount", "monthly_amount"),
        ("annual_amount", "annual_amount"),
        ("management_fee_rate", "management_fee_rate"),
        ("security_deposit", "security_deposit"),
        ("loc_amount", "loc_amount"),
    ]

    for field, key in numeric_fields:
        if key in header_map:
            raw_value = get_cell_text(header_map[key])
            if raw_value:
                normalized = normalize_number(raw_value)
                if normalized is not None:
                    setattr(tenancy_row, field, normalized)
                else:
                    # Add warning for unparseable value
                    tenancy_row.warnings.append(f"Could not parse {field}: {raw_value}")

    return tenancy_row


def _has_meaningful_data(row: TenancyRow) -> bool:
    """Check if a row has any meaningful data (not all None/empty)."""
    return any([
        row.property,
        row.tenant_name,
        row.suite,
        row.lease_from,
        row.lease_to,
        row.area_sqft,
        row.charge_label,
        row.period_from,
        row.period_to,
        row.monthly_amount,
        row.annual_amount,
    ])


# ---------------------------------------------------------------------------
# HTML Export with Reasoning Block
# ---------------------------------------------------------------------------


def export_tenancy_to_html(
    rows: list[TenancyRow],
    output_path: str | Path | None = None,
    warnings_list: list[str] | None = None,
) -> dict[str, Any]:
    """Export tenancy rows to an HTML table and return a structured JSON result.

    The returned dictionary has exactly two top-level keys:

    * ``"html_table"`` – a complete ``<table>…</table>`` string with the
      columns defined in :data:`HTML_SCHEMA_COLUMNS`.
    * ``"reasoning"`` – an object with three keys:
      ``"parsing_strategy"``, ``"normalization_decisions"``, and
      ``"warnings"`` (list of strings).

    If *output_path* is given the JSON result is also written to that file.

    Args:
        rows: List of :class:`TenancyRow` objects to export.
        output_path: Optional path for the ``.json`` output file.
        warnings_list: Additional warnings collected during parsing (appended
            to per-row warnings already stored in each row's ``warnings``
            attribute).

    Returns:
        Dictionary with ``"html_table"`` and ``"reasoning"`` keys.
    """
    # ------------------------------------------------------------------ #
    # Build HTML table
    # ------------------------------------------------------------------ #
    lines: list[str] = ["<table>", "  <thead>", "    <tr>"]
    for col in HTML_SCHEMA_COLUMNS:
        lines.append(f"      <th>{html.escape(col)}</th>")
    lines += ["    </tr>", "  </thead>", "  <tbody>"]

    all_row_warnings: list[str] = list(warnings_list or [])

    for row in rows:
        row_dict = row.to_dict()
        lines.append("    <tr>")
        for col in HTML_SCHEMA_COLUMNS:
            value = row_dict.get(col)
            cell_text = "" if value is None else str(value)
            lines.append(f"      <td>{html.escape(cell_text)}</td>")
        lines.append("    </tr>")

        if row.warnings:
            all_row_warnings.extend(row.warnings)

    lines += ["  </tbody>", "</table>"]
    html_table = "\n".join(lines)

    # ------------------------------------------------------------------ #
    # Build reasoning block
    # ------------------------------------------------------------------ #
    normalization_decisions = [
        "Dates converted to ISO YYYY-MM-DD format (e.g. 4/1/2022 → 2022-04-01)",
        "Numeric amounts: commas removed, dollar signs stripped, OCR 'O'→'0' applied",
        "Amounts in parentheses (e.g. (100)) treated as negative values",
        "row_type assigned: lease_summary, rent_step, charge_schedule, or occupancy_summary",
        "charge_label captures expense codes such as RENT, INSUR, CAM, TAX",
        "period_from / period_to used for sub-lease periods (rent steps, charge schedules)",
        "management_fee_rate stored as a decimal fraction when present",
        "Cells with unparseable dates or numbers recorded in per-row warnings",
    ]

    result: dict[str, Any] = {
        "html_table": html_table,
        "reasoning": {
            "parsing_strategy": (
                "Property name and as_of_date detected from document header lines. "
                "Sections distinguished by keyword headers: 'Rent Steps' → row_type=rent_step, "
                "'Charge Schedule' → row_type=charge_schedule, "
                "'Occupancy Summary' → row_type=occupancy_summary, "
                "main lease table rows → row_type=lease_summary. "
                "Column mapping uses keyword matching against header text; "
                "falls back to positional order when headers are absent."
            ),
            "normalization_decisions": normalization_decisions,
            "warnings": all_row_warnings,
        },
    }

    if output_path is not None:
        output_path = Path(output_path)
        output_path.write_text(html_table, encoding="utf-8")
        logger.info(
            "Wrote tenancy HTML table to %s (%d rows, %d columns)",
            output_path,
            len(rows),
            len(HTML_SCHEMA_COLUMNS),
        )

    return result

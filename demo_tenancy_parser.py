#!/usr/bin/env python3
"""Demo script showing tenancy schedule parser without OCR dependencies.

This demonstrates the multi-column HTML output by creating
a synthetic TableGrid and exporting it using the tenancy parser.
"""

from pathlib import Path

from ocr_extractor.table_detector import CellRegion, TableGrid
from ocr_extractor.tenancy_parser import export_tenancy_to_html, parse_grid_to_rows


def create_demo_grid() -> TableGrid:
    """Create a synthetic tenancy schedule grid with realistic data."""
    cells = [
        # Header row
        CellRegion(row=0, col=0, text="Property", bbox=(0, 0, 100, 30)),
        CellRegion(row=0, col=1, text="Tenant", bbox=(100, 0, 250, 30)),
        CellRegion(row=0, col=2, text="Suite", bbox=(250, 0, 300, 30)),
        CellRegion(row=0, col=3, text="Lease From", bbox=(300, 0, 400, 30)),
        CellRegion(row=0, col=4, text="Lease To", bbox=(400, 0, 500, 30)),
        CellRegion(row=0, col=5, text="Area (SF)", bbox=(500, 0, 600, 30)),
        CellRegion(row=0, col=6, text="Monthly Rent", bbox=(600, 0, 750, 30)),
        CellRegion(row=0, col=7, text="Annual Rent", bbox=(750, 0, 900, 30)),
        CellRegion(row=0, col=8, text="Security", bbox=(900, 0, 1000, 30)),
        # Data row 1
        CellRegion(row=1, col=0, text="AIP KKR Plaza", bbox=(0, 30, 100, 60)),
        CellRegion(row=1, col=1, text="PKP LKT Corp", bbox=(100, 30, 250, 60)),
        CellRegion(row=1, col=2, text="101", bbox=(250, 30, 300, 60)),
        CellRegion(row=1, col=3, text="01/01/2024", bbox=(300, 30, 400, 60)),
        CellRegion(row=1, col=4, text="12/31/2025", bbox=(400, 30, 500, 60)),
        CellRegion(row=1, col=5, text="1,500", bbox=(500, 30, 600, 60)),
        CellRegion(row=1, col=6, text="$2,500.00", bbox=(600, 30, 750, 60)),
        CellRegion(row=1, col=7, text="$30,000.00", bbox=(750, 30, 900, 60)),
        CellRegion(row=1, col=8, text="$5,000", bbox=(900, 30, 1000, 60)),
        # Data row 2
        CellRegion(row=2, col=0, text="KSN Southland", bbox=(0, 60, 100, 90)),
        CellRegion(row=2, col=1, text="Corner Fudge LLC", bbox=(100, 60, 250, 90)),
        CellRegion(row=2, col=2, text="205", bbox=(250, 60, 300, 90)),
        CellRegion(row=2, col=3, text="06/15/2023", bbox=(300, 60, 400, 90)),
        CellRegion(row=2, col=4, text="06/14/2026", bbox=(400, 60, 500, 90)),
        CellRegion(row=2, col=5, text="2,100", bbox=(500, 60, 600, 90)),
        CellRegion(row=2, col=6, text="$3,200.50", bbox=(600, 60, 750, 90)),
        CellRegion(row=2, col=7, text="$38,406.00", bbox=(750, 60, 900, 90)),
        CellRegion(row=2, col=8, text="$6,500", bbox=(900, 60, 1000, 90)),
        # Data row 3
        CellRegion(row=3, col=0, text="Precision Tower", bbox=(0, 90, 100, 120)),
        CellRegion(row=3, col=1, text="Tech Solutions Inc", bbox=(100, 90, 250, 120)),
        CellRegion(row=3, col=2, text="3A", bbox=(250, 90, 300, 120)),
        CellRegion(row=3, col=3, text="03/01/2024", bbox=(300, 90, 400, 120)),
        CellRegion(row=3, col=4, text="02/28/2027", bbox=(400, 90, 500, 120)),
        CellRegion(row=3, col=5, text="3,000", bbox=(500, 90, 600, 120)),
        CellRegion(row=3, col=6, text="$4,500.00", bbox=(600, 90, 750, 120)),
        CellRegion(row=3, col=7, text="$54,000.00", bbox=(750, 90, 900, 120)),
        CellRegion(row=3, col=8, text="$9,000", bbox=(900, 90, 1000, 120)),
    ]

    grid = TableGrid(cells=cells, header_rows=1)
    return grid


def main():
    """Run the demo."""
    import tempfile

    print("=" * 80)
    print("TENANCY SCHEDULE PARSER DEMO")
    print("=" * 80)
    print()
    print("This demo shows how the tenancy parser guarantees multi-column HTML output")
    print("by creating a synthetic tenancy schedule and exporting it to HTML.")
    print()

    # Create synthetic grid
    print("Step 1: Creating synthetic tenancy schedule grid...")
    grid = create_demo_grid()
    print(f"  ✓ Created grid with {grid.num_rows} rows × {grid.num_cols} columns")
    print(f"  ✓ Header rows: {grid.header_rows}")
    print(f"  ✓ Total cells: {len(grid.cells)}")
    print()

    # Parse to structured rows
    print("Step 2: Parsing grid to structured tenancy rows...")
    rows = parse_grid_to_rows(grid)
    print(f"  ✓ Parsed {len(rows)} tenancy rows")
    print()
    for i, row in enumerate(rows, start=1):
        print(f"  Row {i}:")
        print(f"    Property: {row.property}")
        print(f"    Tenant: {row.tenant_name}")
        print(f"    Suite: {row.suite}")
        print(f"    Lease: {row.lease_from} to {row.lease_to}")
        print(f"    Area: {row.area_sqft} sqft")
        print(f"    Monthly: ${row.monthly_amount:,.2f}" if row.monthly_amount else "    Monthly: N/A")
        print(f"    Annual: ${row.annual_amount:,.2f}" if row.annual_amount else "    Annual: N/A")
        if row.warnings:
            print(f"    Warnings: {'; '.join(row.warnings)}")
        print()

    # Export to HTML
    temp_dir = Path(tempfile.gettempdir())
    output_path = temp_dir / "demo_tenancy_output.html"
    print(f"Step 3: Exporting to HTML: {output_path}")
    result = export_tenancy_to_html(rows, output_path=output_path)
    print(f"  ✓ Successfully exported to {output_path}")
    print()

    # Verify HTML structure
    print("Step 4: Verifying HTML table structure...")
    content = output_path.read_text(encoding="utf-8")
    assert "<table>" in content
    assert "<thead>" in content
    assert "<tbody>" in content
    print(f"  ✓ Valid HTML table")
    print(f"  ✓ Content length: {len(content)} bytes")
    print()

    print("=" * 80)
    print("SUCCESS: Multi-column HTML file generated!")
    print("=" * 80)
    print()
    print(f"You can open the file with: xdg-open {output_path}")
    print()


if __name__ == "__main__":
    main()

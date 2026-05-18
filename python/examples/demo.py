#!/usr/bin/env python3
"""
Demo of the Witan Python SDK.

Usage:
    cd python && python3 examples/demo.py
"""

import json
import tempfile
import shutil

from witan import Workbook, Regex


def main():
    # Create a temp directory for our test file
    tmp_dir = tempfile.mkdtemp(prefix="witan-demo-")
    test_file = f"{tmp_dir}/test.xlsx"

    try:
        print("=== Creating new workbook ===")
        with Workbook(test_file, create=True) as wb:
            # List sheets
            sheets = wb.list_sheets()
            print(f"Initial sheets: {[s['sheet'] for s in sheets]}")

            # Write some data
            print("\n=== Writing data ===")
            wb.set_cells([
                {"address": "Sheet1!A1", "value": "Name"},
                {"address": "Sheet1!B1", "value": "Age"},
                {"address": "Sheet1!C1", "value": "Score"},
                {"address": "Sheet1!A2", "value": "Alice"},
                {"address": "Sheet1!B2", "value": 30},
                {"address": "Sheet1!C2", "value": 95.5},
                {"address": "Sheet1!A3", "value": "Bob"},
                {"address": "Sheet1!B3", "value": 25},
                {"address": "Sheet1!C3", "value": 87.0},
                {"address": "Sheet1!A4", "value": "Charlie"},
                {"address": "Sheet1!B4", "value": 35},
                {"address": "Sheet1!C4", "value": 92.0},
            ])
            print("Data written")

            # Add a formula
            wb.set_cells([
                {"address": "Sheet1!C5", "value": None, "formula": "=AVERAGE(C2:C4)"},
                {"address": "Sheet1!A5", "value": "Average:"},
            ])
            print("Formula added")

            # Save
            saved = wb.save()
            print(f"Saved: {saved}")

        print("\n=== Reopening workbook ===")
        with Workbook(test_file) as wb:
            # Read range as TSV
            print("\nData as TSV:")
            tsv = wb.read_range_tsv("Sheet1!A1:C5")
            print(tsv)

            # Read specific cell
            avg_cell = wb.read_cell("Sheet1!C5")
            print("\nAverage cell:")
            print(f"  Value: {avg_cell['value']}")
            print(f"  Formula: {avg_cell.get('formula')}")
            print(f"  Text: {avg_cell['text']}")

            # Find cells
            print("\n=== Search operations ===")
            found = wb.find_cells("Alice")
            print(f"Found \"Alice\": {[c['address'] for c in found]}")

            # Find with regex
            numbers = wb.find_cells(Regex(r"^\d+$"))
            num_strs = [f"{c['address']}={c['value']}" for c in numbers]
            print(f"Found numbers: {num_strs}")

            # Evaluate formula
            print("\n=== Formula evaluation ===")
            result = wb.evaluate_formula("Sheet1", "=SUM(B2:B4)")
            print(f"SUM(B2:B4) = {result['value']}")

            # Get cell precedents
            print("\n=== Dependency tracing ===")
            precedents = wb.get_cell_precedents("Sheet1!C5")
            print(f"C5 depends on: {[c['address'] for c in precedents['cells']]}")

            # Add a new sheet
            print("\n=== Sheet operations ===")
            wb.add_sheet("Summary")
            sheets = wb.list_sheets()
            print(f"Sheets after adding: {[s['sheet'] for s in sheets]}")

            # Copy data to new sheet
            wb.copy_range("Sheet1!A1:C1", "Summary!A1")

            # Read from new sheet
            summary_tsv = wb.read_range_tsv("Summary!A1:C1")
            print(f"Summary sheet header: {summary_tsv.strip()}")

            # Describe sheet
            print("\n=== Sheet description ===")
            desc = wb.describe_sheet("Sheet1")
            print(f"Sheet1 description: {json.dumps(desc, indent=2)[:500]}...")

            # Save changes
            wb.save()

        print("\n=== Done! ===")

    finally:
        # Cleanup
        shutil.rmtree(tmp_dir)
        print("\nCleaned up temp directory")


if __name__ == "__main__":
    main()

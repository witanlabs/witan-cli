#!/usr/bin/env python3
"""
Demo of the Witan Python SDK for Google Sheets.

Prerequisites (one-time CLI setup):
    witan auth login
    witan gsheets connect
    # each existing sheet must also be authorized once (this demo does it for you):
    witan gsheets authorize gs://YOUR_SHEET_REF

Usage:
    cd python && python3 examples/google_sheets_demo.py gs://YOUR_SHEET_REF
    cd python && python3 examples/google_sheets_demo.py https://docs.google.com/spreadsheets/d/ID/edit
    cd python && python3 examples/google_sheets_demo.py --create --title "Witan demo"
"""

from __future__ import annotations

import argparse
import json
import sys

from witan import (
    GoogleSheet,
    Regex,
    WitanRPCError,
    is_google_auth_required,
    is_needs_file_authorization,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Witan Google Sheets SDK demo")
    parser.add_argument(
        "ref",
        nargs="?",
        help="Spreadsheet ref (gs://... or Google Sheets URL). Required unless --create.",
    )
    parser.add_argument(
        "--create",
        action="store_true",
        help="Create a new spreadsheet (uses witan gsheets rpc --create).",
    )
    parser.add_argument(
        "--title",
        default="Witan SDK demo",
        help="Title when using --create (default: %(default)s).",
    )
    return parser.parse_args()


def ensure_authorized(ref: str) -> None:
    """Make sure the app is authorized for this sheet, handing the picker URL
    off to a human if not. Sheets you create are authorized automatically."""
    info = GoogleSheet.authorize_url(ref)
    if info.get("authorized"):
        return
    print(
        "This sheet must be authorized. Open this link in a browser and pick the file:\n"
        f"  {info['picker_url']}\n"
        "(the link expires in a few minutes)",
        file=sys.stderr,
    )
    print("Waiting for authorization...", file=sys.stderr)
    GoogleSheet.wait_until_authorized(ref)


def open_sheet(args: argparse.Namespace) -> GoogleSheet:
    if args.create:
        return GoogleSheet.create(title=args.title)
    if not args.ref:
        print("error: ref is required unless --create is set", file=sys.stderr)
        sys.exit(2)
    ensure_authorized(args.ref)
    return GoogleSheet(args.ref)


def main() -> None:
    args = parse_args()

    try:
        with open_sheet(args) as sheet:
            if sheet.is_create:
                print(f"=== Created spreadsheet: {args.title!r} ===")
            else:
                print(f"=== Opened spreadsheet: {args.ref} ===")

            sheets = sheet.list_sheets()
            print(f"Sheets: {[s['sheet'] for s in sheets]}")

            print("\n=== Writing sample data ===")
            sheet.set_cells(
                [
                    {"address": "Sheet1!A1", "value": "Name"},
                    {"address": "Sheet1!B1", "value": "Score"},
                    {"address": "Sheet1!A2", "value": "Alice"},
                    {"address": "Sheet1!B2", "value": 95},
                    {"address": "Sheet1!A3", "value": "Bob"},
                    {"address": "Sheet1!B3", "value": 87},
                    {"address": "Sheet1!A4", "value": "Average:"},
                    {"address": "Sheet1!B4", "value": None, "formula": "=AVERAGE(B2:B3)"},
                ]
            )
            print("Changes applied (no save() — Google Sheets persists immediately)")

            print("\n=== Reading data ===")
            print(sheet.read_range_tsv("Sheet1!A1:B4"))

            avg = sheet.read_cell("Sheet1!B4")
            print(f"\nAverage cell: value={avg['value']!r} formula={avg.get('formula')!r}")

            print("\n=== Search ===")
            found = sheet.find_cells("Alice")
            print(f"Found Alice at: {[c['address'] for c in found]}")

            numbers = sheet.find_cells(Regex(r"^\d+$"))
            num_strs = [f"{c['address']}={c['value']}" for c in numbers]
            print(f"Numeric cells: {num_strs}")

            print("\n=== Formula evaluation ===")
            print("Skipped: evaluateFormulas is not implemented for Google Sheets")
            print(f"Read formula result from B4 instead: {avg['value']}")

            print("\n=== Sheet description (truncated) ===")
            desc = sheet.describe_sheet("Sheet1")
            print(json.dumps(desc, indent=2)[:500] + "...")

        print("\n=== Done ===")

    except WitanRPCError as err:
        if is_google_auth_required(err):
            print(
                "Google authentication required. Run:\n"
                "  witan auth login\n"
                "  witan gsheets connect",
                file=sys.stderr,
            )
        elif is_needs_file_authorization(err):
            print(
                "This sheet is not authorized. Run:\n"
                f"  witan gsheets authorize {args.ref}",
                file=sys.stderr,
            )
        else:
            print(f"RPC error ({err.code}): {err}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()

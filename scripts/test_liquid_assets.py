#!/usr/bin/env python3
"""
Test script for Liquid Assets extraction. Run on a single PDF to see why liquid assets
was or wasn't picked up (page scores, retirement year, table parse, header column, values summed).

Usage (from project root):
  python scripts/test_liquid_assets.py path/to/report.pdf
"""

import argparse
import sys
from pathlib import Path

try:
    import pymupdf
    sys.modules["fitz"] = pymupdf
except ImportError:
    pass

_project_root = Path(__file__).resolve().parent.parent
if str(_project_root) not in sys.path:
    sys.path.insert(0, str(_project_root))

import fitz
from extractor import extract_liquid_assets_debug


def _fmt(v) -> str:
    if v is None:
        return "—"
    return str(v)


def main() -> None:
    parser = argparse.ArgumentParser(description="Debug Liquid Assets extraction from a Voyant PDF.")
    parser.add_argument("pdf", type=Path, help="Path to Voyant PDF (pre or post)")
    args = parser.parse_args()

    if not args.pdf.exists():
        print(f"Error: PDF not found: {args.pdf}", file=sys.stderr)
        sys.exit(1)

    doc = fitz.open(args.pdf)
    try:
        debug = extract_liquid_assets_debug(doc)
    finally:
        doc.close()

    print("=" * 60)
    print("Liquid Assets extraction debug")
    print("=" * 60)
    print(f"  Retirement year used:    {_fmt(debug.get('retirement_year'))}")
    print(f"  Pages with 'Liquid Assets': {debug.get('liquid_heading_pages') or 'none'}")
    print(f"  Table page chosen:       {_fmt(debug.get('table_page_found'))} (1-based)")
    print()

    scores = debug.get("candidate_page_scores") or []
    if scores:
        print("  Candidate page scores (why each page was included or skipped):")
        for s in scores:
            inc = "✓ included" if s.get("included") else "✗ skipped (no year+£)"
            print(f"    Page {s.get('page')}: years={s.get('year_count')}, £count={s.get('currency_count')}, "
                  f"labels_bonus={s.get('label_bonus')}, has_all_3_labels={s.get('has_all_labels')}, score={s.get('score')} — {inc}")
        print()

    if debug.get("table_page_found"):
        print(f"  On table page: Savings={debug.get('savings_found_on_table_page')}, "
              f"Investments={debug.get('investments_found_on_table_page')}, Pensions={debug.get('pensions_found_on_table_page')}")
        print(f"  Rows parsed: {debug.get('num_rows_parsed')}")
        print(f"  Header row index: {_fmt(debug.get('header_row_index'))}, column index: {_fmt(debug.get('header_col_index'))}")
        print(f"  Year in header: {_fmt(debug.get('year_in_header'))}")
        if debug.get("candidate_years_on_page") is not None:
            print(f"  Candidate years on page: {debug.get('candidate_years_on_page')}")
        print(f"  Row preview (first cells): {debug.get('row_preview', [])[:8]}")
        print(f"  Values summed: {debug.get('values_summed')}")
        print(f"  Total liquid assets: {_fmt(debug.get('total'))}")
        print()

    if debug.get("error"):
        print("  ERROR:", debug.get("error"))
    if debug.get("traceback"):
        print("  TRACEBACK:")
        print(debug.get("traceback"))

    print("=" * 60)


if __name__ == "__main__":
    main()

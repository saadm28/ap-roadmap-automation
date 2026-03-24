#!/usr/bin/env python3
"""
Test script for Liquid Assets extraction. Run on pre and/or post PDFs to see why liquid assets
was or wasn't picked up. Pass both pre and post to compare side-by-side.

Usage (from project root):
  python scripts/test_liquid_assets.py path/to/pre.pdf path/to/post.pdf
  python scripts/test_liquid_assets.py path/to/pre.pdf
  python scripts/test_liquid_assets.py path/to/post.pdf
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


def _print_debug(debug: dict, label: str) -> None:
    """Print full liquid assets debug for one PDF."""
    print()
    print("=" * 70)
    print(f"  {label}")
    print("=" * 70)
    print(f"  retirement_year:         {_fmt(debug.get('retirement_year'))}")
    print(f"  Pages with 'Liquid Assets': {debug.get('liquid_heading_pages') or 'none'}")
    print(f"  table_page_found:         {_fmt(debug.get('table_page_found'))} (1-based)")
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
        print(f"  num_rows_parsed:         {debug.get('num_rows_parsed')}")
        print(f"  header_row_index:       {_fmt(debug.get('header_row_index'))}")
        print(f"  retirement_column_index: {_fmt(debug.get('header_col_index'))}")
        print(f"  retirement_year_ordinal: {_fmt(debug.get('retirement_year_ordinal'))} (0-based among year columns)")
        if debug.get("retirement_year_x_pt") is not None:
            print(f"  retirement_year_x_pt:   {debug.get('retirement_year_x_pt'):.1f} (header year column centre, pt)")
        print(f"  block_start:             {_fmt(debug.get('block_start'))}")
        print(f"  block_end:              {_fmt(debug.get('block_end'))}")
        print(f"  detected_row_labels:    {debug.get('detected_row_labels', [])}")
        print(f"  year_in_header:         {_fmt(debug.get('year_in_header'))}")
        if debug.get("candidate_years_on_page") is not None:
            print(f"  candidate_years_on_page: {debug.get('candidate_years_on_page')}")
        print(f"  row_preview (first 10 rows, first 4 cells):")
        for i, row in enumerate((debug.get("row_preview") or [])[:10]):
            print(f"    [{i}] {row}")
        print(f"  extracted_values:       {debug.get('values_summed')}")
        print(f"  final_sum:              {_fmt(debug.get('total'))}")
        pick_dbg = debug.get("liquid_value_pick_debug") or []
        if pick_dbg:
            print(f"  value_pick (Savings/Inv/Pens): {pick_dbg}")
        block_rows = debug.get("block_rows_scanned")
        if block_rows:
            print(f"  block_rows_scanned (full block rows):")
            for ri, first_cell, full_row in block_rows:
                print(f"    row_index={ri} first_cell={repr(first_cell)} row={full_row}")
        first_cells = debug.get("block_first_cells_as_read")
        if first_cells:
            print(f"  block_first_cells_as_read (exactly as read):")
            for i, fc in enumerate(first_cells):
                print(f"    [{i}] {repr(fc)}")
    else:
        print("  (No table page chosen — see candidate_page_scores and error above.)")

    if debug.get("error"):
        print()
        print("  ERROR:", debug.get("error"))
    if debug.get("traceback"):
        print()
        print("  TRACEBACK:")
        print(debug.get("traceback"))
    print()


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Debug Liquid Assets extraction. Pass pre and post PDFs to compare both."
    )
    parser.add_argument("pre_pdf", type=Path, nargs="?", default=None, help="Path to pre-advice Voyant PDF")
    parser.add_argument("post_pdf", type=Path, nargs="?", default=None, help="Path to post-advice Voyant PDF")
    args = parser.parse_args()

    if not args.pre_pdf and not args.post_pdf:
        parser.error("Provide at least one PDF (pre_pdf and/or post_pdf).")
    if args.pre_pdf and not args.pre_pdf.exists():
        print(f"Error: Pre PDF not found: {args.pre_pdf}", file=sys.stderr)
        sys.exit(1)
    if args.post_pdf and not args.post_pdf.exists():
        print(f"Error: Post PDF not found: {args.post_pdf}", file=sys.stderr)
        sys.exit(1)

    results = {}

    if args.pre_pdf:
        doc = fitz.open(args.pre_pdf)
        try:
            results["pre"] = extract_liquid_assets_debug(doc)
        finally:
            doc.close()
        _print_debug(results["pre"], f"PRE  — {args.pre_pdf.name}")

    if args.post_pdf:
        doc = fitz.open(args.post_pdf)
        try:
            results["post"] = extract_liquid_assets_debug(doc)
        finally:
            doc.close()
        _print_debug(results["post"], f"POST — {args.post_pdf.name}")

    if args.pre_pdf and args.post_pdf:
        print("=" * 70)
        print("  COMPARISON (PRE vs POST)")
        print("=" * 70)
        pre_total = results["pre"].get("total") if results.get("pre") else None
        post_total = results["post"].get("total") if results.get("post") else None
        pre_yr = results["pre"].get("retirement_year") if results.get("pre") else None
        post_yr = results["post"].get("retirement_year") if results.get("post") else None
        print(f"  Retirement year:  PRE={_fmt(pre_yr)}  POST={_fmt(post_yr)}")
        print(f"  Liquid assets:    PRE={_fmt(pre_total)}  POST={_fmt(post_total)}")
        if pre_total is not None and post_total is not None:
            print(f"  Change:           {post_total - pre_total:+,}")
        missing = []
        if results.get("pre") and results["pre"].get("total") is None:
            missing.append("PRE")
        if results.get("post") and results["post"].get("total") is None:
            missing.append("POST")
        if missing:
            print()
            print(f"  ⚠ Liquid assets NOT found for: {', '.join(missing)}")
            print("  Check the debug output above for that PDF (table_page_found, block_start/end, detected_row_labels, error).")
        print("=" * 70)


if __name__ == "__main__":
    main()

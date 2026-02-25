"""
AP Partners RoadMap Automation — Core PDF extraction logic.
Page finding by anchor text only. Charts: fixed crop → render → trim whitespace.
Values: regex on anchor pages; liquid assets at retirement from table (Savings + Investments + Pensions).
"""

import re
from pathlib import Path
from typing import Optional

import fitz  # PyMuPDF
from PIL import Image

from config import (
    CHART_CROPS,
    VALUE_EXTRACTORS,
    RENDER_DPI,
    COMPARISON_CHART_FILENAME,
    SLIDE24_ESTATE_ANALYSIS_ANCHOR,
    SLIDE24_ESTATE_ANALYSIS_CROP,
    SLIDE24_ESTATE_ANALYSIS_FILENAME,
)

# Default retirement year when not found in report (e.g. sample docs use 2038)
DEFAULT_RETIREMENT_YEAR = 2038

# Row labels for Liquid Assets table (must appear on the table page)
LIQUID_TABLE_ROW_LABELS = frozenset({"Savings", "Investments", "Pensions"})
LIQUID_ASSETS_HEADING = "Liquid Assets"


# ---------------------------------------------------------------------------
# Page finding by anchor
# ---------------------------------------------------------------------------

def get_page_text(doc: fitz.Document, page_num: int) -> str:
    """Extract text from a single page."""
    return doc[page_num].get_text()


def find_page_by_anchors(doc: fitz.Document, anchors: list[str]) -> Optional[int]:
    """First page (0-based) that contains any of the anchor strings."""
    for i in range(len(doc)):
        text = get_page_text(doc, i)
        for anchor in anchors:
            if anchor in text:
                return i
    return None


# ---------------------------------------------------------------------------
# Value extraction: regex on page text → first £ amount as int
# ---------------------------------------------------------------------------

def extract_value_from_text(text: str, pattern: str) -> Optional[int]:
    """First match of pattern; capture group 1 is the £ amount (digits/commas). Return as int."""
    m = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
    if not m:
        return None
    raw = m.group(1).replace(",", "").replace("£", "").strip()
    try:
        return int(float(raw))
    except ValueError:
        return None


def extract_values_from_doc(doc: fitz.Document) -> dict[str, Optional[int]]:
    """For each value type: find page by anchors, get page text, run regex. Returns dict of value key -> int or None."""
    result = {}
    for key, cfg in VALUE_EXTRACTORS.items():
        page_num = find_page_by_anchors(doc, cfg["anchors"])
        if page_num is None:
            result[key] = None
            continue
        text = get_page_text(doc, page_num)
        result[key] = extract_value_from_text(text, cfg["regex"])
    # Liquid assets at retirement: from table (Savings + Investments + Pensions) at retirement year, not regex
    result["liquid_assets_retirement"] = _extract_liquid_assets_table_total(doc)
    return result


# ---------------------------------------------------------------------------
# Retirement year: from report (e.g. Events & Goals / Retirement), else default
# ---------------------------------------------------------------------------

def extract_retirement_year(doc: fitz.Document) -> int:
    """Locate retirement year from report (e.g. Retirement event). Default 2038 if not found."""
    # Prefer Events & Goals / timeline-style page that mentions Retirement
    retirement_year_candidates = re.compile(r"\b(20[3-4]\d)\b")
    for i in range(len(doc)):
        text = get_page_text(doc, i)
        if "Retirement" not in text:
            continue
        matches = retirement_year_candidates.findall(text)
        for m in matches:
            y = int(m)
            if 2030 <= y <= 2049:
                return y
    return DEFAULT_RETIREMENT_YEAR


# ---------------------------------------------------------------------------
# Liquid Assets table page: score candidates by year tokens, currency, row labels
# ---------------------------------------------------------------------------

_YEAR_TOKEN_RE = re.compile(r"\b20\d{2}\b")
_CURRENCY_RE = re.compile(r"£[\d,]+")


def _score_liquid_assets_page(text: str) -> tuple[int, int, int, bool]:
    """Return (year_count, currency_count, row_label_bonus, has_all_three_labels)."""
    year_count = len(_YEAR_TOKEN_RE.findall(text))
    currency_count = len(_CURRENCY_RE.findall(text))
    has_savings = "Savings" in text
    has_investments = "Investments" in text
    has_pensions = "Pensions" in text
    has_all_labels = has_savings and has_investments and has_pensions
    return year_count, currency_count, (3 if has_all_labels else 0), has_all_labels


def find_liquid_assets_table_page(doc: fitz.Document) -> Optional[int]:
    """
    When multiple pages have 'Liquid Assets' heading, pick the best table page by content:
    - Must have at least one 4-digit year (e.g. 2038) AND at least one £ value (so we skip intro/chart pages).
    - Score by: year_count, currency_count, and row labels (Savings, Investments, Pensions).
    - Return the 0-based page index with the highest score.
    """
    candidates = []
    for i in range(len(doc)):
        text = get_page_text(doc, i)
        if LIQUID_ASSETS_HEADING not in text:
            continue
        year_count, currency_count, label_bonus, _ = _score_liquid_assets_page(text)
        # Require at least one year token and one currency value (skip intro page like 13)
        if year_count >= 1 and currency_count >= 1:
            score = year_count * 10 + currency_count + label_bonus
            candidates.append((score, i))
    if not candidates:
        return None
    candidates.sort(key=lambda x: (-x[0], x[1]))  # highest score first, then lowest page index
    return candidates[0][1]


def extract_liquid_assets_debug(doc: fitz.Document) -> dict:
    """
    Run liquid assets table extraction and return (total, debug_info).
    Debug info includes: which pages have heading/rows, retirement year, parsed rows, header column, values summed.
    """
    debug = {
        "retirement_year": None,
        "liquid_heading_pages": [],
        "table_page_found": None,
        "savings_found_on_table_page": None,
        "investments_found_on_table_page": None,
        "pensions_found_on_table_page": None,
        "num_rows_parsed": 0,
        "row_preview": [],
        "header_row_index": None,
        "header_col_index": None,
        "year_in_header": None,
        "candidate_years_on_page": None,
        "data_rows_found": [],
        "values_summed": {},
        "total": None,
        "error": None,
    }
    try:
        # Which pages have "Liquid Assets" heading?
        for i in range(len(doc)):
            text = get_page_text(doc, i)
            if LIQUID_ASSETS_HEADING in text:
                debug["liquid_heading_pages"].append(i + 1)  # 1-based for display
        # Retirement year
        debug["retirement_year"] = extract_retirement_year(doc)
        # Table page
        page_num = find_liquid_assets_table_page(doc)
        debug["table_page_found"] = (page_num + 1) if page_num is not None else None
        if page_num is None:
            debug["error"] = (
                "No Liquid Assets page had both year headers (e.g. 2038) and £ values. "
                "Intro/chart-only pages were skipped. Check that the numeric table is on a page with 'Liquid Assets'."
            )
            return debug
        page = doc[page_num]
        text = get_page_text(doc, page_num)
        debug["savings_found_on_table_page"] = "Savings" in text
        debug["investments_found_on_table_page"] = "Investments" in text
        debug["pensions_found_on_table_page"] = "Pensions" in text
        # Parse table
        try:
            raw = page.get_text("dict")
        except Exception as e:
            debug["error"] = f"get_text('dict') failed: {e}"
            return debug
        blocks = raw.get("blocks") or []
        rows = []
        for block in blocks:
            for line in block.get("lines") or []:
                y_vals = []
                cells = []
                for span in line.get("spans") or []:
                    t = (span.get("text") or "").strip()
                    if not t:
                        continue
                    bbox = span.get("bbox") or (0, 0, 0, 0)
                    x_center = (bbox[0] + bbox[2]) / 2
                    y_center = (bbox[1] + bbox[3]) / 2
                    y_vals.append(y_center)
                    cells.append((x_center, t))
                if cells:
                    y_avg = sum(y_vals) / len(y_vals)
                    cells.sort(key=lambda c: c[0])
                    rows.append((y_avg, [t for _, t in cells]))
        rows.sort(key=lambda r: r[0])
        # Merge lines that are on the same visual row (similar y) so we get logical table rows with multiple columns
        rows = _merge_rows_by_y(rows)
        debug["num_rows_parsed"] = len(rows)
        debug["row_preview"] = [
            (cells[:4] if len(cells) >= 4 else cells) for _, cells in rows[:20]
        ]  # first 4 cells of first 20 rows
        # Find year header row that contains retirement year and looks like a row of years (not early 2025-2028)
        year_str = str(debug["retirement_year"])
        header_row_index, header_col_index = _find_year_header_row(rows, debug["retirement_year"])
        debug["header_row_index"] = header_row_index
        debug["header_col_index"] = header_col_index
        debug["year_in_header"] = year_str if header_col_index is not None else None
        if header_col_index is None or header_row_index is None:
            # List candidate years found on the table page (from any cell)
            candidate_years = set()
            for _, cells in rows:
                for c in cells:
                    for m in _YEAR_TOKEN_RE.findall(c):
                        candidate_years.add(int(m))
            debug["candidate_years_on_page"] = sorted(candidate_years) if candidate_years else None
            debug["error"] = (
                f"Retirement year {year_str} not found in any row. "
                f"Candidate years on this page: {debug.get('candidate_years_on_page') or 'none'}. "
                f"Row starts: {[c[:2] for _, c in rows[:5]]}"
            )
            return debug
        # Sum only the three rows immediately below year header: Savings, Investments, Pensions
        total, values_by_label = _sum_liquid_from_three_rows(rows, header_row_index, header_col_index)
        debug["values_summed"] = dict(values_by_label)
        debug["data_rows_found"] = [(k, v) for k, v in values_by_label.items()]
        for offset, expected_label in enumerate(_LIQUID_DATA_ROW_LABELS_ORDER):
            ri = header_row_index + 2 + offset
            if ri < len(rows):
                _, cells = rows[ri]
                if expected_label not in values_by_label and header_col_index < len(cells):
                    debug["data_rows_found"].append(
                        (expected_label, f"cell={cells[header_col_index]!r} (parse failed or label mismatch)")
                    )
        debug["total"] = total if total else None
        if total == 0:
            debug["error"] = "No numeric values summed (check row labels and column alignment for the 3 rows under year header)."
    except Exception as e:
        debug["error"] = str(e)
        import traceback
        debug["traceback"] = traceback.format_exc()
    return debug


def _parse_int_from_cell(s: str) -> Optional[int]:
    """Parse a table cell to int (strip £, commas, spaces)."""
    s = s.replace(",", "").replace("£", "").replace(" ", "").strip()
    if not s:
        return None
    try:
        return int(float(s))
    except ValueError:
        return None


def _merge_rows_by_y(rows: list[tuple[float, list[str]]], y_tolerance: float = 6.0) -> list[tuple[float, list[str]]]:
    """
    Merge lines that share the same visual row (similar y) so one logical row can have many columns.
    Input: list of (y_center, cells). Output: list of (y_center, merged_cells) with cells sorted by position.
    """
    if not rows:
        return []
    # Group by y band (round y to tolerance)
    bands: dict[int, list[tuple[float, list[str]]]] = {}
    for y_avg, cells in rows:
        key = int(round(y_avg / y_tolerance) * y_tolerance)
        bands.setdefault(key, []).append((y_avg, cells))
    out = []
    for key in sorted(bands.keys()):
        group = bands[key]
        y_center = sum(y for y, _ in group) / len(group)
        # Flatten all cells and sort by original order (assume each line's cells are already x-sorted; we keep row order then cell order)
        all_cells = []
        for _, cells in group:
            all_cells.extend(cells)
        # Dedupe by keeping order (in case same text appears; we can't sort without x) - just concatenate
        out.append((y_center, all_cells))
    out.sort(key=lambda r: r[0])
    return out


def _cell_looks_like_year(cell: str) -> bool:
    """True if cell is a 4-digit year (e.g. 2025, 2038)."""
    s = (cell or "").strip().replace(",", "")
    return bool(s and len(s) == 4 and s.isdigit() and s.startswith("20"))


def _find_year_header_row(
    rows: list[tuple[float, list[str]]], retirement_year: int
) -> tuple[Optional[int], Optional[int]]:
    """
    Find the year header row that contains the retirement year and looks like a row of years
    (e.g. 2034, 2035, ..., 2038, ...). Do not use an early header row (e.g. 2025-2028) when
    retirement year is 2038. Returns (header_row_index, header_col_index) or (None, None).
    """
    year_str = str(retirement_year)
    candidates = []  # (row_index, col_index, year_like_cell_count)
    for ri, (_, cells) in enumerate(rows):
        col_index = None
        for col, cell in enumerate(cells):
            if cell == year_str or cell.replace(",", "").strip() == year_str:
                col_index = col
                break
        if col_index is None:
            continue
        year_like_count = sum(1 for c in cells if _cell_looks_like_year(c))
        candidates.append((ri, col_index, year_like_count))
    if not candidates:
        return None, None
    # Prefer the row with the most year-like cells (real year header has many years; early row has fewer)
    candidates.sort(key=lambda x: (-x[2], x[0]))
    return candidates[0][0], candidates[0][1]


# Expected order of the three data rows immediately under the year header (row+2, row+3, row+4 in 1-based = indices +1, +2, +3)
_LIQUID_DATA_ROW_LABELS_ORDER = ("Savings", "Investments", "Pensions")


def _sum_liquid_from_three_rows(
    rows: list[tuple[float, list[str]]],
    header_row_index: int,
    header_col_index: int,
) -> tuple[int, dict[str, int]]:
    """
    Sum values only from the three rows that contain Savings, Investments, Pensions.
    These are the 2nd, 3rd, 4th rows after the year header (row+2, row+3, row+4 in 1-based;
    often row+1 is "Age : Joe | Jane"). So we use header_row_index+2, +3, +4.
    Returns (total, values_by_label). If a row is missing or value unparseable, that label is skipped.
    """
    total = 0
    values_by_label = {}
    for offset, expected_label in enumerate(_LIQUID_DATA_ROW_LABELS_ORDER):
        ri = header_row_index + 2 + offset  # +2,+3,+4: skip first row after year header (e.g. Age)
        if ri >= len(rows):
            break
        _, cells = rows[ri]
        if header_col_index >= len(cells):
            continue
        label = None
        for c in cells:
            if c.strip() in LIQUID_TABLE_ROW_LABELS:
                label = c.strip()
                break
        if label is None or label != expected_label:
            continue
        val = _parse_int_from_cell(cells[header_col_index])
        if val is not None:
            total += val
            values_by_label[label] = val
    return total, values_by_label


def _extract_liquid_assets_table_total(doc: fitz.Document) -> Optional[int]:
    """
    Find Liquid Assets table page by heading; parse table with get_text('dict').
    Identify the year header row that contains the retirement year (e.g. 2038) and looks like
    a row of years; sum only Savings + Investments + Pensions from the three rows immediately
    below that header. Ignore all other blocks on the page.
    """
    page_num = find_liquid_assets_table_page(doc)
    if page_num is None:
        return None
    page = doc[page_num]
    retirement_year = extract_retirement_year(doc)
    try:
        raw = page.get_text("dict")
    except Exception:
        raw = {}
    blocks = raw.get("blocks") or []
    rows = []
    for block in blocks:
        for line in block.get("lines") or []:
            y_vals = []
            cells = []
            for span in line.get("spans") or []:
                t = (span.get("text") or "").strip()
                if not t:
                    continue
                bbox = span.get("bbox") or (0, 0, 0, 0)
                x_center = (bbox[0] + bbox[2]) / 2
                y_center = (bbox[1] + bbox[3]) / 2
                y_vals.append(y_center)
                cells.append((x_center, t))
            if cells:
                y_avg = sum(y_vals) / len(y_vals)
                cells.sort(key=lambda c: c[0])
                rows.append((y_avg, [t for _, t in cells]))
    rows.sort(key=lambda r: r[0])
    rows = _merge_rows_by_y(rows)
    header_row_index, header_col_index = _find_year_header_row(rows, retirement_year)
    if header_row_index is None or header_col_index is None:
        return None
    total, _ = _sum_liquid_from_three_rows(rows, header_row_index, header_col_index)
    return total if total else None


# ---------------------------------------------------------------------------
# Chart export: find page → crop rect → render at DPI → trim whitespace → save
# ---------------------------------------------------------------------------

def render_crop_to_png(
    doc: fitz.Document,
    page_num: int,
    crop_rect: tuple[float, float, float, float],
    out_path: Path,
    dpi: int = RENDER_DPI,
    trim: bool = True,
) -> None:
    """Render the crop region of the page at dpi and save as PNG. Optionally trim whitespace."""
    page = doc[page_num]
    x0, y0, x1, y1 = crop_rect
    rect = fitz.Rect(x0, y0, x1, y1)
    zoom = dpi / 72
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, clip=rect, alpha=False)
    pix.save(str(out_path))
    if trim:
        trim_whitespace(out_path)


def trim_whitespace(png_path: Path, white_threshold: int = 250) -> None:
    """Overwrite PNG with image cropped to non-white content (PIL)."""
    img = Image.open(png_path).convert("RGB")
    w, h = img.size
    pixels = list(img.getdata())
    min_x, min_y = w, h
    max_x, max_y = 0, 0
    for y in range(h):
        for x in range(w):
            r, g, b = pixels[y * w + x]
            if r < white_threshold or g < white_threshold or b < white_threshold:
                min_x = min(min_x, x)
                min_y = min(min_y, y)
                max_x = max(max_x, x)
                max_y = max(max_y, y)
    if min_x > max_x or min_y > max_y:
        return
    bbox = (min_x, min_y, max_x + 1, max_y + 1)
    img.crop(bbox).save(png_path)


# ---------------------------------------------------------------------------
# Extract charts for one PDF (pre or post): 3 chart types, filenames include page number
# ---------------------------------------------------------------------------

def extract_charts_from_pdf(
    pdf_path: Path,
    charts_dir: Path,
    prefix: str,
) -> dict[str, str]:
    """
    prefix is 'pre' or 'post'. Returns dict mapping logical key -> relative path e.g. charts/pre_timeline_page4.png
    Page number in filename is 1-based (for display).
    """
    charts = {}
    doc = fitz.open(pdf_path)
    try:
        for chart_key, cfg in CHART_CROPS.items():
            page_num = find_page_by_anchors(doc, cfg["anchors"])
            if page_num is None:
                continue
            page_display = page_num + 1
            out_name = f"{prefix}_{chart_key}_page{page_display}.png"
            out_path = charts_dir / out_name
            render_crop_to_png(doc, page_num, cfg["crop"], out_path)
            charts[f"{prefix}_{chart_key}"] = f"charts/{out_name}"
    finally:
        doc.close()
    return charts


# ---------------------------------------------------------------------------
# Extract values for one PDF (pre or post): 4 value types from anchor pages
# ---------------------------------------------------------------------------

def extract_values_from_pdf(pdf_path: Path) -> dict[str, Optional[int]]:
    """Returns dict with keys retirement_spending, estate_tax, estate_transfer, net_taxable (values as int or None)."""
    doc = fitz.open(pdf_path)
    try:
        return extract_values_from_doc(doc)
    finally:
        doc.close()


# ---------------------------------------------------------------------------
# Pre-advice: charts + values (same interface for app/summary)
# ---------------------------------------------------------------------------

def extract_pre_advice(pdf_path: Path, charts_dir: Path) -> tuple[dict, dict]:
    """Returns (values_dict, charts_dict). values_dict uses retirement_spending_annual, estate_tax_iht, etc. for compatibility."""
    raw_values = extract_values_from_pdf(pdf_path)
    values = {
        "retirement_spending_annual": raw_values.get("retirement_spending"),
        "estate_tax_iht": raw_values.get("estate_tax"),
        "total_estate_transfer": raw_values.get("estate_transfer"),
        "net_taxable_estate": raw_values.get("net_taxable"),
        "liquid_assets_retirement": raw_values.get("liquid_assets_retirement"),
    }
    charts = extract_charts_from_pdf(pdf_path, charts_dir, "pre")
    return values, charts


# ---------------------------------------------------------------------------
# Post-advice: charts + values
# ---------------------------------------------------------------------------

def extract_post_advice(pdf_path: Path, charts_dir: Path) -> tuple[dict, dict]:
    """Returns (values_dict, charts_dict)."""
    raw_values = extract_values_from_pdf(pdf_path)
    values = {
        "retirement_spending_annual": raw_values.get("retirement_spending"),
        "estate_tax_iht": raw_values.get("estate_tax"),
        "total_estate_transfer": raw_values.get("estate_transfer"),
        "net_taxable_estate": raw_values.get("net_taxable"),
        "liquid_assets_retirement": raw_values.get("liquid_assets_retirement"),
    }
    charts = extract_charts_from_pdf(pdf_path, charts_dir, "post")

    # Slide 24 — Estate Analysis - Detailed (post-advice only)
    doc = fitz.open(pdf_path)
    try:
        page_num = find_page_by_anchors(doc, [SLIDE24_ESTATE_ANALYSIS_ANCHOR])
        if page_num is not None:
            out_path = charts_dir / SLIDE24_ESTATE_ANALYSIS_FILENAME
            render_crop_to_png(doc, page_num, SLIDE24_ESTATE_ANALYSIS_CROP, out_path, trim=False)
            charts["slide24_estate_analysis"] = f"charts/{SLIDE24_ESTATE_ANALYSIS_FILENAME}"
    finally:
        doc.close()

    return values, charts


# ---------------------------------------------------------------------------
# FE Analytics comparison charts (Slide 19): crop from page dimensions, no trim (keep legends)
# ---------------------------------------------------------------------------

def _comparison_chart_crop(page: fitz.Page) -> tuple[float, float, float, float]:
    """Crop rect from page dimensions: exclude header/footer, keep chart + legends; tight margins to minimise white space."""
    r = page.rect
    width = r.x1 - r.x0
    height = r.y1 - r.y0
    margin_top = 115
    margin_bottom = 110
    margin_side = 35
    return (margin_side, margin_top, width - margin_side, height - margin_bottom)


# Max total comparison charts across all files and pages (avoids runaway output)
MAX_COMPARISON_CHARTS = 20


def extract_comparison_charts(pdf_paths: list[Path], charts_dir: Path) -> dict[str, str]:
    """Render every page of each FE Analytics PDF: same crop per page (by page dimensions). Multiple files = each file's pages; multi-page PDF = one chart per page."""
    charts = {}
    chart_num = 0
    for path in pdf_paths[:5]:
        if chart_num >= MAX_COMPARISON_CHARTS:
            break
        doc = fitz.open(path)
        try:
            for page_num in range(len(doc)):
                if chart_num >= MAX_COMPARISON_CHARTS:
                    break
                chart_num += 1
                page = doc[page_num]
                crop = _comparison_chart_crop(page)
                out_name = COMPARISON_CHART_FILENAME.format(n=chart_num)
                out_path = charts_dir / out_name
                render_crop_to_png(doc, page_num, crop, out_path, trim=False)
                charts[f"slide19_comparison_chart_{chart_num}"] = f"charts/{out_name}"
        finally:
            doc.close()
    return charts


# ---------------------------------------------------------------------------
# Console summary and summary file
# ---------------------------------------------------------------------------

def _fmt(v) -> str:
    if v is None:
        return "—"
    if isinstance(v, int):
        return f"£{v:,}"
    return str(v)


def print_console_summary(
    client_name: str,
    report_date: str,
    pre_values: dict,
    post_values: dict,
    pre_charts: dict,
    post_charts: dict,
    comparison_charts: dict,
    output_dir: Path,
) -> None:
    """Print formatted extraction summary to console."""
    def line():
        print("=" * 72)

    line()
    print("PRE-ADVICE EXTRACTION")
    line()
    print("📊 CHARTS EXTRACTED:")
    for name, rel in pre_charts.items():
        print(f"  ✅ {name} → {rel}")
    print("💰 KEY VALUES:")
    print(f"  ✅ Retirement Spending (Annual): {_fmt(pre_values.get('retirement_spending_annual'))}")
    print(f"  ✅ Estate Tax (IHT): {_fmt(pre_values.get('estate_tax_iht'))}")
    print(f"  ✅ Total Estate Transfer: {_fmt(pre_values.get('total_estate_transfer'))}")
    print(f"  ✅ Net Taxable Estate: {_fmt(pre_values.get('net_taxable_estate'))}")
    print(f"  ✅ Liquid assets at retirement: {_fmt(pre_values.get('liquid_assets_retirement'))}")

    line()
    print("POST-ADVICE EXTRACTION")
    line()
    print("📊 CHARTS EXTRACTED:")
    for name, rel in post_charts.items():
        print(f"  ✅ {name} → {rel}")
    print("💰 KEY VALUES:")
    print(f"  ✅ Retirement Spending (Annual): {_fmt(post_values.get('retirement_spending_annual'))}")
    print(f"  ✅ Estate Tax (IHT): {_fmt(post_values.get('estate_tax_iht'))}")
    print(f"  ✅ Total Estate Transfer: {_fmt(post_values.get('total_estate_transfer'))}")
    print(f"  ✅ Net Taxable Estate: {_fmt(post_values.get('net_taxable_estate'))}")
    print(f"  ✅ Liquid assets at retirement: {_fmt(post_values.get('liquid_assets_retirement'))}")

    pre_spend = pre_values.get("retirement_spending_annual") or 0
    post_spend = post_values.get("retirement_spending_annual") or 0
    diff_annual = post_spend - pre_spend
    diff_monthly = round(diff_annual / 12) if diff_annual else 0

    line()
    print("📈 PRE vs POST COMPARISON (for Slide 6)")
    line()
    print(f"  Retirement Spending: {_fmt(pre_spend)} → {_fmt(post_spend)} (↑ {_fmt(diff_annual)}/year, ↑ {_fmt(diff_monthly)}/month)")
    pre_liq = pre_values.get("liquid_assets_retirement")
    post_liq = post_values.get("liquid_assets_retirement")
    print(f"  Liquid assets at retirement: {_fmt(pre_liq)} → {_fmt(post_liq)}")

    line()
    print("COMPARISON CHARTS (Slide 19)")
    line()
    for name, rel in comparison_charts.items():
        print(f"  ✅ {name} → {rel}")

    line()
    print(f"✅ ALL EXTRACTIONS COMPLETE — Output saved to {output_dir}/")
    line()


def write_extraction_summary(
    output_dir: Path,
    client_name: str,
    report_date: str,
    template_type: str,
    pre_values: dict,
    post_values: dict,
    comparison_annual: int,
    comparison_monthly: int,
    all_charts: dict,
) -> None:
    """Write extraction_summary.txt in output folder."""
    path = output_dir / "extraction_summary.txt"
    lines = [
        f"client_name={client_name}",
        f"report_date={report_date}",
        f"template_type={template_type}",
        "",
        "[pre_advice]",
        f"retirement_spending_annual={pre_values.get('retirement_spending_annual') or ''}",
        f"estate_tax_iht={pre_values.get('estate_tax_iht') or ''}",
        f"total_estate_transfer={pre_values.get('total_estate_transfer') or ''}",
        f"net_taxable_estate={pre_values.get('net_taxable_estate') or ''}",
        f"liquid_assets_retirement={pre_values.get('liquid_assets_retirement') or ''}",
        "",
        "[post_advice]",
        f"retirement_spending_annual={post_values.get('retirement_spending_annual') or ''}",
        f"estate_tax_iht={post_values.get('estate_tax_iht') or ''}",
        f"total_estate_transfer={post_values.get('total_estate_transfer') or ''}",
        f"net_taxable_estate={post_values.get('net_taxable_estate') or ''}",
        f"liquid_assets_retirement={post_values.get('liquid_assets_retirement') or ''}",
        "",
        "[comparison]",
        f"retirement_spending_increase_annual={comparison_annual}",
        f"retirement_spending_increase_monthly={comparison_monthly}",
        "",
        "[charts]",
    ]
    for key in sorted(all_charts.keys()):
        lines.append(f"{key}={all_charts[key]}")
    path.write_text("\n".join(lines), encoding="utf-8")


# ---------------------------------------------------------------------------
# Full pipeline
# ---------------------------------------------------------------------------

def run_extraction(
    client_name: str,
    report_date: str,
    template_type: str,
    pre_advice_path: Path,
    post_advice_path: Path,
    comparison_pdf_paths: list[Path],
    output_base: Path,
    print_summary: bool = True,
) -> tuple[Path, dict, dict, dict]:
    """
    Run full extraction. Returns (output_dir, pre_values, post_values, all_charts).
    """
    safe_name = client_name.strip()
    for c in " /\\:*?\"<>|":
        safe_name = safe_name.replace(c, "_")
    output_dir = output_base / f"{safe_name}_{report_date}"
    charts_dir = output_dir / "charts"
    charts_dir.mkdir(parents=True, exist_ok=True)

    pre_values, pre_charts = extract_pre_advice(pre_advice_path, charts_dir)
    post_values, post_charts = extract_post_advice(post_advice_path, charts_dir)
    comparison_charts = extract_comparison_charts(comparison_pdf_paths, charts_dir)

    all_charts = {**pre_charts, **post_charts, **comparison_charts}

    pre_spend = pre_values.get("retirement_spending_annual") or 0
    post_spend = post_values.get("retirement_spending_annual") or 0
    comparison_annual = post_spend - pre_spend
    comparison_monthly = round(comparison_annual / 12) if comparison_annual else 0

    write_extraction_summary(
        output_dir, client_name, report_date, template_type,
        pre_values, post_values, comparison_annual, comparison_monthly, all_charts,
    )

    if print_summary:
        print_console_summary(
            client_name, report_date,
            pre_values, post_values,
            pre_charts, post_charts, comparison_charts,
            output_dir,
        )

    return output_dir, pre_values, post_values, all_charts

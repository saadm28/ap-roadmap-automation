"""
AP Partners RoadMap Automation — Core PDF extraction logic.
Page finding by anchor text only. Charts: fixed crop → render → trim whitespace. Values: regex on anchor page text.
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
    return result


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

    pre_spend = pre_values.get("retirement_spending_annual") or 0
    post_spend = post_values.get("retirement_spending_annual") or 0
    diff_annual = post_spend - pre_spend
    diff_monthly = round(diff_annual / 12) if diff_annual else 0

    line()
    print("📈 PRE vs POST COMPARISON (for Slide 6)")
    line()
    print(f"  Retirement Spending: {_fmt(pre_spend)} → {_fmt(post_spend)} (↑ {_fmt(diff_annual)}/year, ↑ {_fmt(diff_monthly)}/month)")

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
        "",
        "[post_advice]",
        f"retirement_spending_annual={post_values.get('retirement_spending_annual') or ''}",
        f"estate_tax_iht={post_values.get('estate_tax_iht') or ''}",
        f"total_estate_transfer={post_values.get('total_estate_transfer') or ''}",
        f"net_taxable_estate={post_values.get('net_taxable_estate') or ''}",
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

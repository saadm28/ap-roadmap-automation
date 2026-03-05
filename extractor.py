"""
AP Partners RoadMap Automation — Core PDF extraction logic.
Page finding by anchor text only. Charts: fixed crop → render → trim whitespace.
Values: regex on anchor pages; liquid assets at retirement from table (Savings + Investments + Pensions).
"""

import re
import logging
from pathlib import Path
from typing import Optional, Any

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

logger = logging.getLogger(__name__)

# Default retirement year when not found in report (e.g. sample docs use 2038)
DEFAULT_RETIREMENT_YEAR = 2038

# Row labels for Liquid Assets table (must appear on the table page)
LIQUID_TABLE_ROW_LABELS = frozenset({"Savings", "Investments", "Pensions"})
LIQUID_ASSETS_HEADING = "Liquid Assets"

# Slide 14: anchor phrases to find the correct page (NOT "Financial Summary" intro).
SLIDE14_ANCHOR_PHRASES = [
    "Retirement Summary",
    "Expenses are not funded",
    "Expenses are funded",
    "You can afford to spend",
    "shortfall",
    "additional lump sum",
    "save an additional",
]


def _find_slide14_page(doc: fitz.Document) -> tuple[Optional[int], Optional[str], str]:
    """Find page with Slide 14 content: must contain an anchor phrase AND at least one regex match (skip intro pages)."""
    # Patterns we need at least one of (same as in _extract_financial_summary_slide14)
    any_match = re.compile(
        r"shortfall\s+in\s+\d+\s+of\s+\d+\s+retirement\s+years"
        r"|additional\s+lump\s+sum\s+of\s+£[\d,]+\s+in\s+\d{4}"
        r"|(?:save\s+an\s+additional\s+£[\d,]+\s+per\s+year|Save\s+an\s+additional\s+£[\d,]+\s+annually)"
        r"|Expenses\s+are\s+not\s+funded\s+in\s+\d+\s+years"
        r"|Expenses\s+are\s+funded\s+in\s+\d+\s+years"
        r"|You\s+can\s+afford\s+to\s+spend\s+£[\d,]+\s+(?:per\s+year|annually)",
        re.IGNORECASE,
    )
    for i in range(len(doc)):
        text = get_page_text(doc, i)
        if not any_match.search(text):
            continue
        for anchor in SLIDE14_ANCHOR_PHRASES:
            if anchor in text:
                pos = text.find(anchor)
                start = max(0, pos - 200)
                end = min(len(text), pos + 400)
                snippet = text[start:end]
                if start > 0:
                    snippet = "…" + snippet
                if end < len(text):
                    snippet = snippet + "…"
                return (i, anchor, snippet)
    return (None, None, "")


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


def extract_values_from_doc(doc: fitz.Document, retirement_year_override: Optional[int] = None) -> dict[str, Any]:
    """Extract all values from doc. Returns dict with int/str/None values (numeric keys + client_name, report_date, report_month, report_year from first page).
    If retirement_year_override is set, use it for liquid assets table column lookup instead of deriving from PDF (allows different retirement ages per report).
    """
    result = {}
    for key, cfg in VALUE_EXTRACTORS.items():
        page_num = find_page_by_anchors(doc, cfg["anchors"])
        if page_num is None:
            result[key] = None
            continue
        text = get_page_text(doc, page_num)
        result[key] = extract_value_from_text(text, cfg["regex"])
    # Liquid assets at retirement: from table (Savings + Investments + Pensions) at retirement year, not regex.
    # Single code path: same logic used by app (run_extraction) and scripts/test_liquid_assets.py (extract_liquid_assets_debug).
    result["liquid_assets_retirement"] = _extract_liquid_assets_table_total(doc, retirement_year_override=retirement_year_override)
    if result["liquid_assets_retirement"] is None:
        liq_debug = extract_liquid_assets_debug(doc, retirement_year_override=retirement_year_override)
        logger.warning(
            "Liquid assets not extracted. retirement_year=%s, table_page=%s, pages_with_heading=%s, error=%s",
            liq_debug.get("retirement_year"),
            liq_debug.get("table_page_found"),
            liq_debug.get("liquid_heading_pages"),
            liq_debug.get("error"),
        )
    # Slide 14: Retirement Summary / shortfall / funding page
    result.update(_extract_financial_summary_slide14(doc))
    # First page: client name and report date (for {{CLIENT_NAME}}, {{REPORT_MONTH}}, {{REPORT_YEAR}})
    result.update(_extract_first_page_meta(doc))
    return result


def _extract_first_page_meta(doc: fitz.Document) -> dict[str, Optional[str]]:
    """Extract client name and report date from first page. Returns client_name, report_date, report_month, report_year."""
    out = {"client_name": None, "report_date": None, "report_month": None, "report_year": None}
    if len(doc) == 0:
        return out
    text = get_page_text(doc, 0)
    # Client name: line after "Financial Plan for"
    m = re.search(r"Financial\s+Plan\s+for\s*[\r\n]+\s*([^\r\n]+)", text, re.IGNORECASE)
    if m:
        out["client_name"] = m.group(1).strip()
    # Date: after "Prepared:" in DD/MM/YYYY format
    m = re.search(r"Prepared:\s*[\r\n]*(\d{1,2}/\d{1,2}/\d{4})", text, re.IGNORECASE)
    if m:
        date_str = m.group(1).strip()
        out["report_date"] = date_str
        parts = date_str.split("/")
        if len(parts) == 3:
            try:
                day, month, year = int(parts[0]), int(parts[1]), int(parts[2])
                if 1 <= month <= 12 and year > 1900:
                    _MONTH_NAMES = ["", "January", "February", "March", "April", "May", "June",
                                    "July", "August", "September", "October", "November", "December"]
                    out["report_month"] = _MONTH_NAMES[month]
                    out["report_year"] = str(year)
            except ValueError:
                pass
    return out


def _extract_post_retirement_spending_from_any_page(doc: fitz.Document) -> tuple[Optional[int], Optional[int], Optional[str]]:
    """
    Search ALL pages for "You can afford to spend £X" / "Projected Spending Capacity" / "annually excluding taxes in retirement".
    Used when the Retirement Summary (Slide 14) page doesn't contain this text (it often appears on Insights | Retirement Spending or Financial Summary).
    Returns (value, 1_based_page_num, matched_span_for_debug) or (None, None, None).
    """
    flags = re.IGNORECASE | re.DOTALL
    # Patterns that indicate the spending capacity sentence (order: prefer full phrase then anchor)
    anchors = [
        "You can afford to spend",
        "Projected Spending Capacity in Retirement",
        "annually excluding taxes in retirement",
    ]
    patterns = [
        r"You\s+can\s+afford\s+to\s+spend\s*(?:£\s*)?([\d,]+).*?(?:per\s+year|annually|excluding)",
        r"Projected\s+Spending\s+Capacity.*?(?:£\s*)?([\d,]+)",
        r"afford\s+to\s+spend\s*(?:£\s*)?([\d,]+)\s+annually\s+excluding\s+taxes",
    ]
    for page_idx in range(len(doc)):
        text = get_page_text(doc, page_idx)
        for pattern in patterns:
            m = re.search(pattern, text, flags)
            if m:
                try:
                    val = int(m.group(1).replace(",", ""))
                    return (val, page_idx + 1, m.group(0)[:300])
                except (ValueError, IndexError):
                    continue
        # Fallback: any page containing an anchor, then first £ amount in window
        for anchor in anchors:
            idx = text.find(anchor)
            if idx >= 0:
                window = text[idx : idx + 250]
                mm = re.search(r"£\s*([\d,]+)", window)
                if mm:
                    try:
                        val = int(mm.group(1).replace(",", ""))
                        return (val, page_idx + 1, mm.group(0)[:200])
                    except ValueError:
                        continue
    return (None, None, None)


def _extract_financial_summary_slide14(doc: fitz.Document, spans_out: Optional[dict[str, str]] = None) -> dict[str, Optional[int]]:
    """Find page by Slide 14 anchor phrases and extract placeholders. If spans_out is provided, record matched text span for each key."""
    page_num, _anchor, _snippet = _find_slide14_page(doc)
    if page_num is None:
        return {
            "shortfall_years": None,
            "total_retirement_years": None,
            "lump_sum_required": None,
            "retirement_year_lump": None,
            "annual_savings_required": None,
            "post_not_funded_years": None,
            "post_funded_years": None,
            "post_retirement_spending": None,
        }
    text = get_page_text(doc, page_num)
    flags = re.IGNORECASE | re.DOTALL
    out = {
        "shortfall_years": None,
        "total_retirement_years": None,
        "lump_sum_required": None,
        "retirement_year_lump": None,
        "annual_savings_required": None,
        "post_not_funded_years": None,
        "post_funded_years": None,
        "post_retirement_spending": None,
    }
    m = re.search(r"shortfall\s+in\s+(\d+)\s+of\s+(\d+)\s+retirement\s+years", text, flags)
    if m:
        out["shortfall_years"] = int(m.group(1))
        out["total_retirement_years"] = int(m.group(2))
        if spans_out is not None:
            spans_out["shortfall_years"] = m.group(0)[:300]
            spans_out["total_retirement_years"] = m.group(0)[:300]
    # LUMP_SUM_REQUIRED + RETIREMENT_YEAR: first try both in one go (DOTALL allows newlines)
    m = re.search(r"additional\s+lump\s+sum\s+of\s*(?:£\s*)?([\d,]+).*?\s+in\s+(\d{4})", text, flags)
    if m:
        out["lump_sum_required"] = int(m.group(1).replace(",", ""))
        out["retirement_year_lump"] = int(m.group(2))
        if spans_out is not None:
            spans_out["lump_sum_required"] = m.group(0)[:300]
            spans_out["retirement_year_lump"] = m.group(0)[:300]
    else:
        my = re.search(r"lump\s+sum.*?in\s+(\d{4})", text, flags)
        if my:
            out["retirement_year_lump"] = int(my.group(1))
            if spans_out is not None:
                spans_out["retirement_year_lump"] = my.group(0)[:300]
        idx = text.find("additional lump sum")
        if idx >= 0:
            region = text[idx : idx + 300]
            mm = re.search(r"£\s*([\d,]+)", region)
            if mm and out["lump_sum_required"] is None:
                out["lump_sum_required"] = int(mm.group(1).replace(",", ""))
                if spans_out is not None:
                    spans_out["lump_sum_required"] = mm.group(0)[:300]
    # ANNUAL_SAVINGS_REQUIRED (PRE): line-break tolerant — £ can be before or after "per year"
    # (a) save an additional ... £X ... per year
    m = re.search(r"save\s+an\s+additional.*?£\s*([\d,]+).*?per\s+year", text, flags)
    if m:
        out["annual_savings_required"] = int(m.group(1).replace(",", ""))
        if spans_out is not None:
            spans_out["annual_savings_required"] = m.group(0)[:300]
    if out["annual_savings_required"] is None:
        # (b) save an additional ... per year ... £X (Voyant often puts £ on next line after "per year")
        m = re.search(r"save\s+an\s+additional.*?per\s+year.*?£\s*([\d,]+)", text, flags)
        if m:
            out["annual_savings_required"] = int(m.group(1).replace(",", ""))
            if spans_out is not None:
                spans_out["annual_savings_required"] = m.group(0)[:300]
    if out["annual_savings_required"] is None:
        # (c) Option B: first £ amount in window after "save an additional" (250–400 chars)
        idx = text.lower().find("save an additional")
        if idx >= 0:
            window = text[idx : idx + 400]
            mm = re.search(r"£\s*([\d,]+)", window)
            if mm:
                out["annual_savings_required"] = int(mm.group(1).replace(",", ""))
                if spans_out is not None:
                    spans_out["annual_savings_required"] = mm.group(0)[:300]
                    spans_out["annual_savings_required_window"] = window[:400]
    if out["annual_savings_required"] is None:
        # "Save an additional £X annually" (compact form)
        m = re.search(r"Save\s+an\s+additional\s*(?:£\s*)?([\d,]+)\s+annually", text, flags)
        if m:
            out["annual_savings_required"] = int(m.group(1).replace(",", ""))
            if spans_out is not None:
                spans_out["annual_savings_required"] = m.group(0)[:300]
    # POST "not funded" years: robust to PyMuPDF text order (e.g. "funded in not X years" when "not" is jumbled)
    m = re.search(r"Expenses\s+are\s+(?:not\s+)?funded\s+in\s+(?:not\s+)?(\d+)\s+years", text, flags)
    if not m:
        m = re.search(r"Expenses\s+are\s+not\s+funded\s+in\s+(\d+)\s+years", text, flags)
    if not m:
        m = re.search(r"Expenses\s+are\s+funded\s+in\s+not\s+(\d+)\s+years", text, flags)
    if m:
        out["post_not_funded_years"] = int(m.group(1))
        if spans_out is not None:
            spans_out["post_not_funded_years"] = m.group(0)[:300]
    m = re.search(r"Expenses\s+are\s+funded\s+in\s+(?!not\s)(\d+)\s+years", text, flags)
    if m:
        out["post_funded_years"] = int(m.group(1))
        if spans_out is not None:
            spans_out["post_funded_years"] = m.group(0)[:300]
    # POST_RETIREMENT_SPENDING: £value may be on own line after "You can afford to spend"
    m = re.search(r"You\s+can\s+afford\s+to\s+spend\s*(?:£\s*)?([\d,]+).*?(?:per\s+year|annually)", text, flags)
    if m:
        out["post_retirement_spending"] = int(m.group(1).replace(",", ""))
        if spans_out is not None:
            spans_out["post_retirement_spending"] = m.group(0)[:300]
    if out["post_retirement_spending"] is None:
        idx = text.find("You can afford to spend")
        if idx >= 0:
            region = text[idx : idx + 200]
            mm = re.search(r"£\s*([\d,]+)", region)
            if mm:
                out["post_retirement_spending"] = int(mm.group(1).replace(",", ""))
                if spans_out is not None:
                    spans_out["post_retirement_spending"] = mm.group(0)[:300]
    # If still missing, search all pages (Insights | Retirement Spending / Financial Summary often have this, not Slide 14)
    if out["post_retirement_spending"] is None:
        val, from_page, span = _extract_post_retirement_spending_from_any_page(doc)
        if val is not None:
            out["post_retirement_spending"] = val
            if spans_out is not None and span:
                spans_out["post_retirement_spending"] = span
                spans_out["post_retirement_spending_page"] = str(from_page)
    return out


def extract_financial_summary_slide14_debug(doc: fitz.Document) -> tuple[dict[str, Optional[int]], dict]:
    """Extract Slide 14 values and return (values_dict, debug_dict). Debug includes slide14_page, matched_anchor_phrase, page_snippet, matched_spans (exact text matched per field)."""
    debug = {"slide14_page": None, "matched_anchor_phrase": None, "page_snippet": None, "page_text_length": 0, "matched_spans": {}, "regex_results": []}
    page_num, matched_anchor, snippet = _find_slide14_page(doc)
    regex_specs = [
        ("shortfall in X of Y retirement years", r"shortfall\s+in\s+(\d+)\s+of\s+(\d+)\s+retirement\s+years", ("shortfall_years", "total_retirement_years")),
        ("lump sum + year (DOTALL)", r"additional\s+lump\s+sum\s+of\s*(?:£\s*)?([\d,]+).*?\s+in\s+(\d{4})", ("lump_sum_required", "retirement_year_lump")),
        ("save an additional ... £X ... per year (a)", r"save\s+an\s+additional.*?£\s*([\d,]+).*?per\s+year", ("annual_savings_required", None)),
        ("save an additional ... per year ... £X (b)", r"save\s+an\s+additional.*?per\s+year.*?£\s*([\d,]+)", ("annual_savings_required", None)),
        ("save an additional ... annually", r"Save\s+an\s+additional\s*(?:£\s*)?([\d,]+)\s+annually", ("annual_savings_required", None)),
        ("Expenses are not funded in X years", r"Expenses\s+are\s+not\s+funded\s+in\s+(\d+)\s+years", ("post_not_funded_years", None)),
        ("Expenses are funded in not X years (jumbled)", r"Expenses\s+are\s+funded\s+in\s+not\s+(\d+)\s+years", ("post_not_funded_years", None)),
        ("Expenses are [not] funded in [not] X (flexible)", r"Expenses\s+are\s+(?:not\s+)?funded\s+in\s+(?:not\s+)?(\d+)\s+years", ("post_not_funded_years", None)),
        ("Expenses are funded in X years", r"Expenses\s+are\s+funded\s+in\s+(?!not\s)(\d+)\s+years", ("post_funded_years", None)),
        ("You can afford to spend ... (DOTALL)", r"You\s+can\s+afford\s+to\s+spend\s*(?:£\s*)?([\d,]+).*?(?:per\s+year|annually)", ("post_retirement_spending", None)),
    ]
    if page_num is None:
        debug["page_snippet"] = ""
        for name, pat, _ in regex_specs:
            debug["regex_results"].append({"name": name, "pattern": pat, "matched": False, "match_text": None, "groups": [], "extracted": {}})
        return (_extract_financial_summary_slide14(doc, spans_out=debug["matched_spans"]), debug)
    debug["slide14_page"] = page_num + 1
    debug["matched_anchor_phrase"] = matched_anchor
    debug["page_snippet"] = snippet
    text = get_page_text(doc, page_num)
    debug["page_text_length"] = len(text)
    flags_debug = re.IGNORECASE | re.DOTALL
    values = _extract_financial_summary_slide14(doc, spans_out=debug["matched_spans"])
    debug["annual_savings_required_extracted"] = values.get("annual_savings_required")
    for (name, pattern, keys) in regex_specs:
        m = re.search(pattern, text, flags_debug)
        if m:
            groups = list(m.groups())
            extracted = {}
            if keys[0]:
                raw = (groups[0] or "").replace(",", "") if groups else ""
                if not raw and len(groups) > 1:
                    raw = (groups[1] or "").replace(",", "")
                try:
                    extracted[keys[0]] = int(float(raw)) if raw else None
                except ValueError:
                    extracted[keys[0]] = groups[0] or (groups[1] if len(groups) > 1 else None)
            if keys[1] and len(groups) > 1 and keys[1] != keys[0]:
                raw = (groups[1] or "").replace(",", "")
                try:
                    extracted[keys[1]] = int(float(raw)) if raw else None
                except ValueError:
                    extracted[keys[1]] = groups[1]
            debug["regex_results"].append({"name": name, "pattern": pattern, "matched": True, "match_text": m.group(0)[:200], "groups": groups, "extracted": extracted})
        else:
            debug["regex_results"].append({"name": name, "pattern": pattern, "matched": False, "match_text": None, "groups": [], "extracted": {}})
    return (values, debug)


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


def _get_retirement_year_for_liquid(doc: fitz.Document) -> int:
    """Prefer retirement year from Retirement Summary (Slide 14) page; else fall back to extract_retirement_year."""
    page_num, _, _ = _find_slide14_page(doc)
    if page_num is not None:
        text = get_page_text(doc, page_num)
        m = re.search(r"lump\s+sum.*?in\s+(\d{4})", text, re.IGNORECASE | re.DOTALL)
        if m:
            y = int(m.group(1))
            if 2030 <= y <= 2049:
                return y
    return extract_retirement_year(doc)


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


def extract_liquid_assets_debug(doc: fitz.Document, retirement_year_override: Optional[int] = None) -> dict:
    """
    Run liquid assets table extraction and return debug_info dict.
    If retirement_year_override is set, use it for column lookup instead of deriving from PDF.
    Debug includes: pages with heading, page scores (why chosen/skipped), retirement year, parsed rows, header column, values summed, error if any.
    """
    debug = {
        "retirement_year": None,
        "liquid_heading_pages": [],
        "candidate_page_scores": [],
        "table_page_found": None,
        "savings_found_on_table_page": None,
        "investments_found_on_table_page": None,
        "pensions_found_on_table_page": None,
        "num_rows_parsed": 0,
        "row_preview": [],
        "header_row_index": None,
        "header_col_index": None,
        "retirement_year_ordinal": None,
        "year_in_header": None,
        "candidate_years_on_page": None,
        "block_start": None,
        "block_end": None,
        "detected_row_labels": [],
        "block_rows_scanned": [],
        "block_first_cells_as_read": [],
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
        # Retirement year: override (e.g. from retirement age) or from PDF (Slide 14 / fallback)
        debug["retirement_year"] = retirement_year_override if retirement_year_override is not None else _get_retirement_year_for_liquid(doc)
        # Candidate page scores: why each page was chosen or skipped
        for i in range(len(doc)):
            text = get_page_text(doc, i)
            if LIQUID_ASSETS_HEADING not in text:
                continue
            year_count, currency_count, label_bonus, has_all_labels = _score_liquid_assets_page(text)
            score = year_count * 10 + currency_count + label_bonus if (year_count >= 1 and currency_count >= 1) else None
            debug["candidate_page_scores"].append({
                "page": i + 1,
                "year_count": year_count,
                "currency_count": currency_count,
                "label_bonus": label_bonus,
                "has_all_labels": has_all_labels,
                "score": score,
                "included": score is not None,
            })
        # Table page (same logic as find_liquid_assets_table_page)
        page_num = find_liquid_assets_table_page(doc)
        debug["table_page_found"] = (page_num + 1) if page_num is not None else None
        if page_num is None:
            debug["error"] = (
                "No Liquid Assets page had both year headers (e.g. 2038) and £ values. "
                "Intro/chart-only pages were skipped. Check candidate_page_scores to see why each page was excluded."
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
        # Find year header row and column
        header_row_index, retirement_col_index = _find_year_header_row(rows, debug["retirement_year"])
        debug["header_row_index"] = header_row_index
        debug["header_col_index"] = retirement_col_index
        # Year ordinal: position of retirement year among 4-digit-year header cells (avoids header token drift)
        header_cells = rows[header_row_index][1] if header_row_index is not None and header_row_index < len(rows) else []
        year_ordinal = _get_year_ordinal_from_header_row(header_cells, debug["retirement_year"]) if header_cells else None
        debug["retirement_year_ordinal"] = year_ordinal
        debug["year_in_header"] = str(debug["retirement_year"]) if retirement_col_index is not None else None
        if header_row_index is None or retirement_col_index is None:
            # List candidate years found on the table page (from any cell)
            candidate_years = set()
            for _, cells in rows:
                for c in cells:
                    for m in _YEAR_TOKEN_RE.findall(c):
                        candidate_years.add(int(m))
            debug["candidate_years_on_page"] = sorted(candidate_years) if candidate_years else None
            debug["error"] = (
                f"Retirement year {debug['retirement_year']} not found in any row. "
                f"Candidate years on this page: {debug.get('candidate_years_on_page') or 'none'}. "
                f"Row starts: {[c[:2] for _, c in rows[:5]]}"
            )
            return debug
        block_start = header_row_index
        block_end = _find_block_end(rows, header_row_index)
        debug["block_start"] = block_start
        debug["block_end"] = block_end
        total, values_by_label = _sum_liquid_from_block(
            rows, header_row_index, retirement_col_index, block_end, debug_out=debug, year_ordinal=year_ordinal
        )
        debug["detected_row_labels"] = list(values_by_label.keys())
        debug["values_summed"] = dict(values_by_label)
        debug["data_rows_found"] = [(k, v) for k, v in values_by_label.items()]
        debug["total"] = total if total else None
        if total == 0:
            debug["error"] = "No numeric values summed in block (check first-cell row labels and column alignment)."
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


def _get_year_ordinal_from_header_row(cells: list[str], retirement_year: int) -> Optional[int]:
    """
    Count only header cells that are 4-digit years; return 0-based position of retirement_year among them.
    This avoids index drift when header has stray tokens (e.g. "1 2"). Returns None if retirement_year not in year list.
    """
    year_str = str(retirement_year)
    pos = 0
    for cell in cells:
        c = (cell or "").strip()
        if not _cell_looks_like_year(c):
            continue
        if c == year_str:
            return pos
        pos += 1
    return None


def _find_year_header_row(
    rows: list[tuple[float, list[str]]], retirement_year: int
) -> tuple[Optional[int], Optional[int]]:
    """
    Find the year header row that contains the retirement year (exact string match in a cell).
    Column index is determined by: for idx, cell in enumerate(header_row): if cell.strip() == str(retirement_year): column = idx.
    Do not assume consistent column index across PDFs. Returns (header_row_index, retirement_col_index) or (None, None).
    """
    year_str = str(retirement_year)
    candidates = []  # (row_index, col_index, year_like_cell_count)
    for ri, (_, cells) in enumerate(rows):
        col_index = None
        for idx, cell in enumerate(cells):
            if (cell or "").strip() == year_str:
                col_index = idx
                break
        if col_index is None:
            continue
        year_like_count = sum(1 for c in cells if _cell_looks_like_year(c))
        candidates.append((ri, col_index, year_like_count))
    if not candidates:
        return None, None
    candidates.sort(key=lambda x: (-x[2], x[0]))
    return candidates[0][0], candidates[0][1]


# Expected row labels in Liquid Assets block (order may vary; we detect dynamically via startswith)
LIQUID_ROW_LABELS = frozenset({"Savings", "Investments", "Pensions"})


def _liquid_row_label_from_first_cell(first_cell: str) -> Optional[str]:
    """Normalized match: first_cell.strip().lower() startswith savings/investments/pensions. Returns canonical label or None."""
    row_label = (first_cell or "").strip().lower()
    if not row_label:
        return None
    if row_label.startswith("savings"):
        return "Savings"
    if row_label.startswith("investments"):
        return "Investments"
    if row_label.startswith("pensions"):
        return "Pensions"
    return None


def _find_block_end(rows: list[tuple[float, list[str]]], header_row_index: int) -> int:
    """
    Block starts at header_row_index. Block ends at the next row whose first cell is a 4-digit year, or end of table.
    Returns the exclusive end index (first row index not in this block).
    """
    for ri in range(header_row_index + 1, len(rows)):
        _, cells = rows[ri]
        if not cells:
            continue
        first_cell = (cells[0] or "").strip()
        if _cell_looks_like_year(first_cell):
            return ri
    return len(rows)


def _sum_liquid_from_block(
    rows: list[tuple[float, list[str]]],
    header_row_index: int,
    retirement_col_index: int,
    block_end: int,
    debug_out: Optional[dict] = None,
    year_ordinal: Optional[int] = None,
) -> tuple[int, dict[str, int]]:
    """
    Scan rows from header_row_index+1 until block_end (exclusive). Detect rows by first cell using
    normalized startswith(savings|investments|pensions). Column for retirement year: if year_ordinal
    is set, use 1 + year_ordinal (position among year columns); else use retirement_col_index. If
    effective column >= len(row), treat that row as £0 (blank/drawn down for that year), do not use
    the last cell. debug_out: optional dict to add block_rows_scanned, block_first_cells_as_read.
    """
    total = 0
    values_by_label = {}
    block_rows_scanned = []
    first_cells_as_read = []
    for ri in range(header_row_index + 1, block_end):
        if ri >= len(rows):
            break
        _, cells = rows[ri]
        first_cell_raw = (cells[0] if cells else "")
        first_cells_as_read.append(first_cell_raw)
        canonical_label = _liquid_row_label_from_first_cell(first_cell_raw)
        if canonical_label is None:
            if debug_out is not None:
                block_rows_scanned.append((ri, first_cell_raw, list(cells) if cells else []))
            continue
        if year_ordinal is not None:
            effective_col = 1 + year_ordinal
        else:
            effective_col = retirement_col_index
        if effective_col >= len(cells):
            # Column exists in header but this row has no value for that year (e.g. asset drawn down) → £0
            logger.debug(
                "Liquid assets: row index %s retirement column %s >= len(row)=%s, treating as £0; row=%s",
                ri, effective_col, len(cells), cells,
            )
            values_by_label[canonical_label] = 0
            if debug_out is not None:
                block_rows_scanned.append((ri, first_cell_raw, list(cells) if cells else []))
            continue
        val = _parse_int_from_cell(cells[effective_col])
        if val is not None:
            total += val
            values_by_label[canonical_label] = val
        if debug_out is not None:
            block_rows_scanned.append((ri, first_cell_raw, list(cells) if cells else []))
    if debug_out is not None:
        debug_out["block_rows_scanned"] = block_rows_scanned
        debug_out["block_first_cells_as_read"] = first_cells_as_read
    return total, values_by_label


def _extract_liquid_assets_table_total(doc: fitz.Document, retirement_year_override: Optional[int] = None) -> Optional[int]:
    """
    Find Liquid Assets table page; get retirement year from override (if set), else from Retirement Summary when possible.
    Identify header row and column for that year; define block (until next year row or end);
    within block only, detect Savings/Investments/Pensions by first cell and sum values at retirement column.
    """
    page_num = find_liquid_assets_table_page(doc)
    if page_num is None:
        return None
    page = doc[page_num]
    retirement_year = retirement_year_override if retirement_year_override is not None else _get_retirement_year_for_liquid(doc)
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
    header_row_index, retirement_col_index = _find_year_header_row(rows, retirement_year)
    if header_row_index is None or retirement_col_index is None:
        logger.info(
            "Liquid assets: retirement_year=%s, header_row_index=%s, retirement_column_index=%s (not found)",
            retirement_year, header_row_index, retirement_col_index,
        )
        return None
    header_cells = rows[header_row_index][1] if header_row_index < len(rows) else []
    year_ordinal = _get_year_ordinal_from_header_row(header_cells, retirement_year) if header_cells else None
    block_start = header_row_index
    block_end = _find_block_end(rows, header_row_index)
    liquid_debug = {}
    total, values_by_label = _sum_liquid_from_block(
        rows, header_row_index, retirement_col_index, block_end, debug_out=liquid_debug, year_ordinal=year_ordinal
    )
    detected_row_labels = list(values_by_label.keys())
    logger.info(
        "Liquid assets: retirement_year=%s, header_row_index=%s, retirement_column_index=%s, "
        "block_start=%s, block_end=%s, detected_row_labels=%s, extracted_values=%s, final_sum=%s",
        retirement_year, header_row_index, retirement_col_index,
        block_start, block_end, detected_row_labels, values_by_label, total,
    )
    if liquid_debug.get("block_rows_scanned"):
        logger.debug("Liquid assets block_rows_scanned: %s", liquid_debug.get("block_rows_scanned"))
    if liquid_debug.get("block_first_cells_as_read"):
        logger.debug("Liquid assets block_first_cells_as_read: %s", liquid_debug.get("block_first_cells_as_read"))
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

def extract_values_from_pdf(pdf_path: Path, retirement_year_override: Optional[int] = None) -> dict[str, Any]:
    """Returns dict with all extracted values (int/str/None). If retirement_year_override is set, use it for liquid assets column lookup."""
    doc = fitz.open(pdf_path)
    try:
        return extract_values_from_doc(doc, retirement_year_override=retirement_year_override)
    finally:
        doc.close()


# ---------------------------------------------------------------------------
# Pre-advice: charts + values (same interface for app/summary)
# ---------------------------------------------------------------------------

def extract_pre_advice(pdf_path: Path, charts_dir: Path, retirement_year_override: Optional[int] = None) -> tuple[dict, dict]:
    """Returns (values_dict, charts_dict). values_dict uses retirement_spending_annual, estate_tax_iht, etc. for compatibility.
    If retirement_year_override is set, use it for liquid assets column lookup (e.g. birth_year + pre_retirement_age).
    """
    raw_values = extract_values_from_pdf(pdf_path, retirement_year_override=retirement_year_override)
    values = {
        "retirement_spending_annual": raw_values.get("retirement_spending"),
        "estate_tax_iht": raw_values.get("estate_tax"),
        "total_estate_transfer": raw_values.get("estate_transfer"),
        "net_taxable_estate": raw_values.get("net_taxable"),
        "liquid_assets_retirement": raw_values.get("liquid_assets_retirement"),
        "shortfall_years": raw_values.get("shortfall_years"),
        "total_retirement_years": raw_values.get("total_retirement_years"),
        "lump_sum_required": raw_values.get("lump_sum_required"),
        "retirement_year_lump": raw_values.get("retirement_year_lump"),
        "annual_savings_required": raw_values.get("annual_savings_required"),
        "post_not_funded_years": raw_values.get("post_not_funded_years"),
        "post_funded_years": raw_values.get("post_funded_years"),
        "post_retirement_spending": raw_values.get("post_retirement_spending"),
        "client_name": raw_values.get("client_name"),
        "report_date": raw_values.get("report_date"),
        "report_month": raw_values.get("report_month"),
        "report_year": raw_values.get("report_year"),
    }
    charts = extract_charts_from_pdf(pdf_path, charts_dir, "pre")
    return values, charts


# ---------------------------------------------------------------------------
# Post-advice: charts + values
# ---------------------------------------------------------------------------

def extract_post_advice(pdf_path: Path, charts_dir: Path, retirement_year_override: Optional[int] = None) -> tuple[dict, dict]:
    """Returns (values_dict, charts_dict). If retirement_year_override is set, use it for liquid assets column lookup."""
    raw_values = extract_values_from_pdf(pdf_path, retirement_year_override=retirement_year_override)
    values = {
        "retirement_spending_annual": raw_values.get("retirement_spending"),
        "estate_tax_iht": raw_values.get("estate_tax"),
        "total_estate_transfer": raw_values.get("estate_transfer"),
        "net_taxable_estate": raw_values.get("net_taxable"),
        "liquid_assets_retirement": raw_values.get("liquid_assets_retirement"),
        "shortfall_years": raw_values.get("shortfall_years"),
        "total_retirement_years": raw_values.get("total_retirement_years"),
        "lump_sum_required": raw_values.get("lump_sum_required"),
        "retirement_year_lump": raw_values.get("retirement_year_lump"),
        "annual_savings_required": raw_values.get("annual_savings_required"),
        "post_not_funded_years": raw_values.get("post_not_funded_years"),
        "post_funded_years": raw_values.get("post_funded_years"),
        "post_retirement_spending": raw_values.get("post_retirement_spending"),
        "client_name": raw_values.get("client_name"),
        "report_date": raw_values.get("report_date"),
        "report_month": raw_values.get("report_month"),
        "report_year": raw_values.get("report_year"),
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
    pre_retirement_year: Optional[int] = None,
    post_retirement_year: Optional[int] = None,
) -> tuple[Path, dict, dict, dict]:
    """
    Run full extraction. Returns (output_dir, pre_values, post_values, all_charts).
    If pre_retirement_year / post_retirement_year are set, use them for liquid assets column lookup (e.g. birth_year + retirement_age).
    """
    safe_name = client_name.strip()
    for c in " /\\:*?\"<>|":
        safe_name = safe_name.replace(c, "_")
    output_dir = output_base / f"{safe_name}_{report_date}"
    charts_dir = output_dir / "charts"
    charts_dir.mkdir(parents=True, exist_ok=True)

    pre_values, pre_charts = extract_pre_advice(pre_advice_path, charts_dir, retirement_year_override=pre_retirement_year)
    post_values, post_charts = extract_post_advice(post_advice_path, charts_dir, retirement_year_override=post_retirement_year)
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

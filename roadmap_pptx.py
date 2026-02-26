"""
Populate the generic RoadMap template with extracted charts and values.
Uses shape names (Selection Pane names) to find placeholders. Slide numbers are 1-based in spec, 0-based in code.
"""

import logging
from pathlib import Path
from typing import Optional, Tuple, Any

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

logger = logging.getLogger(__name__)


# Slide index (0-based) for each target slide
SLIDE_6 = 5
SLIDE_8 = 7
SLIDE_9 = 8
SLIDE_12 = 11
SLIDE_13 = 12
SLIDE_14 = 13
SLIDE_19 = 18
SLIDE_24 = 23


def _fmt_liquid_millions(value: Optional[int]) -> str:
    """Format liquid total (in £) as c.£Xm for Slide 6. One decimal if needed; no trailing .0 if whole."""
    if value is None:
        return "—"
    millions = value / 1_000_000
    if millions == int(millions):
        return f"c.£{int(millions)}m"
    return f"c.£{millions:.1f}m"


def _fmt_gbp(value: Optional[int]) -> str:
    """Format int as £ with commas for Slide 14 (and similar)."""
    if value is None:
        return "—"
    return f"£{value:,}"


def _get_slide_text(slide) -> str:
    """Collect all text from a slide (for placeholder check)."""
    def collect(shapes):
        parts = []
        for shape in shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        if run.text:
                            parts.append(run.text)
            if getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.GROUP:
                parts.extend(collect(shape.shapes))
        return parts
    return "".join(collect(slide.shapes))


def _check_slide6_placeholders(slide, token_names: list[str]) -> list[str]:
    """If any placeholder token still appears in the slide text, return list of missing names."""
    text = _get_slide_text(slide)
    return [name for name in token_names if name in text]


def _find_shape_by_name(slide, name: str) -> Tuple[Optional[Any], int, int]:
    """
    Find a shape by name on the slide, including inside groups.
    Returns (shape, abs_left, abs_top) in slide coordinates, or (None, 0, 0).
    """
    def search(shapes, offset_left: int, offset_top: int):
        for shape in shapes:
            if shape.name == name:
                return shape, offset_left + shape.left, offset_top + shape.top
            if getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.GROUP:
                found, abs_left, abs_top = search(shape.shapes, offset_left + shape.left, offset_top + shape.top)
                if found is not None:
                    return found, abs_left, abs_top
        return None, 0, 0
    return search(slide.shapes, 0, 0)


def _replace_shape_with_image(slide, shape_name: str, image_path: Path) -> bool:
    shape, abs_left, abs_top = _find_shape_by_name(slide, shape_name)
    if shape is None or not image_path.exists():
        return False
    width, height = shape.width, shape.height
    sp = shape._element.getparent()
    sp.remove(shape._element)
    slide.shapes.add_picture(str(image_path), abs_left, abs_top, width, height)
    return True


def _replace_text_tokens_in_shapes(shapes, tokens: dict[str, str]) -> None:
    """Replace tokens in all text in shapes (and recurse into groups)."""
    for shape in shapes:
        if getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.GROUP:
            _replace_text_tokens_in_shapes(shape.shapes, tokens)
            continue
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            full_text = "".join(run.text for run in para.runs)
            if not full_text:
                continue
            new_text = full_text
            for token, value in tokens.items():
                new_text = new_text.replace(token, value)
            if new_text == full_text:
                continue
            if para.runs:
                para.runs[0].text = new_text
                for i in range(1, len(para.runs)):
                    para.runs[i].text = ""


def _replace_text_tokens(slide, tokens: dict[str, str]) -> None:
    """
    Replace placeholder tokens with values in all text on the slide (including inside groups).
    Works at paragraph level so placeholders split across multiple runs are still replaced.
    """
    _replace_text_tokens_in_shapes(slide.shapes, tokens)


def populate_roadmap_pptx(
    template_path: Path,
    output_path: Path,
    charts_dir: Path,
    all_charts: dict[str, str],
    retirement_annual_diff: int,
    retirement_monthly_diff: int,
    liquid_pre: Optional[int] = None,
    liquid_post: Optional[int] = None,
    shortfall_years: Optional[int] = None,
    total_retirement_years: Optional[int] = None,
    lump_sum_required: Optional[int] = None,
    retirement_year: Optional[int] = None,
    annual_savings_required: Optional[int] = None,
    post_not_funded_years: Optional[int] = None,
    post_funded_years: Optional[int] = None,
    post_retirement_spending: Optional[int] = None,
) -> None:
    """
    charts_dir is the charts folder (output_dir / "charts").
    all_charts maps key -> relative path e.g. "charts/pre_timeline_page4.png"; base is output_dir so full path = charts_dir.parent / rel_path.
    """
    prs = Presentation(str(template_path))
    base_dir = charts_dir.parent

    def chart_path(key: str) -> Optional[Path]:
        if key not in all_charts:
            return None
        return base_dir / all_charts[key]

    # Slide 6 – Financial Implications (text replacement by placeholder name)
    # Retirement: pre/post from "You can afford to spend £X annually…"; diffs formatted as £ per month / per year.
    # Liquid: from Liquid Assets table total at retirement year; format c.£Xm / c.£Ym.
    if len(prs.slides) > SLIDE_6:
        slide = prs.slides[SLIDE_6]
        tokens = {
            "{{RETIREMENT_MONTHLY_DIFF}}": f"£{retirement_monthly_diff:,}",
            "{{RETIREMENT_ANNUAL_DIFF}}": f"£{retirement_annual_diff:,}",
            "{{LIQUID_ASSETS_PRE}}": _fmt_liquid_millions(liquid_pre),
            "{{LIQUID_ASSETS_POST}}": _fmt_liquid_millions(liquid_post),
        }
        _replace_text_tokens(slide, tokens)
        missing = _check_slide6_placeholders(slide, list(tokens.keys()))
        if missing:
            logger.error(
                "Slide 6 — placeholder shape(s) not found or not replaced: %s",
                ", ".join(missing),
            )

    # Slide 8 – timeline (use pre_timeline)
    if len(prs.slides) > SLIDE_8:
        path = chart_path("pre_timeline")
        if path:
            _replace_shape_with_image(prs.slides[SLIDE_8], "[TIMELINE_IMAGE]", path)

    # Slide 9 – pre cashflow (top) + pre liquid (bottom)
    # Placeholder names in template Selection Pane: PRE_CASHFLOW_IMAGE = top, PRE_LIQUID_IMAGE = bottom
    if len(prs.slides) > SLIDE_9:
        s9 = prs.slides[SLIDE_9]
        p = chart_path("pre_cashflow")
        if p:
            _replace_shape_with_image(s9, "PRE_CASHFLOW_IMAGE", p)
        p = chart_path("pre_liquid_assets")
        if p:
            _replace_shape_with_image(s9, "PRE_LIQUID_IMAGE", p)

    # Slide 12 – cashflow pre vs post
    if len(prs.slides) > SLIDE_12:
        s12 = prs.slides[SLIDE_12]
        p = chart_path("pre_cashflow")
        if p:
            _replace_shape_with_image(s12, "PRE_CASHFLOW_COMPARISON", p)
        p = chart_path("post_cashflow")
        if p:
            _replace_shape_with_image(s12, "POST_CASHFLOW_COMPARISON", p)

    # Slide 13 – liquid pre vs post
    if len(prs.slides) > SLIDE_13:
        s13 = prs.slides[SLIDE_13]
        p = chart_path("pre_liquid_assets")
        if p:
            _replace_shape_with_image(s13, "PRE_LIQUID_COMPARISON", p)
        p = chart_path("post_liquid_assets")
        if p:
            _replace_shape_with_image(s13, "POST_LIQUID_COMPARISON", p)

    # Slide 14 – Financial Summary (from pre + post PDFs; replace placeholders in existing text)
    if len(prs.slides) > SLIDE_14:
        slide14 = prs.slides[SLIDE_14]
        pre_funded_years: Optional[int] = None
        if total_retirement_years is not None and shortfall_years is not None:
            pre_funded_years = total_retirement_years - shortfall_years
        tokens_14 = {
            "{{SHORTFALL_YEARS}}": str(shortfall_years) if shortfall_years is not None else "—",
            "{{TOTAL_RETIREMENT_YEARS}}": str(total_retirement_years) if total_retirement_years is not None else "—",
            "{{PRE_FUNDED_YEARS}}": str(pre_funded_years) if pre_funded_years is not None else "—",
            "{{LUMP_SUM_REQUIRED}}": _fmt_gbp(lump_sum_required),
            "{{RETIREMENT_YEAR}}": str(retirement_year) if retirement_year is not None else "—",
            "{{ANNUAL_SAVINGS_REQUIRED}}": _fmt_gbp(annual_savings_required),
            "{{POST_NOT_FUNDED_YEARS}}": str(post_not_funded_years) if post_not_funded_years is not None else "—",
            "{{POST_FUNDED_YEARS}}": str(post_funded_years) if post_funded_years is not None else "—",
            "{{POST_RETIREMENT_SPENDING}}": _fmt_gbp(post_retirement_spending),
        }
        _replace_text_tokens(slide14, tokens_14)

    # Slide 19 – comparison charts 1..4
    if len(prs.slides) > SLIDE_19:
        s19 = prs.slides[SLIDE_19]
        for i in range(1, 5):
            key = f"slide19_comparison_chart_{i}"
            path = chart_path(key)
            if path and path.exists():
                _replace_shape_with_image(s19, f"COMP_CHART_{i}", path)
            # else leave placeholder as-is

    # Slide 24 – post estate
    if len(prs.slides) > SLIDE_24:
        path = chart_path("slide24_estate_analysis")
        if path:
            _replace_shape_with_image(prs.slides[SLIDE_24], "POST_ESTATE_IMAGE", path)

    prs.save(str(output_path))

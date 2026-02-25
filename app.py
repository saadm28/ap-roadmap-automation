"""
AP Partners RoadMap Automation — Streamlit frontend.
Upload Voyant reports and FE Analytics charts to generate RoadMap assets.
"""

import io
import shutil
import tempfile
from pathlib import Path
from datetime import datetime

import streamlit as st
from PIL import Image

from extractor import run_extraction
from roadmap_pptx import populate_roadmap_pptx, _fmt_liquid_millions


# Session state keys for persisting results (so download doesn’t clear the page)
ROADMAP_KEYS = (
    "roadmap_output_dir",
    "roadmap_pre_values",
    "roadmap_post_values",
    "roadmap_all_charts",
    "roadmap_diff_annual",
    "roadmap_diff_monthly",
)
UPLOAD_RESET_KEY = "roadmap_upload_reset_key"


def _clear_roadmap_state():
    for key in ROADMAP_KEYS:
        if key in st.session_state:
            del st.session_state[key]
    st.session_state[UPLOAD_RESET_KEY] = st.session_state.get(UPLOAD_RESET_KEY, 0) + 1


def _render_results(output_dir: Path, pre_values: dict, post_values: dict, all_charts: dict, diff_annual: int, diff_monthly: int, key_prefix: str = ""):
    """Render: message, download/reset buttons, divider, then extracted values and charts."""
    pre_spend = pre_values.get("retirement_spending_annual") or 0
    post_spend = post_values.get("retirement_spending_annual") or 0
    liq_pre = pre_values.get("liquid_assets_retirement")
    liq_post = post_values.get("liquid_assets_retirement")

    # Message and primary actions first
    st.success("Review complete. Your RoadMap is ready.")
    tpl = Path(__file__).resolve().parent / "templates" / "Generic_RoadMap.pptx"
    if tpl.exists():
        pptx_path = output_dir / "RoadMap_GENERATED.pptx"
        if pptx_path.exists():
            st.download_button(
                "Download RoadMap PPTX",
                data=pptx_path.read_bytes(),
                file_name="RoadMap_GENERATED.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                type="primary",
                use_container_width=True,
                key=f"dl_pptx_{key_prefix}",
            )
        else:
            st.warning("Generated PPTX not found. Use the button below to create another RoadMap.")
        if st.button("Create another RoadMap", type="secondary", use_container_width=True, key=f"reset_{key_prefix}"):
            _clear_roadmap_state()
            try:
                shutil.rmtree(output_dir.parent, ignore_errors=True)
            except Exception:
                pass
            st.rerun()
    else:
        st.info("Place **Generic_RoadMap.pptx** in the `templates/` folder to enable Download RoadMap PPTX.")
        if st.button("Create another RoadMap", type="secondary", use_container_width=True, key=f"reset_no_pptx_{key_prefix}"):
            _clear_roadmap_state()
            st.rerun()

    # Divider before optional review (extractions and charts)
    st.divider()
    st.caption("Optional: review extracted values and charts below.")

    st.subheader("Extracted values")
    st.markdown("**Pre-Advice**")
    st.dataframe(
        [
            {"Metric": "Retirement Spending (Annual)", "Value": f"£{pre_spend:,}" if pre_spend else "—"},
            {"Metric": "Estate Tax (IHT)", "Value": f"£{pre_values.get('estate_tax_iht'):,}" if pre_values.get('estate_tax_iht') else "—"},
            {"Metric": "Total Estate Transfer", "Value": f"£{pre_values.get('total_estate_transfer'):,}" if pre_values.get('total_estate_transfer') else "—"},
            {"Metric": "Net Taxable Estate", "Value": f"£{pre_values.get('net_taxable_estate'):,}" if pre_values.get('net_taxable_estate') else "—"},
            {"Metric": "Liquid assets at retirement", "Value": _fmt_liquid_millions(pre_values.get('liquid_assets_retirement'))},
        ],
        use_container_width=True,
        hide_index=True,
    )
    st.markdown("**Post-Advice**")
    st.dataframe(
        [
            {"Metric": "Retirement Spending (Annual)", "Value": f"£{post_spend:,}" if post_spend else "—"},
            {"Metric": "Estate Tax (IHT)", "Value": f"£{post_values.get('estate_tax_iht'):,}" if post_values.get('estate_tax_iht') else "—"},
            {"Metric": "Total Estate Transfer", "Value": f"£{post_values.get('total_estate_transfer'):,}" if post_values.get('total_estate_transfer') else "—"},
            {"Metric": "Net Taxable Estate", "Value": f"£{post_values.get('net_taxable_estate'):,}" if post_values.get('net_taxable_estate') else "—"},
            {"Metric": "Liquid assets at retirement", "Value": _fmt_liquid_millions(post_values.get('liquid_assets_retirement'))},
        ],
        use_container_width=True,
        hide_index=True,
    )
    st.markdown("**Pre vs Post (Slide 6)**")
    st.dataframe(
        [
            {"Metric": "Retirement Spending Increase (Annual)", "Value": f"£{diff_annual:,}"},
            {"Metric": "Retirement Spending Increase (Monthly)", "Value": f"£{diff_monthly:,}"},
            {"Metric": "Liquid assets at retirement (Pre → Post)", "Value": f"{_fmt_liquid_millions(liq_pre)} → {_fmt_liquid_millions(liq_post)}" if (liq_pre is not None or liq_post is not None) else "—"},
        ],
        use_container_width=True,
        hide_index=True,
    )
    st.subheader("Extracted charts")

    def show_chart(rel_path: str, title: str, slide_info: str) -> None:
        path = output_dir / rel_path
        if not path.exists():
            return
        try:
            img = Image.open(path)
            st.caption(f"**{slide_info}** — {title}")
            st.image(img, use_container_width=True)
        except Exception:
            st.caption(f"**{slide_info}** — {title}")
            st.warning("Could not load image.")

    st.markdown("---")
    st.markdown("#### Slide 8 — Events & Goals (Timeline)")
    c8a, c8b = st.columns(2)
    with c8a:
        if "pre_timeline" in all_charts:
            show_chart(all_charts["pre_timeline"], "Pre-advice timeline", "Slide 8")
    with c8b:
        if "post_timeline" in all_charts:
            show_chart(all_charts["post_timeline"], "Post-advice timeline", "Slide 8")
    st.markdown("---")
    st.markdown("#### Slide 9 & 12 — Financial Summary (Cashflow)")
    c9a, c9b = st.columns(2)
    with c9a:
        if "pre_cashflow" in all_charts:
            show_chart(all_charts["pre_cashflow"], "Pre-advice cashflow", "Slide 9")
    with c9b:
        if "post_cashflow" in all_charts:
            show_chart(all_charts["post_cashflow"], "Post-advice cashflow", "Slide 12")
    st.markdown("---")
    st.markdown("#### Slide 9 & 13 — Liquid Assets")
    c13a, c13b = st.columns(2)
    with c13a:
        if "pre_liquid_assets" in all_charts:
            show_chart(all_charts["pre_liquid_assets"], "Pre-advice liquid assets", "Slide 9")
    with c13b:
        if "post_liquid_assets" in all_charts:
            show_chart(all_charts["post_liquid_assets"], "Post-advice liquid assets", "Slide 13")
    if "slide24_estate_analysis" in all_charts:
        st.markdown("---")
        st.markdown("#### Slide 24 — Estate Analysis - Detailed")
        show_chart(all_charts["slide24_estate_analysis"], "Planned Legacy (post-advice)", "Slide 24")
    comp_keys = [k for k in sorted(all_charts.keys()) if k.startswith("slide19_comparison_chart_")]
    if comp_keys:
        st.markdown("---")
        st.markdown("#### Slide 19 — FE Analytics comparison charts")
        for k in comp_keys:
            show_chart(all_charts[k], f"Comparison chart {k.replace('slide19_comparison_chart_', '')}", "Slide 19")


st.set_page_config(
    page_title="AP Partners — RoadMap Generator",
    page_icon="📊",
    layout="centered",
    initial_sidebar_state="collapsed",
)

st.title("AP Partners — RoadMap Generator")
st.caption("Upload your Voyant reports and comparison charts to generate RoadMap assets.")

template_type = st.selectbox(
    "Template Type",
    options=["Generic", "Lawyers"],
    help="Stored in output for later use with .pptx templates.",
)

_upload_key = st.session_state.get(UPLOAD_RESET_KEY, 0)
pre_pdf = st.file_uploader(
    "Pre-Advice Voyant PDF",
    type=["pdf"],
    accept_multiple_files=False,
    help="One PDF only. Required.",
    key=f"pre_pdf_{_upload_key}",
)

post_pdf = st.file_uploader(
    "Post-Advice Voyant PDF",
    type=["pdf"],
    accept_multiple_files=False,
    help="One PDF only. Required.",
    key=f"post_pdf_{_upload_key}",
)

comparison_pdfs = st.file_uploader(
    "FE Analytics / Custom comparison charts",
    type=["pdf"],
    accept_multiple_files=True,
    help="Up to 5 PDF files. Multi-page PDFs are supported (one chart per page).",
    key=f"comparison_pdfs_{_upload_key}",
)

generate = st.button("Generate RoadMap Assets", type="primary", use_container_width=True)


def validate():
    errors = []
    if not pre_pdf:
        errors.append("Please upload the Pre-Advice Voyant PDF.")
    if not post_pdf:
        errors.append("Please upload the Post-Advice Voyant PDF.")
    if comparison_pdfs and len(comparison_pdfs) > 5:
        errors.append("Please upload at most 5 comparison chart PDFs.")
    return errors


# Persisted results: show again after download so client can review or download again
if st.session_state.get("roadmap_output_dir") and not generate:
    _out = Path(st.session_state["roadmap_output_dir"])
    if _out.exists():
        _render_results(
            _out,
            st.session_state["roadmap_pre_values"],
            st.session_state["roadmap_post_values"],
            st.session_state["roadmap_all_charts"],
            st.session_state["roadmap_diff_annual"],
            st.session_state["roadmap_diff_monthly"],
            "persisted",
        )
    else:
        _clear_roadmap_state()


if generate:
    errs = validate()
    if errs:
        for e in errs:
            st.error(e)
    else:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        tmpdir = Path(tempfile.mkdtemp())
        output_base = tmpdir
        pre_path = output_base / "_tmp_pre.pdf"
        post_path = output_base / "_tmp_post.pdf"
        pre_path.write_bytes(pre_pdf.getvalue())
        post_path.write_bytes(post_pdf.getvalue())
        comparison_paths = []
        for i, f in enumerate(comparison_pdfs or []):
            p = output_base / f"_tmp_comp_{i}.pdf"
            p.write_bytes(f.getvalue())
            comparison_paths.append(p)

        with st.spinner("Extracting charts and values…"):
            try:
                output_dir, pre_values, post_values, all_charts = run_extraction(
                    client_name="extraction",
                    report_date=timestamp,
                    template_type=template_type,
                    pre_advice_path=pre_path,
                    post_advice_path=post_path,
                    comparison_pdf_paths=comparison_paths,
                    output_base=output_base,
                    print_summary=False,
                )
            except Exception as e:
                st.exception(e)
                st.error("Extraction failed. Check the error above.")
                st.stop()

        pre_spend = pre_values.get("retirement_spending_annual") or 0
        post_spend = post_values.get("retirement_spending_annual") or 0
        diff_annual = post_spend - pre_spend
        diff_monthly = round(diff_annual / 12) if diff_annual else 0

        st.session_state["roadmap_output_dir"] = str(output_dir)
        st.session_state["roadmap_pre_values"] = pre_values
        st.session_state["roadmap_post_values"] = post_values
        st.session_state["roadmap_all_charts"] = all_charts
        st.session_state["roadmap_diff_annual"] = diff_annual
        st.session_state["roadmap_diff_monthly"] = diff_monthly

        template_path = Path(__file__).resolve().parent / "templates" / "Generic_RoadMap.pptx"
        if template_path.exists():
            generated_pptx = output_dir / "RoadMap_GENERATED.pptx"
            try:
                populate_roadmap_pptx(
                    template_path=template_path,
                    output_path=generated_pptx,
                    charts_dir=output_dir / "charts",
                    all_charts=all_charts,
                    retirement_annual_diff=diff_annual,
                    retirement_monthly_diff=diff_monthly,
                    liquid_pre=pre_values.get("liquid_assets_retirement"),
                    liquid_post=post_values.get("liquid_assets_retirement"),
                )
            except Exception as e:
                st.exception(e)
                st.warning("Could not generate RoadMap PPTX. Check template and placeholder names.")
        _render_results(output_dir, pre_values, post_values, all_charts, diff_annual, diff_monthly, "gen")

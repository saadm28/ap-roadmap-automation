"""
AP Partners RoadMap Automation — Streamlit frontend.
Upload Voyant reports and FE Analytics charts to generate RoadMap assets.
"""

import io
import tempfile
import zipfile
from pathlib import Path
from datetime import datetime

import streamlit as st
from PIL import Image

from extractor import run_extraction


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

pre_pdf = st.file_uploader(
    "Pre-Advice Voyant PDF",
    type=["pdf"],
    accept_multiple_files=False,
    help="One PDF only. Required.",
)

post_pdf = st.file_uploader(
    "Post-Advice Voyant PDF",
    type=["pdf"],
    accept_multiple_files=False,
    help="One PDF only. Required.",
)

comparison_pdfs = st.file_uploader(
    "FE Analytics / Custom comparison charts",
    type=["pdf"],
    accept_multiple_files=True,
    help="Up to 5 PDF files. Multi-page PDFs are supported (one chart per page).",
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


if generate:
    errs = validate()
    if errs:
        for e in errs:
            st.error(e)
    else:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        with tempfile.TemporaryDirectory() as tmpdir:
            output_base = Path(tmpdir)
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

            st.success("Extraction complete.")

            # ---------- Extracted values ----------
            pre_spend = pre_values.get("retirement_spending_annual") or 0
            post_spend = post_values.get("retirement_spending_annual") or 0
            diff_annual = post_spend - pre_spend
            diff_monthly = round(diff_annual / 12) if diff_annual else 0

            st.subheader("Extracted values")

            st.markdown("**Pre-Advice**")
            st.dataframe(
                [
                    {"Metric": "Retirement Spending (Annual)", "Value": f"£{pre_spend:,}" if pre_spend else "—"},
                    {"Metric": "Estate Tax (IHT)", "Value": f"£{pre_values.get('estate_tax_iht'):,}" if pre_values.get('estate_tax_iht') else "—"},
                    {"Metric": "Total Estate Transfer", "Value": f"£{pre_values.get('total_estate_transfer'):,}" if pre_values.get('total_estate_transfer') else "—"},
                    {"Metric": "Net Taxable Estate", "Value": f"£{pre_values.get('net_taxable_estate'):,}" if pre_values.get('net_taxable_estate') else "—"},
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
                ],
                use_container_width=True,
                hide_index=True,
            )

            st.markdown("**Pre vs Post (Slide 6)**")
            st.dataframe(
                [
                    {"Metric": "Retirement Spending Increase (Annual)", "Value": f"£{diff_annual:,}"},
                    {"Metric": "Retirement Spending Increase (Monthly)", "Value": f"£{diff_monthly:,}"},
                ],
                use_container_width=True,
                hide_index=True,
            )

            # ---------- Extracted charts gallery (with titles and slide numbers) ----------
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

            # Slide 8 – Timeline (pre and post side by side)
            st.markdown("---")
            st.markdown("#### Slide 8 — Events & Goals (Timeline)")
            col8a, col8b = st.columns(2)
            with col8a:
                if "pre_timeline" in all_charts:
                    show_chart(all_charts["pre_timeline"], "Pre-advice timeline", "Slide 8")
            with col8b:
                if "post_timeline" in all_charts:
                    show_chart(all_charts["post_timeline"], "Post-advice timeline", "Slide 8")

            # Slide 9 / 12 – Cashflow (pre and post side by side)
            st.markdown("---")
            st.markdown("#### Slide 9 & 12 — Financial Summary (Cashflow)")
            col9a, col9b = st.columns(2)
            with col9a:
                if "pre_cashflow" in all_charts:
                    show_chart(all_charts["pre_cashflow"], "Pre-advice cashflow", "Slide 9")
            with col9b:
                if "post_cashflow" in all_charts:
                    show_chart(all_charts["post_cashflow"], "Post-advice cashflow", "Slide 12")

            # Slide 9 / 13 – Liquid assets (pre and post side by side)
            st.markdown("---")
            st.markdown("#### Slide 9 & 13 — Liquid Assets")
            col13a, col13b = st.columns(2)
            with col13a:
                if "pre_liquid_assets" in all_charts:
                    show_chart(all_charts["pre_liquid_assets"], "Pre-advice liquid assets", "Slide 9")
            with col13b:
                if "post_liquid_assets" in all_charts:
                    show_chart(all_charts["post_liquid_assets"], "Post-advice liquid assets", "Slide 13")

            # Slide 24 — Estate Analysis - Detailed (post-advice only)
            if "slide24_estate_analysis" in all_charts:
                st.markdown("---")
                st.markdown("#### Slide 24 — Estate Analysis - Detailed")
                show_chart(all_charts["slide24_estate_analysis"], "Planned Legacy (post-advice)", "Slide 24")

            # Slide 19 – FE Analytics comparison charts at the end
            comparison_keys = [k for k in sorted(all_charts.keys()) if k.startswith("slide19_comparison_chart_")]
            if comparison_keys:
                st.markdown("---")
                st.markdown("#### Slide 19 — FE Analytics comparison charts")
                for k in comparison_keys:
                    num = k.replace("slide19_comparison_chart_", "")
                    show_chart(all_charts[k], f"Comparison chart {num}", "Slide 19")

            # Download ZIP (built from temp dir before it is removed)
            st.markdown("---")
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for item in output_dir.rglob("*"):
                    if item.is_file():
                        arcname = item.relative_to(output_dir.parent)
                        zf.write(item, arcname)
            zip_buffer.seek(0)
            st.download_button(
                "Download output as ZIP",
                data=zip_buffer,
                file_name=f"extraction_{timestamp}.zip",
                mime="application/zip",
            )

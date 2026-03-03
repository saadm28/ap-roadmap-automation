# AP Partners RoadMap Automation

Streamlit app: upload **pre-** and **post-advice Voyant PDFs** (and optional FE Analytics comparison PDFs), extract charts and values, and generate a populated **RoadMap PowerPoint** (Generic or Lawyers template).

## Setup

```bash
pip install -r requirements.txt
```

Put templates in `templates/`: `Generic_RoadMap.pptx` and/or `Lawyers_Roadmap.pptx`.

## Run

```bash
streamlit run app.py
```

Choose template type, upload PDFs, click **Generate RoadMap Assets**. Download the generated PPTX from the results page.

## How it works

- **Charts & values** — Extracted from Voyant PDFs by anchor text (no fixed page numbers). Charts rendered at 300 DPI to `output/.../charts/`.
- **PPT mapping** — Placeholder-based only: text tokens (e.g. `{{LIQUID_ASSETS_PRE}}`, `{{SHORTFALL_YEARS}}`) and shape names (e.g. `[TIMELINE_IMAGE]`, `COMP_CHART_1`) are replaced **everywhere** in the deck. Same logic for Generic and Lawyers; only the template file differs.
- **Output filename** — `AP Partners - Generic RoadMap - YYYY-MM-DD_HHMMSS.pptx` (or Lawyers), so downloads don’t overwrite.

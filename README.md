# AP Partners RoadMap Automation

Generate RoadMap PowerPoints from Voyant reports. Upload pre- and post-advice PDFs (and optional FE Analytics comparison charts); the app extracts charts and values and fills a Generic or Lawyers RoadMap template.

## Quick start

```bash
pip install -r requirements.txt
```

Add your templates to `templates/`: `Generic_RoadMap.pptx` and/or `Lawyers_Roadmap.pptx`.

```bash
streamlit run app.py
```

Select template type, upload PDFs, click **Generate RoadMap Assets**, then download the PPTX from the results page.

## What it does

- **Extraction** — Finds pages by anchor text (e.g. "Liquid Assets", "Retirement Summary"), extracts values via regex and tables, and pulls retirement years from the Retirement Summary page (first primary). Charts are cropped and rendered at 300 DPI.
- **Templates** — Replaces placeholders (e.g. `{{LIQUID_ASSETS_PRE}}`, `{{SHORTFALL_YEARS}}`) and mapped shape names (e.g. timeline, cashflow, comparison charts) in your PPTX. Output: `AP Partners - Generic/Lawyers RoadMap - YYYY-MM-DD_HHMMSS.pptx`.

## Deploy on Streamlit Cloud

1. Push this repo to GitHub.
2. At [share.streamlit.io](https://share.streamlit.io), sign in with GitHub and click **New app**.
3. Select your repo, branch `main`, and set **Main file path** to `app.py`.
4. Deploy. Add the template files to `templates/` in the repo if you want PPTX download to work in the cloud app.

## Requirements

- Python 3.10+
- See `requirements.txt` for dependencies (Streamlit, PyMuPDF, Pillow, python-pptx).

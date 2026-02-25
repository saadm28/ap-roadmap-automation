# AP Partners RoadMap Automation

A Streamlit app that automates the generation of financial RoadMap report assets for AP Partners Wealth Management. It takes Voyant PDF reports (pre-advice and post-advice) plus optional FE Analytics comparison chart PDFs, extracts specific charts as high-resolution images and key financial values, and saves everything into an organised output folder (with an extraction summary file for later use).

**v1** focuses on extraction and chart export only. PowerPoint template population will be added in v2 once the `.pptx` templates are available from the client.

## Setup

```bash
pip install -r requirements.txt
```

## Run

```bash
streamlit run app.py
```

Then open the URL shown in the terminal (typically http://localhost:8501). Enter the client name and report date, choose the template type, upload the Pre-Advice and Post-Advice Voyant PDFs (and up to 4 FE Analytics comparison PDFs if needed), and click **Generate RoadMap Assets**. The app will extract charts and values, show a summary table, and offer a ZIP download of the output folder.

## Output

- **output/{client_name}_{date}/**
  - **charts/** — PNGs named by target slide (e.g. `slide08_events_goals_timeline.png`, `slide22_comparison_chart_1.png`).
  - **extraction_summary.txt** — Key-value summary of all extracted values and chart paths for use in a future PowerPoint step.

Charts are rendered at 300 DPI. Sections in the Voyant PDFs are found by **heading text**, not by fixed page numbers, so layout changes that shift pages are handled correctly.

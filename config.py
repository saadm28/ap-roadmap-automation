"""
AP Partners RoadMap Automation — Configuration.
Anchor-based page finding; fixed crop rectangles for charts; regex for values.
No hardcoded page numbers.
"""

# PDF coordinates: (x0, y0, x1, y1) in points, origin top-left
CHART_CROPS = {
    "timeline": {
        "anchors": ["Events & Goals"],
        "crop": (20, 130, 575, 720),  # Full timeline from visual bar through event cards
    },
    "cashflow": {
        "anchors": ["Financial Summary"],
        "crop": (20, 235, 575, 465),  # Chart only; extra top crop to remove white space (Slide 9 & 12)
    },
    "liquid_assets": {
        "anchors": ["Liquid Assets"],
        "crop": (20, 140, 575, 620),  # Chart + legend; tight bottom to exclude white space below x-axis/icons
    },
}

# Value extraction: anchor(s) to find page, then regex on that page's text. First £ amount captured.
VALUE_EXTRACTORS = {
    "retirement_spending": {
        "anchors": ["Retirement Spending", "You can afford to spend"],
        "regex": r"You can afford to spend\s+.*?£([\d,]+)",
    },
    "estate_tax": {
        "anchors": ["Estate Analysis - Detailed"],
        "regex": r"Total Tax[^\d]*£([\d,]+)",
    },
    "estate_transfer": {
        "anchors": ["Estate Analysis - Detailed"],
        "regex": r"Total Estate Transfer[^\d]*£([\d,]+)",
    },
    "net_taxable": {
        "anchors": ["Estate Analysis - Detailed"],
        "regex": r"Net Taxable[^\d]*£([\d,]+)",
    },
}

RENDER_DPI = 300

# FE Analytics comparison charts (Slide 19): include top legend (Pricing Spread...) + chart + bottom legend (A/B + date)
# Exclude: header (logos, "Performance Line Chart" title) and footer disclaimer
# PDF points (x0, y0, x1, y1); y0 below header, y1 above disclaimer
FE_ANALYTICS_CHART_CROP = (25, 170, 570, 665)
COMPARISON_CHART_FILENAME = "slide19_comparison_chart_{n}.png"

# Slide 24 — Estate Analysis - Detailed (post-advice only): Planned Legacy flowcharts + summary
SLIDE24_ESTATE_ANALYSIS_ANCHOR = "Estate Analysis - Detailed"
SLIDE24_ESTATE_ANALYSIS_CROP = (20, 115, 575, 470)
SLIDE24_ESTATE_ANALYSIS_FILENAME = "slide24_estate_analysis.png"

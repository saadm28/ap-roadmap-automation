# AP Partners — RoadMap Generator

A Streamlit web application that generates client RoadMap PowerPoint presentations from Voyant financial planning reports. Upload pre- and post-advice PDFs, and the app automatically extracts charts, financial values, and trajectory data to populate a branded PPTX template.

---

## Features

- **Automated extraction** — finds pages by anchor text, extracts financial values via regex and table parsing, crops charts at 300 DPI
- **Smart liquid assets** — sums savings, investments, and pensions at the exact retirement year using x-coordinate alignment across table columns
- **Trajectory charts** — slide 14 bar charts rendered as pixel-perfect coloured rectangles (guaranteed proportions, no PowerPoint caching issues)
- **Two template types** — Generic and Lawyers, each independently versioned and upgradeable via the app
- **Template management** — upload a new `.pptx` via the app UI; previous versions are archived, never deleted
- **MySQL storage** — every PDF upload, generated PPTX, and template version is logged to a database with filesystem storage
- **Login protection** — bcrypt-secured login form; no unauthenticated access to the app
- **Graceful fallback** — if the database is unreachable, the app continues using temp storage and bundled templates

---

## Project Structure

```
ap-roadmap-automation/
├── app.py                  # Streamlit frontend
├── extractor.py            # PDF extraction logic (charts + values)
├── roadmap_pptx.py         # PPTX population and chart rendering
├── database.py             # MySQL DB layer (clients, uploads, templates, outputs)
├── config.py               # Chart crop regions, anchor text, regex patterns
├── requirements.txt
├── .env                    # Local config (DB credentials, auth, storage root)
├── templates/
│   ├── Generic_RoadMap.pptx
│   └── Lawyers_Roadmap.pptx
├── storage/                # Created at runtime — gitignored
│   ├── uploads/            # Saved PDFs per client/run
│   ├── outputs/            # Generated PPTXs per run
│   └── templates/          # DB-managed template versions
└── scripts/
    ├── init_db.py                  # One-time DB setup
    ├── generate_password_hash.py   # Generate bcrypt hash for password changes
    ├── test_liquid_assets.py       # Debug liquid assets extraction
    └── test_chart_output.py        # Debug chart rectangle output
```

---

## Local Setup

### 1. Prerequisites

- Python 3.10+
- MySQL 8.0+ running locally (`brew install mysql` on Mac)

### 2. Install dependencies

```bash
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

### 3. Configure environment

Copy and edit the `.env` file:

```bash
cp .env .env.local  # optional — .env is gitignored
```

```env
# Database
DB_HOST=localhost
DB_PORT=3306
DB_NAME=ap_roadmap
DB_USER=root
DB_PASSWORD=

# Storage root (leave blank to use storage/ next to app.py)
STORAGE_ROOT=

# Login credentials
AUTH_USERNAME=appartners
AUTH_NAME=AP Partners
AUTH_PASSWORD_HASH=<bcrypt hash — see below>
AUTH_COOKIE_KEY=ap_roadmap_secure_key_2026
```

### 4. Initialise the database

Run once to create the schema and seed the bundled templates:

```bash
python scripts/init_db.py --seed-templates
```

### 5. Run the app

```bash
streamlit run app.py
```

---

## Login Credentials

Default credentials to share with the client:

| Field    | Value             |
|----------|-------------------|
| Username | `appartners`      |
| Password | `APPartners2026!` |

**To change the password:**

```bash
python scripts/generate_password_hash.py
```

Paste the generated hash into `.env` as `AUTH_PASSWORD_HASH=...` and restart the app.

---

## Using the App

### Tab 1 — RoadMap Generator

1. Select **Template Type** (Generic or Lawyers)
2. Upload **Pre-Advice Voyant PDF** and **Post-Advice Voyant PDF**
3. Optionally upload up to 5 **FE Analytics comparison chart PDFs**
4. Click **Generate RoadMap Assets**
5. Download the generated `.pptx` from the results page

The client name is extracted automatically from the PDF.

### Tab 2 — Template Settings

- View the currently active template for each type (Generic / Lawyers) with version and upload date
- Upload a new `.pptx` to replace the active version immediately
- Full version history table — old versions are archived, never deleted

---

## PPTX Template Placeholders

The app replaces the following text tokens in the template (anywhere in the deck):

| Token | Description |
|---|---|
| `{{CLIENT_NAME}}` | Client name extracted from PDF |
| `{{REPORT_MONTH}}` / `{{REPORT_YEAR}}` | Report date |
| `{{LIQUID_ASSETS_PRE}}` / `{{LIQUID_ASSETS_POST}}` | Liquid assets at retirement (c.£X.XXm) |
| `{{RETIREMENT_SPENDING_PRE}}` / `{{RETIREMENT_SPENDING_POST}}` | Annual retirement spending |
| `{{DIFF_ANNUAL}}` / `{{DIFF_MONTHLY}}` | Spending increase pre→post |
| `{{SHORTFALL_YEARS}}` / `{{TOTAL_RETIREMENT_YEARS}}` | Pre-advice trajectory years |
| `{{LUMP_SUM_REQUIRED}}` / `{{RETIREMENT_YEAR}}` | Lump sum slide values |
| `{{ANNUAL_SAVINGS_REQUIRED}}` | Annual savings required |
| `{{POST_NOT_FUNDED_YEARS}}` / `{{POST_FUNDED_YEARS}}` | Post-advice trajectory years |
| `{{POST_RETIREMENT_SPENDING}}` | Post-advice affordable spending |
| `{{ONTRACK_FLAG}}` | `"not "` or `""` (for on-track message) |

**Shape names** (Selection Pane) for image/chart placeholders:

| Shape name | Content |
|---|---|
| `PRE_ADVICE_CHART` | Pre-advice trajectory bar (slide 14) |
| `POST_ADVICE_CHART` | Post-advice trajectory bar (slide 14) |
| `PRE_CASHFLOW_IMAGE` | Pre-advice cashflow chart |
| `PRE_LIQUID_IMAGE` | Pre-advice liquid assets chart |
| `PRE_CASHFLOW_COMPARISON` | Pre cashflow (comparison slide) |
| `POST_CASHFLOW_COMPARISON` | Post cashflow (comparison slide) |
| `PRE_LIQUID_COMPARISON` | Pre liquid assets (comparison slide) |
| `POST_LIQUID_COMPARISON` | Post liquid assets (comparison slide) |
| `COMP_CHART_1` – `COMP_CHART_4` | FE Analytics comparison charts |
| `POST_ESTATE_IMAGE` | Estate analysis chart (slide 24) |

---

## VPS Deployment

### 1. Set environment variables

Instead of `.env`, set these in your systemd service or docker-compose:

```env
DB_HOST=<vps-db-host>
DB_PORT=3306
DB_NAME=ap_roadmap
DB_USER=<db-user>
DB_PASSWORD=<db-password>
STORAGE_ROOT=/var/app/storage
AUTH_USERNAME=appartners
AUTH_NAME=AP Partners
AUTH_PASSWORD_HASH=<bcrypt hash>
AUTH_COOKIE_KEY=<long random string>
```

### 2. Initialise the database

```bash
python scripts/init_db.py --seed-templates
```

### 3. Run with systemd or Docker

```bash
streamlit run app.py --server.port 8501 --server.address 0.0.0.0
```

Recommended: put Nginx in front as a reverse proxy with HTTPS.

### Backups

Back up two things:
- MySQL dump: `mysqldump ap_roadmap > backup.sql`
- Storage folder: `$STORAGE_ROOT` (uploads, outputs, templates)

---

## Requirements

- `streamlit>=1.28.0`
- `PyMuPDF>=1.23.0`
- `Pillow>=10.0.0`
- `python-pptx>=0.6.21`
- `openpyxl>=3.1.0`
- `mysql-connector-python>=8.0.0`
- `python-dotenv>=1.0.0`
- `streamlit-authenticator>=0.4.0`
- `bcrypt>=4.0.0`

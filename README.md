# AP Partners — RoadMap Generator

A Streamlit web application that generates client RoadMap PowerPoint presentations from Voyant financial planning reports. Upload pre- and post-advice PDFs, and the app automatically extracts charts, financial values, and trajectory data to populate a branded PPTX template.

---

## Project Structure

```
ap-roadmap-automation/
├── app.py                  # Streamlit frontend (tabs: Generator + Template Settings)
├── extractor.py            # PDF extraction logic (charts + values)
├── roadmap_pptx.py         # PPTX population and chart rendering
├── database.py             # MySQL DB layer — see "Why database.py?" below
├── config.py               # Chart crop regions, anchor text, regex patterns
├── init_db.py              # One-time DB setup script — run before first launch
├── requirements.txt
├── .env                    # Local config — DB credentials, auth, storage root (gitignored)
└── templates/
    ├── Generic_RoadMap.pptx    # Bundled default — used as seed on first deploy
    └── Lawyers_Roadmap.pptx
```

> `storage/`, `uploads/`, `outputs/`, `.env`, and `scripts/` are all gitignored.

---

## Why `database.py`?

`database.py` is the MySQL layer that handles all persistent storage for the app. It manages:

- **`clients`** — one row per client extracted from PDFs (e.g. "John Smith")
- **`templates`** — versioned PPTX templates with active/archived status
- **`uploads`** — every pre/post PDF pair processed, linked to a client
- **`outputs`** — every generated PPTX, linked to its upload run

All actual files (PDFs, PPTXs, templates) live on disk in `STORAGE_ROOT`. The DB stores only metadata and file paths. If the DB is unreachable, the app falls back gracefully to temp storage and bundled templates — nothing breaks.

---

## What `init_db.py` Does

Creates the MySQL database, all tables, and seeds the bundled templates into `storage/templates/` so the app has working templates from day one.

**Run it once — on first setup locally or on VPS:**
```bash
python init_db.py --seed-templates
```

- `--seed-templates` copies `templates/Generic_RoadMap.pptx` and `templates/Lawyers_Roadmap.pptx` into `storage/templates/` and registers them as active in the DB
- Safe to re-run if `storage/templates/` gets deleted — it just adds new active versions
- Never run in production after the client has uploaded their own templates (it would create new versions on top)

---

## Local Setup

### 1. Prerequisites

- Python 3.10+
- MySQL 8.0+ running locally (`brew install mysql && brew services start mysql` on Mac)

### 2. Create virtual environment and install dependencies

```bash
python -m venv venv
source venv/bin/activate    # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

### 3. Configure `.env`

Create a `.env` file in the project root (it is gitignored):

```env
# Database
DB_HOST=localhost
DB_PORT=3306
DB_NAME=ap_roadmap
DB_USER=root
DB_PASSWORD=

# Storage root — all PDFs, templates, outputs stored here
# Leave blank to use storage/ next to app.py
STORAGE_ROOT=

# Login credentials
AUTH_USERNAME=<username>
AUTH_NAME=<display name>
AUTH_PASSWORD_HASH=<bcrypt hash — see Login Credentials section>
AUTH_COOKIE_KEY=<random string — generate with: python -c "import secrets; print(secrets.token_hex(32))">
```

### 4. Initialise the database

```bash
python init_db.py --seed-templates
```

### 5. Run the app

```bash
streamlit run app.py
```

> **No directories need to be created manually.** The app automatically creates `storage/`, `storage/uploads/`, and `storage/templates/` on first use.

---

## Login Credentials

Credentials are set in `.env` via `AUTH_USERNAME` and `AUTH_PASSWORD_HASH`. The actual username and password are shared separately — do not commit them to the repo.

To generate a bcrypt hash for a new password:
```bash
python -c "import bcrypt; print(bcrypt.hashpw(b'YourPassword', bcrypt.gensalt()).decode())"
```
Paste the output into `.env` as `AUTH_PASSWORD_HASH=...` and restart the app.

---

## Using the App

### Tab 1 — RoadMap Generator

1. Select **Template Type** (Generic or Lawyers)
2. Upload **Pre-Advice Voyant PDF** and **Post-Advice Voyant PDF**
3. Optionally upload up to 5 **FE Analytics comparison chart PDFs**
4. Click **Generate RoadMap Assets**
5. Download the generated `.pptx`

Client name is extracted automatically from the PDF and used to organise stored files.

### Tab 2 — Template Settings

- View currently active template for each type with version and upload date
- Upload a new `.pptx` to replace the active version immediately
- Full version history — old versions are archived, never deleted

---

## Storage Structure

All files are stored under `STORAGE_ROOT` (default: `storage/` next to `app.py`):

```
storage/
├── templates/
│   ├── Generic_v1_20260324_212708.pptx   ← archived
│   └── Generic_v2_20260325_100918.pptx   ← active
│
└── uploads/
    └── {client_id}_{ClientName}/
        └── {YYYYMMDD_HHMMSS}/
            ├── client-pre_pre.pdf
            ├── client-post_post.pdf
            └── AP Partners - Generic RoadMap - YYYY-MM-DD_HHMMSS.pptx
```

The folder uses `{id}_{name}` to guarantee uniqueness even if two clients share the same name.

---

## PPTX Template Placeholders

Text tokens replaced anywhere in the deck:

| Token | Description |
|---|---|
| `{{CLIENT_NAME}}` | Client name from PDF |
| `{{REPORT_MONTH}}` / `{{REPORT_YEAR}}` | Report date |
| `{{LIQUID_ASSETS_PRE}}` / `{{LIQUID_ASSETS_POST}}` | Liquid assets at retirement |
| `{{RETIREMENT_SPENDING_PRE}}` / `{{RETIREMENT_SPENDING_POST}}` | Annual retirement spending |
| `{{DIFF_ANNUAL}}` / `{{DIFF_MONTHLY}}` | Spending increase pre→post |
| `{{SHORTFALL_YEARS}}` / `{{TOTAL_RETIREMENT_YEARS}}` | Pre-advice trajectory years |
| `{{LUMP_SUM_REQUIRED}}` / `{{RETIREMENT_YEAR}}` | Lump sum slide values |
| `{{ANNUAL_SAVINGS_REQUIRED}}` | Annual savings required |
| `{{POST_NOT_FUNDED_YEARS}}` / `{{POST_FUNDED_YEARS}}` | Post-advice trajectory years |
| `{{POST_RETIREMENT_SPENDING}}` | Post-advice affordable spending |
| `{{ONTRACK_FLAG}}` | `"not "` or `""` |

Shape names (Selection Pane) for image placeholders:

| Shape name | Content |
|---|---|
| `PRE_ADVICE_CHART` | Pre-advice trajectory bar |
| `POST_ADVICE_CHART` | Post-advice trajectory bar |
| `PRE_CASHFLOW_IMAGE` | Pre-advice cashflow chart |
| `PRE_LIQUID_IMAGE` | Pre-advice liquid assets chart |
| `PRE_CASHFLOW_COMPARISON` | Pre cashflow (comparison slide) |
| `POST_CASHFLOW_COMPARISON` | Post cashflow (comparison slide) |
| `PRE_LIQUID_COMPARISON` | Pre liquid assets (comparison slide) |
| `POST_LIQUID_COMPARISON` | Post liquid assets (comparison slide) |
| `COMP_CHART_1` – `COMP_CHART_4` | FE Analytics comparison charts |
| `POST_ESTATE_IMAGE` | Estate analysis chart |

---

## VPS Deployment

### 1. Server requirements

- Ubuntu 22.04+ (or similar)
- Python 3.10+
- MySQL 8.0+
- Nginx (reverse proxy, recommended)

### 2. Clone and install

```bash
git clone <repo-url> /var/app/ap-roadmap
cd /var/app/ap-roadmap
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

### 3. Create storage directory outside the repo (optional but recommended)

> **You do not need to create any directories manually.** The app auto-creates `storage/`, `storage/uploads/`, and `storage/templates/` on first run.
>
> However, on a VPS it is recommended to place storage **outside the repo folder** so that client files are never affected by a `git pull` or redeploy. If you skip this step, the app will just create a `storage/` folder inside the repo directory instead.

```bash
mkdir -p /var/app/storage
```

### 4. Set environment variables

Create `/var/app/ap-roadmap/.env` (or set in systemd service):

```env
DB_HOST=localhost
DB_PORT=3306
DB_NAME=ap_roadmap
DB_USER=apuser
DB_PASSWORD=<strong-password>
STORAGE_ROOT=/var/app/storage
AUTH_USERNAME=<username>
AUTH_NAME=<display name>
AUTH_PASSWORD_HASH=<bcrypt hash — see Login Credentials section>
AUTH_COOKIE_KEY=<random string — generate with: python -c "import secrets; print(secrets.token_hex(32))">
```

### 5. Initialise the database (first deploy only)

```bash
python init_db.py --seed-templates
```

This creates the schema and copies the bundled templates into `/var/app/storage/templates/`.

### 6. Run with systemd

Create `/etc/systemd/system/ap-roadmap.service`:

```ini
[Unit]
Description=AP Partners RoadMap Generator
After=network.target mysql.service

[Service]
User=ubuntu
WorkingDirectory=/var/app/ap-roadmap
EnvironmentFile=/var/app/ap-roadmap/.env
ExecStart=/var/app/ap-roadmap/venv/bin/streamlit run app.py --server.port 8501 --server.address 127.0.0.1
Restart=always

[Install]
WantedBy=multi-user.target
```

```bash
systemctl daemon-reload
systemctl enable ap-roadmap
systemctl start ap-roadmap
```

### 7. Nginx reverse proxy

```nginx
server {
    listen 80;
    server_name yourdomain.com;

    location / {
        proxy_pass http://127.0.0.1:8501;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection "upgrade";
        proxy_set_header Host $host;
    }
}
```

### 8. Redeployment

```bash
cd /var/app/ap-roadmap
git pull
systemctl restart ap-roadmap
```

`/var/app/storage` is outside the repo — client files, uploads, and templates are never touched by a redeploy.

### Backups

```bash
# Database
mysqldump ap_roadmap > backup_$(date +%Y%m%d).sql

# Files
tar -czf storage_backup_$(date +%Y%m%d).tar.gz /var/app/storage
```

---

## Requirements

- `streamlit>=1.28.0`
- `PyMuPDF>=1.23.0`
- `Pillow>=10.0.0`
- `python-pptx>=0.6.21`
- `mysql-connector-python>=8.0.0`
- `python-dotenv>=1.0.0`
- `streamlit-authenticator>=0.4.0`
- `bcrypt>=4.0.0`

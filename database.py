"""
AP Partners RoadMap — MySQL database layer.

Tables:
  clients   — one row per client (name)
  uploads   — every pre/post PDF pair processed
  templates — versioned PPTX templates (latest active = is_active=1)
  outputs   — every generated PPTX

Connection settings are read from environment variables so the same code
works locally and on VPS without any code changes:

  DB_HOST     (default: localhost)
  DB_PORT     (default: 3306)
  DB_NAME     (default: ap_roadmap)
  DB_USER     (default: root)
  DB_PASSWORD (default: "")

Set these in a .env file locally; on VPS set them in the systemd service or
docker-compose environment.
"""

import os
import logging
from contextlib import contextmanager
from pathlib import Path
from datetime import datetime
from typing import Optional

import mysql.connector
from mysql.connector import Error as MySQLError

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Connection config — read from env, sensible local defaults
# ---------------------------------------------------------------------------

def _cfg() -> dict:
    return {
        "host": os.getenv("DB_HOST", "localhost"),
        "port": int(os.getenv("DB_PORT", "3306")),
        "database": os.getenv("DB_NAME", "ap_roadmap"),
        "user": os.getenv("DB_USER", "root"),
        "password": os.getenv("DB_PASSWORD", ""),
        "autocommit": False,
        "connection_timeout": 10,
    }


@contextmanager
def _conn():
    """Yield a connection, commit on exit, rollback on error."""
    cn = mysql.connector.connect(**_cfg())
    try:
        yield cn
        cn.commit()
    except Exception:
        cn.rollback()
        raise
    finally:
        cn.close()


@contextmanager
def _cursor():
    """Yield a cursor inside a managed connection."""
    with _conn() as cn:
        cur = cn.cursor(dictionary=True)
        try:
            yield cur
        finally:
            cur.close()


# ---------------------------------------------------------------------------
# Schema creation
# ---------------------------------------------------------------------------

_DDL = [
    """
    CREATE TABLE IF NOT EXISTS clients (
        id         INT AUTO_INCREMENT PRIMARY KEY,
        name       VARCHAR(255) NOT NULL,
        created_at DATETIME     NOT NULL DEFAULT CURRENT_TIMESTAMP,
        UNIQUE KEY uq_client_name (name)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
    """,
    """
    CREATE TABLE IF NOT EXISTS templates (
        id          INT AUTO_INCREMENT PRIMARY KEY,
        name        VARCHAR(64)  NOT NULL COMMENT 'Generic or Lawyers',
        version     INT          NOT NULL DEFAULT 1,
        file_path   VARCHAR(512) NOT NULL,
        is_active   TINYINT(1)   NOT NULL DEFAULT 1,
        uploaded_at DATETIME     NOT NULL DEFAULT CURRENT_TIMESTAMP,
        KEY idx_template_name_active (name, is_active)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
    """,
    """
    CREATE TABLE IF NOT EXISTS uploads (
        id            INT AUTO_INCREMENT PRIMARY KEY,
        client_id     INT          NOT NULL,
        pre_pdf_path  VARCHAR(512) NOT NULL,
        post_pdf_path VARCHAR(512) NOT NULL,
        template_type VARCHAR(64)  NOT NULL DEFAULT 'Generic',
        uploaded_at   DATETIME     NOT NULL DEFAULT CURRENT_TIMESTAMP,
        status        VARCHAR(32)  NOT NULL DEFAULT 'pending'
            COMMENT 'pending | processing | done | error',
        CONSTRAINT fk_upload_client FOREIGN KEY (client_id) REFERENCES clients(id)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
    """,
    """
    CREATE TABLE IF NOT EXISTS outputs (
        id          INT AUTO_INCREMENT PRIMARY KEY,
        upload_id   INT          NOT NULL,
        pptx_path   VARCHAR(512) NOT NULL,
        filename    VARCHAR(255) NOT NULL,
        created_at  DATETIME     NOT NULL DEFAULT CURRENT_TIMESTAMP,
        CONSTRAINT fk_output_upload FOREIGN KEY (upload_id) REFERENCES uploads(id)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
    """,
]


def init_db() -> None:
    """Create the database and all tables if they don't exist."""
    # Connect without specifying db to create it if missing
    cfg = _cfg()
    db_name = cfg.pop("database")
    cfg["autocommit"] = True
    cn = mysql.connector.connect(**cfg)
    try:
        cur = cn.cursor()
        cur.execute(f"CREATE DATABASE IF NOT EXISTS `{db_name}` CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;")
        cur.execute(f"USE `{db_name}`;")
        for ddl in _DDL:
            cur.execute(ddl)
        cur.close()
        logger.info("Database '%s' initialised.", db_name)
    finally:
        cn.close()


# ---------------------------------------------------------------------------
# Storage root — all files live under here
# ---------------------------------------------------------------------------

def storage_root() -> Path:
    """
    Base directory for all stored files.
    Uses STORAGE_ROOT env var if set; otherwise sibling 'storage/' folder.
    """
    root = Path(os.getenv("STORAGE_ROOT", Path(__file__).resolve().parent / "storage"))
    root.mkdir(parents=True, exist_ok=True)
    return root


# ---------------------------------------------------------------------------
# Clients
# ---------------------------------------------------------------------------

def get_or_create_client(name: str) -> int:
    """Return client id, creating the row if it doesn't exist."""
    with _cursor() as cur:
        cur.execute("SELECT id FROM clients WHERE name = %s", (name,))
        row = cur.fetchone()
        if row:
            return row["id"]
        cur.execute("INSERT INTO clients (name) VALUES (%s)", (name,))
        return cur.lastrowid


def list_clients() -> list[dict]:
    with _cursor() as cur:
        cur.execute("SELECT id, name, created_at FROM clients ORDER BY name")
        return cur.fetchall()


# ---------------------------------------------------------------------------
# Templates
# ---------------------------------------------------------------------------

def save_template(name: str, file_bytes: bytes) -> int:
    """
    Save a new template file to disk and mark it active.
    Deactivates all previous versions of the same template name.
    Returns the new template id.
    """
    tmpl_dir = storage_root() / "templates"
    tmpl_dir.mkdir(parents=True, exist_ok=True)

    # Get next version number
    with _cursor() as cur:
        cur.execute("SELECT MAX(version) AS v FROM templates WHERE name = %s", (name,))
        row = cur.fetchone()
        version = (row["v"] or 0) + 1

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{name}_v{version}_{ts}.pptx"
    file_path = tmpl_dir / filename
    file_path.write_bytes(file_bytes)

    with _cursor() as cur:
        # Deactivate old versions
        cur.execute("UPDATE templates SET is_active = 0 WHERE name = %s", (name,))
        # Insert new active version
        cur.execute(
            "INSERT INTO templates (name, version, file_path, is_active) VALUES (%s, %s, %s, 1)",
            (name, version, str(file_path)),
        )
        return cur.lastrowid


def get_active_template_path(name: str) -> Optional[Path]:
    """
    Return the filesystem path of the active template for this name.
    Falls back to the bundled templates/ folder if no DB entry exists.
    """
    with _cursor() as cur:
        cur.execute(
            "SELECT file_path FROM templates WHERE name = %s AND is_active = 1 ORDER BY id DESC LIMIT 1",
            (name,),
        )
        row = cur.fetchone()
    if row:
        p = Path(row["file_path"])
        if p.exists():
            return p
        logger.warning("Template file missing on disk: %s — falling back to bundled.", p)

    # Fall back to bundled template
    bundled_name = "Lawyers_Roadmap.pptx" if name == "Lawyers" else "Generic_RoadMap.pptx"
    bundled = Path(__file__).resolve().parent / "templates" / bundled_name
    return bundled if bundled.exists() else None


def list_templates() -> list[dict]:
    with _cursor() as cur:
        cur.execute(
            "SELECT id, name, version, file_path, is_active, uploaded_at FROM templates ORDER BY name, version DESC"
        )
        return cur.fetchall()


# ---------------------------------------------------------------------------
# Uploads (PDF pairs)
# ---------------------------------------------------------------------------

def save_upload(
    client_id: int,
    pre_bytes: bytes,
    post_bytes: bytes,
    pre_filename: str,
    post_filename: str,
    template_type: str,
) -> tuple[int, Path, Path]:
    """
    Save pre/post PDFs to disk, insert upload row.
    Returns (upload_id, pre_path, post_path).
    """
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    upload_dir = storage_root() / "uploads" / str(client_id) / ts
    upload_dir.mkdir(parents=True, exist_ok=True)

    stem_pre = Path(pre_filename).stem[:80]
    stem_post = Path(post_filename).stem[:80]
    pre_path = upload_dir / f"{stem_pre}_pre.pdf"
    post_path = upload_dir / f"{stem_post}_post.pdf"
    pre_path.write_bytes(pre_bytes)
    post_path.write_bytes(post_bytes)

    with _cursor() as cur:
        cur.execute(
            """
            INSERT INTO uploads (client_id, pre_pdf_path, post_pdf_path, template_type, status)
            VALUES (%s, %s, %s, %s, 'processing')
            """,
            (client_id, str(pre_path), str(post_path), template_type),
        )
        upload_id = cur.lastrowid

    return upload_id, pre_path, post_path


def update_upload_status(upload_id: int, status: str) -> None:
    with _cursor() as cur:
        cur.execute("UPDATE uploads SET status = %s WHERE id = %s", (status, upload_id))


def list_uploads(client_id: Optional[int] = None) -> list[dict]:
    with _cursor() as cur:
        if client_id:
            cur.execute(
                """
                SELECT u.id, c.name AS client, u.template_type, u.status, u.uploaded_at
                FROM uploads u JOIN clients c ON c.id = u.client_id
                WHERE u.client_id = %s ORDER BY u.uploaded_at DESC
                """,
                (client_id,),
            )
        else:
            cur.execute(
                """
                SELECT u.id, c.name AS client, u.template_type, u.status, u.uploaded_at
                FROM uploads u JOIN clients c ON c.id = u.client_id
                ORDER BY u.uploaded_at DESC LIMIT 100
                """
            )
        return cur.fetchall()


# ---------------------------------------------------------------------------
# Outputs (generated PPTXs)
# ---------------------------------------------------------------------------

def save_output(upload_id: int, pptx_bytes: bytes, filename: str) -> tuple[int, Path]:
    """
    Save generated PPTX to disk, insert output row.
    Returns (output_id, pptx_path).
    """
    out_dir = storage_root() / "outputs" / str(upload_id)
    out_dir.mkdir(parents=True, exist_ok=True)
    pptx_path = out_dir / filename
    pptx_path.write_bytes(pptx_bytes)

    with _cursor() as cur:
        cur.execute(
            "INSERT INTO outputs (upload_id, pptx_path, filename) VALUES (%s, %s, %s)",
            (upload_id, str(pptx_path), filename),
        )
        return cur.lastrowid, pptx_path


def list_outputs(upload_id: int) -> list[dict]:
    with _cursor() as cur:
        cur.execute(
            "SELECT id, pptx_path, filename, created_at FROM outputs WHERE upload_id = %s ORDER BY created_at DESC",
            (upload_id,),
        )
        return cur.fetchall()


# ---------------------------------------------------------------------------
# Health check
# ---------------------------------------------------------------------------

def db_ping() -> bool:
    """Return True if DB is reachable."""
    try:
        with _cursor() as cur:
            cur.execute("SELECT 1")
            cur.fetchone()
        return True
    except Exception as e:
        logger.error("DB ping failed: %s", e)
        return False

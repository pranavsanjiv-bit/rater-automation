
import os
import sqlite3

# ── DB helpers ────────────────────────────────────────────────────────────────
# BASE_DIR is rater_automation/server/
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.environ.get("RATER_DB", os.path.join(BASE_DIR, "rater.db"))


def get_db() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with get_db() as conn:
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS rate_templates (
                id              TEXT PRIMARY KEY,
                name            TEXT,
                definition_json TEXT NOT NULL,
                content_hash    TEXT,
                created_at      TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS rate_values (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                template_id TEXT NOT NULL REFERENCES rate_templates(id),
                lookup_key  TEXT NOT NULL,
                dims_json   TEXT,
                value       REAL NOT NULL,
                updated_at  TEXT NOT NULL,
                UNIQUE(template_id, lookup_key)
            );
        """)


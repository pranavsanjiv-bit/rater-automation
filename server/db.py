import os
import sqlite3
import json
import uuid
from datetime import datetime, timezone

# ── DB helpers ────────────────────────────────────────────────────────────────
# BASE_DIR is rater_automation/server/
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.environ.get("RATER_DB", os.path.join(BASE_DIR, "rater.db"))


def get_db() -> sqlite3.Connection:
    """
    Create and return a SQLite database connection.

    The connection uses sqlite3.Row as the row factory so query results
    can be accessed by column names as well as indexes.
    """
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """
    Initialize the application database tables if they do not already exist.
    """
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

            CREATE TABLE IF NOT EXISTS dimension_library (
                id          TEXT PRIMARY KEY,
                name        TEXT NOT NULL UNIQUE COLLATE NOCASE,
                type        TEXT NOT NULL,
                config_json TEXT NOT NULL,
                is_system   INTEGER NOT NULL DEFAULT 0,
                created_at  TEXT NOT NULL,
                updated_at  TEXT NOT NULL
            );
        """)
        _seed_system_dimensions(conn)


# ── Dimension Library: system (built-in) dimension seeding ───────────────────
# Mirrors the predefined dimensions currently hardcoded as _LOS_DIMS/_FE_DIMS/
# _PREDEFINED_NAME_TO_SLUG in frontend/app.py (Age, Loan Amount, Loan Type,
# Loan Tenure, Gender, Borrowers count, Sum Insured, Tenure, Cover Type) —
# the same 9 names currently in this codebase, not an assumed/older list.
# Seeded once with is_system=1 so "Choose Existing Dimension" has the same
# options previously offered by that hardcoded list. Since each library entry
# now needs one fixed type + config (unlike the old free-choice-per-template
# behaviour), the values below are reasonable defaults inferred from this
# project's own sample payloads/templates; system dimensions are immutable
# in name/type, so adjust this seed list directly if the business wants
# different canonical defaults before first deploy.
_SYSTEM_DIMENSIONS = [
    ("Age",             "Range", {"min": 18, "max": 65}),
    ("Loan Amount",     "Range", {"min": 0, "max": 10000000}),
    ("Loan Type",       "Enum",  {"values": ["HL", "LAP"]}),
    ("Loan Tenure",     "Range", {"min": 1, "max": 30}),
    ("Gender",          "Enum",  {"values": ["Male", "Female", "Transgender"]}),
    ("Borrowers count", "Range", {"min": 1, "max": 4}),
    ("Sum Insured",     "Range", {"min": 0, "max": 10000000}),
    ("Tenure",          "Range", {"min": 1, "max": 30}),
    ("Cover Type",      "Enum",  {"values": ["Reducing", "Flat"]}),
]


def _seed_system_dimensions(conn: sqlite3.Connection) -> None:
    """Insert the built-in dimensions once. Safe to call on every startup:
    INSERT OR IGNORE keyed on UNIQUE(name COLLATE NOCASE) means it never
    overwrites an admin's edits or duplicates rows on restart."""
    now = datetime.now(timezone.utc).isoformat()
    for name, dim_type, config in _SYSTEM_DIMENSIONS:
        conn.execute(
            "INSERT OR IGNORE INTO dimension_library "
            "(id, name, type, config_json, is_system, created_at, updated_at) "
            "VALUES (?, ?, ?, ?, 1, ?, ?)",
            (str(uuid.uuid4()), name, dim_type, json.dumps(config), now, now),
        )
# Rater Automation

An internal tool for building, filling out, and calculating insurance/loan
**rate tables** — replacing manual Excel-based premium/rate lookups with a
structured, formula-driven system that can also push results straight into
a partner API to create a lead, save a lead, or save a loan.

## What it does

1. **Build a rate template** — define which dimensions (Age, Loan Amount,
   Loan Type, Gender, Sum Insured, Tenure, Cover Type, etc.) make up a
   rate table, plus any formula that combines them.
2. **Generate a fillable Excel workbook** from that template — a
   spreadsheet with one row/column per dimension value, ready for the
   rating/actuarial team to fill in actual rate numbers.
3. **Upload the filled workbook back** (`/flatten`) — every cell gets
   parsed and stored in the database as a `(template, dimension
   combination) → value` lookup row. Blank cells are tracked, not
   silently dropped, and a completion percentage is reported.
4. **Calculate a premium** for a specific customer (single lookup via
   `/calculate`, or an entire Excel of customers at once via
   `/calculate-bulk`) by resolving their dimension values against the
   stored rate table and applying the template's formula.
5. **Optionally submit the result to a partner platform** — the
   `injestion` module can take a calculated premium and inject it into a
   `save_lead`, `create_lead`, or `save_loan` payload and send it to an
   external insurance/loan partner API (`api_service.py`).

## Architecture

Two separate processes, matching the `Procfile`:

| Process | Command | What it is |
|---|---|---|
| `web` | `streamlit run frontend/app.py` | The UI — template builder, formula editor, Excel generation, premium calculator |
| `server` | `python server/run.py` | Flask API (default port `5050`) — templates, dimension library, ingestion, calculation, and partner API integration |

The Streamlit frontend talks to the Flask backend over HTTP.

## Project structure

```
rater_automation/
├── Procfile                    # process definitions (web + server)
├── requirements.txt
│
├── frontend/                   # Streamlit UI
│   ├── app.py                  # main app — template builder, calculator, formula editor
│   ├── app_legacy.py           # previous version, kept for reference
│   ├── dim_parser.py           # dimension name/value slugging, unit parsing (20L, 1Cr, etc.)
│   ├── formula_eval.py         # safe AST-based formula evaluator (+, -, *, /, ^)
│   ├── excel_generator.py      # builds the fillable .xlsx from a template definition
│   └── .streamlit/config.toml
│
└── server/                     # Flask API
    ├── run.py                  # entry point
    ├── __init__.py             # create_app(), blueprint registration
    ├── db.py                   # SQLite schema + system dimension seeding
    ├── api_service.py          # partner API client (auth, leads, loans, quotes)
    ├── templates.py            # loads payload templates from payloads/*.json
    ├── helpers.py               # shared calculation helpers
    ├── payloads/                # JSON payload templates (create_lead, save_lead, save_loan)
    ├── injection_path.json      # maps calculated values to partner payload fields
    └── routes/
        ├── templates.py         # CRUD for rate templates + /calculate, /calculate-bulk
        ├── dimension_library.py # CRUD for rating dimensions (system + custom)
        ├── data_process.py      # /flatten — parses an uploaded workbook into rate_values
        ├── injestion.py         # builds & sends save_lead/create_lead/save_loan payloads
        └── .env                 # partner API credentials (not committed)
```

## Data model (SQLite, `server/rater.db`)

- **`dimension_library`** — every rating dimension (Age, Loan Amount, ...),
  each either a `Range` (min/max) or `Enum` (fixed list of values). Nine
  system dimensions are seeded automatically on first run; custom ones can
  be added through the UI.
- **`rate_templates`** — a template's dimension/axis definition and
  formula, stored as JSON.
- **`rate_values`** — the actual filled-in numbers, one row per unique
  combination of dimension values for a given template (populated by
  `/flatten`, read by `/calculate`).

## Setup

```bash
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

**Note:** `frontend/app.py` imports `code_editor` (for the in-app formula
editor), which isn't currently listed in `requirements.txt` — install it
separately if you hit an import error:
```bash
pip install code-editor
```

Create `server/routes/.env` with the partner API credentials
(`PARTNER_CODE`, `PARTNER_KEY`, `PARTNER_USERID`, and the various
`*_URL` endpoints referenced in `api_service.py`). This file is never
committed — ask whoever set up the partner integration for the values.

## Running it

With [honcho](https://github.com/nickstenning/honcho) (already in the
venv, matches the `Procfile` format):
```bash
honcho start
```

Or run each process manually in two terminals:
```bash
python server/run.py            # backend on :5050
streamlit run frontend/app.py   # frontend, opens in browser
```

## Known gaps / things to double check before relying on this

- `code_editor` isn't pinned in `requirements.txt` (see above).
- `app_legacy.py` exists alongside `app.py` — confirm which one is
  actually in use before making changes; the legacy file may be safe to
  remove once confirmed unused.
- Partner API credentials live in `server/routes/.env` — make sure
  `.gitignore` covers this path so they're never committed.

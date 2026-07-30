# Offline Rater Automation

Build local insurance/loan rate tables from Excel, then validate them against a live partner API — see local vs. partner premium side by side, instantly.

## Quick start

```bash
source rater/bin/activate
pip install -r requirements.txt
# fill in server/routes/.env with PARTNER_CODE / PARTNER_KEY / PARTNER_USERID / *_URL vars
honcho start
```

- UI → http://localhost:8501
- API → http://localhost:5050

## Tabs

| Tab | What it does |
|---|---|
| **Dimension Library** | Reusable rating dimensions (Age, Tenure, Sum Insured…) — Enum or Range |
| **Template Builder** | Assemble a product's rating rules from dimensions + formula |
| **Upload & Flatten** | Turn a pivoted rate-card Excel into queryable rate-table rows |
| **Calculate Premium** | Single Customer or Bulk Upload — run Save Loan / Save Lead / Create Lead, get local + partner premium + diff |
| **Matrix Unpivoter (CSV)** | Standalone CSV unpivot utility |

## Custom Payload

Every partner call normally uses a hardcoded JSON template (fake name/PAN/address/income). Hit **Edit Payload** next to Workflow Action to paste your own — DOB/tenure/sum insured (and Name/Mobile/Email for leads) always stay controlled by the UI widgets, everything else in your JSON is used as-is. Session-only, shared between Single Customer and Bulk modes.

> ⚠️ Save Lead / Create Lead in **Bulk Upload** create a real partner-side lead per row — no dry-run. Use Save Loan for safe bulk rating-math tests.

## How it calculates

- **Local premium** — dimension inputs bucket-match against your uploaded rate table; no match = 404.
- **Partner premium** — dimensions injected into the payload, sent live to the partner API.

Both only ever use the Rater Dimension values — identity/address fields never affect the math.

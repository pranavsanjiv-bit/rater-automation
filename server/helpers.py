from datetime import timedelta, date
import json
from hashlib import sha256
from frontend.formula_eval import FormulaError,validate_formula
from frontend.dim_parser import get_dimension_slug, get_value_slug, operator_to_slug, _parse_value_with_units

# ── Predefined input variables ────────────────────────────────────────────────
# Mirrors _LOS_DIMS / _FE_DIMS / _ALL_DIMS in app.py. These are the single
# predefined set of input variables shown on the Calculate Premium screen and
# always usable inside formulas — independent of whether they are part of the
# workbook's parsed_dimensions / row / col axes.

_LOS_DIMS = ["Age", "Loan Amount", "Loan Type", "Loan Tenure", "Gender",
             "Borrowers count"]
_FE_DIMS = ["Sum Insured", "Tenure", "Cover Type"]
_ALL_DIMS = _LOS_DIMS + _FE_DIMS

# Categorical (non-numeric) predefined variables are not valid inside formulas.
_EXCLUDED_FROM_FORMULA = {"Loan Type", "Gender", "Cover Type"}

# Explicit display-name -> slug map for the predefined variables. This is the
# single source of truth for predefined slugs on the backend and MUST stay
# identical to _PREDEFINED_NAME_TO_SLUG in app.py — do not derive these from
# get_dimension_slug(), which produces slugs without underscores
# (e.g. "suminsured" instead of "sum_insured").
_PREDEFINED_NAME_TO_SLUG = {
    "Age": "age",
    "Loan Amount": "loan_amount",
    "Loan Type": "loan_type",
    "Loan Tenure": "loan_tenure",
    "Gender": "gender",
    "Borrowers count": "borrowers_count",
    "Sum Insured": "sum_insured",
    "Tenure": "tenure",
    "Cover Type": "cover_type",
}

_PREDEFINED_VARIABLE_SLUGS = {
    slug for name, slug in _PREDEFINED_NAME_TO_SLUG.items()
    if name not in _EXCLUDED_FROM_FORMULA
}

# ── Slug resolution helpers ───────────────────────────────────────────────────

def _normalize_float_to_str(value) -> str:
    """
    Convert a numeric value to string, normalizing whole-number floats to integers.
    180.0 → "180"
    10.0 → "10"
    180.5 → "180.5"
    """
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _resolve_input_to_slug(dim_name: str, dim_def: dict, raw_value) -> str:
    """
    Given a dimension definition and a raw user input value,
    return the slug that would appear in the lookup key.

    dim_def example (from definition_json["definition"]["dimensions"][dim_name]):
      Enum:       {"type": "Enum",       "config": ["val1", "val2"]}
      Range:      {"type": "Range",      "config": {"min": 0, "max": 100}}
      Comparison: {"type": "Comparison", "config": [{"op": "<", "val": "20"}, ...]}
    """
    dim_type = dim_def["type"]

    if dim_type == "Enum":
        slug = get_value_slug(str(raw_value))
        return slug

    if dim_type == "Range":
        # For a Range dimension used as an *outer* (sheet-splitting) dimension,
        # the only bucket is min_max.  If it's the row/col axis, the raw value
        # IS the slug (plain integer string).
        return get_value_slug(str(raw_value))

    if dim_type == "Comparison":
        comparisons = dim_def["config"]  # list of {"op":..., "val":...}
        numeric_input = float(raw_value)
        op_map = {"<": lambda a, b: a < b, "<=": lambda a, b: a <= b,
                  ">": lambda a, b: a > b, ">=": lambda a, b: a >= b,
                  "=": lambda a, b: a == b}
        for comp in comparisons:
            op, threshold = comp["op"], _parse_value_with_units(comp["val"])
            if op_map.get(op, lambda a, b: False)(numeric_input, threshold):
                op_slug = operator_to_slug(op)
                val_slug = get_value_slug(comp["val"])
                slug = f"{op_slug}_{val_slug}".strip("_")
                return slug
        raise ValueError(
            f"Input '{raw_value}' for dimension '{dim_name}' did not match any comparison bucket."
        )

    raise ValueError(f"Unknown dimension type '{dim_type}' for '{dim_name}'.")


def _build_lookup_key(template_def: dict, parsed_dimensions: dict, inputs: dict) -> tuple[str, dict]:
    """
    Build the deterministic lookup key and a debug breakdown dict.

    Key format:
      outer dims (sorted alphabetically by slug) joined with '-',
      then row value slug,
      then col value slug (if present).

    Returns (lookup_key_str, resolved_buckets_dict)
    """
    definition = template_def["definition"]
    dimensions = definition["dimensions"]
    axes = definition["axes"]
    row_axis = axes.get("row")
    col_axis = axes.get("col")

    row_dim_name = row_axis["name"] if row_axis else None
    col_dim_name = col_axis["name"] if col_axis else None

    resolved = {}  # slug -> resolved_bucket_slug

    # Outer dimensions (not row/col axis)
    outer_slugs = {}
    for dim_name, dim_def in dimensions.items():
        dim_slug = get_dimension_slug(dim_name)
        raw_val = inputs.get(dim_slug)
        if raw_val is None:
            raise ValueError(f"Missing input for dimension '{dim_name}' (slug: '{dim_slug}').")
        bucket_slug = _resolve_input_to_slug(dim_name, dim_def, raw_val)
        outer_slugs[dim_slug] = bucket_slug
        resolved[dim_slug] = bucket_slug

    # Build key parts: outer dims sorted alphabetically
    key_parts = [f"{ds}_{outer_slugs[ds]}" for ds in sorted(outer_slugs)]

    # Row axis value
    if row_dim_name:
        row_slug = get_dimension_slug(row_dim_name)
        raw_row = inputs.get(row_slug)
        if raw_row is None:
            raise ValueError(f"Missing input for row axis '{row_dim_name}' (slug: '{row_slug}').")
        row_val_slug = get_value_slug(_normalize_float_to_str(raw_row))
        key_parts.append(row_val_slug)
        resolved[row_slug] = row_val_slug

    # Col axis value
    if col_dim_name:
        col_slug = get_dimension_slug(col_dim_name)
        raw_col = inputs.get(col_slug)
        if raw_col is None:
            raise ValueError(f"Missing input for col axis '{col_dim_name}' (slug: '{col_slug}').")
        col_val_slug = get_value_slug(_normalize_float_to_str(raw_col))
        key_parts.append(col_val_slug)
        resolved[col_slug] = col_val_slug

    return "-".join(key_parts), resolved


# ── Template body parsing ─────────────────────────────────────────────────────

def _parse_template_body(body: dict) -> tuple[dict, str, str] | tuple[None, None, str]:
    """Validate request body and return (full_def, name, content_hash) or (None, None, error)."""
    definition = body.get("definition")
    parsed_dimensions = body.get("parsed_dimensions")
    calculation = body.get("calculation", {})
    name = (body.get("name") or "").strip() or "Untitled"

    if definition is None or parsed_dimensions is None:
        return None, None, "'definition' and 'parsed_dimensions' are required."
    
    formula = calculation.get("formula", "").strip()
    if formula:
        constants = {c["name"] for c in calculation.get("constants", [])}
        allowed = _PREDEFINED_VARIABLE_SLUGS | constants | {"rater_val"}

        try:
            validate_formula(formula, allowed)
        except FormulaError as e:
            return None, None, f"Formula validation failed: {e}"

    full_def = {
        "definition": definition,
        "parsed_dimensions": parsed_dimensions,
        "calculation": calculation,
    }
    content_hash = sha256(json.dumps(full_def, sort_keys=True).encode()).hexdigest()
    return full_def, name, content_hash


def _read_workbook_meta(wb) -> dict:
    """Extract the JSON payload from the hidden _meta worksheet."""
    if "_meta" not in wb.sheetnames:
        raise ValueError("Workbook has no _meta sheet; not a valid rater template.")
    ws_meta = wb["_meta"]
    raw = ws_meta.cell(row=1, column=1).value
    if not raw:
        raise ValueError("_meta sheet is empty; cannot read template fingerprint.")
    try:
        return json.loads(raw)
    except json.JSONDecodeError as e:
        raise ValueError(f"_meta payload is not valid JSON: {e}") from e


def _age_to_dob(age: int, format = None) -> str:
    """Convert age in years to a DD-MM-YYYY string (approximate)."""
    today = date.today()
    dob = today - timedelta(days=age * 365.25)
    return dob.strftime(format if format is not None else "%d-%m-%Y")


def _dob_to_age(dob_str: str) -> int:
    """Calculate age from a DOB string in DD-MM-YYYY, YYYY-MM-DD, or DD/MM/YYYY formats."""
    from datetime import datetime, date
    for fmt in ("%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            dob = datetime.strptime(dob_str, fmt).date()
            today = date.today()
            age = today.year - dob.year - ((today.month, today.day) < (dob.month, dob.day))
            return age
        except ValueError:
            continue
    raise ValueError(f"Invalid DOB format: '{dob_str}'. Expected DD-MM-YYYY.")


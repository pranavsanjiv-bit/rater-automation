import math
import re

# ── LOS / FE predefined dimension names ──────────────────────────────────────
# Single source of truth for "predefined" dimension names/slugs that can be
# referenced inside a formula without being an explicit sheet/row/col
# dimension or a Formula Variable — e.g. Sum Insured, Loan Amount. Both
# frontend/app.py (the Single Customer input form, via _PREDEFINED_SLUG_TO_NAME)
# and _template_required_dimensions() below read from this same dict, so a
# template's required-inputs list can never drift out of sync with what the
# Single Customer form actually asks for.

_LOS_DIMS = ["Age", "Loan Amount", "Loan Type", "Loan Tenure", "Gender",
             "Borrowers count"]
_FE_DIMS = ["Sum Insured", "Tenure", "Cover Type"]
_ALL_DIMS = _LOS_DIMS + _FE_DIMS
_DIM_CATEGORY = {name: "LOS" for name in _LOS_DIMS}
_DIM_CATEGORY.update({name: "FE" for name in _FE_DIMS})

_EXCLUDED_FROM_FORMULA = {"Loan Type", "Gender", "Cover Type"}

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

_PREDEFINED_SLUG_TO_NAME = {
    slug: name
    for name, slug in _PREDEFINED_NAME_TO_SLUG.items()
    if name not in _EXCLUDED_FROM_FORMULA
}

def get_dimension_slug(name: str) -> str:
    """Creates a slug from the dimension name (e.g., 'Sum Insured' -> 'suminsured')."""
    return re.sub(r"[^\w]", "", name.strip().lower())

def _parse_value_with_units(val_str:str) -> int:
    """Parses a value like '20L' or '20L and above' and returns the numeric value in rupees."""
    val_str = val_str.lower().strip()
    if val_str.endswith('l'):
        return float(val_str[:-1]) * 100000
    elif val_str.endswith('cr'):
        return float(val_str[:-2]) * 10000000
    elif val_str.endswith('k'):
        return float(val_str[:-1]) * 1000
    elif val_str.endswith('m'):
        return float(val_str[:-1]) * 1000000
    elif val_str.endswith('b'):
        return float(val_str[:-1]) * 1000000000
    elif val_str.endswith('%'):
        return float(val_str[:-1]) / 100
    else:
        return float(val_str)

def get_value_slug(value: str) -> str:
    """Converts a value to a slug (e.g., '20L and above' -> '20l_and_above')."""
    s = str(value).lower().replace(" ", "_")
    return re.sub(r'[^\w]', '', s)

def operator_to_slug(op: str) -> str:
    mapping = {"<": "lt", "<=": "lte", ">": "gt", ">=": "gte", "=": ""}
    return mapping.get(op, "")

def normalize_dimension_value(dim_name: str, dim_type: str, raw_val: dict):
    """
    Returns dict mapping slug -> human_label
    raw_val: dict containing config based on dim_type
      - Enum: {"values": [...]}
      - Range: {"min": X, "max": Y}
      - Comparison: {"comparisons": [{"op": str, "val": str}, ...]}
    """
    dim_slug = get_dimension_slug(dim_name)
    result = {}
    
    if dim_type == "Enum":
        values = raw_val.get("values", [])
        for val in values:
            slug = get_value_slug(val)
            result[slug] = val
        return result
    
    if dim_type == "Range":
        min_val = raw_val.get('min', 0)
        max_val = raw_val.get('max', 100)
        slug = f"{min_val}_{max_val}"
        result[slug] = f"{min_val} to {max_val}"
        return result
    
    if dim_type == "Comparison":
        comparisons = raw_val.get("comparisons", [])
        for comp in comparisons:
            op, val = comp.get('op', ''), comp.get('val', '')
            op_slug = operator_to_slug(op)
            
            # Human Label
            op_map_human = {"<": "Less than", "<=": "Less than or equal to", 
                            ">": "Greater than", ">=": "Greater than or equal to", "=": ""}
            label = f"{op_map_human.get(op, '')} {val}".strip()
            
            # Slug
            val_part = get_value_slug(val)
            slug = f"{op_slug}_{val_part}".replace("__", "_").strip("_")
            result[slug] = label
        
        return result
    
    return result

def _template_required_dimensions(full_def: dict) -> list[dict]:
    """
    Returns [{"name", "slug", "type", "sample"}, ...] for every input a
    template needs a value for: each outer (sheet-level) dimension plus the
    row and column axis (if set).

    Shared by both the backend bulk-calculate column validation
    (server/routes/templates.py) and the frontend "Download Sample
    Template" generator (frontend/app.py) so neither hardcodes a dimension
    list — both always derive it fresh from the template definition, and
    stay in sync with each other automatically.

    Note: outer dimensions and row/col axes store their config in slightly
    different shapes (outer: {"type":..., "config": {...}}; axis:
    {"name":..., "type":..., <type-specific keys directly>}) — this
    function normalizes both into the same {name, slug, type, sample}
    shape rather than leaking that difference to callers.
    """
    definition = full_def["definition"]
    dimensions = definition.get("dimensions", {})
    axes = definition.get("axes", {})

    def _sample(dim_type: str, values) -> str:
        if dim_type == "Enum":
            return values[0] if values else ""
        if dim_type == "Range":
            return str(values.get("min", "")) if isinstance(values, dict) else ""
        if dim_type == "Comparison":
            comps = values or []
            return str(comps[0].get("val", "")) if comps else ""
        return ""

    required = []
    seen = set()

    for dim_name, dim_def in dimensions.items():
        slug = get_dimension_slug(dim_name)
        if slug in seen:
            continue
        dim_type = dim_def["type"]
        config = dim_def.get("config", {})
        if dim_type == "Enum":
            values = config.get("values", [])
        elif dim_type == "Range":
            values = config
        else:
            values = config.get("comparisons", [])
        required.append({"name": dim_name, "slug": slug, "type": dim_type, "sample": _sample(dim_type, values)})
        seen.add(slug)

    for axis_key in ("row", "col"):
        axis = axes.get(axis_key)
        if not axis:
            continue
        slug = get_dimension_slug(axis["name"])
        if slug in seen:
            continue
        dim_type = axis["type"]
        if dim_type == "Enum":
            values = axis.get("config", [])
        elif dim_type == "Range":
            values = {"min": axis.get("min"), "max": axis.get("max")}
        else:
            values = axis.get("config", [])
        required.append({"name": axis["name"], "slug": slug, "type": dim_type, "sample": _sample(dim_type, values)})
        seen.add(slug)

    # Formula Variables — extra inputs used only in the formula, not part
    # of the rate table itself. Same source _render_calc_inputs() reads
    # (full_def["calculation"]["formula_variables"]).
    calculation = full_def.get("calculation", {})
    for fv in calculation.get("formula_variables", []):
        slug = fv.get("slug", "")
        if not slug or slug in seen:
            continue
        default_val = fv.get("default_value", "")
        required.append({
            "name": fv.get("name", slug), "slug": slug, "type": "Comparison",
            "sample": str(default_val) if default_val not in (None, "") else "0",
        })
        seen.add(slug)

    # Predefined slugs (Age, Sum Insured, Loan Amount, ...) referenced
    # directly inside the formula text but not already covered above —
    # mirrors _render_calc_inputs()'s own regex scan over the formula
    # string, so a template needing e.g. "sum_insured" in its formula gets
    # that column here too, not just in the Single Customer form.
    formula_str = calculation.get("formula", "") or ""
    for slug, display_name in _PREDEFINED_SLUG_TO_NAME.items():
        if slug in seen:
            continue
        if not re.search(rf"\b{re.escape(slug)}\b", formula_str):
            continue
        required.append({"name": display_name, "slug": slug, "type": "Comparison", "sample": "0"})
        seen.add(slug)

    return required


# ── Premium comparison (shared by Single Premium's UI and Bulk Premium) ──────
# Pure functions, no Streamlit/Flask dependency, so both the Streamlit
# renderer (frontend/app.py: _render_premium_comparison) and the backend
# bulk-processing route (server/routes/templates.py: calculate_bulk) call
# the exact same comparison logic — one source of truth, not two.

def _round2(val):
    """Round to currency (2dp) precision for comparison, so floating-point
    noise (e.g. 262.580000001 vs 262.58) never produces a false MISMATCH.
    Returns None untouched — that's how a metric is flagged 'not returned'."""
    if val is None:
        return None
    try:
        return round(float(val), 2)
    except (TypeError, ValueError):
        return None


def _apply_rounding_rule(val, rounding_rule: str):
    """Rounds a value to a whole number per the template's chosen rule,
    for MATCH/MISMATCH comparison purposes only — never used for the
    displayed/exported premium itself, which always keeps full precision.
    'none' returns the value unchanged (original exact-to-the-paisa
    comparison behavior)."""
    if val is None or rounding_rule == "none":
        return val
    if rounding_rule == "up":
        return math.ceil(val)
    if rounding_rule == "down":
        return math.floor(val)
    return round(val)  # "nearest" (default) and any unrecognized value


def _build_comparison_row(label: str, local_val, partner_val, rounding_rule: str = "nearest") -> dict:
    """One row of the sanity-check comparison. If either side didn't return
    this metric, the row is 'unavailable' rather than a false MATCH/MISMATCH.

    rounding_rule only affects the MATCH/MISMATCH verdict — the 'local'/
    'partner'/'diff' values returned here always stay at full (2dp)
    precision, so nothing about the displayed or exported premium numbers
    changes regardless of which rule is chosen. This exists because a
    partner API commonly returns whole-rupee premiums while the local
    Excel-derived calculation keeps decimals, which would otherwise show
    as a MISMATCH for a difference that's really just rounding, not a
    rating error."""
    local_r = _round2(local_val)
    partner_r = _round2(partner_val)

    if local_r is None or partner_r is None:
        status, diff = "unavailable", None
    elif _apply_rounding_rule(local_r, rounding_rule) == _apply_rounding_rule(partner_r, rounding_rule):
        status, diff = "match", 0.0
    else:
        status, diff = "mismatch", abs(local_r - partner_r)

    status_label = {"match": "🟢 MATCH", "mismatch": "🔴 MISMATCH", "unavailable": "⚪ Unavailable"}[status]
    return {
        "label": label, "local": local_r, "partner": partner_r, "diff": diff,
        "status": status, "status_label": status_label,
    }


def _overall_premium_status(rows: list) -> tuple[str, str]:
    """(icon, text) summary across a list of _build_comparison_row rows —
    mismatch always wins, then unavailable, then match."""
    if any(r["status"] == "mismatch" for r in rows):
        return "🔴", "Premium Mismatch Detected"
    if any(r["status"] == "unavailable" for r in rows):
        return "⚪", "Some Values Unavailable for Comparison"
    return "🟢", "All Premium Values Match"
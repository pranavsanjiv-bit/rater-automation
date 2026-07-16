# pyrefly: ignore [missing-import]
import streamlit as st
# pyrefly: ignore [missing-import]
from code_editor import code_editor as st_code_editor
import re
import requests
import json
from dim_parser import (
    get_dimension_slug,
    get_value_slug,
    operator_to_slug,
)
from excel_generator import generate_excel
from formula_eval import validate_formula, FormulaError
from concurrent.futures import ThreadPoolExecutor

st.set_page_config(page_title="Offline rater automation", layout="wide")

st.markdown(
    """
    <style>
    /* ── Global Streamlit noise suppression ──────────────────────────────── */
    [data-testid="InputInstructions"] { display: none !important; }

    /* ── Formula bar ─────────────────────────────────────────────────────── */
    input[placeholder*="formula"] {
        font-family: monospace !important;
        font-size: 1rem !important;
        background-color: #1e1e2e !important;
        color: #cdd6f4 !important;
        border: 1px solid #45475a !important;
    }
    .formula-bar {
        font-family: monospace;
        font-size: 1rem;
        background: #1e1e2e;
        color: #cdd6f4;
        padding: 0 0.75rem;
        border-radius: 6px;
        height: 2.4rem;
        line-height: 2.4rem;
        border: 1px solid #45475a;
        word-break: break-all;
        box-sizing: border-box;
    }
    .formula-bar-empty {
        color: #585b70;
        font-style: italic;
    }

    /* ── All buttons: baseline height + flex centering ───────────────────── */
    div[data-testid="stButton"] > button {
        height: 2.4rem;
        padding-top: 0 !important;
        padding-bottom: 0 !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        line-height: 1 !important;
    }

    /* ── Card row columns: vertically center every cell ─────────────────── */
    div[data-testid="stColumn"] > div[data-testid="stVerticalBlock"] {
        display: flex !important;
        flex-direction: column !important;
        justify-content: center !important;
        min-height: 2.4rem;
    }

    /* ── Role pill badges ────────────────────────────────────────────────── */
    .pill {
        display: inline-flex;
        align-items: center;
        padding: 0.2rem 0.6rem;
        border-radius: 12px;
        font-size: 0.82rem;
        white-space: nowrap;
        line-height: 1;
    }
    .pill-green { background: #1a4731; color: #4ade80; }
    .pill-blue  { background: #172554; color: #60a5fa; }

    /* ── FV slug badge ───────────────────────────────────────────────────── */
    .fv-slug {
        font-family: monospace;
        font-size: 0.82rem;
        background: #2a2a3d;
        color: #a6e3a1;
        padding: 0.2rem 0.5rem;
        border-radius: 4px;
        display: inline-block;
        line-height: 1.4;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

OPERATORS = ["<", "<=", ">", ">="]
SERVER_URL = "http://localhost:5050"


# ── Session state init ────────────────────────────────────────────────────────

def init_state():
    defaults = {
        "dims": [],
        "form_open": False,
        "form": _empty_form(),
        "last_result": None,
        # formula builder
        "formula_tokens": [],
        "formula_editor_input": "",
        # formula variables
        "formula_variables": [],   # list of {name, slug, default_value}
        "_fv_add_open": False,
        "_fv_editing": None,       # int index of row being edited, or None
        # cross-tab state
        "last_template_id": None,
        "template_name": "",
        "product_code": "",
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def _empty_form():
    return {
        "category": "LOS",
        "name": "",
        "type": "Enum",
        "values": [],
        "min": 0,
        "max": 100,
        "comparisons": [],
        "_val_input_seq": 0,
        "_comp_val_input_seq": 0,
    }


# ── Role helpers ──────────────────────────────────────────────────────────────

def _row_dim():
    return next((d for d in st.session_state.dims if d.get("role") == "row"), None)


def _col_dim():
    return next((d for d in st.session_state.dims if d.get("role") == "col"), None)


def _set_role(target_dim, role):
    if role == "col" and _row_dim() is None:
        st.warning("Set a Row axis before choosing a Column axis.", icon="⚠️")
        return
    for d in st.session_state.dims:
        if d.get("role") == role:
            d["role"] = None
    target_dim["role"] = role
    st.session_state.last_result = None


def _clear_role(target_dim):
    role = target_dim.get("role")
    target_dim["role"] = None
    if role == "row":
        col = _col_dim()
        if col is not None:
            col["role"] = None
    st.session_state.last_result = None


# ── Helpers ───────────────────────────────────────────────────────────────────

def dim_summary(dim):
    if dim["type"] == "Enum":
        vals = dim["config"].get("values", [])
        return ", ".join(vals) if vals else "—"
    if dim["type"] == "Range":
        return f"{dim['config']['min']} → {dim['config']['max']}"
    if dim["type"] == "Comparison":
        comps = dim["config"].get("comparisons", [])
        return ", ".join(f"{c['op']} {c['val']}" for c in comps) if comps else "—"
    return "—"


def build_dim_from_form(form):
    if form["type"] == "Enum":
        config = {"values": list(form["values"])}
    elif form["type"] == "Range":
        config = {"min": form["min"], "max": form["max"]}
    else:
        config = {"comparisons": [dict(c) for c in form["comparisons"]]}
    return {"name": form["name"].strip(), "type": form["type"], "config": config, "role": None}


# ── Reusable dimension form ───────────────────────────────────────────────────

def render_dim_form(form_key, form, on_confirm_key, on_cancel_key):
    col_name, col_type = st.columns([3, 2])
    with col_name:
        form["name"] = st.text_input("Name", value=form["name"], key=f"{form_key}_name")
    with col_type:
        form["type"] = st.selectbox(
            "Type",
            ["Enum", "Range", "Comparison"],
            index=["Enum", "Range", "Comparison"].index(form["type"]),
            key=f"{form_key}_type",
        )

    if form["type"] == "Enum":
        _render_enum_inputs(form_key, form)
    elif form["type"] == "Range":
        c1, c2 = st.columns(2)
        with c1:
            form["min"] = st.number_input("Min", value=int(form["min"]), step=1, key=f"{form_key}_min")
        with c2:
            form["max"] = st.number_input("Max", value=int(form["max"]), step=1, key=f"{form_key}_max")
    elif form["type"] == "Comparison":
        _render_comparison_inputs(form_key, form)

    st.write("")
    st.write("")
    btn_col1, btn_col2, _ = st.columns([1, 1, 4])
    with btn_col1:
        if st.button("Add dimension", key=f"{form_key}_confirm", type="primary"):
            st.session_state[on_confirm_key] = True
    with btn_col2:
        if st.button("Cancel", key=f"{form_key}_cancel"):
            st.session_state[on_cancel_key] = True


def _render_enum_inputs(form_key, form):
    st.markdown("**Values**")
    seq = form.get("_val_input_seq", 0)
    input_key = f"{form_key}_new_val_{seq}"

    c1, c2 = st.columns([4, 1])
    with c1:
        st.text_input(
            "New value", value="", key=input_key,
            label_visibility="collapsed", placeholder="Type a value and press add →",
        )
    with c2:
        if st.button("Add value", key=f"{form_key}_add_val"):
            v = st.session_state[input_key].strip()
            if v and v not in form["values"]:
                form["values"].append(v)
            form["_val_input_seq"] = seq + 1
            st.rerun()

    vals = form["values"]
    if vals:
        st.write("")
        to_delete = None
        for i, v in enumerate(vals):
            c1, c2 = st.columns([5, 1])
            with c1:
                st.markdown(f"- {v}")
            with c2:
                if st.button("✕", key=f"{form_key}_del_val_{i}"):
                    to_delete = i
        if to_delete is not None:
            form["values"].pop(to_delete)
            st.rerun()
    else:
        st.caption("No values yet.")


def _render_comparison_inputs(form_key, form):
    st.markdown("**Conditions**")
    comps = form["comparisons"]
    to_delete = None
    for i, c in enumerate(comps):
        cc1, cc2, cc3, _ = st.columns([1.5, 3, 0.8, 0.2])
        with cc1:
            comps[i]["op"] = st.selectbox(
                "Op", OPERATORS, index=OPERATORS.index(c["op"]),
                key=f"{form_key}_comp_op_{i}", label_visibility="collapsed",
            )
        with cc2:
            comps[i]["val"] = st.text_input(
                "Val", value=c["val"], key=f"{form_key}_comp_val_{i}",
                label_visibility="collapsed", placeholder="value",
            )
        with cc3:
            if st.button("Remove", key=f"{form_key}_del_comp_{i}"):
                to_delete = i

    if to_delete is not None:
        form["comparisons"].pop(to_delete)
        st.rerun()

    seq = form.get("_comp_val_input_seq", 0)
    new_val_key = f"{form_key}_new_comp_val_{seq}"

    c1, c2, c3 = st.columns([1.5, 3, 2])
    with c1:
        new_op = st.selectbox("New op", OPERATORS, key=f"{form_key}_new_op", label_visibility="collapsed")
    with c2:
        st.text_input(
            "New val", key=new_val_key,
            label_visibility="collapsed", placeholder="value",
        )
    with c3:
        if st.button("＋ Add condition", key=f"{form_key}_add_comp"):
            new_cv = st.session_state[new_val_key]
            form["comparisons"].append({"op": new_op, "val": new_cv.strip()})
            form["_comp_val_input_seq"] = seq + 1
            st.rerun()


# ── LOS / FE predefined dimension names ──────────────────────────────────────

_LOS_DIMS = ["Age", "Loan Amount", "Loan Type", "Loan Tenure", "Gender",
             "Borrowers count", ]
_FE_DIMS  = ["Sum Insured", "Tenure", "Cover Type"]
# All predefined dimensions in one ordered list (LOS first, then FE)
_ALL_DIMS = _LOS_DIMS + _FE_DIMS
# Internal mapping: name → category (not shown in UI)
_DIM_CATEGORY = {name: "LOS" for name in _LOS_DIMS}
_DIM_CATEGORY.update({name: "FE" for name in _FE_DIMS})

# Categorical (non-numeric) predefined variables are not valid inside formulas.
_EXCLUDED_FROM_FORMULA = {"Loan Type", "Gender", "Cover Type"}

# Explicit display-name -> slug map for the predefined variables. This is the
# single source of truth for predefined slugs and MUST stay identical to the
# mapping in helpers.py (_PREDEFINED_NAME_TO_SLUG) — do not derive these from
# get_dimension_slug(), which produces slugs without underscores
# (e.g. "suminsured" instead of "sum_insured") and will not match the backend.
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

# slug -> display name, for predefined numeric variables (used when a formula
# references a predefined variable that isn't part of the workbook itself)
_PREDEFINED_SLUG_TO_NAME = {
    slug: name
    for name, slug in _PREDEFINED_NAME_TO_SLUG.items()
    if name not in _EXCLUDED_FROM_FORMULA
}


# ── Section: unified dimension table + role assignment ────────────────────────

def _render_dim_rows():
    """Render one card per dimension in the single combined dims list."""
    dims = st.session_state.dims
    row_dim = _row_dim()
    col_dim = _col_dim()

    # Column header labels
    lh1, lh3, lh4, lh5 = st.columns([3.5, 1.3, 3, 0.7])
    with lh1:
        st.caption("Dimension")
    with lh3:
        st.caption("Type")
    with lh4:
        st.caption("Axis Role")
    with lh5:
        st.caption("")

    for i, dim in enumerate(dims):
        with st.container(border=True):
            c_name, c_type, c_role, c_del = st.columns([3.5, 1.3, 3, 0.7])

            with c_name:
                st.markdown(f"**{dim['name']}**")

            with c_type:
                st.markdown(f"<small>{dim['type']}</small>", unsafe_allow_html=True)

            role = dim.get("role")
            with c_role:
                if role == "row":
                    rc1, rc2 = st.columns([3, 1])
                    with rc1:
                        st.markdown('<span class="pill pill-green">● Row</span>',
                                    unsafe_allow_html=True)
                    with rc2:
                        if st.button("✕", key=f"unassign_{i}", help="Unassign Row"):
                            _clear_role(dim)
                            st.rerun()
                elif role == "col":
                    rc1, rc2 = st.columns([3, 1])
                    with rc1:
                        st.markdown('<span class="pill pill-blue">● Column</span>',
                                    unsafe_allow_html=True)
                    with rc2:
                        if st.button("✕", key=f"unassign_{i}", help="Unassign Column"):
                            _clear_role(dim)
                            st.rerun()
                else:
                    can_be_row = row_dim is None
                    can_be_col = (row_dim is not None) and (dim is not row_dim) and (col_dim is None)
                    help_col_text = None
                    if row_dim is None:
                        help_col_text = "Set a Row axis first"
                    elif col_dim is not None:
                        help_col_text = "Clear the current Column axis first"
                    elif dim is row_dim:
                        help_col_text = "Already the Row axis"
                    rb1, rb2 = st.columns(2)
                    with rb1:
                        if st.button("→ Row", key=f"set_row_{i}",
                                     disabled=not can_be_row,
                                     help=None if can_be_row else "Clear the current Row axis first",
                                     use_container_width=True):
                            _set_role(dim, "row")
                            st.rerun()
                    with rb2:
                        if st.button("→ Col", key=f"set_col_{i}",
                                     disabled=not can_be_col,
                                     help=help_col_text,
                                     use_container_width=True):
                            _set_role(dim, "col")
                            st.rerun()

            with c_del:
                if st.button("🗑", key=f"del_dim_{i}", help="Remove dimension"):
                    st.session_state.dims.pop(i)
                    st.session_state.last_result = None
                    st.rerun()


def render_dims_section():
    for k in ["do_confirm", "do_cancel"]:
        if k not in st.session_state:
            st.session_state[k] = False

    # ── Header + Add button ───────────────────────────────────────────────────
    hdr_col, btn_col = st.columns([3, 2])
    with hdr_col:
        st.markdown("#### Dimensions")
        st.caption("Dimensions that exist in your rate table / Excel.")
    with btn_col:
        st.write("")
        if not st.session_state.form_open:
            if st.button("＋ Add Dimension", key="open_form", use_container_width=True):
                st.session_state.form_open = True
                st.session_state.form = _empty_form()
                st.rerun()

    # ── Unified dimension table ───────────────────────────────────────────────
    if st.session_state.dims:
        _render_dim_rows()
    else:
        st.caption("No dimensions yet.")

    # ── Add form (opens when button above is clicked) ─────────────────────────
    if st.session_state.form_open:
        form = st.session_state.form

        with st.container(border=True):
            st.markdown("##### New dimension")

            # Create two columns for Dimension and Type side-by-side
            col_dim, col_type = st.columns(2)

            with col_dim:
                # Single dimension dropdown — all predefined names, already-added ones excluded
                already_added = {d["name"] for d in st.session_state.dims}
                available = [n for n in _ALL_DIMS if n not in already_added]

                if available:
                    default_idx = available.index(form["name"]) if form["name"] in available else 0
                    form["name"] = st.selectbox(
                        "Dimension", available, index=default_idx, key="form_name_select",
                    )
                    # Derive category internally — not shown to the user
                    form["category"] = _DIM_CATEGORY.get(form["name"], "LOS")
                else:
                    form["name"] = ""
                    st.info("All predefined dimensions have already been added.", icon="ℹ️")

            with col_type:
                # Type + config (unchanged existing flow)
                form["type"] = st.selectbox(
                    "Type", ["Enum", "Range", "Comparison"],
                    index=["Enum", "Range", "Comparison"].index(form["type"]),
                    key="form_type",
                )
            if form["type"] == "Enum":
                _render_enum_inputs("form", form)
            elif form["type"] == "Range":
                c1, c2 = st.columns(2)
                with c1:
                    form["min"] = st.number_input("Min", value=int(form["min"]), step=1, key="form_min")
                with c2:
                    form["max"] = st.number_input("Max", value=int(form["max"]), step=1, key="form_max")
            elif form["type"] == "Comparison":
                _render_comparison_inputs("form", form)

            st.write("")
            btn_col1, btn_col2, _ = st.columns([3, 2, 3])
            with btn_col1:
                if st.button("Add dimension", key="form_confirm", type="primary",
                             disabled=not available, use_container_width=True):
                    st.session_state.do_confirm = True
            with btn_col2:
                if st.button("Cancel", key="form_cancel", use_container_width=True):
                    st.session_state.do_cancel = True

        if st.session_state.get("do_confirm"):
            st.session_state.do_confirm = False
            f = st.session_state.form
            if f["name"].strip():
                new_dim = build_dim_from_form(f)
                new_dim["category"] = f["category"]
                st.session_state.dims.append(new_dim)
            st.session_state.form_open = False
            st.session_state.form = _empty_form()
            st.session_state.last_result = None
            st.rerun()

        if st.session_state.get("do_cancel"):
            st.session_state.do_cancel = False
            st.session_state.form_open = False
            st.session_state.form = _empty_form()
            st.rerun()

    st.markdown(
        '<div style="background:#1e1e30;border-radius:6px;padding:0.55rem 0.8rem;'
        'margin-top:0.75rem;font-size:0.82rem;color:#888">'
        'ℹ️ &nbsp; Set exactly one dimension as <b>Row</b> (mandatory). '
        'Set zero or one as <b>Column</b> (optional). '
        'All others become sheet-level dimensions.</div>',
        unsafe_allow_html=True,
    )


# ── Section: axis summary ─────────────────────────────────────────────────────

def render_axis_summary_section():
    st.markdown("### 📊 Table axes")
    st.caption(
        "These define the 2D grid (or 1D list, if no Column is set) inside each sheet. "
        "Change them any time using the Row/Column buttons above."
    )

    row_dim = _row_dim()
    col_dim = _col_dim()

    c1, c2 = st.columns(2)
    with c1:
        with st.container(border=True):
            st.markdown("**Row axis**")
            if row_dim:
                st.markdown(f"**{row_dim['name']}** `{row_dim['type']}`")
                st.caption(dim_summary(row_dim))
            else:
                st.caption("Not set — mark a dimension as Row above.")

    with c2:
        with st.container(border=True):
            st.markdown("**Column axis**")
            if col_dim:
                st.markdown(f"**{col_dim['name']}** `{col_dim['type']}`")
                st.caption(dim_summary(col_dim))
            elif row_dim:
                st.caption("Not set — sheets will be 1D (Row values only).")
            else:
                st.caption("Set a Row axis first.")


# ── Section: formula variables table ─────────────────────────────────────────

def _fv_slug_from_name(name: str) -> str:
    """Convert a display name to a formula variable slug."""
    return get_value_slug(name.strip())


def render_formula_variables_table():
    """Editable card list of formula-only variables (not part of workbook)."""
    fvs: list = st.session_state.formula_variables

    hdr_col, add_col = st.columns([3, 2])
    with hdr_col:
        st.markdown("#### Formula Variables (Extra Inputs)")
        st.caption("Additional inputs used only in formulas (not part of the rate table).")
    with add_col:
        st.write("")
        if not st.session_state._fv_add_open and st.session_state._fv_editing is None:
            if st.button("＋ Add Variable", key="fv_open_add", use_container_width=True):
                st.session_state._fv_add_open = True
                st.session_state["_fv_new_name"] = ""
                st.session_state["_fv_new_default"] = ""
                st.rerun()

    # ── Column header labels ──────────────────────────────────────────────────
    if fvs:
        lh1, lh2, lh3, lh4 = st.columns([2.5, 2, 1.5, 1])
        with lh1:
            st.caption("Name")
        with lh2:
            st.caption("Slug")
        with lh3:
            st.caption("Default Value")
        with lh4:
            st.caption("Actions")

    # Existing variable cards
    to_delete = None
    for i, fv in enumerate(fvs):
        if st.session_state._fv_editing == i:
            # ── Inline edit form ──────────────────────────────────────────
            with st.container(border=True):
                st.markdown(f"**Edit variable #{i + 1}**")
                ec1, ec2 = st.columns([5, 3])
                with ec1:
                    edit_name = st.text_input(
                        "Display Name", value=fv["name"],
                        key=f"fv_edit_name_{i}",
                    )
                    edit_slug = _fv_slug_from_name(edit_name) if edit_name else fv["slug"]
                    st.caption(f"Slug: `{edit_slug}`")
                with ec2:
                    edit_default = st.text_input(
                        "Default Value", value=str(fv.get("default_value", "")),
                        key=f"fv_edit_default_{i}",
                        placeholder="—",
                    )

                # Slug collision warning
                existing_slugs = set(_get_dim_slug_options())
                other_fv_slugs = {fvs[j]["slug"] for j in range(len(fvs)) if j != i}
                if edit_slug and (edit_slug in existing_slugs or edit_slug in other_fv_slugs):
                    st.warning(f"⚠️ Slug `{edit_slug}` conflicts with an existing variable.")

                sb1, sb2, _ = st.columns([1, 1, 4])
                with sb1:
                    if st.button("Save", key=f"fv_save_{i}", type="primary"):
                        if edit_name.strip() and edit_slug.strip():
                            fvs[i] = {
                                "name": edit_name.strip(),
                                "slug": edit_slug.strip(),
                                "default_value": edit_default.strip(),
                            }
                            st.session_state._fv_editing = None
                            st.rerun()
                with sb2:
                    if st.button("Cancel", key=f"fv_cancel_edit_{i}"):
                        st.session_state._fv_editing = None
                        st.rerun()
        else:
            # ── Display card ──────────────────────────────────────────────
            with st.container(border=True):
                dc1, dc2, dc3, dc4 = st.columns([2.5, 2, 1.5, 1])
                with dc1:
                    st.markdown(f"**{fv['name']}**")
                with dc2:
                    st.markdown(
                        f'<span class="fv-slug">{fv["slug"]}</span>',
                        unsafe_allow_html=True,
                    )
                with dc3:
                    dv = fv.get("default_value", "")
                    st.markdown(
                        f'<small style="color:#888">{dv if dv != "" else "—"}</small>',
                        unsafe_allow_html=True,
                    )
                with dc4:
                    btn_c1, btn_c2 = st.columns(2)
                    with btn_c1:
                        if st.button("✏️", key=f"fv_edit_btn_{i}", help="Edit"):
                            st.session_state._fv_editing = i
                            st.session_state._fv_add_open = False
                            st.rerun()
                    with btn_c2:
                        if st.button("🗑", key=f"fv_del_btn_{i}", help="Delete"):
                            to_delete = i

    if to_delete is not None:
        fvs.pop(to_delete)
        if st.session_state._fv_editing == to_delete:
            st.session_state._fv_editing = None
        st.rerun()

    # ── Add new variable form ─────────────────────────────────────────────────
    if st.session_state._fv_add_open:
        with st.container(border=True):
            st.markdown("**New formula variable**")
            ac1, ac2 = st.columns([5, 3])
            with ac1:
                new_name = st.text_input(
                    "Display Name", value="",
                    key="fv_add_name",
                    placeholder="e.g. Loan Amount",
                )
                new_slug = _fv_slug_from_name(new_name) if new_name else ""
                st.caption(f"Slug: `{new_slug}`")
            with ac2:
                new_default = st.text_input(
                    "Default Value", value="",
                    key="fv_add_default",
                    placeholder="—",
                )

            # Slug collision warning
            existing_slugs = set(_get_dim_slug_options())
            existing_fv_slugs = {fv["slug"] for fv in fvs}
            if new_slug and (new_slug in existing_slugs or new_slug in existing_fv_slugs):
                st.warning(f"⚠️ Slug `{new_slug}` conflicts with an existing variable.")

            ab1, ab2, _ = st.columns([1, 1, 4])
            with ab1:
                if st.button("Add variable", key="fv_confirm_add", type="primary"):
                    if new_name.strip() and new_slug.strip():
                        fvs.append({
                            "name": new_name.strip(),
                            "slug": new_slug.strip(),
                            "default_value": new_default.strip(),
                        })
                        st.session_state._fv_add_open = False
                        st.rerun()
            with ab2:
                if st.button("Cancel", key="fv_cancel_add"):
                    st.session_state._fv_add_open = False
                    st.rerun()

    if not fvs and not st.session_state._fv_add_open:
        st.caption("No formula variables yet. Add one above.")


# ── Section: formula builder ──────────────────────────────────────────────────

def _get_dim_slug_options() -> list[str]:
    """Return all dimension slugs available as formula variables.
    
    Includes:
    - Outer (non-row/col) dims currently added to the session (excluding categorical ones)
    - Row/col axis dim slugs if set (excluding categorical ones)
    - All predefined numeric LOS + FE dimension slugs

    Excluded (categorical/non-numeric): Loan Type, Gender, Cover Type.
    """
    # Categorical dims that don't belong in a numeric formula
    _EXCLUDED_FROM_FORMULA = {"Loan Type", "Gender", "Cover Type"}

    slugs = []
    seen = set()
    row_dim = _row_dim()
    col_dim = _col_dim()

    # First: dims already added to this template (outer dims)
    for dim in st.session_state.dims:
        if dim.get("role") in ("row", "col"):
            continue
        if dim["name"] in _EXCLUDED_FROM_FORMULA:
            continue
        slug = get_dimension_slug(dim["name"])
        base = slug
        counter = 2
        while slug in seen:
            slug = f"{base}{counter}"
            counter += 1
        seen.add(slug)
        slugs.append(slug)

    # Row/col axis slugs if set (skip categorical)
    if row_dim and row_dim["name"] not in _EXCLUDED_FROM_FORMULA:
        slugs.append(get_dimension_slug(row_dim["name"]))
        seen.add(get_dimension_slug(row_dim["name"]))
    if col_dim and col_dim["name"] not in _EXCLUDED_FROM_FORMULA:
        slugs.append(get_dimension_slug(col_dim["name"]))
        seen.add(get_dimension_slug(col_dim["name"]))

    # All predefined numeric LOS + FE dim slugs (include any not already listed)
    # All predefined numeric LOS + FE dim slugs (include any not already listed)
    for name in _ALL_DIMS:
        if name in _EXCLUDED_FROM_FORMULA:
            continue
        slug = _PREDEFINED_NAME_TO_SLUG.get(name, get_dimension_slug(name))
        if slug not in seen:
            seen.add(slug)
            slugs.append(slug)

    return slugs


# ... Keep your existing _get_dim_slug_options() function as is ...
def render_formula_section():
    st.markdown("### 🧮 Calculation logic")
    st.caption(
        "Build the premium formula by editing the box directly or inserting tokens. "
        "`rater_val` is always the value fetched from the rate table."
    )

    if "formula_editor_input" not in st.session_state:
        st.session_state.formula_editor_input = " ".join(st.session_state.get("formula_tokens", []))

    # Define callbacks for modifying the state before widgets are re-instantiated
    def backspace_callback():
        val = st.session_state.formula_editor_input.strip()
        if " " in val:
            parts = val.split()
            parts.pop()
            st.session_state.formula_editor_input = " ".join(parts)
        elif val:
            st.session_state.formula_editor_input = val[:-1]
        else:
            st.session_state.formula_editor_input = ""

    def clear_callback():
        st.session_state.formula_editor_input = ""

    def insert_var_callback():
        chosen_var = st.session_state.formula_var_select
        current = st.session_state.formula_editor_input.strip()
        if current:
            st.session_state.formula_editor_input = f"{current} {chosen_var}"
        else:
            st.session_state.formula_editor_input = chosen_var

    def insert_op_callback(op: str):
        current = st.session_state.formula_editor_input.strip()
        if current:
            st.session_state.formula_editor_input = f"{current} {op}"
        else:
            st.session_state.formula_editor_input = op

    # ── Formula bar + control buttons ────────────────────────────────────────
    bar_col, bs_col, clr_col = st.columns([11, 1, 1])

    with bar_col:
        # Inputtable formula bar
        st.text_input(
            label="Formula Editor",
            label_visibility="collapsed",
            placeholder="formula will appear here… or type directly",
            key="formula_editor_input",
        )

    with bs_col:
        st.button(
            "⌫",
            key="formula_backspace",
            help="Remove last part/character",
            use_container_width=True,
            on_click=backspace_callback,
        )

    with clr_col:
        st.button(
            "✕",
            key="formula_clear",
            help="Clear all",
            use_container_width=True,
            on_click=clear_callback,
        )

    st.write("")

    # Token palette
    palette_col1, palette_col2 = st.columns(2)

    with palette_col1:
        st.markdown("**Variables**")
        fv_slugs_list = [fv["slug"] for fv in st.session_state.get("formula_variables", [])]
        var_options = ["rater_val"] + _get_dim_slug_options() + fv_slugs_list

        st.selectbox(
            "variable", var_options,
            label_visibility="collapsed",
            key="formula_var_select",
        )
        st.button(
            "Insert variable",
            key="insert_var",
            on_click=insert_var_callback,
        )

    with palette_col2:
        st.markdown("**Operators**")
        op_cols = st.columns(6)
        for idx, op in enumerate(["+", "-", "*", "/", "(", ")"]):
            with op_cols[idx]:
                st.button(
                    op,
                    key=f"op_btn_{op}",
                    on_click=insert_op_callback,
                    args=(op,),
                )

    st.write("")

    # Set formula button with validation
    set_col, status_col = st.columns([2, 2])
    
    with set_col:
        if st.button("Set formula", type="primary", key="set_formula_btn"):
            formula_str = st.session_state.formula_editor_input
            if not formula_str.strip():
                st.session_state["_formula_status"] = ("empty", "No formula entered")
            else:
                dim_slugs = set(_get_dim_slug_options())
                fv_slugs = {fv["slug"] for fv in st.session_state.get("formula_variables", [])}
                allowed = dim_slugs | fv_slugs | {"rater_val"}
                try:
                    validate_formula(formula_str, allowed)
                    st.session_state["_formula_status"] = ("success", "Formula is valid")
                except FormulaError as e:
                    st.session_state["_formula_status"] = ("error", str(e))
            st.rerun()

    with status_col:
        status = st.session_state.get("_formula_status")
        if status:
            status_type, status_msg = status
            if status_type == "success":
                st.success(f"✅ {status_msg}")
            elif status_type == "error":
                st.error(f"❌ {status_msg}")
            else:
                st.warning(f"⚠️ {status_msg}")      

# ── Build result JSON ─────────────────────────────────────────────────────────

def _parse_dim_to_dict(dim_type: str, dim_config: dict) -> dict:
    """Build {slug -> human_label} for a dimension's values."""
    op_map = {
        "<":  "Less than",
        "<=": "Less than or equal to",
        ">":  "Greater than",
        ">=": "Greater than or equal to",
        "=":  "Equal to",
    }
    result = {}

    if dim_type == "Enum":
        for val in dim_config.get("values", []):
            slug = get_value_slug(str(val))
            result[slug] = str(val)

    elif dim_type == "Range":
        lo = dim_config.get("min", 0)
        hi = dim_config.get("max", 100)
        result[f"{lo}_{hi}"] = f"{lo} to {hi}"

    elif dim_type == "Comparison":
        for comp in dim_config.get("comparisons", []):
            op  = comp.get("op", "")
            val = comp.get("val", "")
            op_slug  = operator_to_slug(op)
            val_slug = get_value_slug(val)
            slug  = f"{op_slug}_{val_slug}".strip("_")
            label = f"{op_map.get(op, op)} {val}".strip()
            result[slug] = label

    return result


def _axis_def(dim):
    if dim is None:
        return None
    if dim["type"] == "Range":
        return {
            "name": dim["name"], "type": "Range",
            "min": dim["config"]["min"], "max": dim["config"]["max"],
        }
    if dim["type"] == "Enum":
        return {"name": dim["name"], "type": "Enum", "config": dim["config"].get("values", [])}
    if dim["type"] == "Comparison":
        return {"name": dim["name"], "type": "Comparison", "config": dim["config"].get("comparisons", [])}
    return None


def _build_result() -> dict:
    dimensions_def = {}
    parsed_dims = {}

    outer_dims = [d for d in st.session_state.dims if d.get("role") not in ("row", "col")]
    row_dim = _row_dim()
    col_dim = _col_dim()

    for dim in outer_dims:
        dim_name   = dim["name"]
        dim_type   = dim["type"]
        dim_config = dim["config"]

        dim_slug = get_dimension_slug(dim_name)
        base_slug = dim_slug
        counter = 2
        while dim_slug in parsed_dims:
            dim_slug = f"{base_slug}{counter}"
            counter += 1

        if dim_type == "Enum":
            dimensions_def[dim_name] = {"type": "Enum", "config": dim_config.get("values", [])}
        elif dim_type == "Range":
            dimensions_def[dim_name] = {
                "type": "Range",
                "config": {"min": dim_config.get("min"), "max": dim_config.get("max")},
            }
        elif dim_type == "Comparison":
            dimensions_def[dim_name] = {
                "type": "Comparison",
                "config": dim_config.get("comparisons", []),
            }

        parsed_dims[dim_slug] = _parse_dim_to_dict(dim_type, dim_config)

    row_axis_def = _axis_def(row_dim)
    if row_axis_def is not None:
        row_axis_def["enabled"] = True
    col_axis_def = _axis_def(col_dim)

    # Formula
    formula_str = st.session_state.get("formula_editor_input", "")

    return {
        "definition": {
            "dimensions": dimensions_def,
            "axes": {
                "row": row_axis_def,
                "col": col_axis_def,
            },
        },
        "parsed_dimensions": parsed_dims,
        "calculation": {
            "formula": formula_str,
            "constants": [],
            "formula_variables": [dict(fv) for fv in st.session_state.get("formula_variables", [])],
        },
    }


# ── Section: generate ─────────────────────────────────────────────────────────

def render_generate_section():
    st.divider()

    row_dim = _row_dim()
    col_dim = _col_dim()
    outer_count = len([d for d in st.session_state.dims if d.get("role") not in ("row", "col")])

    if row_dim:
        axis_caption = f"{row_dim['name']} (row)"
    else:
        axis_caption = "no row axis (missing)"
    axis_caption += f" · {col_dim['name']} (column)" if col_dim else " · no column (1D)"
    st.caption(f"{outer_count} sheet dimension(s) · {axis_caption}")

    if st.button("Generate structure & save template", type="primary"):
        has_errors = False
        if not st.session_state.template_name.strip():
            st.error("IC/product name is not provided. Enter a name before generating.", icon="⚠️")
            has_errors = True
        if row_dim is None:
            st.error("Row dimension is not set. Mark a dimension as Row before generating.", icon="⚠️")
            has_errors = True

        if not has_errors:
            result = _build_result()
            st.session_state.last_result = result

            payload = {**result, "name": st.session_state.template_name.strip()}
            tid = st.session_state.last_template_id

            try:
                if tid:
                    resp = requests.put(
                        f"{SERVER_URL}/templates/{tid}",
                        json=payload,
                        timeout=10,
                    )
                else:
                    resp = requests.post(
                        f"{SERVER_URL}/templates",
                        json=payload,
                        timeout=10,
                    )
                if resp.ok:
                    data = resp.json()
                    st.session_state.last_template_id = data["template_id"]
                    st.success(f"Template saved! **{data['name']}** · ID: `{data['template_id']}`")
                else:
                    st.error(f"Server error: {resp.text}")
            except requests.exceptions.ConnectionError:
                st.warning(
                    "Could not reach the Flask server. "
                    "Start it with `python server.py` and try again.",
                    icon="⚠️",
                )

            st.rerun()

    if st.session_state.last_result:
        result = st.session_state.last_result

        st.success("Structure generated!")

        with st.expander("View JSON"):
            st.json(result)

        st.divider()

        with st.spinner("Building Excel workbook…"):
            try:
                excel_data = {**result, "template_id": st.session_state.last_template_id}
                xlsx_bytes = generate_excel(excel_data)

                n_sheets = max(1, 1)
                for vals in result["parsed_dimensions"].values():
                    n_sheets *= len(vals)

                row_axis = result["definition"]["axes"]["row"]
                col_axis = result["definition"]["axes"]["col"]
                n_rows = row_axis["max"] - row_axis["min"] + 1 if row_axis.get("type") == "Range" else None

                if col_axis:
                    n_cols = col_axis["max"] - col_axis["min"] + 1 if col_axis.get("type") == "Range" else None
                    grid_caption = f"**{n_rows or '?'}** rows × **{n_cols or '?'}** columns each"
                else:
                    grid_caption = f"**{n_rows or '?'}**-row 1D list each ({row_axis['name']} only)"

                st.info(f"📊 **{n_sheets}** data sheet(s) · {grid_caption}", icon="📋")

                st.download_button(
                    label="⬇️ Download Excel (.xlsx)",
                    data=xlsx_bytes,
                    file_name=f"{st.session_state.template_name}_rater_template.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )
            except Exception as e:
                st.error(f"Excel generation failed: {e}")


# ── Tab 1: Template Builder ───────────────────────────────────────────────────

def render_tab_template_builder():
    if st.session_state.last_template_id and not st.session_state.template_name:
        detail = _fetch_template_detail(st.session_state.last_template_id)
        if detail:
            st.session_state.template_name = detail.get("name", "")
    
    p_name, p_code = st.columns(2)

    with p_name:
        st.text_input(
            'Product name',
            key="template_name",
            placeholder="e.g. Motor Comprehensive",
        )

    with p_code:
        st.text_input(
            'Product Code',
            key="product_code",
            placeholder="e.g. ICICI001",
        )

    # ── Top panel: Dimensions (LOS + FE) ────────────────────────────────────
    render_dims_section()

    st.divider()
    render_axis_summary_section()
    st.divider()
    render_formula_section()
    render_generate_section()


# ── Tab 2: Upload & Flatten ───────────────────────────────────────────────────

def render_tab_upload():
    st.markdown("### Upload & Flatten")
    st.caption(
        "Upload a filled-in Excel file to store its rate values in the database. "
        "The template is identified automatically from the workbook's hidden _meta sheet."
    )

    uploaded_file = st.file_uploader(
        "Upload filled Excel (.xlsx)", type=["xlsx"], key="upload_file"
    )

    if st.button("Upload & Flatten", type="primary", key="do_flatten"):
        if not uploaded_file:
            st.error("Upload an Excel file first.")
        else:
            with st.spinner("Flattening…"):
                try:
                    resp = requests.post(
                        f"{SERVER_URL}/flatten",
                        files={"file": (uploaded_file.name, uploaded_file.getvalue(),
                                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
                        timeout=120,
                    )
                    if resp.ok:
                        data = resp.json()
                        st.success(
                            f"✅ **{data['cells_written']}** cells written · "
                            f"**{data['blank_cells_count']}** blank · "
                            f"**{data['completion_pct']}%** complete"
                        )
                        if data.get("blank_cells"):
                            with st.expander(f"Blank cells ({len(data['blank_cells'])} shown)"):
                                st.json(data["blank_cells"])
                    else:
                        st.error(f"Server error: {resp.text}")
                except requests.exceptions.ConnectionError:
                    st.warning("Could not reach the Flask server.", icon="⚠️")


# ── Tab 3: Calculate Premium ──────────────────────────────────────────────────

def _fetch_templates() -> list[dict]:
    try:
        resp = requests.get(f"{SERVER_URL}/templates", timeout=10)
        if resp.ok:
            return resp.json()
    except requests.exceptions.ConnectionError:
        pass
    return []


def _fetch_template_detail(template_id: str) -> dict | None:
    try:
        resp = requests.get(f"{SERVER_URL}/templates/{template_id}", timeout=10)
        if resp.ok:
            return resp.json()
    except requests.exceptions.ConnectionError:
        pass
    return None


def _render_calc_inputs(full_def: dict) -> dict:
    """Build input widgets from a template definition and return the inputs dict."""
    definition = full_def.get("definition", {})
    dimensions = definition.get("dimensions", {})
    axes = definition.get("axes", {})
    row_axis = axes.get("row")
    col_axis = axes.get("col")
    inputs = {}

    from datetime import date

    for dim_name, dim_def in dimensions.items():
        dim_slug = get_dimension_slug(dim_name)
        dim_type = dim_def["type"]

        if dim_slug == "age":
            dob_val = st.date_input(
                "Date of Birth (DOB)",
                value=date(1999, 1, 1),
                key=f"calc_input_{dim_slug}_dob"
            )
            inputs["dob"] = dob_val.strftime("%d-%m-%Y")
            continue

        if dim_type == "Enum":
            options = dim_def.get("config", [])
            val = st.selectbox(dim_name, options, key=f"calc_input_{dim_slug}")
        elif dim_type == "Range":
            lo = dim_def["config"]["min"]
            hi = dim_def["config"]["max"]
            val = st.number_input(dim_name, min_value=int(lo), max_value=int(hi),
                                  value=int(lo), step=1, key=f"calc_input_{dim_slug}")
        elif dim_type == "Comparison":
            val = st.number_input(dim_name, value=0.0, step=1.0, key=f"calc_input_{dim_slug}")
        else:
            val = st.text_input(dim_name, key=f"calc_input_{dim_slug}")

        inputs[dim_slug] = val

    if row_axis:
        row_slug = get_dimension_slug(row_axis["name"])
        if row_slug == "age":
            dob_val = st.date_input(
                "DOB (row)",
                value=date(1999, 1, 1),
                key="calc_row_val_dob",
            )
            inputs["dob"] = dob_val.strftime("%d-%m-%Y")
        else:
            rtype = row_axis.get("type", "Range")
            if rtype == "Range":
                row_val = st.number_input(
                    row_axis["name"] + " (row)",
                    min_value=int(row_axis["min"]), max_value=int(row_axis["max"]),
                    value=int(row_axis["min"]), step=1, key="calc_row_val",
                )
            elif rtype == "Enum":
                row_val = st.selectbox(row_axis["name"] + " (row)", row_axis.get("config", []), key="calc_row_val")
            else:
                row_val = st.text_input(row_axis["name"] + " (row)", key="calc_row_val")
            inputs[row_slug] = row_val

    if col_axis:
        col_slug = get_dimension_slug(col_axis["name"])
        if col_slug == "age":
            dob_val = st.date_input(
                "DOB (col)",
                value=date(1999, 1, 1),
                key="calc_col_val_dob",
            )
            inputs["dob"] = dob_val.strftime("%d-%m-%Y")
        else:
            ctype = col_axis.get("type", "Range")
            if ctype == "Range":
                col_val = st.number_input(
                    col_axis["name"] + " (col)",
                    min_value=int(col_axis["min"]), max_value=int(col_axis["max"]),
                    value=int(col_axis["min"]), step=1, key="calc_col_val",
                )
            elif ctype == "Enum":
                col_val = st.selectbox(col_axis["name"] + " (col)", col_axis.get("config", []), key="calc_col_val")
            else:
                col_val = st.text_input(col_axis["name"] + " (col)", key="calc_col_val")
            inputs[col_slug] = col_val
    # Formula variables
    calculation = full_def.get("calculation", {})
    for fv in calculation.get("formula_variables", []):
        fv_slug = fv["slug"]
        fv_default = fv.get("default_value", "")
        try:
            default_num = int(fv_default) if fv_default not in (None, "") else 0.0
        except ValueError:
            default_num = 1
        val = st.number_input(
            fv["name"],
            value=default_num,
            step=1,
            key=f"calc_input_fv_{fv_slug}",
        )
        inputs[fv_slug] = val
    
    # Predefined formula variables referenced in the formula but not covered
    # by any workbook dimension / row / col axis / formula_variables entry.
    # These still need to be collected from the user and sent to /calculate.
    formula_str = full_def.get("calculation", {}).get("formula", "") or ""
    for slug, display_name in _PREDEFINED_SLUG_TO_NAME.items():
        if slug in inputs:
            continue  # already rendered above
        if not re.search(rf"\b{re.escape(slug)}\b", formula_str):
            continue  # not used in this template's formula
        val = st.number_input(
            display_name,
            value=0,
            step=1,
            key=f"calc_input_predef_{slug}",
        )
        inputs[slug] = val

    return inputs

# ── Calculate: concurrent API calls ───────────────────────────────────────────

def _call_template_calculate(tid: str, inputs: dict) -> dict:
    """Call POST /templates/<id>/calculate. Returns a normalized result dict."""
    try:
        resp = requests.post(
            f"{SERVER_URL}/templates/{tid}/calculate",
            json={"inputs": inputs},
            timeout=10,
        )
        if resp.ok:
            return {"ok": True, "data": resp.json()}
        try:
            err_data = resp.json()
        except ValueError:
            err_data = {"error": resp.text}
        return {
            "ok": False,
            "status_code": resp.status_code,
            "data": err_data,
            "error": err_data.get("error", resp.text),
        }
    except requests.exceptions.ConnectionError:
        return {"ok": False, "status_code": None, "data": None,
                "error": "Could not reach the Flask server."}
    except requests.exceptions.RequestException as e:
        return {"ok": False, "status_code": None, "data": None, "error": str(e)}


def _call_api_get_premium(api_payload: dict) -> dict:
    """Call POST /api/get-premium using the payload already built by the frontend."""
    try:
        resp = requests.post(
            f"{SERVER_URL}/api/get-premium",
            json=api_payload,
            timeout=10,
        )
        if resp.ok:
            return {"ok": True, "data": resp.json()}
        try:
            err_data = resp.json()
        except ValueError:
            err_data = {"error": resp.text}
        return {
            "ok": False,
            "status_code": resp.status_code,
            "data": err_data,
            "error": err_data.get("error", resp.text),
        }
    except requests.exceptions.ConnectionError:
        return {"ok": False, "status_code": None, "data": None,
                "error": "Could not reach the API server."}
    except requests.exceptions.RequestException as e:
        return {"ok": False, "status_code": None, "data": None, "error": str(e)}


def _run_calculate_calls(tid: str, inputs: dict, api_payload: dict) -> tuple[dict, dict]:
    """Fire both calls concurrently and wait for both to finish."""
    with ThreadPoolExecutor(max_workers=2) as executor:
        future_template = executor.submit(_call_template_calculate, tid, inputs)
        future_api = executor.submit(_call_api_get_premium, api_payload)
        template_result = future_template.result()
        api_result = future_api.result()
    return template_result, api_result


# ── Calculate: results dashboard ──────────────────────────────────────────────

def _fmt_currency(val) -> str:
    try:
        return f"₹ {float(val):,.2f}"
    except (TypeError, ValueError):
        return "—"


def _render_calculate_results(template_result: dict, api_result: dict) -> None:
    st.markdown("#### Calculation Results")

    template_data = template_result.get("data") if template_result.get("ok") else None
    api_data = api_result.get("data") if api_result.get("ok") else None

    template_base = template_data.get("base_premium") if template_data else None
    template_total = (template_base * 1.18) if isinstance(template_base, (int, float)) else None

    api_base = None
    api_total = None
    if api_data:
        api_base = api_data.get("base_premium")
        api_total = api_data.get("total_premium")
        if api_total is None and isinstance(api_base, (int, float)):
            api_total = api_base * 1.18

    col_excel, col_api = st.columns(2)

    with col_excel:
        with st.container(border=True):
            st.markdown("**Excel / Template**")
            if template_result.get("ok"):
                m1, m2 = st.columns(2)
                with m1:
                    st.metric("Base Premium", _fmt_currency(template_base))
                with m2:
                    st.metric("Total Premium", _fmt_currency(template_total))
            elif template_result.get("status_code") == 404:
                st.error("No rate found for these inputs.")
            else:
                st.error(f"Failed: {template_result.get('error', 'Unknown error')}")

    with col_api:
        with st.container(border=True):
            st.markdown("**API Premium**")
            if api_result.get("ok"):
                m1, m2 = st.columns(2)
                with m1:
                    st.metric("Base Premium", _fmt_currency(api_base))
                with m2:
                    st.metric("Total Premium", _fmt_currency(api_total))
            else:
                st.error(f"Failed: {api_result.get('error', 'Unknown error')}")

    if not template_result.get("ok") and not api_result.get("ok"):
        st.warning("Both calls failed — see errors above.", icon="⚠️")

    # # ── Lookup details (from template call only) ───────────────────────────
    # if template_data:
    #     st.markdown("#### Lookup Details")
    #     d1, d2 = st.columns(2)
    #     with d1:
    #         st.caption("Lookup Key")
    #         st.code(template_data.get("lookup_key", "—"))
    #         st.caption("Rater Value")
    #         st.write(template_data.get("rater_val", "—"))
    #     with d2:
    #         st.caption("Resolved Buckets")
    #         st.json(template_data.get("resolved_buckets", {}))
    #     if template_data.get("formula"):
    #         st.caption("Formula Used")
    #         st.code(template_data["formula"])

    # # ── Raw debug ────────────────────────────────────────────────────────────
    # with st.expander("▼ Raw Responses"):
    #     raw_col1, raw_col2 = st.columns(2)
    #     with raw_col1:
    #         st.caption("`/templates/<id>/calculate`")
    #         st.json(template_result.get("data") or {"error": template_result.get("error")})
    #     with raw_col2:
    #         st.caption("`/api/get-premium`")
    #         st.json(api_result.get("data") or {"error": api_result.get("error")})


def render_tab_calculate():
    st.markdown("### Calculate Premium")
    st.caption("Select a saved product template and enter values for each dimension.")

    # ── Flow Type selector (payload generation only — does not affect lookup) ──
    flow_type_label = st.selectbox(
        "Flow Type",
        ["Save Loan", "Save Lead", "Create Lead"],
        index=0,
        key="calc_flow_type_select",
    )
    flow_type_map = {
        "Save Loan": "save_loan",
        "Save Lead": "save_lead",
        "Create Lead": "create_lead",
    }
    flow_type = flow_type_map[flow_type_label]

    templates = _fetch_templates()
    if not templates:
        st.warning(
            "No saved templates found. Generate and save a template in Tab 1 first, "
            "or check that the Flask server is running.",
            icon="⚠️",
        )
        return

    name_to_id = {t["name"]: t["id"] for t in templates}
    template_names = list(name_to_id.keys())

    default_idx = 0
    if st.session_state.last_template_id:
        for i, t in enumerate(templates):
            if t["id"] == st.session_state.last_template_id:
                default_idx = i
                break

    selected_name = st.selectbox(
        "Product",
        template_names,
        index=default_idx,
        key="calc_product_select",
    )
    tid = name_to_id[selected_name]

    full_def = _fetch_template_detail(tid)
    if not full_def:
        st.error("Could not load template definition from the server.")
        return

    st.caption(f"Template ID: `{tid}`")

    st.markdown("#### Inputs")
    inputs = _render_calc_inputs(full_def)

    # ── API payload ──────────────
    api_payload = {
        "flow_type": flow_type,
        "dimensions": inputs,
    }
    # with st.expander("API Payload (Debug)", expanded=True):
    #     st.json(api_payload)

    if st.button("Calculate", type="primary", key="do_calculate"):
        with st.spinner("Calculating…"):
            template_result, api_result = _run_calculate_calls(tid, inputs, api_payload)
        _render_calculate_results(template_result, api_result)

# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    init_state()

    st.title("Offline rater automation")
    st.write("")

    tab1, tab2, tab3 = st.tabs(["Template Builder", "Upload & Flatten", "Calculate Premium"])

    with tab1:
        render_tab_template_builder()

    with tab2:
        render_tab_upload()

    with tab3:
        render_tab_calculate()


if __name__ == "__main__":
    main()
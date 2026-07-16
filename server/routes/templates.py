# pyrefly: ignore [missing-import]
from flask import Blueprint, jsonify, request

import json 
from frontend.formula_eval import evaluate_formula,FormulaError
from datetime import timezone, datetime
import uuid
from server.db import get_db
from server.helpers import _parse_template_body, _build_lookup_key

# Setting a url_prefix means we don't have to keep repeating '/templates'
templates_bp = Blueprint('templates', __name__, url_prefix='/templates')

# ── POST /templates ───────────────────────────────────────────────────────────
@templates_bp.post("/")
def create_template():
    body = request.get_json(force=True)
    if not body:
        return jsonify({"error": "Empty request body."}), 400

    full_def, name, content_hash = _parse_template_body(body)
    if full_def is None:
        return jsonify({"error": content_hash}), 400

    template_id = str(uuid.uuid4())
    now = datetime.now(timezone.utc).isoformat()

    with get_db() as conn:
        conn.execute(
            "INSERT INTO rate_templates (id, name, definition_json, content_hash, created_at) "
            "VALUES (?, ?, ?, ?, ?)",
            (template_id, name, json.dumps(full_def), content_hash, now),
        )

    return jsonify({"template_id": template_id, "name": name}), 201


# ── PUT /templates/<id> ─────────────────────────────────────────────────────
@templates_bp.put("/<template_id>")
def update_template(template_id: str):
    body = request.get_json(force=True)
    if not body:
        return jsonify({"error": "Empty request body."}), 400

    full_def, name, content_hash = _parse_template_body(body)
    if full_def is None:
        return jsonify({"error": content_hash}), 400

    with get_db() as conn:
        row = conn.execute(
            "SELECT id FROM rate_templates WHERE id = ?", (template_id,)
        ).fetchone()
        if not row:
            return jsonify({"error": f"Template '{template_id}' not found."}), 404

        conn.execute(
            "UPDATE rate_templates SET name = ?, definition_json = ?, content_hash = ? WHERE id = ?",
            (name, json.dumps(full_def), content_hash, template_id),
        )

    return jsonify({"template_id": template_id, "name": name}), 200


# ── GET /templates ────────────────────────────────────────────────────────────

@templates_bp.get("")
def list_templates():
    with get_db() as conn:
        rows = conn.execute(
            "SELECT id, name, created_at FROM rate_templates ORDER BY created_at DESC"
        ).fetchall()
    return jsonify([dict(r) for r in rows])


# ── GET /templates/<id> ───────────────────────────────────────────────────────

@templates_bp.get("/<template_id>")
def get_template(template_id: str):
    with get_db() as conn:
        row = conn.execute(
            "SELECT id, name, definition_json, created_at FROM rate_templates WHERE id = ?",
            (template_id,),
        ).fetchone()
    if not row:
        return jsonify({"error": f"Template '{template_id}' not found."}), 404

    full_def = json.loads(row["definition_json"])
    return jsonify({
        "id": row["id"],
        "name": row["name"],
        "created_at": row["created_at"],
        **full_def,
    })


# ── POST /templates/<id>/calculate ───────────────────────────────────────────

@templates_bp.post("/<template_id>/calculate")
def calculate(template_id: str):
    with get_db() as conn:
        row = conn.execute(
            "SELECT definition_json FROM rate_templates WHERE id = ?", (template_id,)
        ).fetchone()
    if not row:
        return jsonify({"error": f"Template '{template_id}' not found."}), 404

    full_def = json.loads(row["definition_json"])
    calculation = full_def.get("calculation", {})
    formula = calculation.get("formula", "").strip()
    constants_list = calculation.get("constants", [])
    constants = {c["name"]: float(c["value"]) for c in constants_list}

    body = request.get_json(force=True)
    inputs = body.get("inputs", {})

    if "dob" in inputs and "age" not in inputs:
        try:
            from server.helpers import _dob_to_age
            inputs["age"] = _dob_to_age(inputs["dob"])
        except ValueError as e:
            return jsonify({"error": str(e)}), 400

    try:
        lookup_key, resolved_buckets = _build_lookup_key(full_def, full_def.get("parsed_dimensions", {}), inputs)
    except ValueError as e:
        return jsonify({"error": str(e)}), 400

    with get_db() as conn:
        val_row = conn.execute(
            "SELECT value FROM rate_values WHERE template_id = ? AND lookup_key = ?",
            (template_id, lookup_key),
        ).fetchone()

    if not val_row:
        return jsonify({
            "error": "No rate value found for the given inputs.",
            "lookup_key": lookup_key,
            "resolved_buckets": resolved_buckets,
        }), 404

    rater_val = float(val_row["value"])

    if not formula:
        return jsonify({
            "rater_val": rater_val,
            "base_premium": rater_val,
            "lookup_key": lookup_key,
            "resolved_buckets": resolved_buckets,
            "note": "No formula defined; base_premium = rater_val.",
        })

    eval_context = {
        "rater_val": rater_val,
        **{k: float(v) for k, v in inputs.items() if isinstance(v, (int, float, str)) and str(v).replace(".", "").lstrip("-").isdigit()},
        **constants,
    }

    try:
        base_premium = evaluate_formula(formula, eval_context)
    except FormulaError as e:
        return jsonify({"error": f"Formula evaluation failed: {e}"}), 500

    return jsonify({
        "rater_val": rater_val,
        "base_premium": base_premium,
        "lookup_key": lookup_key,
        "resolved_buckets": resolved_buckets,
        "formula": formula,
    })


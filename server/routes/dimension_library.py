# ==========================================================
# Imports
# ==========================================================
# pyrefly: ignore [missing-import]
from flask import Blueprint, jsonify, request

import json
import uuid
from datetime import timezone, datetime

from server.db import get_db
from server.helpers import (
    _parse_dimension_library_body,
    _dimension_in_use,
    _validate_dimension_config,
)

# ==========================================================
# Blueprint Configuration
# ==========================================================
dimension_library_bp = Blueprint(
    "dimension_library", __name__, url_prefix="/dimension-library"
)


def _row_to_dict(row) -> dict:
    return {
        "id": row["id"],
        "name": row["name"],
        "type": row["type"],
        "config": json.loads(row["config_json"]),
        "is_system": bool(row["is_system"]),
        "created_at": row["created_at"],
        "updated_at": row["updated_at"],
    }


# ==========================================================
# API Routes
# ==========================================================

# ── GET /dimension-library ────────────────────────────────────────────────────
# Used by the Template Builder's "Choose Existing Dimension" dropdown and by
# the Dimension Library admin tab's list view.
@dimension_library_bp.get("")
def list_dimensions():
    """Return every dimension in the library, system dimensions first."""
    with get_db() as conn:
        rows = conn.execute(
            "SELECT * FROM dimension_library ORDER BY is_system DESC, name ASC"
        ).fetchall()
    return jsonify([_row_to_dict(r) for r in rows])


# ── GET /dimension-library/<id> ───────────────────────────────────────────────
@dimension_library_bp.get("/<dim_id>")
def get_dimension(dim_id: str):
    """Fetch a single library dimension by ID."""
    with get_db() as conn:
        row = conn.execute(
            "SELECT * FROM dimension_library WHERE id = ?", (dim_id,)
        ).fetchone()
    if not row:
        return jsonify({"error": f"Dimension '{dim_id}' not found."}), 404
    return jsonify(_row_to_dict(row))


# ── POST /dimension-library ───────────────────────────────────────────────────
# Creates an admin-managed dimension. is_system is always forced to 0 here —
# a client can never create or mark a dimension as a system dimension.
@dimension_library_bp.post("")
def create_dimension():
    """Create a new (always custom) library dimension."""
    body = request.get_json(force=True)
    if not body:
        return jsonify({"error": "Empty request body."}), 400

    parsed, name, err = _parse_dimension_library_body(body)
    if parsed is None:
        return jsonify({"error": err}), 400

    with get_db() as conn:
        existing = conn.execute(
            "SELECT id FROM dimension_library WHERE name = ? COLLATE NOCASE",
            (name,),
        ).fetchone()
        if existing:
            return jsonify({"error": f"A dimension named '{name}' already exists."}), 409

        dim_id = str(uuid.uuid4())
        now = datetime.now(timezone.utc).isoformat()
        conn.execute(
            "INSERT INTO dimension_library "
            "(id, name, type, config_json, is_system, created_at, updated_at) "
            "VALUES (?, ?, ?, ?, 0, ?, ?)",
            (dim_id, parsed["name"], parsed["type"], json.dumps(parsed["config"]), now, now),
        )

    return jsonify({"id": dim_id, "name": parsed["name"]}), 201


# ── PUT /dimension-library/<id> ───────────────────────────────────────────────
# Admin-created (is_system=0) dimensions: fully editable, as before.
#
# System dimensions (is_system=1): name, type, and is_system remain
# permanently locked. Only the configuration (Enum values / Range min-max /
# Comparison rules) may be updated. Any 'name'/'type' the client sends is
# simply never written for a system row — enforced here regardless of
# whether the frontend disables those fields.
@dimension_library_bp.put("/<dim_id>")
def update_dimension(dim_id: str):
    """Update a library dimension. System rows: config only."""
    body = request.get_json(force=True)
    if not body:
        return jsonify({"error": "Empty request body."}), 400

    with get_db() as conn:
        row = conn.execute(
            "SELECT * FROM dimension_library WHERE id = ?", (dim_id,)
        ).fetchone()
        if not row:
            return jsonify({"error": f"Dimension '{dim_id}' not found."}), 404

        if row["is_system"]:
            config = body.get("config")
            if not isinstance(config, dict):
                return jsonify({"error": "'config' is required and must be an object."}), 400

            err = _validate_dimension_config(row["type"], config)
            if err:
                return jsonify({"error": err}), 400

            now = datetime.now(timezone.utc).isoformat()
            conn.execute(
                "UPDATE dimension_library SET config_json = ?, updated_at = ? WHERE id = ?",
                (json.dumps(config), now, dim_id),
            )
            return jsonify({"id": dim_id, "name": row["name"]}), 200

        parsed, name, err = _parse_dimension_library_body(body)
        if parsed is None:
            return jsonify({"error": err}), 400

        conflict = conn.execute(
            "SELECT id FROM dimension_library WHERE name = ? COLLATE NOCASE AND id != ?",
            (name, dim_id),
        ).fetchone()
        if conflict:
            return jsonify({"error": f"A dimension named '{name}' already exists."}), 409

        now = datetime.now(timezone.utc).isoformat()
        conn.execute(
            "UPDATE dimension_library SET name = ?, type = ?, config_json = ?, updated_at = ? "
            "WHERE id = ?",
            (parsed["name"], parsed["type"], json.dumps(parsed["config"]), now, dim_id),
        )

    return jsonify({"id": dim_id, "name": parsed["name"]}), 200


# ── DELETE /dimension-library/<id> ────────────────────────────────────────────
# System dimensions can never be deleted. Admin-created dimensions can be
# deleted; if they appear to be referenced by a saved template, the caller
# must pass ?force=true to proceed anyway.
@dimension_library_bp.delete("/<dim_id>")
def delete_dimension(dim_id: str):
    """Delete an admin-created library dimension."""
    force = request.args.get("force", "false").lower() == "true"

    with get_db() as conn:
        row = conn.execute(
            "SELECT * FROM dimension_library WHERE id = ?", (dim_id,)
        ).fetchone()
        if not row:
            return jsonify({"error": f"Dimension '{dim_id}' not found."}), 404
        if row["is_system"]:
            return jsonify({"error": "System dimensions are immutable and cannot be deleted."}), 403

        if not force:
            usage_count = _dimension_in_use(conn, row["name"])
            if usage_count > 0:
                return jsonify({
                    "error": f"Dimension '{row['name']}' appears to be used in "
                             f"{usage_count} saved template(s). Pass force=true to delete anyway.",
                    "usage_count": usage_count,
                }), 409

        conn.execute("DELETE FROM dimension_library WHERE id = ?", (dim_id,))

    return jsonify({"id": dim_id, "deleted": True}), 200
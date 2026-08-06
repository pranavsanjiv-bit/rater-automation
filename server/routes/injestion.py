# ==========================================================
# Imports
# ==========================================================
import os
import json
import copy
from datetime import datetime as dt
# pyrefly: ignore [missing-import]
from flask import Blueprint, request, jsonify
from server.helpers import _age_to_dob
from server.templates import get_payload_template
from server.api_service import APIService


# ==========================================================
# Blueprint Configuration
# ==========================================================

injestion_bp = Blueprint("injestion", __name__, url_prefix="/api")

# ---------------------------------------------------------------------------
# Shared APIService instance (created once at module load time)
# ---------------------------------------------------------------------------
api_service = APIService()

# ---------------------------------------------------------------------------
# Sample request payload (no master_payload required any more):
#
# {
#     "flow_type"   : "save_loan",          # "create_lead" | "save_loan"
#     "product_code": "BAJAJ_LIFE_GCPP",
#     "dimensions"  : {
#         "dob"              : "dd-mm-yyyy",
#         "property_pincode" : 577201,
#         "gender"           : "Male",
#         "loan_amount"      : 2000000,
#         "loan_tenure"      : 20,
#         "loan_type"        : "HL",
#         "coverage_type"    : "Reducing"
#     }
# }
# ---------------------------------------------------------------------------

# ==========================================================
# Helper Functions
# ==========================================================
def _load_injection_paths() -> dict:
    """Loads the injection path configuration file."""
    current_dir = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(current_dir, "..", "injection_path.json")
    with open(config_path, "r") as f:
        return json.load(f)


def _inject_value(target, path: str, val):
    """Inject a value into a nested payload using a dotted path."""
    keys = path.split('.')
    current = target

    for key in keys[:-1]:
        actual_key = int(key) if key.isdigit() else key
        current = current[actual_key]

    last_key = int(keys[-1]) if keys[-1].isdigit() else keys[-1]
    current[last_key] = val


def _process_injection(dims: dict, flow_type: str, payload: dict) -> dict:
    """
    Resolves dimension values, looks up the injection paths for the given
    flow_type, and injects each value into payload in-place.
    Returns the modified payload.
    """
    injection_paths = _load_injection_paths()

    # Resolve age → dob; pass everything else through as-is
    resolved_dims = {}
    for k, v in dims.items():
        resolved_dims[k] = v

    # Fallback: inject sum_insured/tenure if loan_amount/loan_tenure are missing or empty
    if resolved_dims.get("loan_amount") in (None, "") and "sum_insured" in resolved_dims:
        resolved_dims["loan_amount"] = resolved_dims["sum_insured"]
    if resolved_dims.get("loan_tenure") in (None, "") and "tenure" in resolved_dims:
        resolved_dims["loan_tenure"] = resolved_dims["tenure"]

    if resolved_dims.get("dob") in (None, "") and resolved_dims.get("age") not in (None, ""):
        try:
            resolved_dims["dob"] = _age_to_dob(int(float(resolved_dims["age"])))
        except (TypeError, ValueError):
            pass

    if "dob" in resolved_dims and resolved_dims["dob"]:
        dob_val = str(resolved_dims["dob"]).replace("-", "/") if flow_type == "save_lead" else str(resolved_dims["dob"]).replace("/", "-")
        resolved_dims["dob"] = dob_val

    # Inject each resolved dimension into the payload
    for key, val in resolved_dims.items():
        if key in injection_paths:
            path_mapping = injection_paths[key]
            if flow_type in path_mapping:
                _inject_value(payload, path_mapping[flow_type], val)

    if flow_type == "save_lead" and isinstance(payload.get("proposer"), dict):
        if "dob" in resolved_dims and resolved_dims["dob"]:
            payload["proposer"]["dob"] = resolved_dims["dob"]
        if "gender" in resolved_dims and resolved_dims["gender"]:
            payload["proposer"]["gender"] = resolved_dims["gender"]

    return payload

def _apply_lead_details(payload: dict, lead_details: dict) -> dict:
    """Apply incoming lead details into the lead payload template."""
    if not lead_details:
        return payload

    name = (lead_details.get("name") or "").strip()
    mobile = (lead_details.get("mobile") or "").strip()
    email = (lead_details.get("email") or "").strip()

    if name:
        if "proposer" in payload and isinstance(payload["proposer"], dict):
            payload["proposer"]["first_name"] = name
        if "borrowers" in payload and payload["borrowers"]:
            payload["borrowers"][0]["first_name"] = name

    if mobile:
        if "proposer" in payload and isinstance(payload["proposer"], dict):
            payload["proposer"]["phone_number"] = mobile
        if "borrowers" in payload and payload["borrowers"]:
            payload["borrowers"][0]["phone_number"] = mobile

    if email:
        if "proposer" in payload and isinstance(payload["proposer"], dict):
            payload["proposer"]["email"] = email
        if "borrowers" in payload and payload["borrowers"]:
            payload["borrowers"][0]["email"] = email

    return payload


def _extract_lead_payload(data: dict) -> dict | None:
    """Extract a direct partner-style lead payload if one is provided."""
    if not isinstance(data, dict):
        return None

    if "payload" in data:
        return data["payload"]
    if "lead_payload" in data:
        return data["lead_payload"]

    direct_keys = {"loan", "borrowers", "proposer", "property", "assets", "loan_amount", "loan_type"}
    if any(key in data for key in direct_keys):
        return data

    return None


def _build_insured_person(person: dict) -> dict:
    """Build the insured person payload from borrower/proposer details."""
    return {
        "title": person.get("title"),
        "first_name": person.get("first_name"),
        "last_name": person.get("last_name"),
        "email": person.get("email"),
        "phone_number": person.get("phone_number"),
        "gender": person.get("gender"),
        "dob": person.get("dob"),
        "pan": person.get("pan"),
        "occupation": person.get("occupation"),
        "annual_income": person.get("annual_income"),
        "is_primary_borrower": person.get("is_primary_borrower", True),
        "external_user_id": person.get("partner_uid") or person.get("user_id"),
        "address": person.get("address", {}),
    }


def _ensure_insured_details(payload: dict) -> dict:
    """Ensure the payload contains an insured section."""
    if "insured" not in payload or not payload["insured"]:
        source = None
        if "borrowers" in payload and payload["borrowers"]:
            source = payload["borrowers"][0]
        elif "proposer" in payload and isinstance(payload["proposer"], dict):
            source = payload["proposer"]

        if source:
            payload["insured"] = [_build_insured_person(source)]

    return payload


# ==========================================================
# API Routes
# ==========================================================

class _PartnerResponseError(Exception):
    """Raised when the partner API returns an empty response or no usable
    token. Both save_lead and create_lead routes have always mapped this
    to a 502 — this exception just lets the extracted functions signal
    that same condition to a caller that isn't necessarily an HTTP route.
    raw_response is preserved so the route can include it in the error
    body exactly as it always did."""
    def __init__(self, message: str, raw_response: dict = None):
        super().__init__(message)
        self.raw_response = raw_response


def _run_save_lead_calculation(dims: dict, lead_details: dict = None, direct_payload: dict = None) -> dict:
    """Core Save Lead calculation, extracted verbatim from /api/save-lead's
    route body so Bulk Premium can call it directly (no HTTP) per row —
    same payload construction (get_payload_template, _process_injection,
    _apply_lead_details, _ensure_insured_details — all untouched), same
    tenure normalization, same Partner API call this route has always made.
    direct_payload mirrors the route's optional full-payload-override input
    (data["payload"]/data["lead_payload"]/direct keys) — Bulk never uses it,
    but it's preserved since it's genuinely part of the existing behaviour.
    Raises FileNotFoundError (missing payload template, mirrors 404),
    _PartnerResponseError (mirrors 502), or lets any other exception
    propagate (mirrors the route's generic 500 handling)."""
    if direct_payload is not None:
        payload = copy.deepcopy(direct_payload)
    else:
        payload = get_payload_template("save_lead")

    payload = _process_injection(dims, "save_lead", payload)
    payload = _apply_lead_details(payload, lead_details)
    payload = _ensure_insured_details(payload)

    # Normalize tenure to months to match create_lead schema
    raw_tenure = payload.get("loan_tenure")
    raw_tenure_unit = payload.get("loan_tenure_unit", "years")
    if isinstance(raw_tenure_unit, str) and raw_tenure_unit.strip().lower() in {"year", "years", "y"}:
        try:
            raw_tenure = int(float(raw_tenure) * 12) if raw_tenure is not None else raw_tenure
        except (TypeError, ValueError):
            pass

    payload["loan_tenure"] = raw_tenure
    payload["loan_tenure_unit"] = "months"
    payload["tenure_unit"] = "months"

    response = api_service.save_lead(payload, os.getenv("PARTNER_CODE"))
    if not response:
        raise _PartnerResponseError("Save lead API returned empty response.")

    token = response.get("session_token") or response.get("lead_id") or response.get("transaction_id") or response.get("loan_id")
    if not token:
        raise _PartnerResponseError("Save lead API did not return a session token.", raw_response=response)

    response["session_token"] = token
    response["payload"] = payload
    return response


@injestion_bp.post("/save-lead")
def save_lead():
    data = request.get_json() or {}
    lead_details = data.get("lead_details")
    dims = data.get("dimensions", {})

    try:
        direct_payload = _extract_lead_payload(data)
        response = _run_save_lead_calculation(dims, lead_details, direct_payload)
    except _PartnerResponseError as e:
        body = {"error": str(e)}
        if e.raw_response is not None:
            body["raw_response"] = e.raw_response
        return jsonify(body), 502
    except FileNotFoundError as e:
        return jsonify({"error": str(e)}), 404
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    return jsonify(response), 201


def _run_create_lead_calculation(dims: dict, lead_details: dict = None, direct_payload: dict = None) -> dict:
    """Core Create Lead calculation, extracted verbatim from
    /api/create-lead's route body — same shape as
    _run_save_lead_calculation but without the tenure-normalization step
    (create_lead's own payload template already uses months), calling
    api_service.create_lead() instead of api_service.save_lead()."""
    if direct_payload is not None:
        payload = copy.deepcopy(direct_payload)
    else:
        payload = get_payload_template("create_lead")

    payload = _process_injection(dims, "create_lead", payload)
    payload = _apply_lead_details(payload, lead_details)
    payload = _ensure_insured_details(payload)

    response = api_service.create_lead(payload, os.getenv("PARTNER_CODE"))
    if not response:
        raise _PartnerResponseError("Create lead API returned empty response.")

    token = response.get("session_token") or response.get("lead_id") or response.get("transaction_id") or response.get("loan_id")
    if not token:
        raise _PartnerResponseError("Create lead API did not return a session token.", raw_response=response)

    response["session_token"] = token
    response["payload"] = payload
    return response


@injestion_bp.post("/create-lead")
def create_lead():
    data = request.get_json() or {}
    lead_details = data.get("lead_details")
    dims = data.get("dimensions", {})

    try:
        direct_payload = _extract_lead_payload(data)
        response = _run_create_lead_calculation(dims, lead_details, direct_payload)
    except _PartnerResponseError as e:
        body = {"error": str(e)}
        if e.raw_response is not None:
            body["raw_response"] = e.raw_response
        return jsonify(body), 502
    except FileNotFoundError as e:
        return jsonify({"error": str(e)}), 404
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    return jsonify(response), 201


# ==========================================================
# Premium Calculation Route
# ==========================================================
# _run_partner_calculation is the exact body of the original /get-premium
# route, extracted so Bulk Premium can call it directly (no HTTP) for each
# row — reusing the same payload construction (get_payload_template,
# _process_injection — both untouched) and the same Partner API call this
# route has always made. The route itself is now a thin wrapper: same
# validation, same error handling, same response shape, same status codes.

def _run_partner_calculation(flow_type: str, dims: dict, direct_payload: dict = None) -> dict:
    """Core Partner-API premium calculation for flow_type in
    {"create_lead", "save_loan"} — i.e. the /get-premium quote-only path,
    not the fuller lead-creation flows under /save-lead or /create-lead.
    direct_payload mirrors the lead flows' optional full-payload-override
    input — when supplied, it's used as the base instead of the hardcoded
    template, with dimension values still injected on top exactly as they
    are for the default template.
    Raises ValueError for invalid input (mirrors the route's 400s),
    FileNotFoundError if the payload template is missing (mirrors the
    route's 404), and lets any other exception propagate (mirrors the
    route's generic 500 handling) — callers translate these to whatever
    response shape they need."""
    if flow_type not in ["create_lead", "save_loan"]:
        raise ValueError("Invalid flow type.")
    if not dims:
        raise ValueError("Missing dimensions.")

    if isinstance(direct_payload, dict) and direct_payload:
        payload = copy.deepcopy(direct_payload)
    else:
        # Load a fresh deep copy of the template — disk file is never touched
        payload = get_payload_template(flow_type)

    # Inject dimension values into the copy
    mod_payload = _process_injection(
        dims=dims,
        flow_type=flow_type,
        payload=payload,
    )

    # Generate a unique application number and forward to the partner API
    loan_no = api_service.gen_app_no()
    response = api_service.save_loan_details(loan_no, mod_payload)
    response["payload"] = mod_payload
    return response


# ==========================================================
# Payload Template Route (for the frontend "Edit Payload" popup)
# ==========================================================

_VALID_FLOW_TYPES = {"save_loan", "save_lead", "create_lead"}


@injestion_bp.get("/payload-template/<flow_type>")
def payload_template(flow_type: str):
    """Return the default hardcoded payload template for a given flow_type,
    so the frontend can pre-fill the 'Edit Payload' popup with it."""
    if flow_type not in _VALID_FLOW_TYPES:
        return jsonify({"error": f"Invalid flow_type '{flow_type}'."}), 400

    try:
        template = get_payload_template(flow_type)
    except FileNotFoundError as e:
        return jsonify({"error": str(e)}), 404

    return jsonify(template), 200


@injestion_bp.post("/get-premium")
def get_premium():
    data = request.get_json()
    if not data:
        return jsonify({"error": "No JSON payload provided."}), 400

    flow_type = data.get("flow_type")
    dims = data.get("dimensions")
    direct_payload = _extract_lead_payload(data)

    try:
        response = _run_partner_calculation(flow_type, dims, direct_payload)
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except FileNotFoundError as e:
        return jsonify({"error": str(e)}), 404
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    return jsonify(response), 201
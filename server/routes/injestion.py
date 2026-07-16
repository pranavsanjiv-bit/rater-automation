import os
import json
# pyrefly: ignore [missing-import]
from flask import Blueprint, request, jsonify
from server.helpers import _age_to_dob
from server.templates import get_payload_template
from server.api_service import APIService

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


def _load_injection_paths() -> dict:
    """Loads the injection path configuration file."""
    current_dir = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(current_dir, "..", "injection_path.json")
    with open(config_path, "r") as f:
        return json.load(f)


def _inject_value(target, path: str, val):
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

    # Inject each resolved dimension into the payload
    for key, val in resolved_dims.items():
        if key in injection_paths:
            path_mapping = injection_paths[key]
            if flow_type in path_mapping:
                _inject_value(payload, path_mapping[flow_type], val)

    return payload


@injestion_bp.post("/get-premium")
def get_premium():
    data = request.get_json()
    if not data:
        return jsonify({"error": "No JSON payload provided."}), 400

    flow_type = data.get("flow_type")
    dims = data.get("dimensions")

    # Validations
    if flow_type not in ["create_lead", "save_loan"]:
        return jsonify({"error": "Invalid flow type."}), 400

    if not dims:
        return jsonify({"error": "Missing dimensions."}), 400

    try:
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

    except FileNotFoundError as e:
        return jsonify({"error": str(e)}), 404
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    return jsonify(response), 201
import copy
import json
import os

# Directory that holds all payload template files
_PAYLOADS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "payloads")

# In-memory cache: template is loaded once per flow_type and reused read-only
_template_cache: dict[str, dict] = {}


def _load_template(flow_type: str) -> dict:
    """Load and cache the raw template dict for *flow_type* (read once from disk)."""
    if flow_type not in _template_cache:
        path = os.path.join(_PAYLOADS_DIR, f"{flow_type}.json")
        if not os.path.exists(path):
            raise FileNotFoundError(f"No payload template found for flow_type '{flow_type}' at {path}")
        with open(path, "r") as fh:
            _template_cache[flow_type] = json.load(fh)
    return _template_cache[flow_type]


def get_payload_template(flow_type: str) -> dict:
    """Return a deep copy of the payload template for *flow_type*.

    Every caller receives an independent copy so that mutations in one request
    never bleed into another and the cached template (or the JSON file on disk)
    is never modified.

    Args:
        flow_type: Identifies which template file to load, e.g. ``"save_loan"``.

    Returns:
        A fresh deep copy of the template dict.

    Raises:
        FileNotFoundError: If no ``<flow_type>.json`` exists in the payloads directory.
    """
    return copy.deepcopy(_load_template(flow_type))
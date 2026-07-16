import re

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
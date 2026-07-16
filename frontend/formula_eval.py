"""
formula_eval.py
Safe formula evaluator using Python's ast module.
Allowed: numeric literals, +, -, *, /, ^, (, ), and a caller-supplied whitelist of variable names.
"""

import ast


class FormulaError(ValueError):
    pass


# Added ast.Pow to support the '^' operator
_ALLOWED_OPS = (
    ast.Add, ast.Sub, ast.Mult, ast.Div, ast.UAdd, ast.USub, ast.Pow,
)


def _check_node(node: ast.AST, allowed_names: set[str]):
    """Recursively validate that every node in the AST is safe."""
    if isinstance(node, ast.Expression):
        _check_node(node.body, allowed_names)

    elif isinstance(node, ast.BinOp):
        if not isinstance(node.op, _ALLOWED_OPS):
            raise FormulaError(f"Operator {type(node.op).__name__} is not allowed.")
        _check_node(node.left, allowed_names)
        _check_node(node.right, allowed_names)

    elif isinstance(node, ast.UnaryOp):
        if not isinstance(node.op, _ALLOWED_OPS):
            raise FormulaError(f"Unary operator {type(node.op).__name__} is not allowed.")
        _check_node(node.operand, allowed_names)

    elif isinstance(node, ast.Constant):
        if not isinstance(node.value, (int, float)):
            raise FormulaError(f"Non-numeric constant {node.value!r} is not allowed.")

    elif isinstance(node, ast.Name):
        if node.id not in allowed_names:
            raise FormulaError(f"Variable '{node.id}' is not in the allowed list: {sorted(allowed_names)}")

    else:
        raise FormulaError(f"AST node type {type(node).__name__} is not allowed.")


def validate_formula(formula: str, allowed_names: set[str]) -> None:
    """
    Parse and validate a formula string.
    Raises FormulaError with a descriptive message if anything is wrong.
    allowed_names: set of variable names permitted in the formula
    """
    # Translate '^' to Python's native exponentiation operator '**'
    normalized_formula = formula.replace("^", "**").strip()
    
    try:
        tree = ast.parse(normalized_formula, mode="eval")
    except SyntaxError as e:
        raise FormulaError(f"Syntax error in formula: {e}") from e
    _check_node(tree, allowed_names)


def evaluate_formula(formula: str, context: dict[str, float]) -> float:
    """
    Safely evaluate a formula string with the given variable context.
    context: dict mapping variable names to their numeric values.
    Returns the computed float.
    Raises FormulaError for any validation or evaluation issue.
    """
    allowed_names = set(context.keys())
    validate_formula(formula, allowed_names)

    # Translate '^' to '**' for evaluation
    normalized_formula = formula.replace("^", "**").strip()

    try:
        result = eval(  # noqa: S307
            compile(normalized_formula, "<formula>", "eval"),
            {"__builtins__": {}},
            context,
        )
    except ZeroDivisionError:
        raise FormulaError("Formula resulted in division by zero.")
    except Exception as e:
        raise FormulaError(f"Formula evaluation failed: {e}") from e

    if not isinstance(result, (int, float)):
        raise FormulaError(f"Formula did not produce a numeric result: {result!r}")

    return float(result)
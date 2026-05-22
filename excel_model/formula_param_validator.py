"""Formula parameter and cross-reference validation for line items."""

from excel_model.exceptions import FormulaInjectionError
from excel_model.formula_types import FormulaType
from excel_model.injection_guard import validate_custom_formula
from excel_model.spec import ModelSpec

_REQUIRED_PARAMS: dict[str, list[str]] = {
    "growth_projected": ["growth_assumption"],
    "pct_of_revenue": ["revenue_key", "rate_assumption"],
    "sum_of_rows": ["addend_keys"],
    "subtraction": ["minuend_key", "subtrahend_key"],
    "sum_subtraction": ["addend_key", "subtrahend_keys"],
    "ratio": ["numerator_key", "denominator_key"],
    "growth_rate": ["value_key"],
    "discounted_pv": ["cashflow_key", "rate_assumption"],
    "terminal_value": ["cashflow_key", "growth_assumption", "rate_assumption"],
    "npv_sum": ["pv_fcf_key", "pv_terminal_key"],
    "variance": ["plan_key", "actual_key"],
    "variance_pct": ["plan_key", "actual_key"],
    "constant": ["value"],
    "custom": ["formula"],
    "input_ref": ["projected_type"],
    "rank": ["value_key"],
    "index_to_base": ["value_key", "base_entity_key"],
    "bar_chart_text": ["value_key"],
}

_KEY_REF_PARAMS = {
    "revenue_key",
    "minuend_key",
    "subtrahend_key",
    "addend_key",
    "numerator_key",
    "denominator_key",
    "value_key",
    "cashflow_key",
    "pv_fcf_key",
    "pv_terminal_key",
    "plan_key",
    "actual_key",
    "prior_key",
    "base_entity_key",
}
_KEY_LIST_REF_PARAMS = {"addend_keys", "subtrahend_keys"}


def check_formula_params(spec: ModelSpec) -> list[str]:
    """Validate formula types, required params, and custom formula injection."""
    errors: list[str] = []
    valid_formula_types = {ft.value for ft in FormulaType}

    for li in spec.line_items:
        if li.formula_type not in valid_formula_types:
            errors.append(
                f"Line item {li.key!r} has unknown formula_type: {li.formula_type!r}. "
                f"Valid types: {sorted(valid_formula_types)}"
            )

        if li.formula_type in _REQUIRED_PARAMS:
            required = _REQUIRED_PARAMS[li.formula_type]
            for param in required:
                if param not in li.formula_params:
                    errors.append(
                        f"Line item {li.key!r} (formula_type={li.formula_type!r}) missing required param {param!r}"
                    )

        if li.formula_type == "custom" and "formula" in li.formula_params:
            raw = li.formula_params["formula"]
            if isinstance(raw, str):
                try:
                    validate_custom_formula(raw, li.key)
                except FormulaInjectionError as e:
                    errors.append(str(e))

    return errors


def check_cross_refs(spec: ModelSpec) -> list[str]:
    """Validate that line item cross-references point to existing keys."""
    errors: list[str] = []
    line_item_keys = {li.key for li in spec.line_items}

    for li in spec.line_items:
        for param_name, param_value in li.formula_params.items():
            if (
                param_name in _KEY_REF_PARAMS
                and isinstance(param_value, str)
                and param_value
                and param_value not in line_item_keys
            ):
                errors.append(f"Line item {li.key!r} references unknown key {param_value!r} via {param_name!r}")
            if param_name in _KEY_LIST_REF_PARAMS and isinstance(param_value, list):
                for ref_key in param_value:
                    if isinstance(ref_key, str) and ref_key not in line_item_keys:
                        errors.append(f"Line item {li.key!r} references unknown key {ref_key!r} via {param_name!r}")

    return errors

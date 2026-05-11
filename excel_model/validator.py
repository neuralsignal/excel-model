"""Validation for ModelSpec and InputData."""

import re

from excel_model.exceptions import FormulaInjectionError
from excel_model.formula_types import FormulaType
from excel_model.injection_guard import validate_custom_formula, validate_text_field
from excel_model.loader import InputData
from excel_model.spec import ModelSpec

_VALID_NAMED_RANGE_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_.]*$")


_VALID_MODEL_TYPES = {"p_and_l", "dcf", "budget_vs_actuals", "scenario", "comparison", "custom"}
_VALID_GRANULARITIES = {"monthly", "quarterly", "annual", "auto"}
_VALID_FORMATS = {"number", "percent", "currency", "integer"}

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

# Params that reference other line item keys (for cross-reference validation)
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


def validate_spec(spec: ModelSpec) -> list[str]:
    """Return a list of error strings. Empty list means valid."""
    errors: list[str] = []
    errors.extend(_validate_top_level_fields(spec))
    errors.extend(_validate_assumptions(spec))
    errors.extend(_validate_drivers(spec))
    errors.extend(_validate_line_items(spec))
    errors.extend(_validate_model_type_rules(spec))
    errors.extend(_validate_text_fields(spec))
    return errors


def _validate_top_level_fields(spec: ModelSpec) -> list[str]:
    """Validate model_type, title, currency, granularity, start_period, and period counts."""
    errors: list[str] = []

    if spec.model_type not in _VALID_MODEL_TYPES:
        errors.append(f"Invalid model_type: {spec.model_type!r}. Must be one of {sorted(_VALID_MODEL_TYPES)}")

    if not spec.title:
        errors.append("title must not be empty")

    if not spec.currency:
        errors.append("currency must not be empty")

    if spec.granularity not in _VALID_GRANULARITIES:
        errors.append(f"Invalid granularity: {spec.granularity!r}. Must be one of {sorted(_VALID_GRANULARITIES)}")

    if not spec.start_period:
        errors.append("start_period must not be empty")

    # comparison models don't require n_periods >= 1
    if spec.model_type != "comparison" and spec.n_periods < 1:
        errors.append(f"n_periods must be >= 1, got {spec.n_periods}")

    if spec.n_history_periods < 0:
        errors.append(f"n_history_periods must be >= 0, got {spec.n_history_periods}")

    return errors


def _validate_assumptions(spec: ModelSpec) -> list[str]:
    """Validate assumption name uniqueness, named-range format, and value formats."""
    errors: list[str] = []

    seen_names: set[str] = set()
    for assumption in spec.assumptions:
        if assumption.name in seen_names:
            errors.append(f"Duplicate assumption name: {assumption.name!r}")
        seen_names.add(assumption.name)

        if not _VALID_NAMED_RANGE_RE.match(assumption.name):
            errors.append(
                f"Assumption name {assumption.name!r} is not a valid Excel named range. "
                f"Must start with a letter or underscore and contain only "
                f"letters, digits, underscores, or periods."
            )

        if assumption.format not in _VALID_FORMATS:
            errors.append(
                f"Assumption {assumption.name!r} has invalid format: {assumption.format!r}. "
                f"Valid formats: {sorted(_VALID_FORMATS)}"
            )

    return errors


def _validate_drivers(spec: ModelSpec) -> list[str]:
    """Validate driver uniqueness, named-range format, format validity, and namespace collisions."""
    errors: list[str] = []

    assumption_names = {a.name for a in spec.assumptions}
    seen_driver_names: set[str] = set()

    for driver in spec.drivers:
        if driver.name in seen_driver_names:
            errors.append(f"Duplicate driver name: {driver.name!r}")
        seen_driver_names.add(driver.name)

        if driver.name in assumption_names:
            errors.append(
                f"Driver name {driver.name!r} collides with assumption name. Assumptions and drivers share a namespace."
            )

        if not _VALID_NAMED_RANGE_RE.match(driver.name):
            errors.append(
                f"Driver name {driver.name!r} is not a valid Excel named range. "
                f"Must start with a letter or underscore and contain only "
                f"letters, digits, underscores, or periods."
            )

        if driver.format not in _VALID_FORMATS:
            errors.append(
                f"Driver {driver.name!r} has invalid format: {driver.format!r}. Valid formats: {sorted(_VALID_FORMATS)}"
            )

    # Validate driver_overrides keys reference actual driver names
    driver_name_set = set(seen_driver_names)
    for scenario in spec.scenarios:
        for key in scenario.driver_overrides:
            if key not in driver_name_set:
                errors.append(
                    f"Scenario {scenario.name!r} has driver_overrides key {key!r} that does not match any driver name"
                )

    return errors


def _validate_line_items(spec: ModelSpec) -> list[str]:
    """Validate line item key uniqueness, formula types, required params, and cross-references."""
    errors: list[str] = []
    errors.extend(_check_key_uniqueness(spec))
    errors.extend(_check_formula_params(spec))
    errors.extend(_check_cross_refs(spec))
    return errors


def _check_key_uniqueness(spec: ModelSpec) -> list[str]:
    """Check for duplicate line item keys."""
    errors: list[str] = []
    seen_keys: set[str] = set()
    for li in spec.line_items:
        if li.key in seen_keys:
            errors.append(f"Duplicate line item key: {li.key!r}")
        seen_keys.add(li.key)
    return errors


def _check_formula_params(spec: ModelSpec) -> list[str]:
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


def _check_cross_refs(spec: ModelSpec) -> list[str]:
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


def _validate_model_type_rules(spec: ModelSpec) -> list[str]:
    """Validate model-type-specific rules (scenario, bva, comparison, dcf)."""
    errors: list[str] = []

    if spec.model_type == "scenario" and not spec.scenarios:
        errors.append("Scenario model must have at least one scenario defined")

    if spec.model_type == "budget_vs_actuals" and not spec.column_groups:
        errors.append("Budget vs Actuals model must have column_groups defined")

    if spec.model_type == "comparison":
        if not spec.entities:
            errors.append("Comparison model must have at least one entity defined")
        entity_keys: set[str] = set()
        for entity in spec.entities:
            if entity.key in entity_keys:
                errors.append(f"Duplicate entity key: {entity.key!r}")
            entity_keys.add(entity.key)

    if spec.model_type == "dcf":
        errors.extend(_validate_wacc_tgr(spec))

    return errors


def _validate_wacc_tgr(spec: ModelSpec) -> list[str]:
    """For DCF models: ensure WACC and terminal growth rate are not equal."""
    errors: list[str] = []
    assumption_values = {a.name: a.value for a in spec.assumptions}

    for li in spec.line_items:
        if li.formula_type == "terminal_value":
            rate_name = li.formula_params.get("rate_assumption")
            growth_name = li.formula_params.get("growth_assumption")
            if rate_name and growth_name:
                rate_val = assumption_values.get(rate_name)
                growth_val = assumption_values.get(growth_name)
                if rate_val is not None and growth_val is not None and rate_val == growth_val:
                    errors.append(
                        f"DCF terminal value: discount rate ({rate_name}={rate_val}) "
                        f"equals growth rate ({growth_name}={growth_val}). "
                        f"This causes division by zero in the Gordon Growth Model."
                    )
    return errors


def _validate_text_fields(spec: ModelSpec) -> list[str]:
    """Reject user-controlled text fields that start with formula-injection characters."""
    errors: list[str] = []

    def _check(value: str, context: str) -> None:
        try:
            validate_text_field(value, context)
        except FormulaInjectionError as exc:
            errors.append(str(exc))

    _check(spec.title, "title")
    _check(spec.metadata.preparer, "metadata preparer")

    for assumption in spec.assumptions:
        _check(assumption.label, f"assumption {assumption.name!r} label")
        _check(assumption.group, f"assumption {assumption.name!r} group")

    for driver in spec.drivers:
        _check(driver.label, f"driver {driver.name!r} label")
        _check(driver.group, f"driver {driver.name!r} group")

    for li in spec.line_items:
        _check(li.label, f"line item {li.key!r} label")
        _check(li.section, f"line item {li.key!r} section")

    for scenario in spec.scenarios:
        _check(scenario.label, f"scenario {scenario.name!r} label")

    for entity in spec.entities:
        _check(entity.label, f"entity {entity.key!r} label")

    for cg in spec.column_groups:
        _check(cg.label, f"column group {cg.key!r} label")

    return errors


def validate_inputs_against_spec(spec: ModelSpec, inputs: InputData) -> list[str]:
    """Validate that InputData columns match what the spec expects. Return error list."""
    errors: list[str] = []

    if inputs.period_col not in inputs.df.columns:
        errors.append(f"period_col {inputs.period_col!r} not found in input data columns: {inputs.df.columns}")

    for key, col_name in spec.inputs.value_cols.items():
        if col_name not in inputs.df.columns:
            errors.append(f"value_col for {key!r} ({col_name!r}) not found in input data columns: {inputs.df.columns}")

    return errors

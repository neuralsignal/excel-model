"""Validation for ModelSpec."""

import re

from excel_model.exceptions import FormulaInjectionError
from excel_model.formula_param_validator import check_cross_refs, check_formula_params
from excel_model.injection_guard import validate_text_field
from excel_model.spec import ModelSpec

_VALID_NAMED_RANGE_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_.]*$")
_HEX_COLOR_RE = re.compile(r"^#?[0-9A-Fa-f]{6}$")


_VALID_MODEL_TYPES = {"p_and_l", "dcf", "budget_vs_actuals", "scenario", "comparison", "custom"}
_VALID_GRANULARITIES = {"monthly", "quarterly", "annual", "auto"}
_VALID_FORMATS = {"number", "percent", "currency", "integer"}


def validate_spec(spec: ModelSpec) -> list[str]:
    """Return a list of error strings. Empty list means valid."""
    errors: list[str] = []
    errors.extend(_validate_top_level_fields(spec))
    errors.extend(_validate_assumptions(spec))
    errors.extend(_validate_drivers(spec))
    errors.extend(_validate_line_items(spec))
    errors.extend(_validate_model_type_rules(spec))
    errors.extend(_validate_text_fields(spec))
    errors.extend(_validate_colors(spec))
    return errors


def _validate_colors(spec: ModelSpec) -> list[str]:
    """Reject column group color_hex values that are not well-formed 6-digit hex."""
    errors: list[str] = []
    for cg in spec.column_groups:
        if not _HEX_COLOR_RE.match(cg.color_hex):
            errors.append(
                f"Column group {cg.key!r} has invalid color_hex {cg.color_hex!r}. "
                f"Must be a 6-digit hex color (e.g. 'FF8800' or '#FF8800')."
            )
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
    errors.extend(check_formula_params(spec))
    errors.extend(check_cross_refs(spec))
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

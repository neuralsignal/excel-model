"""Extended validator tests — formula_params, cross-ref, WACC≠TGR, comparison, drivers."""

from excel_model.spec import (
    AssumptionDef,
    DriverDef,
    EntityDef,
    InputsDef,
    LineItemDef,
    MetadataDef,
    ModelSpec,
    ScenarioDef,
)
from excel_model.validator import validate_spec


def make_spec(**overrides) -> ModelSpec:
    defaults = dict(
        model_type="p_and_l",
        title="Test",
        currency="CHF",
        granularity="annual",
        start_period="2025",
        n_periods=3,
        n_history_periods=0,
        assumptions=(),
        drivers=(),
        line_items=(),
        metadata=MetadataDef(preparer="", date="", version="1.0"),
        scenarios=(),
        column_groups=(),
        inputs=InputsDef(source="", period_col="period", sheet="", value_cols={}),
        entities=(),
    )
    defaults.update(overrides)
    return ModelSpec(**defaults)


class TestFormulaParamsValidation:
    def test_missing_growth_assumption(self):
        spec = make_spec(
            line_items=(
                LineItemDef(
                    key="rev",
                    label="Rev",
                    formula_type="growth_projected",
                    formula_params={},
                    is_subtotal=False,
                    is_total=False,
                    section="",
                    format="",
                ),
            )
        )
        errors = validate_spec(spec)
        assert any("growth_assumption" in e for e in errors)

    def test_missing_subtraction_params(self):
        spec = make_spec(
            line_items=(
                LineItemDef(
                    key="x",
                    label="X",
                    formula_type="subtraction",
                    formula_params={"minuend_key": "a"},
                    is_subtotal=False,
                    is_total=False,
                    section="",
                    format="",
                ),
            )
        )
        errors = validate_spec(spec)
        assert any("subtrahend_key" in e for e in errors)

    def test_valid_params_no_error(self):
        spec = make_spec(
            line_items=(
                LineItemDef(
                    key="rev",
                    label="Rev",
                    formula_type="constant",
                    formula_params={"value": 42},
                    is_subtotal=False,
                    is_total=False,
                    section="",
                    format="",
                ),
            )
        )
        errors = validate_spec(spec)
        assert not errors

    def test_sum_subtraction_missing_subtrahend_keys(self):
        spec = make_spec(
            line_items=(
                LineItemDef(
                    key="fcf",
                    label="FCF",
                    formula_type="sum_subtraction",
                    formula_params={"addend_key": "nopat"},
                    is_subtotal=False,
                    is_total=False,
                    section="",
                    format="",
                ),
            )
        )
        errors = validate_spec(spec)
        assert any("subtrahend_keys" in e for e in errors)


class TestCrossReferenceValidation:
    def test_unknown_revenue_key(self):
        spec = make_spec(
            line_items=(
                LineItemDef(
                    key="cogs",
                    label="COGS",
                    formula_type="pct_of_revenue",
                    formula_params={"revenue_key": "nonexistent", "rate_assumption": "X"},
                    is_subtotal=False,
                    is_total=False,
                    section="",
                    format="",
                ),
            )
        )
        errors = validate_spec(spec)
        assert any("nonexistent" in e for e in errors)

    def test_valid_key_reference(self):
        spec = make_spec(
            line_items=(
                LineItemDef(
                    key="revenue",
                    label="Revenue",
                    formula_type="constant",
                    formula_params={"value": 100},
                    is_subtotal=False,
                    is_total=False,
                    section="",
                    format="",
                ),
                LineItemDef(
                    key="cogs",
                    label="COGS",
                    formula_type="pct_of_revenue",
                    formula_params={"revenue_key": "revenue", "rate_assumption": "X"},
                    is_subtotal=False,
                    is_total=False,
                    section="",
                    format="",
                ),
            )
        )
        errors = validate_spec(spec)
        assert not any("unknown key" in e for e in errors)

    def test_unknown_key_in_addend_keys(self):
        spec = make_spec(
            line_items=(
                LineItemDef(
                    key="total",
                    label="Total",
                    formula_type="sum_of_rows",
                    formula_params={"addend_keys": ["a", "missing_key"]},
                    is_subtotal=False,
                    is_total=False,
                    section="",
                    format="",
                ),
                LineItemDef(
                    key="a",
                    label="A",
                    formula_type="constant",
                    formula_params={"value": 1},
                    is_subtotal=False,
                    is_total=False,
                    section="",
                    format="",
                ),
            )
        )
        errors = validate_spec(spec)
        assert any("missing_key" in e for e in errors)


class TestWaccTgrGuard:
    def test_equal_wacc_tgr_fails(self):
        spec = make_spec(
            model_type="dcf",
            assumptions=(
                AssumptionDef(name="WACC", label="WACC", value=0.10, format="percent", group="V"),
                AssumptionDef(name="TGR", label="TGR", value=0.10, format="percent", group="V"),
            ),
            line_items=(
                LineItemDef(
                    key="tv",
                    label="TV",
                    formula_type="terminal_value",
                    formula_params={"cashflow_key": "fcf", "growth_assumption": "TGR", "rate_assumption": "WACC"},
                    is_subtotal=False,
                    is_total=False,
                    section="",
                    format="",
                ),
                LineItemDef(
                    key="fcf",
                    label="FCF",
                    formula_type="constant",
                    formula_params={"value": 100},
                    is_subtotal=False,
                    is_total=False,
                    section="",
                    format="",
                ),
            ),
        )
        errors = validate_spec(spec)
        assert any("division by zero" in e.lower() for e in errors)

    def test_different_wacc_tgr_ok(self):
        spec = make_spec(
            model_type="dcf",
            assumptions=(
                AssumptionDef(name="WACC", label="WACC", value=0.10, format="percent", group="V"),
                AssumptionDef(name="TGR", label="TGR", value=0.02, format="percent", group="V"),
            ),
            line_items=(
                LineItemDef(
                    key="tv",
                    label="TV",
                    formula_type="terminal_value",
                    formula_params={"cashflow_key": "fcf", "growth_assumption": "TGR", "rate_assumption": "WACC"},
                    is_subtotal=False,
                    is_total=False,
                    section="",
                    format="",
                ),
                LineItemDef(
                    key="fcf",
                    label="FCF",
                    formula_type="constant",
                    formula_params={"value": 100},
                    is_subtotal=False,
                    is_total=False,
                    section="",
                    format="",
                ),
            ),
        )
        errors = validate_spec(spec)
        assert not any("division by zero" in e.lower() for e in errors)


class TestAssumptionNameValidation:
    def test_spaces_rejected(self):
        spec = make_spec(
            assumptions=(AssumptionDef(name="Patient Count", label="PC", value=100, format="number", group="A"),)
        )
        errors = validate_spec(spec)
        assert any("valid Excel named range" in e for e in errors)

    def test_starts_with_digit_rejected(self):
        spec = make_spec(assumptions=(AssumptionDef(name="1stYear", label="1Y", value=1, format="number", group="A"),))
        errors = validate_spec(spec)
        assert any("valid Excel named range" in e for e in errors)

    def test_special_characters_rejected(self):
        spec = make_spec(
            assumptions=(AssumptionDef(name="Price/Unit", label="PU", value=50, format="number", group="A"),)
        )
        errors = validate_spec(spec)
        assert any("valid Excel named range" in e for e in errors)

    def test_valid_camel_case_accepted(self):
        spec = make_spec(
            assumptions=(AssumptionDef(name="PatientCount", label="PC", value=100, format="number", group="A"),)
        )
        errors = validate_spec(spec)
        assert not any("valid Excel named range" in e for e in errors)

    def test_underscore_and_period_accepted(self):
        spec = make_spec(
            assumptions=(AssumptionDef(name="_Growth.Rate", label="GR", value=0.1, format="percent", group="A"),)
        )
        errors = validate_spec(spec)
        assert not any("valid Excel named range" in e for e in errors)


class TestComparisonValidation:
    def test_comparison_without_entities(self):
        spec = make_spec(model_type="comparison", n_periods=0, entities=())
        errors = validate_spec(spec)
        assert any("entity" in e.lower() for e in errors)

    def test_comparison_duplicate_entity_keys(self):
        spec = make_spec(
            model_type="comparison",
            n_periods=0,
            entities=(
                EntityDef(key="a", label="A"),
                EntityDef(key="a", label="A Duplicate"),
            ),
        )
        errors = validate_spec(spec)
        assert any("Duplicate entity" in e for e in errors)

    def test_valid_comparison(self):
        spec = make_spec(
            model_type="comparison",
            n_periods=0,
            entities=(
                EntityDef(key="a", label="A"),
                EntityDef(key="b", label="B"),
            ),
        )
        errors = validate_spec(spec)
        # Should not have entity-related errors
        assert not any("entity" in e.lower() for e in errors)


class TestDriverValidation:
    def test_duplicate_driver_names(self):
        spec = make_spec(
            drivers=(
                DriverDef(name="PatientCount", label="PC", value=100, format="integer", group="V"),
                DriverDef(name="PatientCount", label="PC2", value=200, format="integer", group="V"),
            )
        )
        errors = validate_spec(spec)
        assert any("Duplicate driver name" in e for e in errors)

    def test_driver_assumption_collision(self):
        spec = make_spec(
            assumptions=(AssumptionDef(name="PatientCount", label="PC", value=100, format="integer", group="G"),),
            drivers=(DriverDef(name="PatientCount", label="PC Driver", value=200, format="integer", group="V"),),
        )
        errors = validate_spec(spec)
        assert any("collides with assumption" in e for e in errors)

    def test_invalid_driver_name_rejected(self):
        spec = make_spec(drivers=(DriverDef(name="Patient Count", label="PC", value=100, format="integer", group="V"),))
        errors = validate_spec(spec)
        assert any("valid Excel named range" in e for e in errors)

    def test_driver_overrides_nonexistent_driver(self):
        spec = make_spec(
            model_type="scenario",
            drivers=(DriverDef(name="PatientCount", label="PC", value=100, format="integer", group="V"),),
            scenarios=(
                ScenarioDef(name="base", label="Base", assumption_overrides={}, driver_overrides={"NonExistent": 999}),
            ),
        )
        errors = validate_spec(spec)
        assert any("does not match any driver" in e for e in errors)

    def test_valid_spec_with_drivers(self):
        spec = make_spec(
            model_type="scenario",
            drivers=(
                DriverDef(name="PatientCount", label="PC", value=100, format="integer", group="V"),
                DriverDef(name="PricePerPatient", label="PPP", value=50, format="currency", group="V"),
            ),
            assumptions=(AssumptionDef(name="CROPerPatient", label="CRO", value=2992, format="currency", group="B"),),
            scenarios=(
                ScenarioDef(name="base", label="Base", assumption_overrides={}, driver_overrides={}),
                ScenarioDef(
                    name="premium", label="Premium", assumption_overrides={}, driver_overrides={"PricePerPatient": 100}
                ),
            ),
        )
        errors = validate_spec(spec)
        assert not any("driver" in e.lower() for e in errors)

"""Strictyaml schema for model spec YAML files."""

from strictyaml import (
    Any,
    Bool,
    Enum,
    Float,
    Int,
    Map,
    MapPattern,
    Optional,
    OrValidator,
    Seq,
    Str,
)

_ASSUMPTION_SCHEMA = Map(
    {
        "name": Str(),
        "label": Str(),
        "value": OrValidator(Int(), OrValidator(Float(), Str())),
        "format": Enum(["number", "percent", "currency", "integer"]),
        Optional("group", default="General"): Str(),
    }
)

_DRIVER_SCHEMA = Map(
    {
        "name": Str(),
        "label": Str(),
        "value": OrValidator(Int(), OrValidator(Float(), Str())),
        "format": Enum(["number", "percent", "currency", "integer"]),
        Optional("group", default="General"): Str(),
    }
)

_METADATA_SCHEMA = Map(
    {
        Optional("preparer", default=""): Str(),
        Optional("date", default=""): Str(),
        Optional("version", default="1.0"): Str(),
    }
)

_FORMULA_PARAMS_SCHEMA = MapPattern(Str(), Any())

_LINE_ITEM_SCHEMA = Map(
    {
        "key": Str(),
        "label": Str(),
        "formula_type": Str(),
        Optional("formula_params"): _FORMULA_PARAMS_SCHEMA,
        Optional("format"): Enum(["number", "percent", "currency", "integer"]),
        Optional("group"): Str(),
        Optional("indent"): Int(),
        Optional("is_subtotal", default="false"): Bool(),
        Optional("is_total", default="false"): Bool(),
        Optional("hide_in_model"): Bool(),
        Optional("section", default=""): Str(),
    }
)

_SCENARIO_SCHEMA = Map(
    {
        "name": Str(),
        "label": Str(),
        Optional("assumption_overrides"): MapPattern(Str(), OrValidator(Int(), OrValidator(Float(), Str()))),
        Optional("driver_overrides"): MapPattern(Str(), OrValidator(Int(), OrValidator(Float(), Str()))),
    }
)

_COLUMN_GROUP_SCHEMA = Map(
    {
        "key": Str(),
        "label": Str(),
        Optional("color_hex", default="FFFFFF"): Str(),
    }
)

_ENTITY_SCHEMA = Map(
    {
        "key": Str(),
        "label": Str(),
    }
)

_INPUTS_SCHEMA = Map(
    {
        Optional("source", default=""): Str(),
        Optional("period_col", default="period"): Str(),
        Optional("sheet", default=""): Str(),
        Optional("value_cols"): MapPattern(Str(), Str()),
        Optional("entity_col"): Str(),
    }
)

SPEC_SCHEMA = Map(
    {
        "model_type": Enum(["p_and_l", "dcf", "budget_vs_actuals", "scenario", "comparison", "custom"]),
        "title": Str(),
        "currency": Str(),
        "granularity": Enum(["monthly", "quarterly", "annual", "auto"]),
        "start_period": Str(),
        "n_periods": Int(),
        "n_history_periods": Int(),
        Optional("metadata"): _METADATA_SCHEMA,
        Optional("assumptions"): Seq(_ASSUMPTION_SCHEMA),
        Optional("drivers"): Seq(_DRIVER_SCHEMA),
        Optional("line_items"): Seq(_LINE_ITEM_SCHEMA),
        Optional("scenarios"): Seq(_SCENARIO_SCHEMA),
        Optional("column_groups"): Seq(_COLUMN_GROUP_SCHEMA),
        Optional("inputs"): _INPUTS_SCHEMA,
        Optional("entities"): Seq(_ENTITY_SCHEMA),
    }
)

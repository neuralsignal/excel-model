"""Input data validation against model spec."""

from excel_model.loader import InputData
from excel_model.spec import ModelSpec


def validate_inputs_against_spec(spec: ModelSpec, inputs: InputData) -> list[str]:
    """Validate that InputData columns match what the spec expects. Return error list."""
    errors: list[str] = []

    if inputs.period_col not in inputs.df.columns:
        errors.append(f"period_col {inputs.period_col!r} not found in input data columns: {inputs.df.columns}")

    for key, col_name in spec.inputs.value_cols.items():
        if col_name not in inputs.df.columns:
            errors.append(f"value_col for {key!r} ({col_name!r}) not found in input data columns: {inputs.df.columns}")

    return errors

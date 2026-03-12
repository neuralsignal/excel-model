"""Orchestrator: build_workbook() — creates the full Excel workbook."""
from pathlib import Path

from openpyxl import Workbook

from excel_model.loader import InputData
from excel_model.spec import ModelSpec
from excel_model.style import StyleConfig
from excel_model.time_engine import generate_periods


def build_workbook(
    spec: ModelSpec,
    inputs: InputData | None,
    output_path: str,
    style: StyleConfig,
) -> None:
    """Build and save a complete Excel workbook from the model spec.

    Creates Assumptions, Inputs, and Model sheets.
    Dispatches to the appropriate model builder based on spec.model_type.
    """
    wb = Workbook()
    # Remove default sheet
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    if spec.model_type == "comparison":
        # Comparison models don't use time_engine periods
        from excel_model.models.comparison import build_comparison
        build_comparison(wb, spec, inputs, style)
    else:
        periods = generate_periods(
            start_period=spec.start_period,
            n_periods=spec.n_periods,
            n_history=spec.n_history_periods,
            granularity=spec.granularity,
        )

        if spec.model_type == "p_and_l":
            from excel_model.models.p_and_l import build_p_and_l
            build_p_and_l(wb, spec, inputs, style, periods)

        elif spec.model_type == "dcf":
            from excel_model.models.dcf import build_dcf
            build_dcf(wb, spec, inputs, style, periods)

        elif spec.model_type == "budget_vs_actuals":
            from excel_model.models.budget_vs_actuals import build_budget_vs_actuals
            build_budget_vs_actuals(wb, spec, inputs, style, periods)

        elif spec.model_type == "scenario":
            from excel_model.models.scenario import build_scenario
            build_scenario(wb, spec, inputs, style, periods)

        elif spec.model_type == "custom":
            from excel_model.models.p_and_l import build_p_and_l
            build_p_and_l(wb, spec, inputs, style, periods)

        else:
            raise ValueError(f"Unknown model_type: {spec.model_type!r}")

    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))

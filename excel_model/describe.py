"""Build and render model description for the describe command."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from excel_model.spec import ModelSpec
    from excel_model.time_engine import Period


def build_description(spec: ModelSpec, periods: list[Period], errors: list[str]) -> dict[str, Any]:
    assumption_groups: dict[str, list[dict[str, Any]]] = {}
    for a in spec.assumptions:
        assumption_groups.setdefault(a.group, []).append({"name": a.name, "value": a.value, "format": a.format})
    sections: dict[str, list[dict[str, Any]]] = {}
    for li in spec.line_items:
        sections.setdefault(li.section, []).append(
            {
                "key": li.key,
                "label": li.label,
                "formula_type": li.formula_type,
                "is_subtotal": li.is_subtotal,
                "is_total": li.is_total,
            }
        )
    return {
        "model_type": spec.model_type,
        "title": spec.title,
        "currency": spec.currency,
        "granularity": spec.granularity,
        "start_period": spec.start_period,
        "n_history_periods": spec.n_history_periods,
        "n_periods": spec.n_periods,
        "total_periods": len(periods),
        "period_labels": [p.label for p in periods],
        "history_labels": [p.label for p in periods if p.is_history],
        "projection_labels": [p.label for p in periods if not p.is_history],
        "metadata": {"preparer": spec.metadata.preparer, "date": spec.metadata.date, "version": spec.metadata.version},
        "assumptions_count": len(spec.assumptions),
        "assumption_groups": assumption_groups,
        "line_items_count": len(spec.line_items),
        "sections": sections,
        "scenarios": [
            {"name": s.name, "label": s.label, "assumption_overrides": dict(s.assumption_overrides)}
            for s in spec.scenarios
        ],
        "column_groups": [{"key": cg.key, "label": cg.label} for cg in spec.column_groups],
        "sheets_to_create": ["Assumptions", "Inputs", "Model"],
        "validation_errors": errors,
        "inputs": {
            "source": spec.inputs.source,
            "period_col": spec.inputs.period_col,
            "value_cols": dict(spec.inputs.value_cols),
        },
    }


def _render_periods_lines(description: dict[str, Any]) -> list[str]:
    lines = [f"Periods ({description['granularity']}):"]
    if description["n_history_periods"] > 0:
        lines.append(f"  History ({description['n_history_periods']}): {', '.join(description['history_labels'])}")
    lines.append(f"  Projection ({description['n_periods']}): {', '.join(description['projection_labels'])}")
    return lines


def _render_assumptions_lines(description: dict[str, Any]) -> list[str]:
    lines = [f"Assumptions ({description['assumptions_count']}):"]
    for group, assumptions in description["assumption_groups"].items():
        lines.append(f"  [{group}]")
        for a in assumptions:
            lines.append(f"    {a['name']}: {a['value']} ({a['format']})")
    return lines


def _render_line_items_lines(description: dict[str, Any]) -> list[str]:
    lines = [f"Line Items ({description['line_items_count']}):"]
    for section, items in description["sections"].items():
        if section:
            lines.append(f"  [{section}]")
        for li in items:
            marker = " [subtotal]" if li["is_subtotal"] else (" [total]" if li["is_total"] else "")
            lines.append(f"    {li['label'].strip()}: {li['formula_type']}{marker}")
    return lines


def _render_scenarios_lines(description: dict[str, Any]) -> list[str]:
    if not description["scenarios"]:
        return []
    lines = [f"Scenarios ({len(description['scenarios'])}):"]
    for s in description["scenarios"]:
        overrides = ", ".join(f"{k}={v}" for k, v in s["assumption_overrides"].items())
        override_str = f" (overrides: {overrides})" if overrides else ""
        lines.append(f"  {s['label']}{override_str}")
    return lines


def _render_validation_lines(description: dict[str, Any]) -> list[str]:
    if not description["validation_errors"]:
        return ["Validation: OK"]
    lines = [f"Validation errors ({len(description['validation_errors'])}):"]
    for e in description["validation_errors"]:
        lines.append(f"  - {e}")
    return lines


def render_description_text(description: dict[str, Any]) -> str:
    sections: list[list[str]] = [
        [
            f"Model: {description['title']}",
            f"Type:  {description['model_type']}",
            f"Currency: {description['currency']}",
        ],
        _render_periods_lines(description),
        _render_assumptions_lines(description),
        _render_line_items_lines(description),
        _render_scenarios_lines(description),
        _render_validation_lines(description),
    ]
    return "\n\n".join("\n".join(s) for s in sections if s)

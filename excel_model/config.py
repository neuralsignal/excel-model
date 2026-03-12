"""Configuration loading with bundled defaults and deep-merge."""

from importlib import resources
from pathlib import Path

import yaml

from excel_model.exceptions import StyleConfigError
from excel_model.style import StyleConfig


def _load_default_style_yaml() -> dict:
    """Load the bundled default style config."""
    ref = resources.files("excel_model") / "defaults" / "default_style.yaml"
    return yaml.safe_load(ref.read_text(encoding="utf-8"))


def _deep_merge(base: dict, override: dict) -> dict:
    """Recursively merge override into base. Override values win."""
    result = dict(base)
    for key, value in override.items():
        if key in result and isinstance(result[key], dict) and isinstance(value, dict):
            result[key] = _deep_merge(result[key], value)
        else:
            result[key] = value
    return result


def load_style(style_path: str | None) -> StyleConfig:
    """Load StyleConfig from YAML, deep-merged with bundled defaults.

    If style_path is None, returns the bundled defaults only.
    If style_path is provided, user overrides are merged on top of defaults.
    """
    defaults = _load_default_style_yaml()

    if style_path is not None:
        p = Path(style_path)
        if not p.exists():
            raise StyleConfigError(f"Style config not found: {style_path}")
        with p.open() as f:
            user_data = yaml.safe_load(f) or {}
        merged = _deep_merge(defaults, user_data)
    else:
        merged = defaults

    required = [
        "header_fill_hex",
        "header_font_color",
        "subtotal_fill_hex",
        "total_fill_hex",
        "history_col_fill_hex",
        "section_header_fill_hex",
        "font_name",
        "font_size",
        "number_format_currency",
        "number_format_percent",
        "number_format_integer",
        "number_format_number",
    ]
    missing = [k for k in required if k not in merged]
    if missing:
        raise StyleConfigError(f"Style config missing required keys: {missing}")

    return StyleConfig(
        header_fill_hex=merged["header_fill_hex"],
        header_font_color=merged["header_font_color"],
        subtotal_fill_hex=merged["subtotal_fill_hex"],
        total_fill_hex=merged["total_fill_hex"],
        history_col_fill_hex=merged["history_col_fill_hex"],
        section_header_fill_hex=merged["section_header_fill_hex"],
        font_name=merged["font_name"],
        font_size=int(merged["font_size"]),
        number_format_currency=merged["number_format_currency"],
        number_format_percent=merged["number_format_percent"],
        number_format_integer=merged["number_format_integer"],
        number_format_number=merged["number_format_number"],
    )

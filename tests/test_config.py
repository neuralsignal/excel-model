"""Tests for excel_model.config: _deep_merge, _load_default_style_yaml, load_style."""

import pytest
from hypothesis import given, settings
from hypothesis import strategies as st

from excel_model.config import _deep_merge, _load_default_style_yaml, load_style
from excel_model.exceptions import StyleConfigError
from excel_model.style import StyleConfig

# ---------------------------------------------------------------------------
# Required keys present in default_style.yaml
# ---------------------------------------------------------------------------
REQUIRED_KEYS = [
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


# ---------------------------------------------------------------------------
# _load_default_style_yaml
# ---------------------------------------------------------------------------
class TestLoadDefaultStyleYaml:
    def test_returns_dict(self) -> None:
        result = _load_default_style_yaml()
        assert isinstance(result, dict)

    def test_contains_all_required_keys(self) -> None:
        result = _load_default_style_yaml()
        for key in REQUIRED_KEYS:
            assert key in result, f"Missing required key: {key}"


# ---------------------------------------------------------------------------
# _deep_merge — concrete tests
# ---------------------------------------------------------------------------
class TestDeepMerge:
    def test_empty_base_returns_override(self) -> None:
        assert _deep_merge({}, {"a": 1}) == {"a": 1}

    def test_empty_override_returns_base(self) -> None:
        assert _deep_merge({"a": 1}, {}) == {"a": 1}

    def test_both_empty(self) -> None:
        assert _deep_merge({}, {}) == {}

    def test_override_wins_for_scalar(self) -> None:
        assert _deep_merge({"a": 1}, {"a": 2}) == {"a": 2}

    def test_base_only_keys_preserved(self) -> None:
        result = _deep_merge({"a": 1, "b": 2}, {"a": 10})
        assert result == {"a": 10, "b": 2}

    def test_override_only_keys_added(self) -> None:
        result = _deep_merge({"a": 1}, {"b": 2})
        assert result == {"a": 1, "b": 2}

    def test_nested_dicts_merged_recursively(self) -> None:
        base = {"outer": {"a": 1, "b": 2}}
        override = {"outer": {"b": 99, "c": 3}}
        result = _deep_merge(base, override)
        assert result == {"outer": {"a": 1, "b": 99, "c": 3}}

    def test_override_replaces_dict_with_scalar(self) -> None:
        result = _deep_merge({"a": {"nested": 1}}, {"a": "flat"})
        assert result == {"a": "flat"}

    def test_override_replaces_scalar_with_dict(self) -> None:
        result = _deep_merge({"a": "flat"}, {"a": {"nested": 1}})
        assert result == {"a": {"nested": 1}}

    def test_does_not_mutate_base(self) -> None:
        base = {"a": {"x": 1}}
        _deep_merge(base, {"a": {"y": 2}})
        assert base == {"a": {"x": 1}}

    def test_does_not_mutate_override(self) -> None:
        override = {"a": {"y": 2}}
        _deep_merge({"a": {"x": 1}}, override)
        assert override == {"a": {"y": 2}}


# ---------------------------------------------------------------------------
# _deep_merge — property-based tests
# ---------------------------------------------------------------------------
json_values = st.recursive(
    st.one_of(st.integers(), st.text(max_size=10), st.booleans(), st.none()),
    lambda children: st.dictionaries(st.text(min_size=1, max_size=5), children, max_size=5),
    max_leaves=20,
)

json_dicts = st.dictionaries(st.text(min_size=1, max_size=5), json_values, max_size=5)


class TestDeepMergeProperties:
    @given(d=json_dicts)
    @settings(max_examples=50)
    def test_merge_with_empty_is_identity(self, d: dict) -> None:
        assert _deep_merge(d, {}) == d

    @given(d=json_dicts)
    @settings(max_examples=50)
    def test_empty_base_equals_override(self, d: dict) -> None:
        assert _deep_merge({}, d) == d

    @given(base=json_dicts, override=json_dicts)
    @settings(max_examples=50)
    def test_all_override_keys_present(self, base: dict, override: dict) -> None:
        result = _deep_merge(base, override)
        for key in override:
            assert key in result

    @given(base=json_dicts, override=json_dicts)
    @settings(max_examples=50)
    def test_all_base_keys_present(self, base: dict, override: dict) -> None:
        result = _deep_merge(base, override)
        for key in base:
            assert key in result


# ---------------------------------------------------------------------------
# load_style
# ---------------------------------------------------------------------------
class TestLoadStyle:
    def test_none_returns_defaults(self) -> None:
        result = load_style(None)
        assert isinstance(result, StyleConfig)
        defaults = _load_default_style_yaml()
        assert result.header_fill_hex == defaults["header_fill_hex"]
        assert result.font_name == defaults["font_name"]
        assert result.font_size == int(defaults["font_size"])

    def test_missing_file_raises_style_config_error(self, tmp_path: object) -> None:
        with pytest.raises(StyleConfigError, match="Style config not found"):
            load_style(str(tmp_path) + "/nonexistent.yaml")  # type: ignore[operator]

    def test_user_override_merges_on_top(self, tmp_path: object) -> None:
        import pathlib

        p = pathlib.Path(str(tmp_path)) / "override.yaml"
        p.write_text('font_name: "Comic Sans"\nfont_size: 14\n')
        result = load_style(str(p))
        assert result.font_name == "Comic Sans"
        assert result.font_size == 14
        # non-overridden keys come from defaults
        defaults = _load_default_style_yaml()
        assert result.header_fill_hex == defaults["header_fill_hex"]

    def test_empty_user_yaml_returns_defaults(self, tmp_path: object) -> None:
        import pathlib

        p = pathlib.Path(str(tmp_path)) / "empty.yaml"
        p.write_text("")
        result = load_style(str(p))
        assert isinstance(result, StyleConfig)

    def test_user_yaml_missing_required_keys_after_merge_raises(self, tmp_path: object) -> None:
        """If bundled defaults somehow lacked keys AND user doesn't supply them,
        StyleConfigError should be raised. We simulate by monkeypatching."""
        import pathlib

        import excel_model.config as config_module

        original = config_module._load_default_style_yaml

        def _broken_defaults() -> dict:
            return {}  # no keys at all

        config_module._load_default_style_yaml = _broken_defaults  # type: ignore[assignment]
        try:
            p = pathlib.Path(str(tmp_path)) / "partial.yaml"
            p.write_text('font_name: "Arial"\n')
            with pytest.raises(StyleConfigError, match="missing required keys"):
                load_style(str(p))
        finally:
            config_module._load_default_style_yaml = original  # type: ignore[assignment]

"""Formula injection validation — rejects dangerous Excel patterns in custom formulas and text labels."""

import re

from excel_model.exceptions import FormulaInjectionError

_FORMULA_INJECTION_PREFIXES = ("=", "+", "-", "@")

# Patterns that indicate potential Excel formula injection (DDE, external data exfiltration)
_DANGEROUS_FORMULA_PATTERNS: tuple[re.Pattern[str], ...] = (
    re.compile(r"\bCMD\b", re.IGNORECASE),
    re.compile(r"\bDDE\b", re.IGNORECASE),
    re.compile(r"\bDDEAUTO\b", re.IGNORECASE),
    re.compile(r"\bWEBSERVICE\s*\(", re.IGNORECASE),
    re.compile(r"\bFILTERXML\s*\(", re.IGNORECASE),
    re.compile(r"\bIMPORTDATA\s*\(", re.IGNORECASE),
    re.compile(r"\bIMPORTFEED\s*\(", re.IGNORECASE),
    re.compile(r"\bIMPORTHTML\s*\(", re.IGNORECASE),
    re.compile(r"\bIMPORTRANGE\s*\(", re.IGNORECASE),
    re.compile(r"\bIMPORTXML\s*\(", re.IGNORECASE),
    re.compile(r"\bCALL\s*\(", re.IGNORECASE),
    re.compile(r"\bREGISTER\.ID\s*\(", re.IGNORECASE),
    re.compile(r"\bEXEC\s*\(", re.IGNORECASE),
    re.compile(r"\bHYPERLINK\s*\(", re.IGNORECASE),
    re.compile(r"\bRTD\s*\(", re.IGNORECASE),
    re.compile(r"\bINDIRECT\s*\(", re.IGNORECASE),
    re.compile(r"\bENCODEURL\s*\(", re.IGNORECASE),
)

# Pipe-based DDE invocations like =CMD|'/c calc'!A0
_DDE_PIPE_RE = re.compile(r"\|.*!", re.IGNORECASE)

# UNC paths that could trigger external network requests
_UNC_PATH_RE = re.compile(r"\\\\[^\\]+\\", re.IGNORECASE)


def validate_custom_formula(formula: str, line_item_key: str) -> None:
    """Reject custom formulas containing dangerous Excel patterns.

    Raises FormulaInjectionError if the formula matches a known injection pattern.
    """
    for pattern in _DANGEROUS_FORMULA_PATTERNS:
        if pattern.search(formula):
            raise FormulaInjectionError(
                f"Line item {line_item_key!r}: custom formula contains dangerous pattern "
                f"{pattern.pattern!r}. Formula: {formula!r}"
            )
    if _DDE_PIPE_RE.search(formula):
        raise FormulaInjectionError(
            f"Line item {line_item_key!r}: custom formula contains a pipe-based DDE pattern. Formula: {formula!r}"
        )
    if _UNC_PATH_RE.search(formula):
        raise FormulaInjectionError(
            f"Line item {line_item_key!r}: custom formula contains a UNC path. Formula: {formula!r}"
        )


def sanitize_cell_text(value: str) -> str:
    """Escape leading formula-injection characters in text written to Excel cells.

    Prefixes with a single-quote so openpyxl stores the value as a plain string
    instead of interpreting it as a formula.  This is a defense-in-depth measure;
    the primary defense is validation via validate_text_field / validate_spec.
    """
    if value and value[0] in _FORMULA_INJECTION_PREFIXES:
        return "'" + value
    return value


def validate_text_field(value: str, field_description: str) -> None:
    """Reject a text value that starts with a formula-injection character.

    Raises FormulaInjectionError with a message identifying the offending field.
    """
    if value and value[0] in _FORMULA_INJECTION_PREFIXES:
        raise FormulaInjectionError(
            f"{field_description}: value {value!r} starts with {value[0]!r}, "
            f"which triggers Excel formula injection. "
            f"Remove or escape the leading character."
        )

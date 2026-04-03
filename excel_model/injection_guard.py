"""Formula injection validation — rejects dangerous Excel patterns in custom formulas."""

import re

from excel_model.exceptions import FormulaInjectionError

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

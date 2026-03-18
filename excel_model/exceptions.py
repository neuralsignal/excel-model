"""Custom exception classes for excel-model."""


class ExcelModelError(Exception):
    """Base exception for all excel-model errors."""


class SpecValidationError(ExcelModelError):
    """Raised when a model spec fails validation."""


class FormulaError(ExcelModelError):
    """Raised when a formula cannot be rendered."""


class InputDataError(ExcelModelError):
    """Raised when input data is missing, malformed, or incompatible."""


class StyleConfigError(ExcelModelError):
    """Raised when style configuration is invalid or missing."""


class FormulaInjectionError(ExcelModelError):
    """Raised when a custom formula contains a potentially dangerous pattern."""

"""
Custom domain exceptions for Excel AI Assistant.
"""

class ExcelAssistantException(Exception):
    """Base exception for all Excel Assistant errors."""
    pass


class ExcelConnectionError(ExcelAssistantException):
    """Raised when connecting to Microsoft Excel fails."""
    pass


class ExcelNotRunningError(ExcelConnectionError):
    """Raised when Microsoft Excel application is not currently running."""
    pass


class WorkbookNotFoundError(ExcelAssistantException):
    """Raised when a specified workbook is not found."""
    pass


class SheetNotFoundError(ExcelAssistantException):
    """Raised when a specified sheet is not found."""
    pass


class RangeInvalidError(ExcelAssistantException):
    """Raised when a cell range coordinate is invalid."""
    pass


class ActionExecutionError(ExcelAssistantException):
    """Raised when an operation on Excel fails during execution."""
    pass


class AmbiguousCommandError(ExcelAssistantException):
    """Raised when user Persian command is ambiguous and needs clarification."""
    def __init__(self, message: str, clarification_prompt: str = "", suggested_options: list = None):
        super().__init__(message)
        self.clarification_prompt = clarification_prompt
        self.suggested_options = suggested_options or []


class SafetyConfirmationRequiredError(ExcelAssistantException):
    """Raised when an action is dangerous (e.g. deleting columns/rows) and needs user confirmation."""
    def __init__(self, action_name: str, details: str, affected_range: str = ""):
        super().__init__(f"Confirmation required for '{action_name}': {details}")
        self.action_name = action_name
        self.details = details
        self.affected_range = affected_range

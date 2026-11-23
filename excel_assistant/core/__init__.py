"""
Core Excel Connection and Driver Modules.
"""
from excel_assistant.core.interfaces import ISheet, IExcelDriver, WorkbookInfo, SheetStructure
from excel_assistant.core.connection import ExcelConnectionManager
from excel_assistant.core.openpyxl_driver import OpenpyxlDriver, OpenpyxlSheet
from excel_assistant.core.xlwings_driver import XlwingsDriver, XlwingsSheet
from excel_assistant.core.exceptions import (
    ExcelAssistantException, ExcelConnectionError, ExcelNotRunningError,
    WorkbookNotFoundError, SheetNotFoundError, RangeInvalidError,
    ActionExecutionError, AmbiguousCommandError, SafetyConfirmationRequiredError
)

__all__ = [
    "ISheet",
    "IExcelDriver",
    "WorkbookInfo",
    "SheetStructure",
    "ExcelConnectionManager",
    "OpenpyxlDriver",
    "OpenpyxlSheet",
    "XlwingsDriver",
    "XlwingsSheet",
    "ExcelAssistantException",
    "ExcelConnectionError",
    "ExcelNotRunningError",
    "WorkbookNotFoundError",
    "SheetNotFoundError",
    "RangeInvalidError",
    "ActionExecutionError",
    "AmbiguousCommandError",
    "SafetyConfirmationRequiredError",
]

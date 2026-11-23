"""
Base definitions and result containers for Excel actions.
"""
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

from excel_assistant.core.interfaces import ISheet
from excel_assistant.core.connection import ExcelConnectionManager
from excel_assistant.analyzer.inspector import SheetAnalysis


@dataclass
class ActionResult:
    success: bool
    message_persian: str
    affected_range: str = ""
    data: Any = None
    requires_confirmation: bool = False
    confirmation_prompt: str = ""


@dataclass
class ActionContext:
    manager: ExcelConnectionManager
    sheet: ISheet
    sheet_analysis: Optional[SheetAnalysis] = None
    confirmed: bool = False

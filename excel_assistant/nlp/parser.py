"""
Unified Persian Intent Parser for Excel Assistant.
Combines fast deterministic Persian semantic rules with optional LLM reasoning.
"""
from typing import Any, Dict, Optional, Tuple

from excel_assistant.nlp.intent_models import (
    ExcelAction,
    ColumnArithmeticAction,
    ColumnStatsAction,
    ApplyDiscountAction,
    InsertDateAction,
    PreventDuplicatesAction,
    HighlightConditionAction,
    HighlightDuplicatesAction,
    DeleteEmptyRowsAction,
    DeleteDuplicatesAction,
    DeleteRowsOrColsAction,
    SortDataAction,
    AutoFitColumnsAction,
    CreateTableAction,
    FormatCellsAction,
    ApplyFormulaAction,
    AnalyzeSheetAction,
    UndoLastAction,
)
from excel_assistant.nlp.persian_rules import PersianRuleParser
from excel_assistant.nlp.llm_client import LLMClient
from excel_assistant.analyzer.inspector import SheetAnalysis
from excel_assistant.core.exceptions import AmbiguousCommandError


ACTION_CLASS_MAP = {
    "column_arithmetic": ColumnArithmeticAction,
    "column_stats": ColumnStatsAction,
    "apply_discount": ApplyDiscountAction,
    "insert_date": InsertDateAction,
    "prevent_duplicates": PreventDuplicatesAction,
    "highlight_condition": HighlightConditionAction,
    "highlight_duplicates": HighlightDuplicatesAction,
    "delete_empty_rows": DeleteEmptyRowsAction,
    "delete_duplicates": DeleteDuplicatesAction,
    "delete_rows_or_cols": DeleteRowsOrColsAction,
    "sort_data": SortDataAction,
    "autofit_columns": AutoFitColumnsAction,
    "create_table": CreateTableAction,
    "format_cells": FormatCellsAction,
    "apply_formula": ApplyFormulaAction,
    "analyze_sheet": AnalyzeSheetAction,
    "undo_last_action": UndoLastAction,
}


class PersianIntentParser:
    """Central parser resolving user Persian instructions into executable actions."""

    def __init__(self, llm_client: Optional[LLMClient] = None):
        self.llm_client = llm_client or LLMClient()

    def parse(self, command: str, sheet_analysis: Optional[SheetAnalysis] = None) -> ExcelAction:
        if not command or not command.strip():
            raise AmbiguousCommandError("دستور نمی‌تواند خالی باشد.")

        header_map = sheet_analysis.header_to_col if sheet_analysis else {}

        # 1. Deterministic Rule Parser
        action = PersianRuleParser.parse(command, header_map=header_map)
        if action:
            return action

        # 2. LLM fallback if configured
        if self.llm_client.is_configured:
            ctx = None
            if sheet_analysis:
                ctx = {
                    "sheet_name": sheet_analysis.sheet_name,
                    "used_range": sheet_analysis.used_range_str,
                    "headers": sheet_analysis.headers
                }
            llm_result = self.llm_client.parse_instruction(command, sheet_context=ctx)
            if llm_result and "action_type" in llm_result:
                act_type = llm_result["action_type"]
                cls = ACTION_CLASS_MAP.get(act_type)
                if cls:
                    try:
                        return cls(**llm_result)
                    except Exception:
                        pass

        # 3. If could not parse, raise AmbiguousCommandError with Persian suggestions
        suggestions = [
            "ستون A را در ستون B ضرب کن و جواب را در C بنویس",
            "میانگین ستون D را حساب کن",
            "داده‌های تکراری را زرد کن",
            "سلول‌های بیشتر از 100 را سبز کن",
            "ردیف‌های خالی را حذف کن",
            "اطلاعات را از بیشترین به کمترین مرتب کن",
            "یک جدول بساز"
        ]
        raise AmbiguousCommandError(
            message=f"متوجه دستور «{command}» نشدم.",
            clarification_prompt="لطفاً دستور را شفاف‌تر بیان کنید یا از نمونه‌های زیر استفاده نمایید:",
            suggested_options=suggestions
        )

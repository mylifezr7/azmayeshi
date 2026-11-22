"""
Central Action Executor and Pipeline for Excel AI Assistant.
Orchestrates safety validation, transaction snapshotting, action dispatch, and error handling.
"""
from typing import Optional

from excel_assistant.actions.base import ActionResult, ActionContext
from excel_assistant.actions.math_actions import MathActionHandler
from excel_assistant.actions.format_actions import FormatActionHandler
from excel_assistant.actions.data_actions import DataActionHandler
from excel_assistant.actions.undo_manager import UndoManager
from excel_assistant.analyzer.inspector import SheetInspector, WorkbookInspector
from excel_assistant.analyzer.suggestions import generate_smart_suggestions
from excel_assistant.safety.security import SecurityPolicy
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
    SortDataAction,
    AutoFitColumnsAction,
    CreateTableAction,
    FormatCellsAction,
    AnalyzeSheetAction,
    SalesSummaryAction,
    ProfitCalculationAction,
    UndoLastAction,
)
from excel_assistant.core.exceptions import (
    SafetyConfirmationRequiredError, ActionExecutionError
)


class ActionExecutor:
    """Executes validated ExcelAction instances with undo tracking and security guards."""

    def __init__(self, undo_manager: Optional[UndoManager] = None):
        self.undo_manager = undo_manager or UndoManager()

    def execute(self, action: ExcelAction, ctx: ActionContext) -> ActionResult:
        # 1. Handle Undo action separately
        if isinstance(action, UndoLastAction):
            if self.undo_manager.can_undo():
                success = self.undo_manager.undo(ctx.sheet)
                if success:
                    return ActionResult(
                        success=True,
                        message_persian="↩️ آخرین عملیات با موفقیت لغو شد و داده‌های اکسل به حالت قبل بازگشتند."
                    )
            return ActionResult(
                success=False,
                message_persian="⚠️ هیچ عملیات قبلی برای لغو کردن در تاریخچه موجود نیست."
            )

        # 2. Handle Analyze Sheet action
        if isinstance(action, AnalyzeSheetAction):
            sa = SheetInspector.inspect_sheet(ctx.sheet)
            wb_analysis = WorkbookInspector.inspect(ctx.manager.active_driver)
            summary = WorkbookInspector.format_persian_summary(wb_analysis)
            suggestions = generate_smart_suggestions(sa)
            if suggestions:
                summary += "\n\n💡 **پیشنهادهای هوشمند:**\n" + "\n".join(f"• {s}" for s in suggestions)
            return ActionResult(success=True, message_persian=summary, data=sa)

        # 3. Security & Confirmation check for destructive actions
        try:
            SecurityPolicy.check_permission(action, confirmed=ctx.confirmed)
        except SafetyConfirmationRequiredError as e:
            return ActionResult(
                success=False,
                message_persian=f"⚠️ {e.details}\nآیا مطمئن هستید که می‌خواهید این عملیات انجام شود؟",
                requires_confirmation=True,
                confirmation_prompt=e.details
            )

        # 4. Save state snapshot for undo before mutating
        self.undo_manager.save_state(ctx.sheet)

        # 5. Dispatch action
        try:
            if isinstance(action, ColumnArithmeticAction):
                return MathActionHandler.execute_arithmetic(action, ctx)

            elif isinstance(action, ApplyDiscountAction):
                return MathActionHandler.execute_discount(action, ctx)

            elif isinstance(action, SalesSummaryAction):
                return MathActionHandler.execute_sales_summary(action, ctx)

            elif isinstance(action, ProfitCalculationAction):
                return MathActionHandler.execute_profit_calculation(action, ctx)

            elif isinstance(action, ColumnStatsAction):
                return MathActionHandler.execute_stats(action, ctx)

            elif isinstance(action, HighlightConditionAction):
                return FormatActionHandler.execute_highlight_condition(action, ctx)

            elif isinstance(action, HighlightDuplicatesAction):
                return FormatActionHandler.execute_highlight_duplicates(action, ctx)

            elif isinstance(action, DeleteEmptyRowsAction):
                return DataActionHandler.execute_delete_empty_rows(action, ctx)

            elif isinstance(action, DeleteDuplicatesAction):
                return DataActionHandler.execute_delete_duplicates(action, ctx)

            elif isinstance(action, SortDataAction):
                return DataActionHandler.execute_sort(action, ctx)

            elif isinstance(action, InsertDateAction):
                return DataActionHandler.execute_insert_date(action, ctx)

            elif isinstance(action, PreventDuplicatesAction):
                return DataActionHandler.execute_prevent_duplicates(action, ctx)

            elif isinstance(action, AutoFitColumnsAction):
                return FormatActionHandler.execute_autofit(action, ctx)

            elif isinstance(action, CreateTableAction):
                return FormatActionHandler.execute_create_table(action, ctx)

            elif isinstance(action, FormatCellsAction):
                return FormatActionHandler.execute_format_cells(action, ctx)

            else:
                return ActionResult(
                    success=False,
                    message_persian=f"عملیات ناشناخته: {action.action_type}"
                )

        except Exception as e:
            return ActionResult(
                success=False,
                message_persian=f"❌ خطا در اجرای عملیات روی اکسل: {e}"
            )

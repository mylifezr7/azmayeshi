"""
Central Action Executor and Pipeline for Excel AI Assistant.
Orchestrates safety validation, transaction snapshotting, action dispatch, and error handling.
"""
from typing import Optional

from excel_assistant.actions.base import ActionResult, ActionContext
from excel_assistant.actions.math_actions import MathActionHandler
from excel_assistant.actions.format_actions import FormatActionHandler
from excel_assistant.actions.data_actions import DataActionHandler
from excel_assistant.actions.accounting_actions import AccountingActionHandler
from excel_assistant.actions.business_actions import BusinessActionHandler
from excel_assistant.actions.undo_manager import UndoManager
from excel_assistant.analyzer.inspector import SheetInspector, WorkbookInspector
from excel_assistant.analyzer.suggestions import generate_smart_suggestions
from excel_assistant.analyzer.deep_auditor import DeepFileAuditor
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
    FilterDataAction,
    ConvertTextToNumberAction,
    SplitTextAction,
    MergeCellsAction,
    SetCellFormatAction,
    PriceAdjustmentAction,
    ReconciliationAction,
    InvoiceAction,
    InventoryAlertAction,
    TopSellingAction,
    CustomerLedgerAction,
    TrialBalanceAction,
    AccountBalanceAction,
    PayrollAction,
    DepreciationAction,
    FinancialStatementAction,
    DeepFileAuditAction,
    AutoChartAction,
    BackupAction,
    SaveWorkbookAction,
    PdfExportAction,
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
        is_mutating = not isinstance(action, (
            AnalyzeSheetAction, DeepFileAuditAction, SaveWorkbookAction, BackupAction, PdfExportAction
        ))
        if is_mutating:
            self.undo_manager.begin_action(ctx.sheet)

        # 5. Dispatch action
        try:
            res = self._dispatch(action, ctx)
            if is_mutating:
                if res.success:
                    self.undo_manager.commit_action(ctx.sheet)
                else:
                    self.undo_manager.discard_pending()
            return res
        except Exception as e:
            if is_mutating:
                self.undo_manager.discard_pending()
            return ActionResult(
                success=False,
                message_persian=f"❌ خطا در اجرای عملیات روی اکسل: {e}"
            )

    def _dispatch(self, action: ExcelAction, ctx: ActionContext) -> ActionResult:
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

        elif isinstance(action, FilterDataAction):
            return DataActionHandler.execute_filter(action, ctx)

        elif isinstance(action, ConvertTextToNumberAction):
            return DataActionHandler.execute_convert_text_to_number(action, ctx)

        elif isinstance(action, SplitTextAction):
            return DataActionHandler.execute_split_text(action, ctx)

        elif isinstance(action, MergeCellsAction):
            return DataActionHandler.execute_merge_cells(action, ctx)

        elif isinstance(action, SetCellFormatAction):
            return DataActionHandler.execute_set_cell_format(action, ctx)

        elif isinstance(action, PriceAdjustmentAction):
            return BusinessActionHandler.execute_price_adjustment(action, ctx)

        elif isinstance(action, ReconciliationAction):
            return AccountingActionHandler.execute_reconciliation(action, ctx)

        elif isinstance(action, InvoiceAction):
            return BusinessActionHandler.execute_invoice(action, ctx)

        elif isinstance(action, InventoryAlertAction):
            return BusinessActionHandler.execute_inventory_alert(action, ctx)

        elif isinstance(action, TopSellingAction):
            return BusinessActionHandler.execute_top_selling(action, ctx)

        elif isinstance(action, CustomerLedgerAction):
            return BusinessActionHandler.execute_customer_ledger(action, ctx)

        elif isinstance(action, TrialBalanceAction):
            return AccountingActionHandler.execute_trial_balance(action, ctx)

        elif isinstance(action, AccountBalanceAction):
            return AccountingActionHandler.execute_account_balance(action, ctx)

        elif isinstance(action, PayrollAction):
            return AccountingActionHandler.execute_payroll(action, ctx)

        elif isinstance(action, DepreciationAction):
            return AccountingActionHandler.execute_depreciation(action, ctx)

        elif isinstance(action, FinancialStatementAction):
            return AccountingActionHandler.execute_financial_statement(action, ctx)

        elif isinstance(action, DeepFileAuditAction):
            audit_res = DeepFileAuditor.audit_sheet(ctx.sheet)
            return ActionResult(success=True, message_persian=audit_res.persian_summary, data=audit_res)

        elif isinstance(action, AutoChartAction):
            return BusinessActionHandler.execute_auto_chart(action, ctx)

        elif isinstance(action, BackupAction):
            return BusinessActionHandler.execute_backup(action, ctx)

        elif isinstance(action, SaveWorkbookAction):
            saved_path = ctx.manager.save(action.file_path)
            return ActionResult(
                success=True,
                message_persian=f"💾 فایل اکسل با موفقیت ذخیره شد:\n**`{saved_path}`**",
                data=saved_path
            )

        elif isinstance(action, PdfExportAction):
            return BusinessActionHandler.execute_pdf_export(action, ctx)

        else:
            return ActionResult(
                success=False,
                message_persian=f"عملیات ناشناخته: {action.action_type}"
            )

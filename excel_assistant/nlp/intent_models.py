"""
Pydantic action schemas for Excel AI Assistant operations.
"""
from typing import Any, Dict, List, Literal, Optional, Union
from pydantic import BaseModel, Field


class BaseActionModel(BaseModel):
    action_type: str
    description_persian: str = ""
    is_destructive: bool = False
    affected_range: str = ""


class ColumnArithmeticAction(BaseActionModel):
    action_type: Literal["column_arithmetic"] = "column_arithmetic"
    operation: Literal["multiply", "add", "subtract", "divide", "average"]  # ضرب، جمع، تفریق، تقسیم، میانگین سطری
    col1: Optional[str] = "A"  # e.g. "A" or header name
    col2: Optional[str] = "B"  # e.g. "B"
    columns: Optional[List[str]] = None  # multi-column operands e.g. ["2", "4", "5"] or ["B", "D", "E"]
    target_col: Optional[str] = None  # e.g. "C" or "H"
    as_formula: bool = True  # generate Excel formula like =A2*B2 or compute values


class ColumnStatsAction(BaseActionModel):
    action_type: Literal["column_stats"] = "column_stats"
    stat_type: Literal["sum", "average", "median", "max", "min", "count", "counta", "percentage"]  # میانگین، مجموع، میانه، بیشترین، کمترین، تعداد، درصد
    source_col: str  # e.g. "D"
    target_cell: Optional[str] = None  # cell to write result into, or just return result in chat


class HighlightConditionAction(BaseActionModel):
    action_type: Literal["highlight_condition"] = "highlight_condition"
    target_col: Optional[str] = None
    target_range: Optional[str] = None
    condition: Literal["gt", "gte", "lt", "lte", "eq", "neq", "between", "duplicates"]  # بیشتر از، کمتر از، مساوی، تکراری
    value: Optional[Union[float, int, str]] = None
    value2: Optional[Union[float, int, str]] = None
    color_hex: str = "#90EE90"  # default light green


class HighlightDuplicatesAction(BaseActionModel):
    action_type: Literal["highlight_duplicates"] = "highlight_duplicates"
    target_col: Optional[str] = None
    target_range: Optional[str] = None
    color_hex: str = "#FFFF00"  # default yellow
    text_only: bool = True  # Only highlight textual/typographic duplicates (skip pure numbers)
    distinct_colors: bool = True  # Assign distinct color to each duplicate value group


class PreventDuplicatesAction(BaseActionModel):
    action_type: Literal["prevent_duplicates"] = "prevent_duplicates"
    target_col: str = "A"  # e.g. "L" or "12"
    error_title: str = "خطای ورود داده تکراری"
    error_message: str = "این مقدار قبلاً در این ستون ثبت شده است و تکراری می‌باشد. لطفاً کد یکتای دیگری وارد کنید."
    highlight_existing: bool = True


class RemoveUniqueValidationAction(BaseActionModel):
    action_type: Literal["remove_unique_validation"] = "remove_unique_validation"
    target_col: str = "A"
    clear_highlight: bool = True


class DeleteEmptyRowsAction(BaseActionModel):
    action_type: Literal["delete_empty_rows"] = "delete_empty_rows"
    is_destructive: bool = True


class DeleteDuplicatesAction(BaseActionModel):
    action_type: Literal["delete_duplicates"] = "delete_duplicates"
    target_col: Optional[str] = None
    is_destructive: bool = True


class DeleteRowsOrColsAction(BaseActionModel):
    action_type: Literal["delete_rows_or_cols"] = "delete_rows_or_cols"
    target_type: Literal["row", "column"]
    targets: List[Union[int, str]]
    is_destructive: bool = True


class SortDataAction(BaseActionModel):
    action_type: Literal["sort_data"] = "sort_data"
    key_col: str
    ascending: bool = False  # False = descending (بیشترین به کمترین)


class AutoFitColumnsAction(BaseActionModel):
    action_type: Literal["autofit_columns"] = "autofit_columns"
    columns: Optional[List[str]] = None


class CreateTableAction(BaseActionModel):
    action_type: Literal["create_table"] = "create_table"
    range_str: Optional[str] = None
    table_name: Optional[str] = None


class FormatCellsAction(BaseActionModel):
    action_type: Literal["format_cells"] = "format_cells"
    target_range: str
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    font_size: Optional[float] = None
    font_name: Optional[str] = None
    bg_color: Optional[str] = None
    font_color: Optional[str] = None


class ApplyDiscountAction(BaseActionModel):
    action_type: Literal["apply_discount"] = "apply_discount"
    price_col: str  # Original price column e.g. "B" or "2"
    discount_col: Optional[str] = None  # Discount percentage column e.g. "C" or "3"
    discount_percent: Optional[float] = None  # Fixed discount percentage e.g. 15.0
    target_col: Optional[str] = None  # Final price after discount column e.g. "D" or "4"
    discount_amount_col: Optional[str] = None  # Optional discount amount column e.g. "E"


class InsertDateAction(BaseActionModel):
    action_type: Literal["insert_date"] = "insert_date"
    target_col: str = "A"  # e.g. "A" or "1"
    start_row: int = 1     # e.g. 1
    end_row: int = 30      # e.g. 30
    year: int = 1403      # e.g. 1403
    month: int = 8        # e.g. 8
    day: int = 29         # e.g. 29
    mode: Literal["fixed", "sequential"] = "fixed"
    add_header: bool = False


class ApplyFormulaAction(BaseActionModel):
    action_type: Literal["apply_formula"] = "apply_formula"
    target_cell_or_range: str
    formula_template: str  # e.g. "=SUM(B2:B20)" or "=A{row}*B{row}"


class AnalyzeSheetAction(BaseActionModel):
    action_type: Literal["analyze_sheet"] = "analyze_sheet"
    sheet_name: Optional[str] = None


class SalesSummaryAction(BaseActionModel):
    action_type: Literal["sales_summary"] = "sales_summary"
    date_col: Optional[str] = None  # Column containing dates, e.g. "A"
    amount_col: Optional[str] = None  # Column containing sales amount, e.g. "D"
    period: Literal["all", "daily", "monthly", "yearly", "total"] = "all"
    create_report_sheet: bool = True
    description_persian: str = "محاسبه فروش روزانه، ماهانه، سالانه و کل"


class ProfitCalculationAction(BaseActionModel):
    action_type: Literal["calculate_profit"] = "calculate_profit"
    sales_sheet: Optional[str] = None  # e.g. "Sheet1" or "شیت 1"
    purchase_sheet: Optional[str] = None  # e.g. "Sheet2" or "شیت 2"
    sales_qty_col: Optional[str] = None
    sales_price_col: Optional[str] = None
    purchase_qty_col: Optional[str] = None
    purchase_price_col: Optional[str] = None
    date_col: Optional[str] = None
    profit_col: Optional[str] = None
    sales_col: Optional[str] = None
    purchase_col: Optional[str] = None
    valuation_method: Literal["fifo", "lifo", "average", "compare"] = "fifo"
    period: Literal["all", "daily", "monthly", "yearly", "total", "by_item"] = "all"
    create_report_sheet: bool = True
    description_persian: str = "محاسبه سود (روزانه، ماهانه، سالانه و کل)"


class FilterDataAction(BaseActionModel):
    action_type: Literal["filter_data"] = "filter_data"
    target_col: Optional[str] = None
    condition: Literal["gt", "gte", "lt", "lte", "eq", "neq", "contains", "autofilter", "clear"] = "autofilter"
    value: Optional[Union[float, int, str]] = None
    range_str: Optional[str] = None
    description_persian: str = "فیلتر هوشمند داده‌ها"


class ConvertTextToNumberAction(BaseActionModel):
    action_type: Literal["convert_text_to_number"] = "convert_text_to_number"
    target_col: Optional[str] = "A"
    range_str: Optional[str] = None
    description_persian: str = "تبدیل متن به عدد واقعی و اصلاح داده‌های عددی"


class SplitTextAction(BaseActionModel):
    action_type: Literal["split_text"] = "split_text"
    source_col: str = "A"
    delimiter: str = " "
    target_col1: Optional[str] = None
    target_col2: Optional[str] = None
    header1: str = "نام"
    header2: str = "نام خانوادگی"
    description_persian: str = "جدا کردن متن و تفکیک ستون"


class MergeCellsAction(BaseActionModel):
    action_type: Literal["merge_cells"] = "merge_cells"
    range_str: str = "A1:B1"
    align_center: bool = True
    description_persian: str = "ادغام سلول‌های محدوده"


class SetCellFormatAction(BaseActionModel):
    action_type: Literal["set_cell_format"] = "set_cell_format"
    target: str = "A"  # e.g. "B" or "A1:C10"
    format_type: Literal["currency", "percentage", "date", "number", "text", "thousands"] = "currency"
    custom_format: Optional[str] = None
    description_persian: str = "تغییر فرمت نمایش سلول‌ها"


class PriceAdjustmentAction(BaseActionModel):
    action_type: Literal["price_adjustment"] = "price_adjustment"
    target_col: str = "C"
    percentage: float = 15.0  # e.g. 15.0 for 15% increase, -10.0 for 10% decrease
    operation: Literal["increase", "decrease"] = "increase"
    in_place: bool = False
    target_result_col: Optional[str] = None
    description_persian: str = "افزایش یا کاهش درصدی قیمت‌ها"


class ReconciliationAction(BaseActionModel):
    action_type: Literal["reconcile_columns"] = "reconcile_columns"
    col1: str = "A"
    col2: str = "B"
    highlight_mismatches: bool = True
    create_report: bool = True
    description_persian: str = "مغایرت‌گیری دو ستون و شناسایی اختلافات"


class InvoiceAction(BaseActionModel):
    action_type: Literal["create_invoice"] = "create_invoice"
    invoice_type: Literal["sales", "purchase"] = "sales"  # فروش یا خرید
    invoice_number: Optional[str] = None
    buyer_name: Optional[str] = None
    seller_name: Optional[str] = None
    date_str: Optional[str] = None
    tax_percent: float = 10.0  # 10% VAT
    include_sample_items: bool = True
    description_persian: str = "ثبت و صدور فاکتور رسمی"


class InventoryAlertAction(BaseActionModel):
    action_type: Literal["inventory_alert"] = "inventory_alert"
    product_col: Optional[str] = None
    stock_col: Optional[str] = None
    min_threshold: float = 5.0
    purchase_price_col: Optional[str] = None
    sales_sheet: Optional[str] = None
    purchase_sheet: Optional[str] = None
    valuation_method: Literal["fifo", "lifo", "average", "latest", "compare"] = "fifo"
    highlight_low_stock: bool = True
    description_persian: str = "کنترل موجودی انبار و هشدار کمبود موجودی"


class InventoryValuationAction(BaseActionModel):
    action_type: Literal["inventory_valuation"] = "inventory_valuation"
    method: Literal["fifo", "lifo", "average", "compare"] = "compare"
    product_col: Optional[str] = None
    stock_col: Optional[str] = None
    purchase_price_col: Optional[str] = None
    sales_sheet: Optional[str] = None
    purchase_sheet: Optional[str] = None
    min_threshold: float = 5.0
    create_report_sheet: bool = True
    highlight_low_stock: bool = True
    description_persian: str = "ارزش‌گذاری موجودی انبار به روش‌های فایفو، لایفو و میانگین موزون"


class TopSellingAction(BaseActionModel):
    action_type: Literal["top_selling"] = "top_selling"
    product_col: Optional[str] = None
    qty_col: Optional[str] = None
    amount_col: Optional[str] = None
    top_n: int = 5
    description_persian: str = "گزارش پرفروش‌ترین کالاها"


class CustomerLedgerAction(BaseActionModel):
    action_type: Literal["customer_ledger"] = "customer_ledger"
    customer_col: Optional[str] = None
    debit_col: Optional[str] = None   # بدهی / فاکتورها
    credit_col: Optional[str] = None  # طلب / واریزی‌ها
    description_persian: str = "محاسبه بدهی و طلب مشتریان"


class TrialBalanceAction(BaseActionModel):
    action_type: Literal["trial_balance"] = "trial_balance"
    account_col: Optional[str] = None
    debit_col: Optional[str] = None
    credit_col: Optional[str] = None
    description_persian: str = "تهیه تراز آزمایشی ۴ ستونه"


class AccountBalanceAction(BaseActionModel):
    action_type: Literal["account_balance"] = "account_balance"
    debit_col: Optional[str] = None
    credit_col: Optional[str] = None
    target_balance_col: Optional[str] = None
    description_persian: str = "محاسبه مانده حساب بدهکار و بستانکار"


class PayrollAction(BaseActionModel):
    action_type: Literal["payroll_calculation"] = "payroll_calculation"
    name_col: Optional[str] = None
    base_salary_col: Optional[str] = None
    overtime_hours_col: Optional[str] = None
    insurance_rate_employee: float = 7.0   # 7% سهم کارگر
    insurance_rate_employer: float = 23.0  # 23% سهم کارفرما
    tax_bracket: float = 10.0              # 10% مالیات
    tax_exempt_ceiling: float = 12000000.0 # سقف معافیت
    create_payslip_sheet: bool = True
    description_persian: str = "محاسبه حقوق و دستمزد، اضافه‌کاری، بیمه و مالیات"


class AttendanceAction(BaseActionModel):
    action_type: Literal["attendance_calculation"] = "attendance_calculation"
    employee_name: Optional[str] = None
    month_name: Optional[str] = "آبان ۱۴۰۵"
    standard_daily_hours: float = 8.0
    date_col: Optional[str] = None
    day_col: Optional[str] = None
    entry_col: Optional[str] = None
    exit_col: Optional[str] = None
    status_col: Optional[str] = None
    create_summary_sheet: bool = True
    description_persian: str = "محاسبه ورود و خروج، کارکرد، اضافه‌کار، کسری کار و حضور و غیاب"


class DepreciationAction(BaseActionModel):
    action_type: Literal["asset_depreciation"] = "asset_depreciation"
    cost: Optional[float] = 100000000.0
    salvage_value: float = 10000000.0
    useful_life_years: int = 5
    method: Literal["straight_line", "declining_balance"] = "straight_line"
    depreciation_rate: Optional[float] = None
    asset_name: str = "دارایی ثابت"
    description_persian: str = "محاسبه استهلاک دارایی‌ها (خط مستقیم و نزولی)"


class FinancialStatementAction(BaseActionModel):
    action_type: Literal["financial_statement"] = "financial_statement"
    statement_type: Literal["income_statement", "balance_sheet"] = "income_statement"
    description_persian: str = "تهیه صورت سود و زیان یا ترازنامه استاندارد"


class DeepFileAuditAction(BaseActionModel):
    action_type: Literal["deep_file_audit"] = "deep_file_audit"
    description_persian: str = "تحلیل هوشمند و عمیق فایل اکسل (سلول‌های خالی، خطاهای فرمول، تکراری، ستون ناقص)"


class AutoChartAction(BaseActionModel):
    action_type: Literal["auto_chart"] = "auto_chart"
    chart_type: Literal["bar", "col", "line", "pie"] = "col"
    title: str = "نمودار تحلیل داده‌ها"
    label_col: Optional[str] = None
    value_col: Optional[str] = None
    description_persian: str = "رسم نمودار گرافیکی و گزارش‌ساز خودکار"


class BackupAction(BaseActionModel):
    action_type: Literal["backup_workbook"] = "backup_workbook"
    description_persian: str = "ذخیره نسخه پشتیبان از فایل اکسل"


class SaveWorkbookAction(BaseActionModel):
    action_type: Literal["save_workbook"] = "save_workbook"
    file_path: Optional[str] = None
    description_persian: str = "ذخیره تغییرات فایل اکسل"


class PdfExportAction(BaseActionModel):
    action_type: Literal["export_pdf"] = "export_pdf"
    output_filename: Optional[str] = None
    sheet_name: Optional[str] = None
    description_persian: str = "تولید خروجی PDF از فایل اکسل"


class UndoLastAction(BaseActionModel):
    action_type: Literal["undo_last_action"] = "undo_last_action"


ExcelAction = Union[
    ColumnArithmeticAction,
    ColumnStatsAction,
    ApplyDiscountAction,
    InsertDateAction,
    PreventDuplicatesAction,
    RemoveUniqueValidationAction,
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
    InventoryValuationAction,
    TopSellingAction,
    CustomerLedgerAction,
    TrialBalanceAction,
    AccountBalanceAction,
    PayrollAction,
    AttendanceAction,
    DepreciationAction,
    FinancialStatementAction,
    DeepFileAuditAction,
    AutoChartAction,
    BackupAction,
    SaveWorkbookAction,
    PdfExportAction,
    UndoLastAction,
]

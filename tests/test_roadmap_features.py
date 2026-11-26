"""
Comprehensive Test Suite for MUYUM Next Version Features:
1. Accounting (قابلیت‌های حسابداری):
   - مغایرت‌گیری دو ستون (Reconciliation)
   - تراز آزمایشی (Trial Balance)
   - محاسبه مانده حساب (Running Balance)
   - محاسبه حقوق و دستمزد، اضافه‌کاری، بیمه و مالیات (Payroll)
   - استهلاک دارایی‌ها (Depreciation)
   - صورت سود و زیان و ترازنامه (Financial Statements)
2. Sales & Inventory (قابلیت‌های فروش و انبارداری):
   - ثبت فاکتور فروش و خرید (Invoices)
   - مدیریت موجودی، هشدار کمبود و ارزش کل انبار (Inventory Alerts)
   - پرفروش‌ترین کالاها (Top-Selling)
   - بدهی و طلب مشتریان (Customer Balances)
   - افزایش و کاهش درصدی قیمت‌ها (Price Adjustment)
3. Smart & Management Capabilities (امکانات هوشمند و مدیریتی):
   - میانگین سطری بین ستون‌ها
   - تحلیل هوشمند سلامت و خطاهای فایل (Deep File Health Audit)
   - رسم خودکار نمودار گرافیکی (Auto Chart)
   - ذخیره نسخه پشتیبان (Backup)
   - خروجی PDF استاندارد (ReportLab)
   - Natural Persian Command Parsing
"""
import os
import pytest

from excel_assistant.core.openpyxl_driver import OpenpyxlDriver
from excel_assistant.core.connection import ExcelConnectionManager
from excel_assistant.actions.base import ActionContext
from excel_assistant.actions.accounting_actions import AccountingActionHandler
from excel_assistant.actions.business_actions import BusinessActionHandler
from excel_assistant.actions.executor import ActionExecutor
from excel_assistant.analyzer.deep_auditor import DeepFileAuditor
from excel_assistant.nlp.persian_rules import PersianRuleParser
from excel_assistant.nlp.intent_models import (
    ReconciliationAction,
    TrialBalanceAction,
    AccountBalanceAction,
    PayrollAction,
    DepreciationAction,
    FinancialStatementAction,
    InvoiceAction,
    InventoryAlertAction,
    TopSellingAction,
    CustomerLedgerAction,
    PriceAdjustmentAction,
    ColumnArithmeticAction,
    DeepFileAuditAction,
    AutoChartAction,
    BackupAction,
    PdfExportAction,
)


@pytest.fixture
def test_manager():
    driver = OpenpyxlDriver()
    driver.connect()
    manager = ExcelConnectionManager()
    manager.active_driver = driver
    return manager


# 1. Reconciliation Test (مغایرت‌گیری دو ستون)
def test_reconciliation_action(test_manager):
    sheet = test_manager.get_active_sheet()
    sheet.set_value("A1", "صورت حساب بانک")
    sheet.set_value("B1", "دفاتر مالی شرکت")
    # Row 2: match
    sheet.set_value("A2", 1000000)
    sheet.set_value("B2", 1000000)
    # Row 3: discrepancy
    sheet.set_value("A3", 2500000)
    sheet.set_value("B3", 2300000)
    # Row 4: match
    sheet.set_value("A4", 500000)
    sheet.set_value("B4", 500000)

    action = ReconciliationAction(col1="A", col2="B", highlight_mismatches=True, create_report=True)
    ctx = ActionContext(manager=test_manager, sheet=sheet)
    res = AccountingActionHandler.execute_reconciliation(action, ctx)

    assert res.success is True
    assert "گزارش مغایرت‌گیری" in res.message_persian
    assert "2" in res.message_persian  # 2 matched
    assert "1" in res.message_persian  # 1 mismatch

    # Check status column values
    assert sheet.get_value("C2") == "منطبق ✅"
    assert sheet.get_value("C3") == "⚠️ مغایرت عددی"
    assert sheet.get_value("D3") == 200000  # 2500000 - 2300000
    assert "مغایرت_گیری" in test_manager.get_sheet_names()


# 2. Trial Balance Test (تراز آزمایشی)
def test_trial_balance_action(test_manager):
    sheet = test_manager.get_active_sheet()
    action = TrialBalanceAction()
    ctx = ActionContext(manager=test_manager, sheet=sheet)
    res = AccountingActionHandler.execute_trial_balance(action, ctx)

    assert res.success is True
    assert "تراز آزمایشی چهار ستونه" in res.message_persian
    assert "تراز_آزمایشی" in test_manager.get_sheet_names()

    tb = test_manager.get_sheet("تراز_آزمایشی")
    assert tb.get_value("A1") == "تراز آزمایشی چهار ستونه حساب‌ها"
    assert "IF(C5>D5" in str(tb.get_formula("E5"))
    assert "=SUM(C5:" in str(tb.get_formula("C17"))


# 3. Account Balance Test (محاسبه مانده حساب جاری)
def test_account_balance_action(test_manager):
    sheet = test_manager.get_active_sheet()
    sheet.set_value("A1", "تاریخ")
    sheet.set_value("B1", "شرح سند")
    sheet.set_value("C1", "بدهکار")
    sheet.set_value("D1", "بستانکار")

    sheet.set_value("C2", 1500000)
    sheet.set_value("D2", 0)
    sheet.set_value("C3", 500000)
    sheet.set_value("D3", 800000)

    action = AccountBalanceAction(debit_col="C", credit_col="D")
    ctx = ActionContext(manager=test_manager, sheet=sheet)
    res = AccountingActionHandler.execute_account_balance(action, ctx)

    assert res.success is True
    assert "مانده حساب" in res.message_persian
    # Check formulas
    assert str(sheet.get_formula("E2")) == "=C2-D2"
    assert str(sheet.get_formula("E3")) == "=E2+C3-D3"
    assert "IF(E2>0" in str(sheet.get_formula("F2"))


def test_account_balance_with_title_banner_rows(test_manager):
    """
    Test running account balance on a sheet with title rows at the top (like trial balance)
    to ensure headers are placed on the real header row and no #VALUE! errors occur.
    """
    sheet = test_manager.active_driver.create_sheet("آزمایش_مانده")
    sheet.set_value("A1", "تراز آزمایشی چهار ستونه حساب‌ها")
    sheet.set_value("A2", "تاریخ گزارش: 1405/01/01")
    # Row 4 is the real header
    sheet.set_value("A4", "کد حساب")
    sheet.set_value("B4", "نام حساب")
    sheet.set_value("C4", "گردش بدهکار")
    sheet.set_value("D4", "گردش بستانکار")

    sheet.set_value("A5", "101")
    sheet.set_value("B5", "بانک")
    sheet.set_value("C5", 450000000)
    sheet.set_value("D5", 120000000)

    sheet.set_value("A6", "102")
    sheet.set_value("B6", "مشتریان")
    sheet.set_value("C6", 280000000)
    sheet.set_value("D6", 95000000)

    sheet.set_value("A7", "جمع کل گردش‌ها")
    sheet.set_value("C7", 730000000)
    sheet.set_value("D7", 215000000)

    action = AccountBalanceAction()
    ctx = ActionContext(manager=test_manager, sheet=sheet)
    res = AccountingActionHandler.execute_account_balance(action, ctx)

    assert res.success is True
    # Real header was at row 4, so balance headers must be placed on row 4, not row 1!
    assert sheet.get_value("E4") == "مانده حساب"
    assert sheet.get_value("F4") == "تشخیص"
    # Row 1 and 2 must remain intact
    assert sheet.get_value("A1") == "تراز آزمایشی چهار ستونه حساب‌ها"
    assert sheet.get_value("E1") is None
    # Formulas must start at row 5 (first data row)
    assert str(sheet.get_formula("E5")) == "=C5-D5"
    assert str(sheet.get_formula("E6")) == "=E5+C6-D6"
    assert "IF(E5>0" in str(sheet.get_formula("F5"))


# 4. Payroll Test (محاسبه حقوق و مزایا، اضافه‌کاری، بیمه و مالیات)
def test_payroll_action(test_manager):
    sheet = test_manager.get_active_sheet()
    action = PayrollAction()
    ctx = ActionContext(manager=test_manager, sheet=sheet)
    res = AccountingActionHandler.execute_payroll(action, ctx)

    assert res.success is True
    assert "فیش_حقوق_و_دستمزد" in test_manager.get_sheet_names()
    ps = test_manager.get_sheet("فیش_حقوق_و_دستمزد")

    assert ps.get_value("A1") == "لیست و محاسبات جامع حقوق و دستمزد و بیمه و مالیات"
    # Check formulas for overtime, gross, insurance 7%, net pay
    assert "C4/220)*1.4*E4" in str(ps.get_formula("F4"))
    assert "=C4+D4+F4" == str(ps.get_formula("G4"))
    assert "*0.07" in str(ps.get_formula("H4"))
    assert "=G4-H4-J4" == str(ps.get_formula("K4"))


def test_payroll_from_existing_sheet_with_deductions_and_rate(test_manager):
    """Verify that execute_payroll dynamically reads existing 'اطلاعات حقوق' sheet."""
    info_sheet = test_manager.create_sheet("اطلاعات حقوق")
    info_sheet.set_value("A1", "نام و نام خانوادگی")
    info_sheet.set_value("B1", "روز کارکرد")
    info_sheet.set_value("C1", "حقوق پایه ماهانه (ریال)")
    info_sheet.set_value("D1", "ساعت اضافه کاری")
    info_sheet.set_value("E1", "ضریب اضافه کاری")
    info_sheet.set_value("F1", "حق مسکن/مزایا (ریال)")
    info_sheet.set_value("G1", "سایر کسورات")

    # Employee 1: 30 days
    info_sheet.set_value("A2", "علی رضایی")
    info_sheet.set_value("B2", 30)
    info_sheet.set_value("C2", 250000000)
    info_sheet.set_value("D2", 12)
    info_sheet.set_value("E2", 2)
    info_sheet.set_value("F2", 35000000)
    info_sheet.set_value("G2", 3000000)

    # Employee 2: 28 days (prorated base)
    info_sheet.set_value("A3", "مهدی کریمی")
    info_sheet.set_value("B3", 28)
    info_sheet.set_value("C3", 165000000)
    info_sheet.set_value("D3", 15)
    info_sheet.set_value("E3", 2)
    info_sheet.set_value("F3", 45000000)
    info_sheet.set_value("G3", 2000000)

    action = PayrollAction()
    ctx = ActionContext(manager=test_manager, sheet=info_sheet)
    res = AccountingActionHandler.execute_payroll(action, ctx)

    assert res.success is True
    assert "2 نفر پرسنل" in res.message_persian
    ps = test_manager.get_sheet([s for s in test_manager.get_sheet_names() if s.startswith("فیش_حقوق_و_دستمزد")][-1])

    # Check Ali Rezaei
    assert ps.get_value("B4") == "علی رضایی"
    assert ps.get_value("C4") == 250000000
    assert ps.get_value("D4") == 35000000.0
    assert ps.get_value("E4") == 12.0
    assert "C4/220)*2*E4" in str(ps.get_formula("F4"))
    assert ps.get_value("K4") == 3000000.0
    assert "=G4-H4-J4-K4" == str(ps.get_formula("L4"))

    # Check Mehdi Karimi (28 days prorated: 165M * 28/30 = 154M)
    assert ps.get_value("B5") == "مهدی کریمی"
    assert ps.get_value("C5") == 154000000
    assert ps.get_value("D5") == 45000000.0
    assert ps.get_value("E5") == 15.0
    assert ps.get_value("K5") == 2000000.0


# 5. Depreciation Test (محاسبه استهلاک دارایی‌ها)
def test_depreciation_action(test_manager):
    sheet = test_manager.get_active_sheet()
    action = DepreciationAction(
        cost=120000000,
        salvage_value=20000000,
        useful_life_years=5,
        method="straight_line",
        asset_name="ماشین‌آلات تولید"
    )
    ctx = ActionContext(manager=test_manager, sheet=sheet)
    res = AccountingActionHandler.execute_depreciation(action, ctx)

    assert res.success is True
    assert "جدول_استهلاک_دارایی" in test_manager.get_sheet_names()
    ds = test_manager.get_sheet("جدول_استهلاک_دارایی")
    assert ds.get_value("A5") == "سال 1"
    assert ds.get_value("C5") == 20000000.0  # (120M - 20M) / 5 = 20M
    assert ds.get_value("E9") == 20000000.0  # Ending book value after 5 years = salvage value


# 6. Financial Statements Test (صورت سود و زیان و ترازنامه)
def test_financial_statements(test_manager):
    sheet = test_manager.get_active_sheet()
    # Income statement
    inc_action = FinancialStatementAction(statement_type="income_statement")
    ctx = ActionContext(manager=test_manager, sheet=sheet)
    res_inc = AccountingActionHandler.execute_financial_statement(inc_action, ctx)
    assert res_inc.success is True
    assert "صورت_سود_و_زیان" in test_manager.get_sheet_names()

    # Balance sheet
    bs_action = FinancialStatementAction(statement_type="balance_sheet")
    res_bs = AccountingActionHandler.execute_financial_statement(bs_action, ctx)
    assert res_bs.success is True
    assert "ترازنامه" in test_manager.get_sheet_names()


# 7. Sales & Purchase Invoice Test (صدور و ثبت فاکتور)
def test_invoice_action(test_manager):
    sheet = test_manager.get_active_sheet()
    action = InvoiceAction(
        invoice_type="sales",
        invoice_number="98234",
        seller_name="شرکت رایان داده",
        buyer_name="فروشگاه مرکزی",
        tax_percent=10.0
    )
    ctx = ActionContext(manager=test_manager, sheet=sheet)
    res = BusinessActionHandler.execute_invoice(action, ctx)

    assert res.success is True
    assert "فاکتور_فروش_98234" in test_manager.get_sheet_names()
    inv = test_manager.get_sheet("فاکتور_فروش_98234")
    assert inv.get_value("A1") == "فاکتور فروش کالا و خدمات"
    assert "D8*F8" in str(inv.get_formula("G8"))
    assert "=SUM(" in str(inv.get_formula("G12"))


# 8. Inventory Alert & Valuation Test (مدیریت موجودی انبار و هشدار کسری)
def test_inventory_alert_action(test_manager):
    sheet = test_manager.get_active_sheet()
    sheet.set_value("A1", "کد کالا")
    sheet.set_value("B1", "نام کالا")
    sheet.set_value("C1", "موجودی")
    sheet.set_value("D1", "قیمت خرید")

    sheet.set_value("B2", "مانیتور")
    sheet.set_value("C2", 2)  # Low stock <= 5
    sheet.set_value("D2", 5000000)

    sheet.set_value("B3", "کیبورد")
    sheet.set_value("C3", 40) # Sufficient stock
    sheet.set_value("D3", 800000)

    action = InventoryAlertAction(min_threshold=5.0, highlight_low_stock=True)
    ctx = ActionContext(manager=test_manager, sheet=sheet)
    res = BusinessActionHandler.execute_inventory_alert(action, ctx)

    assert res.success is True
    assert "کسری موجودی" in res.message_persian
    assert "مانیتور" in res.message_persian
    assert "42,000,000" in res.message_persian  # 2*5M + 40*800k = 42M
    assert "IF(C2<=5" in str(sheet.get_formula("E2"))


def test_inventory_alert_ignores_empty_spacer_and_summary_rows(test_manager):
    """Verify that inventory control ignores spacer blank rows and summary rows,
    places status columns immediately next to data, and doesn't falsely warn on empty rows."""
    sheet = test_manager.get_active_sheet()
    sheet.set_value("A1", "نام کالا")
    sheet.set_value("B1", "قیمت")
    sheet.set_value("C1", "تعداد")
    sheet.set_value("D1", "جمع")
    sheet.set_value("E1", "تاریخ")

    products = [
        ("موز", 150000, 10, 1500000),
        ("بیسکویت", 50000, 30, 1500000),
        ("نان", 6000, 50, 300000),
        ("مودمس", 280000, 100, 28000000),
        ("سوسیس", 150000, 10, 1500000),
        ("کالباس", 160000, 10, 1600000),
        ("قند", 250000, 500, 125000000),
    ]

    for idx, (p, price, qty, total) in enumerate(products, start=2):
        sheet.set_value(f"A{idx}", p)
        sheet.set_value(f"B{idx}", price)
        sheet.set_value(f"C{idx}", qty)
        sheet.set_value(f"D{idx}", total)
        sheet.set_value(f"E{idx}", "1405/08/02")

    # Spacer rows: rows 9 and 10 are completely blank
    # (Leaving A9:E10 empty)

    # Summary rows: rows 11 and 12
    sheet.set_value("A11", "مجموع کل ار")
    sheet.set_value("A12", "مجموع کل ار")

    action = InventoryAlertAction(min_threshold=5.0, highlight_low_stock=True)
    ctx = ActionContext(manager=test_manager, sheet=sheet)
    res = BusinessActionHandler.execute_inventory_alert(action, ctx)

    assert res.success is True
    # None of the 7 products are <= 5
    assert "کلیه اقلام دارای موجودی کافی" in res.message_persian
    assert "کسری موجودی: **0**" in res.message_persian

    # Status column is placed at F (next to E), and valuation at G
    assert sheet.get_value("F1") == "وضعیت موجودی انبار"
    assert sheet.get_value("G1") == "ارزش ریالی موجودی"

    # Row 2 (موز) has formula with safe blank check
    f2 = str(sheet.get_formula("F2"))
    assert 'IF(C2<=5' in f2 or 'IF(C2<=5.0' in f2

    # Spacer rows (row 9, 10) must be None / blank (NOT pink, NOT shortage)
    assert sheet.get_value("F9") is None
    assert sheet.get_value("G9") is None
    assert sheet.get_cell_color("F9") is None
    assert sheet.get_cell_color("A9") is None

    # Summary row (row 11) must have SUM formula for valuation
    f_sum = str(sheet.get_formula("G11"))
    assert "=SUM(G2:G8)" in f_sum


# 9. Top-Selling Products Test (پرفروش‌ترین کالاها)
def test_top_selling_action(test_manager):
    sheet = test_manager.get_active_sheet()
    sheet.set_value("A1", "کالا")
    sheet.set_value("B1", "تعداد فروش")
    sheet.set_value("C1", "مبلغ فروش")

    sheet.set_value("A2", "ماوس گیمینگ")
    sheet.set_value("B2", 150)
    sheet.set_value("C2", 150000000)

    sheet.set_value("A3", "لپ‌تاپ گیمینگ")
    sheet.set_value("B3", 10)
    sheet.set_value("C3", 600000000)

    action = TopSellingAction(top_n=3)
    ctx = ActionContext(manager=test_manager, sheet=sheet)
    res = BusinessActionHandler.execute_top_selling(action, ctx)

    assert res.success is True
    assert "ماوس گیمینگ" in res.message_persian
    assert "150" in res.message_persian


# 10. Customer Balances Test (محاسبه بدهی و طلب مشتریان)
def test_customer_ledger_action(test_manager):
    sheet = test_manager.get_active_sheet()
    sheet.set_value("A1", "نام مشتری")
    sheet.set_value("B1", "فاکتور (بدهی)")
    sheet.set_value("C1", "واریزی (طلب)")

    sheet.set_value("A2", "شرکت سپهر")
    sheet.set_value("B2", 80000000)
    sheet.set_value("C2", 50000000)

    sheet.set_value("A3", "فروشگاه آریا")
    sheet.set_value("B3", 30000000)
    sheet.set_value("C3", 40000000)

    action = CustomerLedgerAction()
    ctx = ActionContext(manager=test_manager, sheet=sheet)
    res = BusinessActionHandler.execute_customer_ledger(action, ctx)

    assert res.success is True
    assert "مانده_حساب_مشتریان" in test_manager.get_sheet_names()
    cs = test_manager.get_sheet("مانده_حساب_مشتریان")
    assert cs.get_value("B4") == "شرکت سپهر"
    assert str(cs.get_formula("E4")) == "=C4-D4"


# 11. Price Adjustment Test (افزایش یا کاهش درصدی قیمت‌ها)
def test_price_adjustment_action(test_manager):
    sheet = test_manager.get_active_sheet()
    sheet.set_value("A1", "کالا")
    sheet.set_value("B1", "قیمت فعلی")
    sheet.set_value("B2", 100000)
    sheet.set_value("B3", 200000)

    # 15% increase into column C
    action = PriceAdjustmentAction(target_col="B", percentage=15.0, operation="increase", in_place=False)
    ctx = ActionContext(manager=test_manager, sheet=sheet)
    res = BusinessActionHandler.execute_price_adjustment(action, ctx)

    assert res.success is True
    assert "تغییر درصدی قیمت‌ها" in res.message_persian
    assert str(sheet.get_formula("C2")) == "=ROUND(B2*1.1500, 0)"
    assert str(sheet.get_formula("C3")) == "=ROUND(B3*1.1500, 0)"


# 12. Row-by-Row Multi-Column Average Test (میانگین سطری بین دو یا چند ستون)
def test_row_by_row_average(test_manager):
    sheet = test_manager.get_active_sheet()
    sheet.set_value("A1", "ردیف")
    sheet.set_value("B1", "نمره میانترم")
    sheet.set_value("C1", "نمره پایانترم")
    sheet.set_value("B2", 16)
    sheet.set_value("C2", 18)
    sheet.set_value("B3", 12)
    sheet.set_value("C3", 14)

    action = ColumnArithmeticAction(
        operation="average",
        col1="B",
        col2="C",
        columns=["B", "C"],
        target_col="D",
        as_formula=True
    )
    ctx = ActionContext(manager=test_manager, sheet=sheet)
    executor = ActionExecutor()
    res = executor.execute(action, ctx)

    assert res.success is True
    assert str(sheet.get_formula("D2")) == "=AVERAGE(B2, C2)"
    assert str(sheet.get_formula("D3")) == "=AVERAGE(B3, C3)"


# 13. Deep File Health Audit Test (تحلیل هوشمند و سلامت فایل)
def test_deep_file_auditor(test_manager):
    sheet = test_manager.get_active_sheet()
    sheet.set_value("A1", "کد")
    sheet.set_value("B1", "تعداد")
    sheet.set_value("C1", "قیمت")

    sheet.set_value("A2", 101)
    sheet.set_value("B2", 5)
    sheet.set_value("C2", None)  # Empty cell

    sheet.set_value("A3", 102)
    sheet.set_value("B3", "#DIV/0!")  # Formula error
    sheet.set_value("C3", 2000)

    sheet.set_value("A4", 101)
    sheet.set_value("B4", 5)
    sheet.set_value("C4", None)  # Duplicate row

    audit_res = DeepFileAuditor.audit_sheet(sheet)
    assert audit_res.health_score < 100
    assert len(audit_res.empty_cells) >= 1
    assert len(audit_res.formula_errors) >= 1
    assert "گزارش تحلیل هوشمند و سلامت فایل اکسل" in audit_res.persian_summary


# 14. Auto Chart Generation Test (رسم نمودار گرافیکی خودکار)
def test_auto_chart_action(test_manager):
    sheet = test_manager.get_active_sheet()
    sheet.set_value("A1", "محصول")
    sheet.set_value("B1", "فروش")
    sheet.set_value("A2", "ماوس")
    sheet.set_value("B2", 500)
    sheet.set_value("A3", "کیبورد")
    sheet.set_value("B3", 750)

    action = AutoChartAction(chart_type="col", label_col="A", value_col="B", title="فروش ماهانه")
    ctx = ActionContext(manager=test_manager, sheet=sheet)
    res = BusinessActionHandler.execute_auto_chart(action, ctx)

    assert res.success is True
    assert "نمودار تحلیلی گرافیکی" in res.message_persian
    # Check chart in openpyxl sheet
    ws = getattr(sheet, "_ws", None)
    assert len(ws._charts) > 0


# 15. Backup & PDF Export Tests
def test_backup_and_pdf(test_manager, tmp_path):
    sheet = test_manager.get_active_sheet()
    sheet.set_value("A1", "نام")
    sheet.set_value("B1", "مبلغ")
    sheet.set_value("A2", "سود خالص")
    sheet.set_value("B2", 15000000)

    # Test Backup
    backup_act = BackupAction()
    ctx = ActionContext(manager=test_manager, sheet=sheet)
    res_b = BusinessActionHandler.execute_backup(backup_act, ctx)
    assert res_b.success is True
    assert "نسخه پشتیبان" in res_b.message_persian

    # Test PDF Export
    pdf_path = str(tmp_path / "test_report.pdf")
    pdf_act = PdfExportAction(output_filename=pdf_path)
    res_p = BusinessActionHandler.execute_pdf_export(pdf_act, ctx)
    assert res_p.success is True
    assert "خروجی PDF" in res_p.message_persian
    assert os.path.exists(pdf_path)
    assert os.path.getsize(pdf_path) > 500


# 16. Natural Persian NLP Rules Parser Coverage
def test_persian_rule_parser_new_features():
    header_map = {"نام کالا": "A", "قیمت": "B", "تعداد": "C", "فروش": "D"}

    # Reconciliation
    a1 = PersianRuleParser.parse("مغایرت‌گیری ستون A و B", header_map)
    assert isinstance(a1, ReconciliationAction)
    assert a1.col1 == "A" and a1.col2 == "B"

    # Invoice
    a2 = PersianRuleParser.parse("ثبت فاکتور فروش رسمی", header_map)
    assert isinstance(a2, InvoiceAction)
    assert a2.invoice_type == "sales"

    # Payroll
    a3 = PersianRuleParser.parse("محاسبه حقوق و دستمزد و بیمه", header_map)
    assert isinstance(a3, PayrollAction)

    # Inventory Alert
    a4 = PersianRuleParser.parse("موجودی انبار و هشدار کمبود موجودی", header_map)
    assert isinstance(a4, InventoryAlertAction)

    # Trial Balance
    a5 = PersianRuleParser.parse("تهیه تراز آزمایشی چهار ستونه", header_map)
    assert isinstance(a5, TrialBalanceAction)

    # Account Balance
    a6 = PersianRuleParser.parse("محاسبه مانده حساب بدهکار و بستانکار", header_map)
    assert isinstance(a6, AccountBalanceAction)

    # Depreciation
    a7 = PersianRuleParser.parse("محاسبه جدول استهلاک دارایی", header_map)
    assert isinstance(a7, DepreciationAction)

    # Financial Statements
    a8 = PersianRuleParser.parse("تهیه صورت سود و زیان جامع", header_map)
    assert isinstance(a8, FinancialStatementAction)
    assert a8.statement_type == "income_statement"

    a9 = PersianRuleParser.parse("تهیه ترازنامه مالی", header_map)
    assert isinstance(a9, FinancialStatementAction)
    assert a9.statement_type == "balance_sheet"

    # Price Adjustment
    a10 = PersianRuleParser.parse("قیمت‌های ستون C را ۱۵ درصد افزایش بده", header_map)
    assert isinstance(a10, PriceAdjustmentAction)
    assert a10.percentage == 15.0
    assert a10.operation == "increase"
    assert a10.target_col == "C"

    # Multi-Column Row Average
    a11 = PersianRuleParser.parse("از ستون B و C میانگین بگیر و نتیجه را در D بنویس", header_map)
    assert isinstance(a11, ColumnArithmeticAction)
    assert a11.operation == "average"
    assert "B" in a11.columns and "C" in a11.columns
    assert a11.target_col == "D"

    # Deep File Health Audit
    a12 = PersianRuleParser.parse("سلامت فایل و بررسی خطاهای اکسل", header_map)
    assert isinstance(a12, DeepFileAuditAction)

    # Auto Chart
    a13 = PersianRuleParser.parse("رسم نمودار ستونی فروش", header_map)
    assert isinstance(a13, AutoChartAction)

    # Backup & PDF
    a14 = PersianRuleParser.parse("ذخیره نسخه پشتیبان از فایل", header_map)
    assert isinstance(a14, BackupAction)

    a15 = PersianRuleParser.parse("خروجی PDF از شیت", header_map)
    assert isinstance(a15, PdfExportAction)


# 19. Inventory Alert Multi-Sheet Sales Deduction Tests
def test_inventory_alert_multi_sheet_sales_deduction(test_manager):
    """Test inventory alert deducting sales from purchase across separate 'خرید' and 'فروش' sheets."""
    p_sheet = test_manager.create_sheet("خرید")
    s_sheet = test_manager.create_sheet("فروش")

    p_sheet.set_value("A1", "نام کالا")
    p_sheet.set_value("B1", "قیمت خرید")
    p_sheet.set_value("C1", "تعداد خرید")

    p_sheet.set_value("A2", "لپ‌تاپ")
    p_sheet.set_value("B2", 20000000)
    p_sheet.set_value("C2", 10)

    p_sheet.set_value("A3", "کیبورد")
    p_sheet.set_value("B3", 500000)
    p_sheet.set_value("C3", 8)

    p_sheet.set_value("A4", "ماوس")
    p_sheet.set_value("B4", 200000)
    p_sheet.set_value("C4", 50)

    s_sheet.set_value("A1", "نام کالا")
    s_sheet.set_value("B1", "قیمت فروش")
    s_sheet.set_value("C1", "تعداد فروش")

    s_sheet.set_value("A2", "لپ‌تاپ")
    s_sheet.set_value("B2", 25000000)
    s_sheet.set_value("C2", 7)

    s_sheet.set_value("A3", "لپ‌تاپ")
    s_sheet.set_value("B3", 25000000)
    s_sheet.set_value("C3", 1)  # Total sold laptop = 8 (Remaining = 10 - 8 = 2 <= 5 -> Alert!)

    s_sheet.set_value("A4", "کیبورد")
    s_sheet.set_value("B4", 700000)
    s_sheet.set_value("C4", 5)  # Total sold keyboard = 5 (Remaining = 8 - 5 = 3 <= 5 -> Alert!)

    s_sheet.set_value("A5", "ماوس")
    s_sheet.set_value("B5", 300000)
    s_sheet.set_value("C5", 15) # Total sold mouse = 15 (Remaining = 50 - 15 = 35 > 5 -> OK)

    action = InventoryAlertAction(min_threshold=5.0, highlight_low_stock=True)
    ctx = ActionContext(manager=test_manager, sheet=p_sheet)
    res = BusinessActionHandler.execute_inventory_alert(action, ctx)

    assert res.success is True
    assert "کسر فروش از خرید" in res.message_persian
    assert "48,500,000" in res.message_persian  # 2*20M + 3*500k + 35*200k = 48.5M
    assert "لپ‌تاپ" in res.message_persian
    assert "کیبورد" in res.message_persian

    # Check generated headers in purchase sheet
    assert p_sheet.get_value("D1") == "تعداد فروش (فروش)"
    assert p_sheet.get_value("E1") == "مانده موجودی انبار"
    assert p_sheet.get_value("F1") == "وضعیت موجودی انبار"
    assert p_sheet.get_value("G1") == "ارزش ریالی موجودی"

    # Check formulas for Row 2 (Laptop)
    f_sold = str(p_sheet.get_formula("D2"))
    assert "SUMIF('فروش'!$A$2:$A$5, A2, 'فروش'!$C$2:$C$5)" in f_sold

    f_rem = str(p_sheet.get_formula("E2"))
    assert "C2-D2" in f_rem

    f_status = str(p_sheet.get_formula("F2"))
    assert "E2<=5" in f_status or "E2<=5.0" in f_status

    f_val = str(p_sheet.get_formula("G2"))
    assert "E2*B2" in f_val

    # Check highlighting: Laptop (row 2) and Keyboard (row 3) highlighted pink, Mouse (row 4) not highlighted
    assert p_sheet.get_cell_color("F2") == "#FEE2E2"
    assert p_sheet.get_cell_color("F3") == "#FEE2E2"
    assert p_sheet.get_cell_color("F4") is None

    # Check summary row at row 5
    assert "مجموع کل ارزش موجودی انبار" in str(p_sheet.get_value("A5"))
    assert "=SUM(G2:G4)" in str(p_sheet.get_formula("G5"))
    assert "=SUM(D2:D4)" in str(p_sheet.get_formula("D5"))
    assert "=SUM(E2:E4)" in str(p_sheet.get_formula("E5"))


def test_inventory_alert_auto_switches_from_sales_to_purchase_sheet(test_manager):
    """Test running inventory alert while the user is currently viewing the sales sheet."""
    p_sheet = test_manager.create_sheet("خرید")
    s_sheet = test_manager.create_sheet("فروش")

    p_sheet.set_value("A1", "نام کالا")
    p_sheet.set_value("B1", "قیمت خرید")
    p_sheet.set_value("C1", "تعداد خرید")

    p_sheet.set_value("A2", "مانیتور")
    p_sheet.set_value("B2", 6000000)
    p_sheet.set_value("C2", 10)

    s_sheet.set_value("A1", "نام کالا")
    s_sheet.set_value("B1", "قیمت فروش")
    s_sheet.set_value("C1", "تعداد فروش")

    s_sheet.set_value("A2", "مانیتور")
    s_sheet.set_value("B2", 8000000)
    s_sheet.set_value("C2", 8)  # Remaining = 10 - 8 = 2 <= 5

    # Context active sheet is sales sheet!
    action = InventoryAlertAction(min_threshold=5.0)
    ctx = ActionContext(manager=test_manager, sheet=s_sheet)
    res = BusinessActionHandler.execute_inventory_alert(action, ctx)

    assert res.success is True
    # Successfully switched to purchase sheet and applied deduction
    assert p_sheet.get_value("E1") == "مانده موجودی انبار"
    assert "C2-D2" in str(p_sheet.get_formula("E2"))
    assert "12,000,000" in res.message_persian  # 2 * 6,000,000 = 12M


def test_inventory_alert_same_sheet_sales_deduction(test_manager):
    """Test inventory alert with purchase and sales columns in the same sheet."""
    sheet = test_manager.get_active_sheet()
    sheet.set_value("A1", "نام کالا")
    sheet.set_value("B1", "قیمت خرید")
    sheet.set_value("C1", "تعداد خرید")
    sheet.set_value("D1", "تعداد فروش")

    sheet.set_value("A2", "پرینتر")
    sheet.set_value("B2", 10000000)
    sheet.set_value("C2", 12)
    sheet.set_value("D2", 10)  # Remaining = 2 <= 5 -> Alert

    sheet.set_value("A3", "اسکنر")
    sheet.set_value("B3", 5000000)
    sheet.set_value("C3", 20)
    sheet.set_value("D3", 5)   # Remaining = 15 > 5 -> OK

    action = InventoryAlertAction(min_threshold=5.0)
    ctx = ActionContext(manager=test_manager, sheet=sheet)
    res = BusinessActionHandler.execute_inventory_alert(action, ctx)

    assert res.success is True
    assert sheet.get_value("E1") == "مانده موجودی انبار"
    assert sheet.get_value("F1") == "وضعیت موجودی انبار"
    assert sheet.get_value("G1") == "ارزش ریالی موجودی"

    # Row 2 remaining formula: C2 - D2
    assert "C2-D2" in str(sheet.get_formula("E2"))
    # Valuation formula: E2 * B2
    assert "E2*B2" in str(sheet.get_formula("G2"))
    # Valuation: (2 * 10M) + (15 * 5M) = 20M + 75M = 95M
    assert "95,000,000" in res.message_persian
    assert "پرینتر" in res.message_persian


def test_nlp_parsing_inventory_sales_deduction_prompts():
    """Verify NLP intent parsing for colloquial user requests about deducting sales from purchases."""
    from excel_assistant.nlp.parser import PersianIntentParser
    parser = PersianIntentParser()

    prompts = [
        "باید تعداد فروش از خرید کم کنه تا موجودی باقی مانده رو به دست بیاره",
        "کسر فروش از خرید برای موجودی انبار",
        "تعداد فروش رو از خرید کسر کن تا مانده انبار مشخص بشه",
        "مانده موجودی انبار رو حساب کن",
        "کنترل موجودی انبار",
    ]

    for p in prompts:
        action = parser.parse(p)
        assert isinstance(action, InventoryAlertAction), f"Failed to parse '{p}' as InventoryAlertAction, got {type(action).__name__}"


# 20. FIFO, LIFO & Weighted Average Inventory Valuation Tests
def test_inventory_valuation_calculator_fifo_lifo_average():
    """Verify core math for FIFO, LIFO, and Weighted Average under rising purchase prices."""
    from excel_assistant.actions.business_actions import InventoryValuationCalculator

    # 3 batches of purchases:
    # 1. 10 units @ 20,000,000 = 200,000,000
    # 2. 5 units @ 22,000,000 = 110,000,000
    # 3. 8 units @ 25,000,000 = 200,000,000
    # Total purchases = 23 units, total cost = 510,000,000
    batches = [(10, 20000000), (5, 22000000), (8, 25000000)]
    total_sold = 14  # Remaining = 23 - 14 = 9 units

    # FIFO: Sales consume earliest batches (10 @ 20M + 4 @ 22M = 288M COGS)
    # Remaining comes from latest batches (8 @ 25M + 1 @ 22M = 222M remaining value)
    fifo = InventoryValuationCalculator.calculate_valuation(batches, total_sold, "fifo")
    assert fifo["remaining_qty"] == 9
    assert fifo["remaining_value"] == 222000000
    assert fifo["cogs"] == 288000000
    assert abs(fifo["effective_unit_price"] - (222000000 / 9)) < 1.0

    # LIFO: Sales consume latest batches (8 @ 25M + 5 @ 22M + 1 @ 20M = 330M COGS)
    # Remaining comes from earliest batch (9 @ 20M = 180M remaining value)
    lifo = InventoryValuationCalculator.calculate_valuation(batches, total_sold, "lifo")
    assert lifo["remaining_qty"] == 9
    assert lifo["remaining_value"] == 180000000
    assert lifo["cogs"] == 330000000
    assert lifo["effective_unit_price"] == 20000000

    # Weighted Average: Unit price = 510,000,000 / 23 = 22,173,913.04
    avg = InventoryValuationCalculator.calculate_valuation(batches, total_sold, "average")
    assert avg["remaining_qty"] == 9
    expected_avg_p = 510000000 / 23
    assert abs(avg["effective_unit_price"] - expected_avg_p) < 0.01
    assert abs(avg["remaining_value"] - (9 * expected_avg_p)) < 0.01
    assert abs(avg["cogs"] - (14 * expected_avg_p)) < 0.01


def test_inventory_valuation_action_multi_batch_excel_sheet(test_manager):
    """Test full execution of Inventory Valuation comparing FIFO, LIFO, and Average in Excel."""
    from excel_assistant.nlp.intent_models import InventoryValuationAction

    p_sheet = test_manager.create_sheet("خرید")
    s_sheet = test_manager.create_sheet("فروش")

    p_sheet.set_value("A1", "نام کالا")
    p_sheet.set_value("B1", "قیمت خرید")
    p_sheet.set_value("C1", "تعداد خرید")

    # Laptop with 3 purchase batches at increasing prices
    p_sheet.set_value("A2", "لپ‌تاپ")
    p_sheet.set_value("B2", 20000000)
    p_sheet.set_value("C2", 10)

    p_sheet.set_value("A3", "لپ‌تاپ")
    p_sheet.set_value("B3", 22000000)
    p_sheet.set_value("C3", 5)

    p_sheet.set_value("A4", "لپ‌تاپ")
    p_sheet.set_value("B4", 25000000)
    p_sheet.set_value("C4", 8)

    # Keyboard with 1 purchase batch
    p_sheet.set_value("A5", "کیبورد")
    p_sheet.set_value("B5", 500000)
    p_sheet.set_value("C5", 8)

    # Mouse with 1 purchase batch
    p_sheet.set_value("A6", "ماوس")
    p_sheet.set_value("B6", 200000)
    p_sheet.set_value("C6", 50)

    # Sales sheet
    s_sheet.set_value("A1", "نام کالا")
    s_sheet.set_value("B1", "تعداد فروش")
    s_sheet.set_value("A2", "لپ‌تاپ")
    s_sheet.set_value("B2", 14)  # Remaining Laptop = 23 - 14 = 9

    s_sheet.set_value("A3", "کیبورد")
    s_sheet.set_value("B3", 5)   # Remaining Keyboard = 8 - 5 = 3

    s_sheet.set_value("A4", "ماوس")
    s_sheet.set_value("B4", 15)  # Remaining Mouse = 50 - 15 = 35

    action = InventoryValuationAction(method="compare", create_report_sheet=True)
    ctx = ActionContext(manager=test_manager, sheet=p_sheet)
    res = BusinessActionHandler.execute_inventory_valuation(action, ctx)

    assert res.success is True
    assert "گزارش تخصصی ارزش‌گذاری موجودی انبار" in res.message_persian
    assert "فایفو" in res.message_persian
    assert "لایفو" in res.message_persian
    assert "میانگین موزون" in res.message_persian

    # Total remaining valuation:
    # FIFO = 222M (laptop) + 1.5M (keyboard) + 7M (mouse) = 230,500,000
    # LIFO = 180M (laptop) + 1.5M (keyboard) + 7M (mouse) = 188,500,000
    assert "230,500,000" in res.message_persian
    assert "188,500,000" in res.message_persian

    # Check generated report sheet
    assert "ارزش‌گذاری_موجودی" in test_manager.get_sheet_names()
    rep_sheet = test_manager.get_sheet("ارزش‌گذاری_موجودی")

    assert "ارزش‌گذاری موجودی انبار" in str(rep_sheet.get_value("A1"))
    assert rep_sheet.get_value("A5") == 230500000  # Card: FIFO
    assert rep_sheet.get_value("E5") == 188500000  # Card: LIFO

    # Row 8 is Laptop
    assert rep_sheet.get_value("B8") == "لپ‌تاپ"
    assert rep_sheet.get_value("C8") == 23
    assert rep_sheet.get_value("E8") == 14
    assert rep_sheet.get_value("H8") == 222000000  # FIFO value
    assert rep_sheet.get_value("J8") == 180000000  # LIFO value


def test_nlp_parsing_fifo_lifo_average_prompts():
    """Verify Persian intent parser for FIFO, LIFO, and Weighted Average inventory valuation."""
    from excel_assistant.nlp.parser import PersianIntentParser
    from excel_assistant.nlp.intent_models import InventoryValuationAction
    parser = PersianIntentParser()

    prompts_and_methods = [
        ("ارزش‌گذاری انبار به روش فایفو", "fifo"),
        ("ارزش موجودی به روش لایفو", "lifo"),
        ("ارزش انبار با میانگین موزون", "average"),
        ("مقایسه روش‌های فایفو و لایفو و میانگین در انبار", "compare"),
        ("ارزش‌گذاری موجودی انبار", "compare"),
    ]

    for prompt, expected_method in prompts_and_methods:
        action = parser.parse(prompt)
        assert isinstance(action, InventoryValuationAction), f"Failed for '{prompt}', got {type(action).__name__}"
        assert action.method == expected_method, f"Expected {expected_method} for '{prompt}', got {action.method}"


def test_consecutive_inventory_alert_no_circular_reference(test_manager):
    """
    Verify that re-running inventory check on an already-processed sheet
    does not cause a circular reference (e.g. G2 subtracting from itself: G2-F2).
    """
    p_sheet = test_manager.create_sheet("خرید_دوباره")
    p_sheet.set_value("A1", "ردیف")
    p_sheet.set_value("B1", "نام کالا")
    p_sheet.set_value("C1", "تعداد")
    p_sheet.set_value("D1", "قیمت واحد")
    p_sheet.set_value("E1", "مبلغ کل")

    p_sheet.set_value("A2", 1)
    p_sheet.set_value("B2", "لپ تاپ")
    p_sheet.set_value("C2", 50)
    p_sheet.set_value("D2", 30000000)
    p_sheet.set_value("E2", 1500000000)

    p_sheet.set_value("A3", 2)
    p_sheet.set_value("B3", "موس")
    p_sheet.set_value("C3", 200)
    p_sheet.set_value("D3", 250000)
    p_sheet.set_value("E3", 50000000)

    s_sheet = test_manager.create_sheet("فروش_دوباره")
    s_sheet.set_value("A1", "نام کالا")
    s_sheet.set_value("B1", "قیمت فروش")
    s_sheet.set_value("C1", "تعداد فروش")

    s_sheet.set_value("A2", "لپ تاپ")
    s_sheet.set_value("B2", 35000000)
    s_sheet.set_value("C2", 48)

    s_sheet.set_value("A3", "موس")
    s_sheet.set_value("B3", 300000)
    s_sheet.set_value("C3", 50)

    action = InventoryAlertAction(min_threshold=5, highlight_low_stock=True)
    ctx = ActionContext(manager=test_manager, sheet=p_sheet)

    # First run
    res1 = BusinessActionHandler.execute_inventory_alert(action, ctx)
    assert res1.success is True
    g2_run1 = p_sheet.get_formula("G2")
    assert "=IF(OR(B2=\"\", C2=\"\"), \"\", C2-F2)" == g2_run1

    # Second run (simulating user re-running on processed sheet)
    res2 = BusinessActionHandler.execute_inventory_alert(action, ctx)
    assert res2.success is True
    g2_run2 = p_sheet.get_formula("G2")
    assert "=IF(OR(B2=\"\", C2=\"\"), \"\", C2-F2)" == g2_run2
    assert "G2-" not in g2_run2, "Circular reference detected in G2!"

    # Third run
    res3 = BusinessActionHandler.execute_inventory_alert(action, ctx)
    assert res3.success is True
    g2_run3 = p_sheet.get_formula("G2")
    assert "=IF(OR(B2=\"\", C2=\"\"), \"\", C2-F2)" == g2_run3
    assert "G2-" not in g2_run3, "Circular reference detected in G2!"


def test_multi_batch_inventory_fifo_allocation(test_manager):
    """
    Test multiple purchase batches of the same product (e.g. 2 banana batches, 2 sausage batches).
    Verifies:
    1. Sales is allocated via FIFO to oldest batch first; second batch is NOT double-deducted.
    2. Total remaining stock across batches is preserved (10+10 - 1.7 = 18.3, not 16.6).
    3. Status checks total remaining stock and does not trigger false alerts.
    """
    p_sheet = test_manager.create_sheet("خرید_چندپارتی")
    p_sheet.set_value("A1", "نام کالا")
    p_sheet.set_value("B1", "فی خرید")
    p_sheet.set_value("C1", "تعداد خرید")
    p_sheet.set_value("D1", "مبلغ کل")

    # Batch 1 of موز
    p_sheet.set_value("A2", "موز")
    p_sheet.set_value("B2", 150000)
    p_sheet.set_value("C2", 10)
    p_sheet.set_value("D2", 1500000)

    # Batch 1 of سوسیس
    p_sheet.set_value("A3", "سوسیس")
    p_sheet.set_value("B3", 150000)
    p_sheet.set_value("C3", 10)
    p_sheet.set_value("D3", 1500000)

    # Batch 2 of سوسیس
    p_sheet.set_value("A4", "سوسیس")
    p_sheet.set_value("B4", 155000)
    p_sheet.set_value("C4", 10)

    # Batch 2 of موز
    p_sheet.set_value("A5", "موز")
    p_sheet.set_value("B5", 155000)
    p_sheet.set_value("C5", 10)

    s_sheet = test_manager.create_sheet("فروش_چندپارتی")
    s_sheet.set_value("A1", "نام کالا")
    s_sheet.set_value("B1", "قیمت فروش")
    s_sheet.set_value("C1", "تعداد فروش")

    s_sheet.set_value("A2", "موز")
    s_sheet.set_value("B2", 170000)
    s_sheet.set_value("C2", 1.7)

    action = InventoryAlertAction(min_threshold=5.0, highlight_low_stock=True)
    ctx = ActionContext(manager=test_manager, sheet=p_sheet)

    res = BusinessActionHandler.execute_inventory_alert(action, ctx)
    assert res.success is True

    # Column D in rows 4 and 5 should be auto-filled:
    assert p_sheet.get_formula("D4") == '=IF(OR(B4="", C4=""), "", B4*C4)'
    assert p_sheet.get_formula("D5") == '=IF(OR(B5="", C5=""), "", B5*C5)'

    # Formulas in Column E (تعداد فروش):
    f_sold_batch1 = p_sheet.get_formula("E2")
    f_sold_batch2 = p_sheet.get_formula("E5")

    # Batch 1 takes up to 10
    assert "MAX(0, MIN(C2" in f_sold_batch1
    # Batch 2 subtracts prior purchases:
    assert "SUMIF($A$2:$A4, A5, $C$2:$C4)" in f_sold_batch2

    # Status formula checks total remaining of the product
    f_status_batch1 = p_sheet.get_formula("G2")
    assert "SUMIF($A$2:$A$5, A2, $F$2:$F$5)<=5.0" in f_status_batch1


def test_financial_statement_from_trial_balance(test_manager):
    """Verify that execute_financial_statement dynamically parses a trial balance sheet."""
    from excel_assistant.actions.accounting_actions import AccountingActionHandler
    from excel_assistant.nlp.intent_models import FinancialStatementAction

    tb = test_manager.create_sheet("تراز آزمایشی")
    tb.set_value("A1", "عنوان حساب")
    tb.set_value("B1", "مانده بدهکار (ریال)")
    tb.set_value("C1", "مانده بستانکار (ریال)")

    data = [
        ("فروش کالا", 0, 1850000000),
        ("درآمد خدمات", 0, 125000000),
        ("برگشت از فروش", 42000000, 0),
        ("تخفیفات فروش", 16000000, 0),
        ("موجودی کالا اول دوره", 210000000, 0),
        ("خرید طی دوره", 620000000, 0),
        ("برگشت از خرید", 0, 18000000),
        ("تخفیفات خرید", 0, 7000000),
        ("موجودی کالا پایان دوره", 265000000, 0),
        ("هزینه حقوق و دستمزد", 310000000, 0),
        ("هزینه اجاره", 84000000, 0),
        ("هزینه آب، برق و گاز", 22000000, 0),
        ("هزینه مالی", 28000000, 0),
        ("درآمد سود سپرده", 0, 14000000),
    ]
    for idx, (acc, d, c) in enumerate(data, start=2):
        tb.set_value(f"A{idx}", acc)
        tb.set_value(f"B{idx}", d)
        tb.set_value(f"C{idx}", c)

    ctx = ActionContext(manager=test_manager, sheet=tb)
    res = AccountingActionHandler.execute_financial_statement(
        FinancialStatementAction(statement_type="income_statement"), ctx
    )
    assert res.success is True
    assert "تراز آزمایشی" in res.message_persian

    inc_sheet = test_manager.get_sheet("صورت_سود_و_زیان")
    # Gross revenue: 1,850,000,000 + 125,000,000 = 1,975,000,000
    assert inc_sheet.get_value("B4") == 1975000000.0
    # Sales contra: - (42M + 16M) = -58M
    assert inc_sheet.get_value("B5") == -58000000.0
    # COGS: 210M + (620M - 25M) - 265M = 540M
    assert inc_sheet.get_value("B7") == -540000000.0


def test_depreciation_from_trial_balance(test_manager):
    """Verify that execute_depreciation dynamically extracts fixed assets from trial balance."""
    from excel_assistant.actions.accounting_actions import AccountingActionHandler
    from excel_assistant.nlp.intent_models import DepreciationAction

    tb = test_manager.create_sheet("تراز آزمایشی")
    tb.set_value("A1", "عنوان حساب")
    tb.set_value("B1", "مانده بدهکار (ریال)")
    tb.set_value("C1", "مانده بستانکار (ریال)")

    data = [
        ("ساختمان", 900000000, 0),
        ("استهلاک انباشته ساختمان", 0, 90000000),
        ("تجهیزات", 350000000, 0),
        ("استهلاک انباشته تجهیزات", 0, 70000000),
        ("هزینه استهلاک", 45000000, 0),
    ]
    for idx, (acc, d, c) in enumerate(data, start=2):
        tb.set_value(f"A{idx}", acc)
        tb.set_value(f"B{idx}", d)
        tb.set_value(f"C{idx}", c)

    ctx = ActionContext(manager=test_manager, sheet=tb)
    res = AccountingActionHandler.execute_depreciation(
        DepreciationAction(), ctx
    )
    assert res.success is True
    assert "ساختمان" in res.message_persian
    assert "تجهیزات" in res.message_persian

    dep_sheet = test_manager.get_sheet("جدول_استهلاک_دارایی")
    assert dep_sheet.get_value("A5") == "ساختمان"
    assert dep_sheet.get_value("B5") == 900000000.0
    assert dep_sheet.get_value("C5") == 90000000.0
    assert dep_sheet.get_formula("D5") == "=B5-C5"


def test_profit_calculation_multibatch_fifo_with_discounts(test_manager):
    """Verify profit calculation with multi-batch purchases and invoice discounts."""
    from excel_assistant.actions.math_actions import MathActionHandler
    from excel_assistant.nlp.intent_models import ProfitCalculationAction

    ws_sales = test_manager.create_sheet("فاکتورهای فروش")
    sales_rows = [
        ["تاریخ", "شماره فاکتور", "نام مشتری", "شرح کالا", "تعداد", "قیمت واحد (ریال)", "تخفیف درصد"],
        ["1405/01/05", "ف-1001", "شرکت آریا", "کالای A", 12, 18500000, 0],
        ["1405/01/07", "ف-1002", "فروشگاه سپهر", "کالای B", 25, 7200000, 5],
    ]
    for r_idx, row in enumerate(sales_rows, start=1):
        for c_idx, val in enumerate(row, start=1):
            ws_sales.set_value(f"{chr(64+c_idx)}{r_idx}", val)

    ws_purch = test_manager.create_sheet("خرید و موجودی کالا")
    purch_rows = [
        ["شرح کالا", "تاریخ", "نوع خرید", "تعداد خرید", "بهای واحد خرید (ریال)"],
        ["کالای A", "1405/01/01", "خرید اول دوره", 30, 12500000],
        ["کالای B", "1405/01/01", "خرید اول دوره", 60, 4800000],
    ]
    for r_idx, row in enumerate(purch_rows, start=1):
        for c_idx, val in enumerate(row, start=1):
            ws_purch.set_value(f"{chr(64+c_idx)}{r_idx}", val)

    ctx = ActionContext(manager=test_manager, sheet=ws_sales)
    action = ProfitCalculationAction(sales_sheet="فاکتورهای فروش", purchase_sheet="خرید و موجودی کالا")
    res = MathActionHandler.execute_profit_calculation(action, ctx)

    assert res.success is True
    # کالای A: 12 * 18.5M = 222M
    # کالای B: 25 * 7.2M * 0.95 = 171M
    # Total sales: 222M + 171M = 393M
    assert res.data["total_sales"] == 393000000.0
    # COGS کالای A: 12 * 12.5M = 150M
    # COGS کالای B: 25 * 4.8M = 120M
    # Total COGS: 270M
    assert res.data["total_purchase"] == 270000000.0
    # Profit: 393M - 270M = 123M
    assert res.data["total_profit"] == 123000000.0


def test_profit_calculation_lifo_and_average_multibatch(test_manager):
    """Verify LIFO and Weighted Average valuation methods in profit calculation."""
    from excel_assistant.actions.math_actions import MathActionHandler
    from excel_assistant.nlp.intent_models import ProfitCalculationAction
    from excel_assistant.nlp.parser import PersianIntentParser

    parser = PersianIntentParser()
    act_lifo = parser.parse("محاسبه سود به روش لایفو")
    assert isinstance(act_lifo, ProfitCalculationAction)
    assert act_lifo.valuation_method == "lifo"

    act_avg = parser.parse("محاسبه سود با روش میانگین موزون")
    assert isinstance(act_avg, ProfitCalculationAction)
    assert act_avg.valuation_method == "average"

    ws_sales = test_manager.create_sheet("فروش_کالا")
    sales_rows = [
        ["تاریخ", "کالا", "تعداد", "قیمت فروش"],
        ["1405/01/10", "کالای X", 15, 30000],
    ]
    for r_idx, row in enumerate(sales_rows, start=1):
        for c_idx, val in enumerate(row, start=1):
            ws_sales.set_value(f"{chr(64+c_idx)}{r_idx}", val)

    ws_purch = test_manager.create_sheet("خرید_کالا")
    purch_rows = [
        ["کالا", "تاریخ", "تعداد", "قیمت خرید"],
        ["کالای X", "1405/01/01", 10, 10000],
        ["کالای X", "1405/01/05", 10, 20000],
    ]
    for r_idx, row in enumerate(purch_rows, start=1):
        for c_idx, val in enumerate(row, start=1):
            ws_purch.set_value(f"{chr(64+c_idx)}{r_idx}", val)

    ctx = ActionContext(manager=test_manager, sheet=ws_sales)

    # 1. FIFO: 10 @ 10,000 + 5 @ 20,000 = 200,000 COGS -> Profit = 250,000
    res_fifo = MathActionHandler.execute_profit_calculation(
        ProfitCalculationAction(sales_sheet="فروش_کالا", purchase_sheet="خرید_کالا", valuation_method="fifo"),
        ctx
    )
    assert res_fifo.success is True
    assert res_fifo.data["total_sales"] == 450000.0
    assert res_fifo.data["total_purchase"] == 200000.0
    assert res_fifo.data["total_profit"] == 250000.0

    # 2. LIFO: 10 @ 20,000 + 5 @ 10,000 = 250,000 COGS -> Profit = 200,000
    res_lifo = MathActionHandler.execute_profit_calculation(
        ProfitCalculationAction(sales_sheet="فروش_کالا", purchase_sheet="خرید_کالا", valuation_method="lifo"),
        ctx
    )
    assert res_lifo.success is True
    assert res_lifo.data["total_sales"] == 450000.0
    assert res_lifo.data["total_purchase"] == 250000.0
    assert res_lifo.data["total_profit"] == 200000.0

    # 3. Average: Total cost = 300,000 / 20 units = 15,000/unit -> COGS = 15 * 15,000 = 225,000 -> Profit = 225,000
    res_avg = MathActionHandler.execute_profit_calculation(
        ProfitCalculationAction(sales_sheet="فروش_کالا", purchase_sheet="خرید_کالا", valuation_method="average"),
        ctx
    )
    assert res_avg.success is True
    assert res_avg.data["total_sales"] == 450000.0
    assert res_avg.data["total_purchase"] == 225000.0
    assert res_avg.data["total_profit"] == 225000.0


def test_payroll_with_overtime_amount_and_personnel_code(test_manager):
    """Verify that overtime in Rials and personnel codes like E001 are parsed properly."""
    sh = test_manager.create_sheet("لیست_پرسنل_اردیبهشت")
    sh.set_value('A1', 'کد پرسنلی')
    sh.set_value('B1', 'نام')
    sh.set_value('C1', 'واحد')
    sh.set_value('D1', 'ماه')
    sh.set_value('E1', 'روز کارکرد')
    sh.set_value('F1', 'حقوق پایه (ریال)')
    sh.set_value('G1', 'اضافه کاری (ریال)')
    sh.set_value('H1', 'پاداش (ریال)')
    sh.set_value('I1', 'مزایای غیرنقدی (ریال)')
    sh.set_value('J1', 'کسورات ثابت/بیمه سهم کارمند (ریال)')

    sh.set_value('A2', 'E001')
    sh.set_value('B2', 'علی رضایی')
    sh.set_value('C2', 'فروش')
    sh.set_value('D2', '1405/01')
    sh.set_value('E2', 30)
    sh.set_value('F2', 85000000)
    sh.set_value('G2', 12000000)
    sh.set_value('H2', 5000000)
    sh.set_value('I2', 0)
    sh.set_value('J2', 2500000)

    ctx = ActionContext(manager=test_manager, sheet=sh)
    action = PayrollAction()
    res = AccountingActionHandler.execute_payroll(action, ctx)

    assert res.success is True
    ps = test_manager.get_sheet([s for s in test_manager.get_sheet_names() if s.startswith("فیش_حقوق_و_دستمزد")][-1])
    assert ps.get_value('A4') == 'E001'
    assert ps.get_value('B4') == 'علی رضایی'
    assert ps.get_value('C4') == 85000000
    assert ps.get_value('D4') == 5000000.0
    assert ps.get_value('F4') == 12000000.0
    assert "ROUND(F4/((C4/220)*1.4), 1)" in str(ps.get_formula('E4'))
    assert ps.get_value('K4') == 2500000.0


def test_payroll_employee_name_with_tavakoli_not_skipped(test_manager):
    """
    Ensure Persian names containing substrings like 'کل' (e.g. 'مریم توکلی')
    are never falsely treated as summary rows ('کل' / 'جمع کل') in payroll.
    """
    sheet = test_manager.get_active_sheet()
    sheet.set_value("A1", "کد")
    sheet.set_value("B1", "نام")
    sheet.set_value("C1", "حقوق پایه")
    sheet.set_value("D1", "اضافه کاری")

    sheet.set_value("A2", "W04")
    sheet.set_value("B2", "مریم توکلی")
    sheet.set_value("C2", 155000000)
    sheet.set_value("D2", 22000000)

    sheet.set_value("A3", "")
    sheet.set_value("B3", "جمع کل")
    sheet.set_value("C3", 155000000)
    sheet.set_value("D3", 22000000)

    ctx = ActionContext(manager=test_manager, sheet=sheet)
    action = PayrollAction()
    res = AccountingActionHandler.execute_payroll(action, ctx)

    assert res.success is True
    ps = test_manager.get_sheet([s for s in test_manager.get_sheet_names() if s.startswith("فیش_حقوق_و_دستمزد")][-1])
    # W04 Maryam Tavakoli must be included in row 4
    assert ps.get_value('A4') == 'W04'
    assert ps.get_value('B4') == 'مریم توکلی'
    assert ps.get_value('C4') == 155000000
    # Row 5 is the generated total row
    assert ps.get_value('A5') == 'جمع کل لیست حقوق:'


def test_attendance_calculation(test_manager):
    """Test attendance, timesheet, clock-in/out, overtime, shortage, and monthly summaries."""
    from excel_assistant.nlp.intent_models import AttendanceAction
    from excel_assistant.nlp.parser import PersianIntentParser
    from excel_assistant.actions.accounting_actions import AccountingActionHandler

    parser = PersianIntentParser()
    action = parser.parse("محاسبه ورود و خروج کارمندان")
    assert isinstance(action, AttendanceAction)

    sheet = test_manager.get_active_sheet()
    ctx = ActionContext(manager=test_manager, sheet=sheet)
    res = AccountingActionHandler.execute_attendance(action, ctx)

    assert res.success is True
    assert "ثبت_ماهانه" in test_manager.get_sheet_names()
    assert "خلاصه_کارمندان" in test_manager.get_sheet_names()

    # Verify daily sheet content (Image 2)
    daily = test_manager.get_sheet("ثبت_ماهانه")
    assert daily.get_value("A2") == "تاریخ"
    assert daily.get_value("B2") == "روز"
    assert daily.get_value("C2") == "ورود"
    assert daily.get_value("D2") == "خروج"
    assert daily.get_value("E2") == "ساعت کار"
    assert daily.get_value("F2") == "اضافه‌کار"
    assert daily.get_value("G2") == "کسری کار"
    assert daily.get_value("H2") == "وضعیت"

    # Day 1: 1 Aban, Friday, Holiday
    assert daily.get_value("A3") == "۱ آبان"
    assert daily.get_value("B3") == "جمعه"
    assert daily.get_value("H3") == "تعطیل"

    # Day 2: 2 Aban, Saturday, Present
    assert daily.get_value("A4") == "۲ آبان"
    assert daily.get_value("C4") == "08:14"
    assert daily.get_value("D4") == "16:16"
    assert daily.get_value("E4") == "8:02"
    assert daily.get_value("F4") == "0:02"

    # Total row (row 33)
    assert daily.get_value("A33") == "جمع کل"
    assert daily.get_value("E33") == "181:38"
    assert daily.get_value("F33") == "8:50"
    assert daily.get_value("G33") == "3:12"

    # Verify summary sheet content (Image 1)
    summary = test_manager.get_sheet("خلاصه_کارمندان")
    assert "خلاصه عملکرد ماهانه" in str(summary.get_value("A1"))
    assert summary.get_value("A2") == "کارمند"
    assert summary.get_value("B2") == "روز کارکرد"
    assert summary.get_value("C2") == "ساعت کار"
    assert summary.get_value("D2") == "اضافه‌کار"
    assert summary.get_value("E2") == "کسری کار"
    assert summary.get_value("F2") == "مرخصی"

    # Check employee rows
    assert summary.get_value("A3") == "علی رضایی"
    assert summary.get_value("B3") == "۲۲ روز"
    assert summary.get_value("C3") == "181:38"
    assert summary.get_value("D3") == "8:50"
    assert summary.get_value("E3") == "3:12"
    assert summary.get_value("F3") == "۳ روز"

    assert summary.get_value("A10") == "زهرا موسوی"
    assert summary.get_value("B10") == "۱۶ روز"


def test_attendance_user_sheet_and_phrases(test_manager):
    """Verify that colloquial attendance phrases parse correctly and user active sheet attendance columns are computed."""
    from excel_assistant.nlp.parser import PersianIntentParser
    from excel_assistant.nlp.intent_models import AttendanceAction
    from excel_assistant.actions.accounting_actions import AccountingActionHandler

    parser = PersianIntentParser()
    for phrase in ["محاسبه ساعت کاری", "ساعت کاری", "محاسبه ساعات کاری", "ساعت کارکرد", "حضور غیاب", "ورود خروج"]:
        act = parser.parse(phrase)
        assert isinstance(act, AttendanceAction), f"Failed for phrase: {phrase}"

    # Active sheet calculation test
    sheet = test_manager.get_active_sheet()
    sheet.set_value("A1", "تاریخ")
    sheet.set_value("B1", "ساعت ورود")
    sheet.set_value("C1", "ساعت خروج")
    sheet.set_value("A2", "1405/08/01")
    sheet.set_value("B2", "08:14")
    sheet.set_value("C2", "16:16")

    action = AttendanceAction()
    ctx = ActionContext(manager=test_manager, sheet=sheet)
    res = AccountingActionHandler.execute_attendance(action, ctx)

    assert res.success is True
    assert sheet.get_value("D1") == "ساعت کار"
    assert sheet.get_value("E1") == "اضافه‌کار"
    assert sheet.get_value("F1") == "کسری کار"
    assert sheet.get_value("G1") == "وضعیت"
    assert sheet.get_value("D2") == "8:02"
    assert sheet.get_value("E2") == "0:02"
    assert sheet.get_value("G2") == "حاضر"


def test_profit_compare_fifo_lifo(test_manager):
    """Verify that comparing profit across FIFO, LIFO, and Average creates comparison sheet and returns side-by-side metrics."""
    from excel_assistant.nlp.parser import PersianIntentParser
    from excel_assistant.nlp.intent_models import ProfitCalculationAction
    from excel_assistant.actions.math_actions import MathActionHandler

    parser = PersianIntentParser()
    act = parser.parse("مقایسه سود به روش فایفو و لایفو")
    assert isinstance(act, ProfitCalculationAction)
    assert act.valuation_method == "compare"

    # Setup sales and purchase sheets
    s1 = test_manager.get_sheet(test_manager.get_sheet_names()[0])
    s1.set_value("A1", "تاریخ")
    s1.set_value("B1", "کالا")
    s1.set_value("C1", "تعداد فروش")
    s1.set_value("D1", "قیمت فروش")
    s1.set_value("A2", "1405/01/05")
    s1.set_value("B2", "ماوس")
    s1.set_value("C2", 10)
    s1.set_value("D2", 150000)

    s2 = test_manager.create_sheet("خرید")
    s2.set_value("A1", "تاریخ")
    s2.set_value("B1", "کالا")
    s2.set_value("C1", "تعداد خرید")
    s2.set_value("D1", "قیمت خرید")
    s2.set_value("A2", "1405/01/01")
    s2.set_value("B2", "ماوس")
    s2.set_value("C2", 5)
    s2.set_value("D2", 100000)
    s2.set_value("A3", "1405/01/03")
    s2.set_value("B3", "ماوس")
    s2.set_value("C3", 10)
    s2.set_value("D3", 120000)

    ctx = ActionContext(manager=test_manager, sheet=s1)
    res = MathActionHandler.execute_profit_calculation(act, ctx)

    assert res.success is True
    assert "مقایسه تطبیقی سود" in res.message_persian
    assert "مقایسه_روش‌های_سود" in test_manager.get_sheet_names()
    cmp_sheet = test_manager.get_sheet("مقایسه_روش‌های_سود")
    assert "فایفو" in str(cmp_sheet.get_value("A1"))
    assert res.data["profit_fifo"] > res.data["profit_lifo"]
    assert res.data["profit_difference"] == 100000.0













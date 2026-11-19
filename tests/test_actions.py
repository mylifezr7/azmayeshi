"""
Tests for Phase 4: Excel Action Execution, Calculations, Data Management, Formatting, and Undo.
"""
import pytest

from excel_assistant.core.connection import ExcelConnectionManager
from excel_assistant.actions.base import ActionContext
from excel_assistant.actions.executor import ActionExecutor
from excel_assistant.nlp.parser import PersianIntentParser
from excel_assistant.analyzer.inspector import SheetInspector


@pytest.fixture
def test_session():
    manager = ExcelConnectionManager(preferred_driver="openpyxl")
    manager.connect()
    sheet = manager.get_active_sheet()

    # Fill sample data
    # Row 1 (Header)
    sheet.set_value("A1", "کد")
    sheet.set_value("B1", "تعداد")
    sheet.set_value("C1", "قیمت")
    sheet.set_value("D1", "امتیاز")

    # Row 2
    sheet.set_value("A2", "P1")
    sheet.set_value("B2", 10)
    sheet.set_value("C2", 50)
    sheet.set_value("D2", 120)

    # Row 3
    sheet.set_value("A3", "P2")
    sheet.set_value("B3", 20)
    sheet.set_value("C3", 30)
    sheet.set_value("D3", 80)

    # Row 4 (Duplicate of P2)
    sheet.set_value("A4", "P2")
    sheet.set_value("B4", 20)
    sheet.set_value("C4", 30)
    sheet.set_value("D4", 80)

    # Row 5 (Empty row)

    # Row 6
    sheet.set_value("A6", "P3")
    sheet.set_value("B6", 5)
    sheet.set_value("C6", 200)
    sheet.set_value("D6", 150)

    executor = ActionExecutor()
    parser = PersianIntentParser()
    yield manager, sheet, executor, parser
    manager.close()


def test_column_arithmetic_execution(test_session):
    manager, sheet, executor, parser = test_session
    analysis = SheetInspector.inspect_sheet(sheet)
    ctx = ActionContext(manager=manager, sheet=sheet, sheet_analysis=analysis, confirmed=True)

    # User command: ستون B را در ستون C ضرب کن و جواب را در E بنویس
    cmd = "ستون B را در ستون C ضرب کن و جواب را در E بنویس"
    action = parser.parse(cmd, sheet_analysis=analysis)
    result = executor.execute(action, ctx)

    assert result.success is True
    assert "ضرب" in result.message_persian
    # Check that formulas are written in column E
    assert sheet.get_formula("E2") == "=B2*C2"
    assert sheet.get_formula("E3") == "=B3*C3"
    assert sheet.get_value("E1") is not None


def test_column_stats_execution(test_session):
    manager, sheet, executor, parser = test_session
    analysis = SheetInspector.inspect_sheet(sheet)
    ctx = ActionContext(manager=manager, sheet=sheet, sheet_analysis=analysis, confirmed=True)

    # میانگین ستون D را حساب کن
    cmd = "میانگین ستون D را حساب کن"
    action = parser.parse(cmd, sheet_analysis=analysis)
    result = executor.execute(action, ctx)

    assert result.success is True
    assert "میانگین" in result.message_persian
    assert result.data is not None  # python mean calculated


def test_highlight_condition_execution(test_session):
    manager, sheet, executor, parser = test_session
    analysis = SheetInspector.inspect_sheet(sheet)
    ctx = ActionContext(manager=manager, sheet=sheet, sheet_analysis=analysis, confirmed=True)

    # سلول‌های بیشتر از 100 را سبز کن
    cmd = "سلول‌های بیشتر از 100 را سبز کن"
    action = parser.parse(cmd, sheet_analysis=analysis)
    result = executor.execute(action, ctx)

    assert result.success is True
    assert len(result.data) > 0  # D2 (120), C6 (200), D6 (150)


def test_highlight_duplicates_execution(test_session):
    manager, sheet, executor, parser = test_session
    analysis = SheetInspector.inspect_sheet(sheet)
    ctx = ActionContext(manager=manager, sheet=sheet, sheet_analysis=analysis, confirmed=True)

    # داده‌های تکراری را زرد کن
    cmd = "داده‌های تکراری را زرد کن"
    action = parser.parse(cmd, sheet_analysis=analysis)
    result = executor.execute(action, ctx)

    assert result.success is True
    assert len(result.data) > 0


def test_delete_empty_rows_with_safety_and_undo(test_session):
    manager, sheet, executor, parser = test_session
    analysis = SheetInspector.inspect_sheet(sheet)

    cmd = "ردیف‌های خالی را حذف کن"
    action = parser.parse(cmd, sheet_analysis=analysis)

    # 1. Unconfirmed execution should require confirmation
    unconfirmed_ctx = ActionContext(manager=manager, sheet=sheet, sheet_analysis=analysis, confirmed=False)
    res_unconf = executor.execute(action, unconfirmed_ctx)
    assert res_unconf.requires_confirmation is True

    # 2. Confirmed execution succeeds
    confirmed_ctx = ActionContext(manager=manager, sheet=sheet, sheet_analysis=analysis, confirmed=True)
    res_conf = executor.execute(action, confirmed_ctx)
    assert res_conf.success is True
    assert res_conf.data == 1  # 1 empty row deleted

    # 3. Test Undo
    undo_cmd = "لغو آخرین عملیات"
    undo_act = parser.parse(undo_cmd)
    undo_res = executor.execute(undo_act, confirmed_ctx)
    assert undo_res.success is True
    assert "لغو شد" in undo_res.message_persian


def test_sort_data_execution(test_session):
    manager, sheet, executor, parser = test_session
    analysis = SheetInspector.inspect_sheet(sheet)
    ctx = ActionContext(manager=manager, sheet=sheet, sheet_analysis=analysis, confirmed=True)

    # اطلاعات را از بیشترین به کمترین مرتب کن
    cmd = "اطلاعات را از بیشترین به کمترین مرتب کن"
    action = parser.parse(cmd, sheet_analysis=analysis)
    result = executor.execute(action, ctx)

    assert result.success is True
    assert "مرتب" in result.message_persian


def test_create_table_execution(test_session):
    manager, sheet, executor, parser = test_session
    analysis = SheetInspector.inspect_sheet(sheet)
    ctx = ActionContext(manager=manager, sheet=sheet, sheet_analysis=analysis, confirmed=True)

    cmd = "یک جدول بساز"
    action = parser.parse(cmd, sheet_analysis=analysis)
    result = executor.execute(action, ctx)

    assert result.success is True
    assert "جدول" in result.message_persian

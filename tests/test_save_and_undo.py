"""
Tests for Granular Cell Delta Undo and Save Workbook functionality.
Ensures manual user data is never wiped out on undo, and save options work across all interfaces.
"""
import openpyxl
import pytest
from pathlib import Path

from excel_assistant.core.openpyxl_driver import OpenpyxlSheet
from excel_assistant.core.connection import ExcelConnectionManager
from excel_assistant.actions.undo_manager import UndoManager
from excel_assistant.actions.executor import ActionExecutor
from excel_assistant.actions.base import ActionContext
from excel_assistant.nlp.intent_models import (
    ColumnArithmeticAction,
    UndoLastAction,
    SaveWorkbookAction,
)
from excel_assistant.nlp.parser import PersianIntentParser
from excel_assistant.ui.style import QSS_STYLE
from excel_assistant.ui.web_server import HTML_TEMPLATE


def test_granular_undo_preserves_user_manual_entries(tmp_path):
    """Verify that Undo only reverts cells touched by the AI action, leaving manual entries intact."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = 10
    ws["B1"] = 20
    ws["A2"] = 30
    ws["B2"] = 40
    test_file = str(tmp_path / "test_undo.xlsx")
    wb.save(test_file)

    manager = ExcelConnectionManager()
    manager.open_workbook(test_file)
    sheet = manager.get_active_sheet()

    undo_mgr = UndoManager()
    executor = ActionExecutor(undo_mgr)
    ctx = ActionContext(manager=manager, sheet=sheet, confirmed=True)

    # 1. AI Action: multiply A and B into C
    action = ColumnArithmeticAction(col_left="A", col_right="B", target_col="C", operation="multiply")
    res = executor.execute(action, ctx)
    assert res.success is True
    assert sheet.get_value("C1") == "=A1*B1"
    assert sheet.get_value("C2") == "=A2*B2"

    # 2. User manually enters data in unrelated cells and rows
    sheet.set_value("D1", "User Note 1")
    sheet.set_value("E1", 12345)
    sheet.set_value("A3", "Manual Row 3")
    sheet.set_value("B3", 500)
    sheet.set_value("Z10", "Far Away Cell")

    # 3. User triggers Undo
    undo_act = UndoLastAction()
    undo_res = executor.execute(undo_act, ctx)
    assert undo_res.success is True

    # 4. Assert: C1 and C2 (touched by AI) are reverted to None
    assert sheet.get_value("C1") is None
    assert sheet.get_value("C2") is None

    # 5. Assert: All manual entries are 100% PRESERVED!
    assert sheet.get_value("A1") == 10
    assert sheet.get_value("B1") == 20
    assert sheet.get_value("A2") == 30
    assert sheet.get_value("B2") == 40
    assert sheet.get_value("D1") == "User Note 1"
    assert sheet.get_value("E1") == 12345
    assert sheet.get_value("A3") == "Manual Row 3"
    assert sheet.get_value("B3") == 500
    assert sheet.get_value("Z10") == "Far Away Cell"


def test_save_workbook_nlp_parsing():
    """Verify that Persian parser resolves various saving phrases into SaveWorkbookAction."""
    parser = PersianIntentParser()
    commands = [
        "ذخیره فایل",
        "سیو فایل",
        "فایل را ذخیره کن",
        "ذخیره تغییرات",
        "تغییرات را ذخیره کن",
        "ذخیره کن",
        "سیو کن",
        "ذخیره",
        "سیو",
        "save",
    ]
    for cmd in commands:
        action = parser.parse(cmd)
        assert isinstance(action, SaveWorkbookAction), f"Failed for command: {cmd}"


def test_save_workbook_execution(tmp_path):
    """Verify that executing SaveWorkbookAction saves the workbook to disk."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "SavedData"
    test_file = str(tmp_path / "test_save.xlsx")
    wb.save(test_file)

    manager = ExcelConnectionManager()
    manager.open_workbook(test_file)
    sheet = manager.get_active_sheet()
    sheet.set_value("B1", "NewData")

    undo_mgr = UndoManager()
    executor = ActionExecutor(undo_mgr)
    ctx = ActionContext(manager=manager, sheet=sheet, confirmed=True)

    action = SaveWorkbookAction()
    res = executor.execute(action, ctx)
    assert res.success is True
    assert "با موفقیت ذخیره شد" in res.message_persian

    # Verify on disk
    wb_reloaded = openpyxl.load_workbook(test_file)
    ws_reloaded = wb_reloaded.active
    assert ws_reloaded["B1"].value == "NewData"


def test_ui_save_elements():
    """Verify that Desktop and Web UI have the save button, styles, and shortcuts."""
    # Desktop QSS
    assert "QPushButton#saveButton" in QSS_STYLE

    # Web Template
    assert "btn-save" in HTML_TEMPLATE
    assert "saveWorkbook" in HTML_TEMPLATE
    assert "/api/save" in HTML_TEMPLATE
    assert "Ctrl+S" in HTML_TEMPLATE

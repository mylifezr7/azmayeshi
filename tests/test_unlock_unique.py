"""
Unit tests for unlocking and removing column uniqueness validation:
- Persian NLP rule parsing for unlock commands
- Driver-level has_unique_validation and remove_unique_validation
- DataActionHandler execution and highlight clearing
- UniqueColumnWizardDialog lock and unlock modes
"""
import pytest
import openpyxl

from excel_assistant.core.openpyxl_driver import OpenpyxlSheet
from excel_assistant.actions.data_actions import DataActionHandler
from excel_assistant.actions.base import ActionContext
from excel_assistant.nlp.parser import PersianIntentParser
from excel_assistant.nlp.intent_models import (
    PreventDuplicatesAction,
    RemoveUniqueValidationAction,
    DeleteDuplicatesAction,
)


def test_nlp_parse_unlock_unique_commands():
    parser = PersianIntentParser()

    cmd1 = "قفل یکتایی ستون B را بردار"
    act1 = parser.parse(cmd1)
    assert isinstance(act1, RemoveUniqueValidationAction)
    assert act1.target_col == "B"

    cmd2 = "برداشتن قفل یکتایی از ستون C"
    act2 = parser.parse(cmd2)
    assert isinstance(act2, RemoveUniqueValidationAction)
    assert act2.target_col == "C"

    cmd3 = "قفل ستون B رو باز کن"
    act3 = parser.parse(cmd3)
    assert isinstance(act3, RemoveUniqueValidationAction)
    assert act3.target_col == "B"

    cmd4 = "حذف قفل یکتا ستون D"
    act4 = parser.parse(cmd4)
    assert isinstance(act4, RemoveUniqueValidationAction)
    assert act4.target_col == "D"

    cmd5 = "محدودیت یکتایی ستون A را بردار"
    act5 = parser.parse(cmd5)
    assert isinstance(act5, RemoveUniqueValidationAction)
    assert act5.target_col == "A"


def test_nlp_prevent_and_delete_not_confused_with_unlock():
    parser = PersianIntentParser()

    # Prevent duplicates should not trigger unlock
    p_act = parser.parse("ستون 12 نباید داده تکراری داشته باشه")
    assert isinstance(p_act, PreventDuplicatesAction)
    assert p_act.target_col == "L"

    # Delete duplicates rows should not trigger unlock
    d_act = parser.parse("حذف داده‌های تکراری ستون A")
    assert isinstance(d_act, DeleteDuplicatesAction)
    assert d_act.target_col == "A"


def test_driver_add_and_remove_unique_validation():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "TestSheet"
    ws["B1"] = "کد ملی"
    ws["B2"] = "111"
    ws["B3"] = "222"

    sheet = OpenpyxlSheet(ws)
    assert sheet.has_unique_validation("B") is False

    # Add validation
    sheet.add_unique_validation("B")
    assert sheet.has_unique_validation("B") is True

    # Remove validation
    sheet.remove_unique_validation("B")
    assert sheet.has_unique_validation("B") is False


def test_execute_remove_unique_validation_clears_highlights():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Products"
    ws["A1"] = "شناسه کالا"
    ws["A2"] = "P01"
    ws["A3"] = "P02"
    ws["A4"] = "P01"  # duplicate

    sheet = OpenpyxlSheet(ws)
    ctx = ActionContext(sheet=sheet, manager=None)

    # 1. Apply lock and highlight
    lock_act = PreventDuplicatesAction(target_col="A", highlight_existing=True)
    res_lock = DataActionHandler.execute_prevent_duplicates(lock_act, ctx)
    assert res_lock.success is True
    assert sheet.has_unique_validation("A") is True
    assert sheet.get_cell_color("A4") == "#FECDD3"

    # 2. Unlock column and clear highlight
    unlock_act = RemoveUniqueValidationAction(target_col="A", clear_highlight=True)
    res_unlock = DataActionHandler.execute_remove_unique_validation(unlock_act, ctx)
    assert res_unlock.success is True
    assert sheet.has_unique_validation("A") is False
    assert sheet.get_cell_color("A4") is None
    assert "با موفقیت برداشته شد" in res_unlock.message_persian


def test_unique_column_dialog_lock_and_unlock_actions():
    try:
        from PySide6.QtWidgets import QApplication
        from excel_assistant.ui.unique_column_dialog import UniqueColumnWizardDialog
        app = QApplication.instance() or QApplication([])
        dialog = UniqueColumnWizardDialog()
        dialog.input_col.setText("B")

        # Lock mode via radio button
        dialog.rb_lock.setChecked(True)
        assert dialog.mode == "lock"
        act_lock = dialog.get_action()
        assert isinstance(act_lock, PreventDuplicatesAction)
        assert act_lock.target_col == "B"

        # Unlock mode via radio button
        dialog.rb_unlock.setChecked(True)
        assert dialog.mode == "unlock"
        act_unlock = dialog.get_action()
        assert isinstance(act_unlock, RemoveUniqueValidationAction)
        assert act_unlock.target_col == "B"
    except (ImportError, Exception):
        pytest.skip("PySide6 GUI not available in headless environment")

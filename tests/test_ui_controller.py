"""
Tests for UI Controller / Presenter (Phase 5).
"""
import pytest

from excel_assistant.core.connection import ExcelConnectionManager
from excel_assistant.ui.controller import AssistantUIController


def test_ui_controller_lifecycle():
    manager = ExcelConnectionManager(preferred_driver="openpyxl")
    controller = AssistantUIController(manager=manager)

    info = controller.connect()
    assert controller.get_status()["connected"] is True
    assert len(controller.messages) >= 1  # Welcome message
    assert "Microsoft Excel" in controller.messages[0].text or "اکسل" in controller.messages[0].text

    # Setup sample data in active sheet
    sheet = manager.get_active_sheet()
    sheet.set_value("A1", "محصول")
    sheet.set_value("B1", "تعداد")
    sheet.set_value("C1", "قیمت")
    sheet.set_value("A2", "ماوس")
    sheet.set_value("B2", 10)
    sheet.set_value("C2", 25)

    # 1. Test standard command
    res = controller.process_command("ستون B را در ستون C ضرب کن و جواب را در D بنویس")
    assert res.success is True
    assert len(controller.messages) >= 3  # Welcome + User + Assistant
    assert sheet.get_formula("D2") == "=B2*C2"

    # 2. Test destructive command confirmation flow
    sheet.set_value("A3", None)
    sheet.set_value("B3", None)
    res_destr = controller.process_command("ردیف‌های خالی را حذف کن")
    assert res_destr.requires_confirmation is True
    assert controller.pending_action is not None
    last_msg = controller.messages[-1]
    assert last_msg.is_confirmation is True

    # 3. Confirm pending action
    conf_res = controller.confirm_pending_action()
    assert conf_res.success is True
    assert controller.pending_action is None

    # 4. Test undo
    undo_res = controller.undo_last_action()
    assert undo_res.success is True
    assert "لغو شد" in undo_res.message_persian

    # 5. Test switch sheet
    manager.active_driver.create_sheet("ReportSheet")
    controller.switch_sheet("ReportSheet")
    assert controller.get_status()["active_sheet"] == "ReportSheet"

    # 6. Test ambiguous command
    amb_res = controller.process_command("چطوری هوای امروز چطوره؟")
    assert amb_res.success is False
    assert "متوجه دستور" in controller.messages[-1].text

    manager.close()

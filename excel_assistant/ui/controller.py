"""
UI Controller / Presenter for Excel AI Assistant.
Decouples UI business logic, command flow, and state from rendering frameworks (PySide6 & Web UI).
"""
from dataclasses import dataclass, field
from typing import Any, Callable, Dict, List, Optional

from excel_assistant.core.connection import ExcelConnectionManager
from excel_assistant.core.interfaces import WorkbookInfo
from excel_assistant.analyzer.inspector import SheetInspector, SheetAnalysis, WorkbookInspector
from excel_assistant.nlp.parser import PersianIntentParser
from excel_assistant.nlp.intent_models import ExcelAction, UndoLastAction, AnalyzeSheetAction
from excel_assistant.actions.base import ActionContext, ActionResult
from excel_assistant.actions.executor import ActionExecutor
from excel_assistant.actions.undo_manager import UndoManager
from excel_assistant.core.exceptions import AmbiguousCommandError


@dataclass
class ChatMessage:
    id: int
    sender: str  # "user" or "assistant"
    text: str
    time_str: str
    is_confirmation: bool = False
    action_payload: Optional[Dict[str, Any]] = None


class AssistantUIController:
    """Manages chat messages, command processing, confirmations, and workbook states."""

    def __init__(self, manager: Optional[ExcelConnectionManager] = None):
        self.manager = manager or ExcelConnectionManager()
        self.parser = PersianIntentParser()
        self.undo_manager = UndoManager()
        self.executor = ActionExecutor(self.undo_manager)

        self.messages: List[ChatMessage] = []
        self._next_id = 1
        self.pending_action: Optional[ExcelAction] = None
        self.history: List[str] = []

        # Callbacks for UI updates
        self.on_message_added: Optional[Callable[[ChatMessage], None]] = None
        self.on_status_changed: Optional[Callable[[WorkbookInfo], None]] = None

    def connect(self, workbook_path_or_name: Optional[str] = None) -> WorkbookInfo:
        info = self.manager.connect(workbook_path_or_name)
        if self.on_status_changed:
            self.on_status_changed(info)

        desc = "Microsoft Excel 2024 (COM)" if self.manager.driver_type == "xlwings" else "موتور مستقل اکسل"
        welcome_msg = (
            f"سلام! من دستیار هوشمند Microsoft Excel هستم.\n"
            f"✅ وضعیت: متصل به **{desc}**\n"
            f"📁 فایل فعال: **`{info.name}`**\n"
            f"📑 شیت فعلی: **`{info.active_sheet}`**\n"
            f"شما می‌توانید هر دستوری را به زبان فارسی مطرح کنید تا مستقیماً روی اکسل اجرا نمایم."
        )
        self.add_message("assistant", welcome_msg)
        return info

    def get_status(self) -> Dict[str, Any]:
        if not self.manager.is_connected():
            return {
                "connected": False,
                "driver": None,
                "file_name": "-",
                "active_sheet": "-",
                "sheets": [],
                "open_workbooks": []
            }
        info = self.manager.get_workbook_info()
        return {
            "connected": True,
            "driver": self.manager.driver_type,
            "file_name": info.name,
            "active_sheet": info.active_sheet,
            "sheets": info.sheets,
            "open_workbooks": self.manager.list_open_workbooks()
        }

    def open_workbook(self, file_path: str) -> str:
        info = self.manager.open_workbook(file_path)
        msg = f"📂 فایل اکسل **`{info.name}`** باز شد و دستیار به آن متصل گردید.\n📑 شیت فعال: **`{info.active_sheet}`**"
        self.add_message("assistant", msg)
        return msg

    def switch_workbook(self, wb_name: str) -> str:
        info = self.manager.switch_workbook(wb_name)
        msg = f"🔄 ورک‌بوک فعال به **`{info.name}`** تغییر یافت."
        self.add_message("assistant", msg)
        return msg

    def switch_sheet(self, sheet_name: str) -> str:
        self.manager.activate_sheet(sheet_name)
        msg = f"📌 شیت فعال به **`{sheet_name}`** تغییر یافت."
        self.add_message("assistant", msg)
        return msg

    def add_message(self, sender: str, text: str, is_confirmation: bool = False,
                    action_payload: Optional[Dict[str, Any]] = None) -> ChatMessage:
        import time
        msg = ChatMessage(
            id=self._next_id,
            sender=sender,
            text=text,
            time_str=time.strftime("%H:%M"),
            is_confirmation=is_confirmation,
            action_payload=action_payload
        )
        self._next_id += 1
        self.messages.append(msg)
        if self.on_message_added:
            self.on_message_added(msg)
        return msg

    def process_command(self, raw_command: str) -> ActionResult:
        cmd = raw_command.strip()
        if not cmd:
            return ActionResult(success=False, message_persian="دستور خالی است.")

        # Record in history
        self.history.append(cmd)
        self.add_message("user", cmd)

        if not self.manager.is_connected():
            msg = "❌ ابتدا باید اتصال به فایل اکسل برقرار باشد."
            self.add_message("assistant", msg)
            return ActionResult(success=False, message_persian=msg)

        sheet = self.manager.get_active_sheet()
        sa = SheetInspector.inspect_sheet(sheet)

        # Parse command
        try:
            action = self.parser.parse(cmd, sheet_analysis=sa)
        except AmbiguousCommandError as ace:
            opts = "\n".join(f"• {opt}" for opt in ace.suggested_options)
            resp = f"⚠️ {ace}\n{ace.clarification_prompt}\n{opts}"
            self.add_message("assistant", resp)
            return ActionResult(success=False, message_persian=resp)

        # Execute
        ctx = ActionContext(manager=self.manager, sheet=sheet, sheet_analysis=sa, confirmed=False)
        result = self.executor.execute(action, ctx)

        if result.requires_confirmation:
            self.pending_action = action
            self.add_message(
                sender="assistant",
                text=result.message_persian,
                is_confirmation=True,
                action_payload=action.model_dump() if hasattr(action, "model_dump") else (action.dict() if hasattr(action, "dict") else None)
            )
            return result

        self.add_message("assistant", result.message_persian)
        return result

    def confirm_pending_action(self) -> ActionResult:
        if not self.pending_action:
            msg = "⚠️ هیچ عملیاتی در انتظار تأیید نیست."
            self.add_message("assistant", msg)
            return ActionResult(success=False, message_persian=msg)

        action = self.pending_action
        self.pending_action = None

        sheet = self.manager.get_active_sheet()
        sa = SheetInspector.inspect_sheet(sheet)
        ctx = ActionContext(manager=self.manager, sheet=sheet, sheet_analysis=sa, confirmed=True)

        result = self.executor.execute(action, ctx)
        self.add_message("assistant", result.message_persian)
        return result

    def cancel_pending_action(self) -> None:
        self.pending_action = None
        self.add_message("assistant", "عملیات موردنظر به درخواست شما لغو شد.")

    def undo_last_action(self) -> ActionResult:
        return self.process_command("لغو آخرین عملیات")

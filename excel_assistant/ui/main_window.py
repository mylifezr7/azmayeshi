"""
Main Window for Excel AI Assistant (PySide6 / RTL Persian).
"""
import sys
from typing import Optional

from PySide6.QtCore import Qt, QThread, Signal, Slot
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLineEdit,
    QPushButton, QLabel, QFrame, QMessageBox, QApplication, QDialog
)

from excel_assistant.core.connection import ExcelConnectionManager
from excel_assistant.analyzer.inspector import SheetInspector, WorkbookInspector
from excel_assistant.nlp.parser import PersianIntentParser
from excel_assistant.nlp.intent_models import ExcelAction, UndoLastAction, AnalyzeSheetAction
from excel_assistant.actions.base import ActionContext, ActionResult
from excel_assistant.actions.executor import ActionExecutor
from excel_assistant.actions.undo_manager import UndoManager
from excel_assistant.ui.style import QSS_STYLE
from excel_assistant.ui.status_bar import ExcelStatusBar
from excel_assistant.ui.chat_widget import ChatWidget
from excel_assistant.ui.history_panel import HistoryPanel
from excel_assistant.core.exceptions import AmbiguousCommandError


class MainWindow(QMainWindow):
    """Main Application Window."""

    def __init__(self, manager: Optional[ExcelConnectionManager] = None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle("دستیار هوشمند اکسل - Excel AI Assistant")
        self.resize(1050, 720)
        self.setLayoutDirection(Qt.RightToLeft)
        self.setStyleSheet(QSS_STYLE)
        self.setAcceptDrops(True)

        # Core components
        self.manager = manager or ExcelConnectionManager()
        self.parser = PersianIntentParser()
        self.undo_manager = UndoManager()
        self.executor = ActionExecutor(self.undo_manager)

        self.pending_destructive_action: Optional[ExcelAction] = None

        self._init_ui()
        self._connect_signals()
        self._perform_initial_connection()

    def _init_ui(self):
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # 1. Top status bar
        self.status_bar_widget = ExcelStatusBar(self)
        main_layout.addWidget(self.status_bar_widget)

        # 2. Content area (Chat + History Panel)
        content_layout = QHBoxLayout()
        content_layout.setContentsMargins(12, 12, 12, 12)
        content_layout.setSpacing(12)

        self.chat_widget = ChatWidget(self)
        self.history_panel = HistoryPanel(self)

        content_layout.addWidget(self.chat_widget, stretch=4)
        content_layout.addWidget(self.history_panel, stretch=1)
        main_layout.addLayout(content_layout)

        # 3. Quick Chips Bar
        chips_frame = QFrame()
        chips_layout = QHBoxLayout(chips_frame)
        chips_layout.setContentsMargins(12, 0, 12, 4)
        chips_layout.setSpacing(8)

        quick_prompts = [
            "➕ جمع ستون‌ها",
            "➖ تفریق ستون‌ها",
            "✖️ ضرب ستون‌ها",
            "➗ تقسیم ستون‌ها",
            "🏷️ اعمال تخفیف کالاها",
            "📊 تحلیل شیت",
            "🧹 ردیف‌های خالی را حذف کن",
            "🟨 داده‌های تکراری را زرد کن",
            "🟩 سلول‌های بیشتر از 100 را سبز کن",
            "📐 عرض ستون‌ها را تنظیم کن",
            "📋 یک جدول بساز"
        ]
        for prompt in quick_prompts:
            btn = QPushButton(prompt)
            btn.setProperty("class", "quickChip")
            clean_cmd = prompt.split(" ", 1)[-1]
            btn.clicked.connect(lambda _, c=clean_cmd: self._on_quick_chip_clicked(c))
            chips_layout.addWidget(btn)

        chips_layout.addStretch()
        main_layout.addWidget(chips_frame)

        # 4. Input bar
        input_frame = QFrame()
        input_frame.setObjectName("inputPanel")
        input_layout = QHBoxLayout(input_frame)
        input_layout.setContentsMargins(12, 8, 12, 12)
        input_layout.setSpacing(10)

        self.input_edit = QLineEdit()
        self.input_edit.setObjectName("chatInput")
        self.input_edit.setPlaceholderText("دستور خود را به زبان فارسی وارد کنید... (مثلاً: ستون A را در ستون B ضرب کن و در C بنویس)")
        self.input_edit.returnPressed.connect(self._on_send_clicked)

        self.btn_send = QPushButton("ارسال دستور")
        self.btn_send.setObjectName("sendButton")
        self.btn_send.clicked.connect(self._on_send_clicked)

        self.btn_stop = QPushButton("توقف")
        self.btn_stop.setObjectName("stopButton")
        self.btn_stop.setEnabled(False)

        input_layout.addWidget(self.input_edit)
        input_layout.addWidget(self.btn_send)
        input_layout.addWidget(self.btn_stop)

        main_layout.addWidget(input_frame)

    def _connect_signals(self):
        self.status_bar_widget.reconnect_requested.connect(self._perform_initial_connection)
        self.status_bar_widget.sheet_switched.connect(self._on_sheet_switched)
        self.status_bar_widget.workbook_switched.connect(self._on_workbook_switched)
        self.status_bar_widget.open_file_requested.connect(self._on_open_file_clicked)
        self.history_panel.command_selected.connect(self._on_history_item_selected)
        self.history_panel.undo_requested.connect(self._on_undo_clicked)
        self.history_panel.analyze_requested.connect(self._on_analyze_clicked)
        self.history_panel.arithmetic_requested.connect(lambda: self._open_arithmetic_dialog("add"))
        self.history_panel.discount_requested.connect(self._open_discount_dialog)

    def _perform_initial_connection(self):
        try:
            if not self.manager.is_connected():
                info = self.manager.connect()
            else:
                info = self.manager.get_workbook_info()
            desc = "Microsoft Excel 2024 (COM)" if self.manager.driver_type == "xlwings" else "موتور مستقل اکسل"
            open_wbs = self.manager.list_open_workbooks()
            self.status_bar_widget.set_connected(True, desc)
            self.status_bar_widget.update_workbook_info(info.name, info.sheets, info.active_sheet, open_wbs)
            self.chat_widget.add_assistant_message(
                f"سلام! من دستیار هوشمند Microsoft Excel هستم.\n"
                f"اتصال به فایل **`{info.name}`** و شیت **`{info.active_sheet}`** برقرار شد.\n"
                f"💡 برای کار روی فایل دیگر، روی دکمهٔ **«📂 باز کردن فایل اکسل...»** در نوار بالا کلیک کنید."
            )
        except Exception as e:
            self.status_bar_widget.set_connected(False)
            self.chat_widget.add_assistant_message(
                f"⚠️ خطا در اتصال به اکسل: {e}\n"
                f"لطفاً مطمئن شوید Microsoft Excel باز است یا فایلی برای کار انتخاب شده است."
            )

    def _on_open_file_clicked(self):
        from PySide6.QtWidgets import QFileDialog
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "انتخاب و اتصال به فایل اکسل",
            "",
            "فایل‌های اکسل (*.xlsx *.xls *.xlsm *.csv);;همه فایل‌ها (*)"
        )
        if file_path:
            try:
                info = self.manager.open_workbook(file_path)
                open_wbs = self.manager.list_open_workbooks()
                self.status_bar_widget.update_workbook_info(info.name, info.sheets, info.active_sheet, open_wbs)
                self.chat_widget.add_assistant_message(
                    f"📂 فایل اکسل **`{info.name}`** باز شد و دستیار به آن متصل گردید.\n"
                    f"📑 شیت فعال: **`{info.active_sheet}`**"
                )
            except Exception as e:
                self.chat_widget.add_assistant_message(f"❌ خطا در باز کردن فایل اکسل: {e}")

    def _on_workbook_switched(self, wb_name: str):
        try:
            info = self.manager.switch_workbook(wb_name)
            open_wbs = self.manager.list_open_workbooks()
            self.status_bar_widget.update_workbook_info(info.name, info.sheets, info.active_sheet, open_wbs)
            self.chat_widget.add_assistant_message(f"🔄 اتصال به فایل **`{info.name}`** برقرار گردید.")
        except Exception as e:
            self.chat_widget.add_assistant_message(f"❌ خطا در جابه‌جایی فایل: {e}")

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if urls:
            local_path = urls[0].toLocalFile()
            if local_path.lower().endswith(('.xlsx', '.xls', '.xlsm', '.csv')):
                try:
                    info = self.manager.open_workbook(local_path)
                    open_wbs = self.manager.list_open_workbooks()
                    self.status_bar_widget.update_workbook_info(info.name, info.sheets, info.active_sheet, open_wbs)
                    self.chat_widget.add_assistant_message(f"📂 فایل اکسل **`{info.name}`** بارگذاری شد.")
                except Exception as e:
                    self.chat_widget.add_assistant_message(f"❌ خطا در بارگذاری فایل: {e}")

    def _on_sheet_switched(self, sheet_name: str):
        try:
            self.manager.activate_sheet(sheet_name)
            self.chat_widget.add_assistant_message(f"📌 شیت فعال به **`{sheet_name}`** تغییر یافت.")
        except Exception as e:
            self.chat_widget.add_assistant_message(f"❌ خطا در تغییر شیت: {e}")

    def _on_quick_chip_clicked(self, command: str):
        if "تخفیف" in command:
            self._open_discount_dialog()
        elif "جمع" in command:
            self._open_arithmetic_dialog("add")
        elif "تفریق" in command or "منها" in command:
            self._open_arithmetic_dialog("subtract")
        elif "ضرب" in command:
            self._open_arithmetic_dialog("multiply")
        elif "تقسیم" in command:
            self._open_arithmetic_dialog("divide")
        else:
            self.input_edit.setText(command)
            self._on_send_clicked()

    def _open_discount_dialog(self):
        """Opens interactive wizard dialog for calculating and applying item discounts."""
        from excel_assistant.ui.discount_dialog import DiscountWizardDialog
        dialog = DiscountWizardDialog(manager=self.manager, parent=self)
        if dialog.exec() == QDialog.Accepted:
            action = dialog.get_action()
            self.chat_widget.add_user_message(action.description_persian)
            self.history_panel.add_command(action.description_persian)
            self._execute_action_pipeline(action)

    def _open_arithmetic_dialog(self, initial_op: str = "add"):
        """Opens interactive wizard dialog for 4 basic arithmetic operations."""
        from excel_assistant.ui.arithmetic_dialog import ArithmeticWizardDialog
        dialog = ArithmeticWizardDialog(manager=self.manager, initial_op=initial_op, parent=self)
        if dialog.exec() == QDialog.Accepted:
            action = dialog.get_action()
            self.chat_widget.add_user_message(action.description_persian)
            self.history_panel.add_command(action.description_persian)
            self._execute_action_pipeline(action)

    def _on_history_item_selected(self, command: str):
        self.input_edit.setText(command)

    def _on_undo_clicked(self):
        self.process_command("لغو آخرین عملیات")

    def _on_analyze_clicked(self):
        self.process_command("تحلیل ساختار فایل و شیت")

    def _on_send_clicked(self):
        text = self.input_edit.text().strip()
        if not text:
            return
        self.input_edit.clear()

        # Check for standalone arithmetic keywords to launch wizard
        lower = text.lower()
        if "تخفیف" in lower and len(lower.split()) <= 2:
            self._open_discount_dialog()
            return

        if lower in ["جمع", "ضرب", "تفریق", "تقسیم", "4 عمل اصلی", "۴ عمل اصلی", "محاسبه", "محاسبات", "ماشین حساب"]:
            op_map = {"جمع": "add", "تفریق": "subtract", "ضرب": "multiply", "تقسیم": "divide"}
            self._open_arithmetic_dialog(op_map.get(lower, "add"))
            return

        self.btn_send.setEnabled(False)
        self.btn_send.setText("⏳ در حال اجرا...")
        QApplication.processEvents()

        try:
            self.chat_widget.add_user_message(text)
            self.history_panel.add_command(text)
            QApplication.processEvents()
            self.process_command(text)
        finally:
            self.btn_send.setEnabled(True)
            self.btn_send.setText("ارسال دستور")
            self.input_edit.setFocus()

    def process_command(self, raw_command: str, confirmed: bool = False):
        if not self.manager.is_connected():
            self.chat_widget.add_assistant_message("❌ ابتدا باید اتصال به اکسل برقرار باشد.")
            return

        sheet = self.manager.get_active_sheet()
        sa = SheetInspector.inspect_sheet(sheet)

        try:
            action = self.parser.parse(raw_command, sheet_analysis=sa)
        except AmbiguousCommandError as e:
            msg = f"{e}\n{e.clarification_prompt}\n" + "\n".join(f"• {opt}" for opt in e.suggested_options)
            self.chat_widget.add_assistant_message(msg)
            return

        self._execute_action_pipeline(action, confirmed=confirmed)

    def _execute_action_pipeline(self, action: ExcelAction, confirmed: bool = False):
        if not self.manager.is_connected():
            self.chat_widget.add_assistant_message("❌ ابتدا باید اتصال به اکسل برقرار باشد.")
            return

        sheet = self.manager.get_active_sheet()
        sa = SheetInspector.inspect_sheet(sheet)
        ctx = ActionContext(manager=self.manager, sheet=sheet, sheet_analysis=sa, confirmed=confirmed)
        res = self.executor.execute(action, ctx)

        if res.requires_confirmation:
            self.pending_destructive_action = action
            self.chat_widget.add_confirmation_request(
                prompt_text=res.message_persian,
                on_confirm=self._on_confirm_destructive_action,
                on_cancel=self._on_cancel_destructive_action
            )
            return

        self.chat_widget.add_assistant_message(res.message_persian)

        # Update status bar sheets and workbooks list
        info = self.manager.get_workbook_info()
        open_wbs = self.manager.list_open_workbooks()
        self.status_bar_widget.update_workbook_info(info.name, info.sheets, info.active_sheet, open_wbs)

    def _on_confirm_destructive_action(self):
        if self.pending_destructive_action:
            action = self.pending_destructive_action
            self.pending_destructive_action = None
            sheet = self.manager.get_active_sheet()
            sa = SheetInspector.inspect_sheet(sheet)
            ctx = ActionContext(manager=self.manager, sheet=sheet, sheet_analysis=sa, confirmed=True)
            res = self.executor.execute(action, ctx)
            self.chat_widget.add_assistant_message(res.message_persian)

    def _on_cancel_destructive_action(self):
        self.pending_destructive_action = None
        self.chat_widget.add_assistant_message("عملیات به درخواست شما لغو شد.")

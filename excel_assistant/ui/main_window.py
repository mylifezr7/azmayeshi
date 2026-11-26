"""
Main Window for Excel AI Assistant (PySide6 / RTL Persian).
"""
import sys
from typing import Optional

from PySide6.QtCore import Qt, QThread, Signal, Slot, QEvent, QObject, QStringListModel
from PySide6.QtGui import QShortcut, QKeySequence
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLineEdit,
    QPushButton, QLabel, QFrame, QMessageBox, QApplication, QDialog,
    QScrollArea, QComboBox, QSizePolicy, QCompleter
)

from excel_assistant.ui.command_search import search_commands, COMMAND_SEARCH_REGISTRY

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


class HorizontalWheelFilter(QObject):
    """Enables smooth horizontal scrolling with mouse wheel on horizontal scroll areas."""

    def __init__(self, scroll_area: QScrollArea):
        super().__init__(scroll_area)
        self.scroll_area = scroll_area

    def eventFilter(self, obj, event):
        if event.type() == QEvent.Wheel:
            delta = event.angleDelta().y() or event.angleDelta().x()
            sb = self.scroll_area.horizontalScrollBar()
            sb.setValue(sb.value() - delta)
            return True
        return super().eventFilter(obj, event)


class DropUpComboBox(QComboBox):
    """QComboBox that ensures its popup list opens upwards (Drop-Up) above the control."""

    def showPopup(self):
        super().showPopup()
        popup = self.findChild(QFrame)
        if popup:
            popup_h = popup.height()
            if popup_h < 50:
                popup_h = popup.sizeHint().height()
            if popup_h < 50:
                popup_h = min(420, max(80, self.count() * 40 + 16))
            top_left = self.mapToGlobal(self.rect().topLeft())
            new_y = top_left.y() - popup_h - 4
            popup.move(top_left.x(), max(10, new_y))


class MainWindow(QMainWindow):
    """Main Application Window."""

    def __init__(self, manager: Optional[ExcelConnectionManager] = None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle("MUYUM | دستیار هوشمند مالی و حسابداری مایکروسافت اکسل")
        self.resize(1050, 720)
        self.setLayoutDirection(Qt.RightToLeft)
        self.setStyleSheet(QSS_STYLE)
        self.setAcceptDrops(True)

        app = QApplication.instance()
        if app:
            from excel_assistant.ui.app import setup_application_font
            setup_application_font(app)

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

        # 3. Cascading 5-Category Dropdown + Action Selector + Run Button (Bottom Action Bar)
        chips_frame = QFrame()
        chips_frame.setObjectName("quickActionBar")
        chips_frame.setStyleSheet("background: transparent; border: none;")
        chips_frame_layout = QHBoxLayout(chips_frame)
        chips_frame_layout.setContentsMargins(12, 4, 12, 6)
        chips_frame_layout.setSpacing(12)

        # 3.1. Structure of the 5 requested categories
        self.five_categories = {
            "essential": {
                "title": "⚡ ۱. قابلیت‌های ضروری",
                "commands": [
                    ("🔍 فیلتر هوشمند اطلاعات", "فیلتر هوشمند جدول"),
                    ("🔄 مرتب‌سازی صعودی داده‌ها", "مرتب‌سازی صعودی ستون A"),
                    ("🔽 مرتب‌سازی نزولی داده‌ها", "مرتب‌سازی نزولی ستون A"),
                    ("🗑️ حذف موارد تکراری", "حذف داده‌های تکراری"),
                    ("🧹 حذف سطرهای خالی", "ردیف‌های خالی را حذف کن"),
                    ("➕ جمع خودکار (AutoSum)", "جمع خودکار ستون A"),
                    ("📊 محاسبه میانگین داده‌ها", "میانگین ستون A"),
                    ("🔺 بیشترین و کمترین (MAX / MIN)", "بیشترین مقدار ستون A"),
                    ("🔢 شمارش اطلاعات و سلول‌های پر", "شمارش سلول‌های پر ستون A"),
                    ("🔢 تبدیل متن به عدد (اصلاح داده‌ها)", "تبدیل متن به عدد در ستون A"),
                    ("🧮 ۴ عمل اصلی ستون‌ها...", "۴ عمل اصلی"),
                    ("📅 درج تاریخ در شیت...", "درج تاریخ در شیت"),
                ]
            },
            "accounting": {
                "title": "💼 ۲. قابلیت‌های حسابداری",
                "commands": [
                    ("🏷️ محاسبه بهای تمام شده (فایفو / لایفو / میانگین)...", "روش بهای تمام شده"),
                    ("💵 محاسبه و مقایسه سود (فایفو / لایفو / میانگین موزون)...", "روش سود"),
                    ("🏷️ بهای تمام شده به روش فایفو (FIFO)", "محاسبه بهای تمام شده به روش فایفو"),
                    ("🏷️ بهای تمام شده به روش لایفو (LIFO)", "محاسبه بهای تمام شده به روش لایفو"),
                    ("⚖️ مقایسه بهای تمام شده فایفو و لایفو", "مقایسه بهای تمام شده فایفو و لایفو"),
                    ("⏱️ محاسبه ساعت کاری، ورود و خروج کارمندان (تایم‌شیت)", "محاسبه ساعت کاری"),
                    ("💰 گزارش جامع فروش (تشخیص خودکار ستون‌ها)", "محاسبه فروش"),
                    ("💼 لیست حقوق و دستمزد، بیمه و مالیات (تشخیص خودکار)", "محاسبه حقوق و دستمزد"),
                    ("⚖️ مغایرت‌گیری دو ستون (Reconciliation)", "مغایرت‌گیری ستون A و B"),
                    ("📊 تهیه تراز آزمایشی چهار ستونه", "تهیه تراز آزمایشی"),
                    ("💳 محاسبه مانده حساب جاری (بدهکار/بستانکار)", "محاسبه مانده حساب"),
                    ("📉 محاسبه جدول استهلاک دارایی‌ها", "محاسبه استهلاک"),
                    ("📈 تهیه صورت سود و زیان جامع دوره", "صورت سود و زیان جامع"),
                    ("🏛️ تهیه ترازنامه مالی (بیلان)", "ترازنامه"),
                    ("⚙️ تنظیمات دستی ستون‌های محاسبه سود...", "تنظیمات سود"),
                    ("⚙️ تنظیمات دستی ستون‌های گزارش فروش...", "تنظیمات فروش"),
                    ("🏷️ محاسبه تخفیف کالاها...", "محاسبه تخفیف"),
                ]
            },
            "sales_inventory": {
                "title": "🛒 ۳. قابلیت‌های فروش و انبارداری",
                "commands": [
                    ("🟢 ارزش‌گذاری موجودی انبار به روش فایفو (FIFO)", "ارزش‌گذاری انبار به روش فایفو"),
                    ("🔵 ارزش‌گذاری موجودی انبار به روش لایفو (LIFO)", "ارزش‌گذاری انبار به روش لایفو"),
                    ("🟣 ارزش‌گذاری انبار به روش میانگین موزون (Average)", "ارزش‌گذاری انبار به روش میانگین موزون"),
                    ("📑 مقایسه همزمان هر ۳ روش ارزش‌گذاری در انبار", "مقایسه روش‌های فایفو و لایفو و میانگین در انبار"),
                    ("🏷️ محاسبه تخفیف کالاها...", "محاسبه تخفیف"),
                    ("🧾 صدور و ثبت فاکتور رسمی فروش", "ثبت فاکتور فروش"),
                    ("🛒 صدور و ثبت فاکتور رسمی خرید", "ثبت فاکتور خرید"),
                    ("📦 کنترل موجودی و هشدار کمبود انبار", "موجودی انبار و هشدار کمبود موجودی"),
                    ("💰 محاسبه ارزش ریالی کل موجودی انبار", "ارزش کل موجودی انبار"),
                    ("👥 محاسبه بدهی و طلب مشتریان", "محاسبه بدهی و طلب مشتریان"),
                    ("🏆 گزارش پرفروش‌ترین کالاها", "گزارش پرفروش‌ترین کالاها"),
                    ("🏷️ افزایش ۱۵ درصدی قیمت‌های ستون C", "قیمت‌های ستون C را ۱۵ درصد افزایش بده"),
                ]
            },
            "smart": {
                "title": "🤖 ۴. قابلیت‌های هوشمند و کاربردی",
                "commands": [
                    ("🔒 مدیریت و قفل یکتاسازی ستون...", "قفل یکتایی ستون"),
                    ("🔓 لغو و برداشتن قفل یکتایی ستون", "قفل یکتایی ستون را بردار"),
                    ("📊 تحلیل هوشمند شیت و داده‌ها", "تحلیل ساختار فایل و شیت"),
                    ("🩺 چکاپ عمیق سلامت و خطاهای فایل", "سلامت فایل و بررسی خطاهای اکسل"),
                    ("🧮 میانگین سطری بین ستون‌ها (B و C در D)", "از ستون B و C میانگین بگیر و نتیجه را در D بنویس"),
                    ("📊 گزارش‌ساز خودکار با نمودار گرافیکی", "رسم نمودار ستونی"),
                    ("📋 تشخیص و ساخت جدول اکسل", "یک جدول بساز"),
                    ("✂️ تفکیک نام و نام خانوادگی", "تفکیک نام و نام خانوادگی در ستون A"),
                    ("🔗 ادغام سلول‌ها (Merge)", "سلول‌های A1:B1 را ادغام کن"),
                ]
            },
            "management": {
                "title": "⚙️ ۵. امکانات مدیریتی خود برنامه",
                "commands": [
                    ("🔒 قفل و مدیریت یکتایی ستون...", "قفل یکتایی ستون"),
                    ("🔓 لغو و برداشتن قفل یکتایی ستون", "قفل یکتایی ستون را بردار"),
                    ("📅 درج تاریخ در شیت (شمسی/میلادی)...", "درج تاریخ در شیت"),
                    ("📊 تحلیل هوشمند شیت و داده‌ها", "تحلیل ساختار فایل و شیت"),
                    ("💾 ذخیره تغییرات فایل اکسل (Ctrl+S)", "ذخیره فایل"),
                    ("💾 ذخیره نسخه پشتیبان (Backup)", "ذخیره نسخه پشتیبان"),
                    ("📄 تولید خروجی PDF استاندارد", "خروجی PDF از شیت"),
                    ("↩️ لغو آخرین عملیات (Undo چندمرحله‌ای)", "لغو آخرین عملیات"),
                    ("📐 تنظیم خودکار عرض ستون‌ها (AutoFit)", "عرض ستون‌ها را خودکار تنظیم کن"),
                    ("🎨 رنگ‌آمیزی متون تکراری", "داده‌های متنی تکراری را رنگ کن"),
                    ("💲 تغییر فرمت به واحد پول (تومان)", "فرمت ستون B را به تومان تغییر بده"),
                    ("📈 تغییر فرمت به درصد (%)", "فرمت ستون C را درصدی کن"),
                    ("📅 تغییر فرمت به تاریخ", "فرمت ستون A را به تاریخ تغییر بده"),
                ]
            }
        }

        # Dropdown 1: انتخاب دسته‌بندی اصلی (۵ بخش) - بازشونده به بالا
        self.category_combo = DropUpComboBox()
        self.category_combo.setObjectName("quickCategoryCombo")
        self.category_combo.setToolTip("۱. ابتدا حوزه یا دسته‌بندی کار خود را انتخاب کنید")
        self.category_combo.addItem("📁 ۱. انتخاب حوزه (۵ دسته‌بندی اصلی)...", None)
        for cat_key, cat_data in self.five_categories.items():
            self.category_combo.addItem(cat_data["title"], cat_key)
        self.category_combo.currentIndexChanged.connect(self._on_category_selected)
        chips_frame_layout.addWidget(self.category_combo, stretch=2)

        # Dropdown 2: انتخاب عملیات زیرمجموعه - بازشونده به بالا
        self.action_combo = DropUpComboBox()
        self.action_combo.setObjectName("quickActionCombo")
        self.action_combo.setToolTip("۲. عملیات مورد نظر در این حوزه را انتخاب کنید")
        self.action_combo.addItem("🎯 ۲. ابتدا حوزه را انتخاب کنید...", None)
        self.action_combo.currentIndexChanged.connect(self._on_action_selected)
        chips_frame_layout.addWidget(self.action_combo, stretch=3)

        # دکمه اجرای سریع عملیات انتخابی
        self.btn_run_action = QPushButton("اجرا 🚀")
        self.btn_run_action.setObjectName("runQuickActionButton")
        self.btn_run_action.setToolTip("اجرای عملیات انتخاب‌شده")
        self.btn_run_action.clicked.connect(self._on_run_action_clicked)
        chips_frame_layout.addWidget(self.btn_run_action, stretch=0)

        # 3. Search Bar for instant command discovery (جستجوی سریع قابلیت‌ها)
        search_frame = QFrame()
        search_frame.setObjectName("searchFrame")
        search_layout = QHBoxLayout(search_frame)
        search_layout.setContentsMargins(12, 4, 12, 4)
        search_layout.setSpacing(8)

        self.search_input = QLineEdit()
        self.search_input.setObjectName("commandSearchInput")
        self.search_input.setPlaceholderText("🔍 جستجوی سریع قابلیت‌ها (مثلاً: سود، سود سالانه، کارکرد، فروش، فیلتر)...")
        self.search_input.setLayoutDirection(Qt.RightToLeft)
        search_layout.addWidget(self.search_input, stretch=1)

        self.btn_clear_search = QPushButton("✕")
        self.btn_clear_search.setObjectName("clearSearchButton")
        self.btn_clear_search.setToolTip("پاک کردن جستجو")
        self.btn_clear_search.setVisible(False)
        self.btn_clear_search.clicked.connect(self._clear_search)
        search_layout.addWidget(self.btn_clear_search)

        # Setup Auto-Completer for Search Input
        self._search_cmd_map = {}
        self.search_model = QStringListModel(self)
        self.search_completer = QCompleter(self.search_model, self)
        self.search_completer.setCaseSensitivity(Qt.CaseInsensitive)
        self.search_completer.setFilterMode(Qt.MatchContains)
        self.search_completer.setCompletionMode(QCompleter.PopupCompletion)
        if self.search_completer.popup():
            self.search_completer.popup().setObjectName("searchCompleterPopup")
            self.search_completer.popup().setLayoutDirection(Qt.RightToLeft)
        self.search_input.setCompleter(self.search_completer)

        self.search_input.textChanged.connect(self._on_search_text_changed)
        self.search_completer.activated[str].connect(self._on_search_item_activated)
        self.search_input.returnPressed.connect(self._on_search_return_pressed)

        main_layout.addWidget(search_frame)

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

        # Setup Completer for Chat Input as well
        self.chat_completer = QCompleter([item["cmd"] for item in COMMAND_SEARCH_REGISTRY], self)
        self.chat_completer.setCaseSensitivity(Qt.CaseInsensitive)
        self.chat_completer.setFilterMode(Qt.MatchContains)
        self.chat_completer.setCompletionMode(QCompleter.PopupCompletion)
        if self.chat_completer.popup():
            self.chat_completer.popup().setObjectName("searchCompleterPopup")
            self.chat_completer.popup().setLayoutDirection(Qt.RightToLeft)
        self.input_edit.setCompleter(self.chat_completer)

        self.btn_send = QPushButton("ارسال دستور ✈️")
        self.btn_send.setObjectName("sendButton")
        self.btn_send.clicked.connect(self._on_send_clicked)

        self.btn_stop = QPushButton("توقف ⏹")
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
        self.history_panel.save_requested.connect(self._on_save_clicked)
        self.history_panel.undo_requested.connect(self._on_undo_clicked)
        self.history_panel.analyze_requested.connect(self._on_analyze_clicked)
        self.history_panel.arithmetic_requested.connect(lambda: self._open_arithmetic_dialog("add"))
        self.history_panel.discount_requested.connect(self._open_discount_dialog)
        self.history_panel.date_requested.connect(self._open_date_dialog)
        self.history_panel.unique_col_requested.connect(self._open_unique_col_dialog)
        self.history_panel.sales_requested.connect(lambda: self._on_message_submitted("محاسبه فروش"))
        self.history_panel.profit_requested.connect(self._open_profit_method_dialog)

        # Keyboard shortcuts (میانبرهای کیبورد)
        self.shortcut_save = QShortcut(QKeySequence("Ctrl+S"), self)
        self.shortcut_save.activated.connect(self._on_save_clicked)

        # Ctrl+P: انتخاب سریع روش سود (فایفو / لایفو / میانگین) و اجرای خودکار
        self.shortcut_profit = QShortcut(QKeySequence("Ctrl+P"), self)
        self.shortcut_profit.activated.connect(self._open_profit_method_dialog)
        # Ctrl+Shift+P: باز کردن ویزارد تنظیمات دستی ستون‌ها
        self.shortcut_profit_shift = QShortcut(QKeySequence("Ctrl+Shift+P"), self)
        self.shortcut_profit_shift.activated.connect(self._open_profit_dialog)

        # Ctrl+Shift+S: تشخیص و محاسبه خودکار فروش
        self.shortcut_sales = QShortcut(QKeySequence("Ctrl+Shift+S"), self)
        self.shortcut_sales.activated.connect(lambda: self._on_message_submitted("محاسبه فروش"))

    def _perform_initial_connection(self):
        try:
            from excel_assistant.core.xlwings_driver import check_xlwings, XLWINGS_AVAILABLE
            check_xlwings()

            target_path = self.manager.file_path if (self.manager and self.manager.file_path) else None
            if not self.manager.is_connected() or (sys.platform == 'win32' and self.manager.driver_type != 'xlwings'):
                info = self.manager.connect(target_path)
            else:
                info = self.manager.get_workbook_info()

            desc = "Microsoft Excel 2024 (COM)" if self.manager.driver_type == "xlwings" else "موتور مستقل اکسل"
            open_wbs = self.manager.list_open_workbooks()
            self.status_bar_widget.set_connected(True, desc)
            self.status_bar_widget.update_workbook_info(info.name, info.sheets, info.active_sheet, open_wbs)

            if self.manager.driver_type == "xlwings":
                self.chat_widget.add_assistant_message(
                    f"سلام **MUYUM** عزیز! من دستیار هوشمند Microsoft Excel شما هستم.\n"
                    f"🟢 **اتصال زنده به نرم‌افزار Microsoft Excel برقرار شد.**\n"
                    f"فایل فعال: **`{info.name}`** | شیت: **`{info.active_sheet}`**\n"
                    f"💡 هر دستوری بدهید، تغییرات مستقیماً و به صورت زنده در نرم‌افزار اکسل شما اعمال خواهد شد."
                )
            else:
                warn_text = (
                    f"سلام **MUYUM** عزیز! من دستیار هوشمند Microsoft Excel شما هستم.\n"
                    f"📌 برنامه در حال حاضر با **موتور مستقل اکسل (آفلاین)** به فایل **`{info.name}`** متصل است.\n"
                )
                if sys.platform == "win32":
                    warn_text += (
                        "\n⚠️ **نرم‌افزار مایکروسافت اکسل باز نشد!**\n"
                        "برای اینکه نرم‌افزار رسمی اکسل هم‌زمان با برنامه باز شود و تغییرات را زنده ببینید، پکیج `xlwings` لازم است.\n"
                        "💡 **راه‌حل:** در پنجره CMD دستور زیر را بزنید:\n"
                        "```cmd\n"
                        "pip install xlwings\n"
                        "```\n"
                        "سپس دکمهٔ **«📶 اتصال مجدد»** در بالای پنجره را بزنید تا اکسل فوراً روی صفحه باز شود."
                    )
                self.chat_widget.add_assistant_message(warn_text)

        except Exception as e:
            self.status_bar_widget.set_connected(False)
            self.chat_widget.add_assistant_message(
                f"⚠️ خطا در اتصال به اکسل: {e}\n"
                f"لطفاً مطمئن شوید Microsoft Excel نصب و باز است."
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

    def _on_search_text_changed(self, text: str):
        q = text.strip()
        if not q:
            self.search_model.setStringList([])
            self.btn_clear_search.setVisible(False)
            return
        self.btn_clear_search.setVisible(True)
        results = search_commands(q, limit=15)
        self._search_cmd_map = {r["title"]: r["cmd"] for r in results}
        titles = [r["title"] for r in results]
        self.search_model.setStringList(titles)
        self.search_completer.complete()

    def _clear_search(self):
        self.search_input.clear()
        self.btn_clear_search.setVisible(False)
        self.search_input.setFocus()

    def _on_search_item_activated(self, text: str):
        cmd = self._search_cmd_map.get(text, text)
        if cmd:
            self.search_input.blockSignals(True)
            self.search_input.clear()
            self.btn_clear_search.setVisible(False)
            self.search_input.blockSignals(False)

            self.input_edit.setText(cmd)
            self._on_quick_chip_clicked(cmd)

    def _on_search_return_pressed(self):
        text = self.search_input.text().strip()
        if not text:
            return
        results = search_commands(text, limit=1)
        if results:
            self._on_search_item_activated(results[0]["title"])
        else:
            self.input_edit.setText(text)
            self._on_send_clicked()
            self._clear_search()

    def _on_category_selected(self, index: int):
        cat_key = self.category_combo.itemData(index)
        self.action_combo.blockSignals(True)
        self.action_combo.clear()
        if cat_key and cat_key in self.five_categories:
            self.action_combo.addItem("🎯 ۲. انتخاب عملیات...", None)
            for label, cmd in self.five_categories[cat_key]["commands"]:
                self.action_combo.addItem(label, cmd)
        else:
            self.action_combo.addItem("🎯 ۲. ابتدا حوزه را انتخاب کنید...", None)
        self.action_combo.setCurrentIndex(0)
        self.action_combo.blockSignals(False)

        # باز شدن خودکار منوی عملیات پس از انتخاب حوزه/قابلیت
        if cat_key and cat_key in self.five_categories:
            from PySide6.QtCore import QTimer
            QTimer.singleShot(100, self.action_combo.showPopup)

    def _on_action_selected(self, index: int):
        if index <= 0:
            return
        cmd = self.action_combo.itemData(index) or self.action_combo.itemText(index)
        if cmd:
            self.input_edit.setText(cmd)
            self._on_quick_chip_clicked(cmd)
            self.action_combo.blockSignals(True)
            self.action_combo.setCurrentIndex(0)
            self.action_combo.blockSignals(False)

    def _on_run_action_clicked(self):
        curr_idx = self.action_combo.currentIndex()
        if curr_idx > 0:
            cmd = self.action_combo.itemData(curr_idx) or self.action_combo.itemText(curr_idx)
            if cmd:
                self.input_edit.setText(cmd)
                self._on_quick_chip_clicked(cmd)
        elif self.input_edit.text().strip():
            self._on_send_clicked()
        else:
            self.chat_widget.add_assistant_message("⚠️ لطفاً ابتدا یک عملیات را از منوی کشویی انتخاب کنید یا دستور خود را بنویسید.")

    def _on_quick_combo_selected(self, index: int):
        if index <= 0:
            return
        cmd = self.quick_combo.itemData(index) or self.quick_combo.itemText(index)
        self.input_edit.setText(cmd)
        self.input_edit.setFocus()
        self._on_send_clicked()
        self.quick_combo.blockSignals(True)
        self.quick_combo.setCurrentIndex(0)
        self.quick_combo.blockSignals(False)

    def _on_quick_chip_clicked(self, command: str):
        if command in ["روش سود", "انتخاب روش سود", "روش محاسبه سود"]:
            self._open_profit_method_dialog()
        elif command in ["روش بهای تمام شده", "بهای تمام شده", "محاسبه بهای تمام شده", "انتخاب روش بهای تمام شده", "cogs"]:
            self._open_cogs_method_dialog()
        elif any(w in command for w in ["تنظیمات دستی ستون‌های محاسبه سود", "تنظیمات سود", "ویزارد سود", "دستی ستون سود"]):
            self._open_profit_dialog()
        elif any(w in command for w in ["تنظیمات دستی ستون‌های گزارش فروش", "تنظیمات فروش", "ویزارد فروش", "دستی ستون فروش"]):
            self._open_sales_dialog()
        elif "جلوگیری" in command or "یکتا" in command:
            if any(w in command for w in ["بردار", "حذف", "برداشتن", "لغو", "باز"]):
                self._open_unique_col_dialog(initial_mode="unlock")
            else:
                self._open_unique_col_dialog(initial_mode="lock")
        elif "تاریخ" in command and any(w in command for w in ["درج", "ثبت", "شیت"]):
            self._open_date_dialog()
        elif "تحلیل" in command:
            self._on_analyze_clicked()
        elif "تخفیف" in command and any(w in command for w in ["ویزارد", "پنجره", "دستی"]):
            self._open_discount_dialog()
        elif "۴ عمل" in command or "4 عمل" in command or "عمل اصلی" in command:
            self._open_arithmetic_dialog("add")
        elif command.strip() in ["جمع", "تفریق", "ضرب", "تقسیم"]:
            op_map = {"جمع": "add", "تفریق": "subtract", "ضرب": "multiply", "تقسیم": "divide"}
            self._open_arithmetic_dialog(op_map.get(command.strip(), "add"))
        else:
            self.input_edit.setText(command)
            self._on_send_clicked()

    def _open_unique_col_dialog(self, initial_mode: str = "lock"):
        """Opens interactive wizard dialog for locking or unlocking a column to prevent duplicate values."""
        from excel_assistant.ui.unique_column_dialog import UniqueColumnWizardDialog
        dialog = UniqueColumnWizardDialog(manager=self.manager, initial_mode=initial_mode, parent=self)
        if dialog.exec() == QDialog.Accepted:
            action = dialog.get_action()
            self.chat_widget.add_user_message(action.description_persian)
            self.history_panel.add_command(action.description_persian)
            self._execute_action_pipeline(action)

    def _open_sales_dialog(self):
        """Opens interactive wizard dialog for calculating daily, monthly, yearly, and total sales."""
        from excel_assistant.ui.sales_dialog import SalesSummaryWizardDialog
        dialog = SalesSummaryWizardDialog(manager=self.manager, parent=self)
        if dialog.exec() == QDialog.Accepted:
            action = dialog.get_action()
            self.chat_widget.add_user_message(action.description_persian)
            self.history_panel.add_command(action.description_persian)
            self._execute_action_pipeline(action)

    def _open_profit_method_dialog(self):
        """Opens quick 1-click method selector for FIFO, LIFO, Weighted Average, or Side-by-Side Comparison."""
        from excel_assistant.ui.profit_dialog import ProfitMethodSelectorDialog
        dialog = ProfitMethodSelectorDialog(parent=self)
        if dialog.exec() == QDialog.Accepted:
            if dialog.open_advanced:
                self._open_profit_dialog()
            elif dialog.selected_method == "compare":
                self._on_message_submitted("مقایسه سود به روش فایفو و لایفو")
            else:
                method_names = {
                    "fifo": "فایفو (FIFO)",
                    "lifo": "لایفو (LIFO)",
                    "average": "میانگین موزون (Average)"
                }
                m_str = method_names.get(dialog.selected_method, "فایفو (FIFO)")
                self._on_message_submitted(f"محاسبه سود به روش {m_str}")

    def _open_cogs_method_dialog(self):
        """Opens quick 1-click method selector for COGS: FIFO, LIFO, Weighted Average, Comparison, or Periodic."""
        from excel_assistant.ui.cogs_dialog import CogsMethodSelectorDialog
        dialog = CogsMethodSelectorDialog(parent=self)
        if dialog.exec() == QDialog.Accepted:
            if dialog.selected_method == "fifo":
                self._on_message_submitted("محاسبه بهای تمام شده به روش فایفو")
            elif dialog.selected_method == "lifo":
                self._on_message_submitted("محاسبه بهای تمام شده به روش لایفو")
            elif dialog.selected_method == "average":
                self._on_message_submitted("محاسبه بهای تمام شده با میانگین موزون")
            elif dialog.selected_method == "compare":
                self._on_message_submitted("مقایسه بهای تمام شده فایفو و لایفو")
            else:
                self._on_message_submitted("محاسبه بهای تمام شده کالای فروش رفته")

    def _open_profit_dialog(self):
        """Opens interactive wizard dialog for calculating profit by subtracting purchase from sales."""
        from excel_assistant.ui.profit_dialog import ProfitWizardDialog
        dialog = ProfitWizardDialog(manager=self.manager, parent=self)
        if dialog.exec() == QDialog.Accepted:
            action = dialog.get_action()
            self.chat_widget.add_user_message(action.description_persian)
            self.history_panel.add_command(action.description_persian)
            self._execute_action_pipeline(action)

    def _open_date_dialog(self):
        """Opens interactive wizard dialog for inserting dates into Excel ranges."""
        from excel_assistant.ui.date_dialog import DateWizardDialog
        dialog = DateWizardDialog(manager=self.manager, parent=self)
        if dialog.exec() == QDialog.Accepted:
            action = dialog.get_action()
            self.chat_widget.add_user_message(action.description_persian)
            self.history_panel.add_command(action.description_persian)
            self._execute_action_pipeline(action)

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

    def _on_save_clicked(self):
        try:
            if not self.manager.is_connected():
                self.chat_widget.add_assistant_message("⚠️ هیچ فایل اکسلی در حال حاضر متصل نیست.")
                return
            saved_path = self.manager.save()
            self.chat_widget.add_assistant_message(f"💾 تغییرات فایل اکسل با موفقیت ذخیره شد:\n**`{saved_path}`**")
            self.history_panel.add_command("ذخیره فایل اکسل")
        except Exception as e:
            self.chat_widget.add_assistant_message(f"❌ خطا در ذخیره‌سازی فایل اکسل: {e}")

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
        if any(k in lower for k in ["یکتا", "یکتاسازی", "کد ملی", "کد محصول", "جلوگیری از تکرار"]) and len(lower.split()) <= 4:
            if any(w in lower for w in ["بردار", "حذف", "برداشتن", "لغو", "باز"]):
                self._open_unique_col_dialog(initial_mode="unlock")
            else:
                self._open_unique_col_dialog(initial_mode="lock")
            return

        if any(k in lower for k in ["تنظیمات سود", "ویزارد سود", "پنجره سود", "دستی سود", "تنظیمات محاسبه سود"]):
            self._open_profit_dialog()
            return

        if any(k in lower for k in ["روش سود", "روش محاسبه سود", "انتخاب روش سود", "انتخاب روش محاسبه سود"]) or lower in ["محاسبه سود", "سود", "محاسبه سود و زیان"]:
            self._open_profit_method_dialog()
            return

        if any(k in lower for k in ["تنظیمات فروش", "ویزارد فروش", "پنجره فروش", "دستی فروش", "تنظیمات گزارش فروش"]):
            self._open_sales_dialog()
            return

        if "تخفیف" in lower and len(lower.split()) <= 4:
            self._open_discount_dialog()
            return

        if "تاریخ" in lower and any(k in lower for k in ["درج", "ثبت", "شیت"]) and len(lower.split()) <= 4:
            self._open_date_dialog()
            return

        if "تحلیل" in lower and any(k in lower for k in ["شیت", "ساختار", "فایل", "داده", "هوشمند"]) and len(lower.split()) <= 5:
            self._on_analyze_clicked()
            return

        if lower in ["جمع", "ضرب", "تفریق", "تقسیم", "4 عمل اصلی", "۴ عمل اصلی", "محاسبه", "محاسبات", "ماشین حساب"]:
            op_map = {"جمع": "add", "تفریق": "subtract", "ضرب": "multiply", "تقسیم": "divide"}
            self._open_arithmetic_dialog(op_map.get(lower, "add"))
            return

        self._on_message_submitted(text)

    def _on_message_submitted(self, text: str):
        text = text.strip()
        if not text:
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


if __name__ == "__main__":
    from excel_assistant.ui.app import run_gui
    sys.exit(run_gui())


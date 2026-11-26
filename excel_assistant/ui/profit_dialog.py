"""
Interactive Profit Summary Wizard Dialog for Persian Excel Assistant.
Allows users to visually select sales sheet & columns and purchase sheet & columns
(e.g., Sales in Sheet 1 and Purchase in Sheet 2),
choose reporting periods (daily, monthly, yearly, total, or comprehensive all-in-one),
and generate styled Excel profit reports with live KPI cards and formulas.
Features a responsive, scrollable layout with pinned action buttons for all screen heights.
"""
from typing import List, Optional
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QRadioButton, QButtonGroup, QLineEdit,
    QComboBox, QFrame, QCheckBox, QMessageBox, QScrollArea
)

from excel_assistant.core.utils import col_index_to_letter, col_letter_to_index
from excel_assistant.nlp.intent_models import ProfitCalculationAction
from excel_assistant.core.connection import ExcelConnectionManager


class ProfitWizardDialog(QDialog):
    """Modern RTL dialog for calculating daily, monthly, yearly, and total profit across sheets."""

    def __init__(self, manager: Optional[ExcelConnectionManager] = None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.manager = manager
        self.setWindowTitle("💵 محاسبه سود (روزانه، ماهانه، سالانه و کل)")
        self.resize(560, 530)
        self.setMinimumSize(480, 380)
        self.setLayoutDirection(Qt.RightToLeft)

        self._init_ui()
        self._load_sheets_and_columns()
        self._update_preview()

    def _init_ui(self):
        root_layout = QVBoxLayout(self)
        root_layout.setContentsMargins(16, 12, 16, 12)
        root_layout.setSpacing(8)

        # 1. Header (pinned at top)
        header_box = QHBoxLayout()
        header_lbl = QLabel("💵 محاسبه سود روزانه، ماهانه، سالانه و کل")
        header_lbl.setStyleSheet("font-size: 15px; font-weight: bold; color: #10B981;")
        header_box.addWidget(header_lbl)
        header_box.addStretch()
        root_layout.addLayout(header_box)

        desc_lbl = QLabel("شیت‌ها، ستون‌های فروش و خرید و دوره زمانی را مشخص کنید:")
        desc_lbl.setStyleSheet("font-size: 11.5px; color: #94A3B8;")
        root_layout.addWidget(desc_lbl)

        # 2. Scrollable Content Area (Ensures all fields are accessible on any screen)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setStyleSheet("QScrollArea { background: transparent; border: none; }")

        content_widget = QWidget()
        content_widget.setStyleSheet("background: transparent;")
        main_layout = QVBoxLayout(content_widget)
        main_layout.setContentsMargins(2, 4, 8, 4)
        main_layout.setSpacing(10)

        # Section 1: Sales Sheet & Columns (اطلاعات فروش - شیت ۱)
        s_box = QFrame()
        s_box.setStyleSheet("background-color: #060E22; border: 1px solid #1E3A8A; border-radius: 9px; padding: 8px;")
        s_layout = QVBoxLayout(s_box)
        s_layout.setSpacing(6)

        s_header = QLabel("🛒 ۱. اطلاعات فروش (شیت ۱ / شیلد فروش):")
        s_header.setStyleSheet("font-weight: bold; color: #38BDF8; font-size: 12.5px;")
        s_layout.addWidget(s_header)

        # Sales Sheet selector
        s_sheet_row = QHBoxLayout()
        s_sheet_lbl = QLabel("انتخاب شیت فروش:")
        s_sheet_lbl.setFixedWidth(120)
        s_sheet_lbl.setStyleSheet("color: #E2E8F0; font-size: 11.5px; font-weight: bold;")
        self.combo_s_sheet = QComboBox()
        self.combo_s_sheet.setStyleSheet("background-color: #081024; color: #FFFFFF; padding: 4px 8px; border: 1px solid #1E3A8A; border-radius: 6px;")
        self.combo_s_sheet.currentTextChanged.connect(self._on_s_sheet_changed)
        s_sheet_row.addWidget(s_sheet_lbl)
        s_sheet_row.addWidget(self.combo_s_sheet)
        s_layout.addLayout(s_sheet_row)

        # Date column
        date_row = QHBoxLayout()
        d_lbl = QLabel("ستون تاریخ فاکتور:")
        d_lbl.setFixedWidth(120)
        d_lbl.setStyleSheet("color: #CBD5E1; font-size: 11.5px;")
        self.combo_date = QComboBox()
        self.combo_date.setStyleSheet("background-color: #081024; color: #FFFFFF; padding: 4px 8px; border: 1px solid #1E3A8A; border-radius: 6px;")
        self.combo_date.currentTextChanged.connect(lambda: self._update_preview())
        self.input_date = QLineEdit()
        self.input_date.setPlaceholderText("مثلاً: A یا 1")
        self.input_date.setFixedWidth(70)
        self.input_date.setStyleSheet("background-color: #081024; color: #FFFFFF; padding: 4px 6px; border: 1px solid #1E3A8A; border-radius: 6px;")
        self.input_date.textChanged.connect(self._on_manual_date_changed)
        date_row.addWidget(d_lbl)
        date_row.addWidget(self.combo_date)
        date_row.addWidget(self.input_date)
        s_layout.addLayout(date_row)

        # Sales amount column
        sales_row = QHBoxLayout()
        s_lbl = QLabel("ستون مبلغ / قیمت فروش:")
        s_lbl.setFixedWidth(120)
        s_lbl.setStyleSheet("color: #CBD5E1; font-size: 11.5px;")
        self.combo_sales = QComboBox()
        self.combo_sales.setStyleSheet("background-color: #081024; color: #FFFFFF; padding: 4px 8px; border: 1px solid #1E3A8A; border-radius: 6px;")
        self.combo_sales.currentTextChanged.connect(lambda: self._update_preview())
        self.input_sales = QLineEdit()
        self.input_sales.setPlaceholderText("مثلاً: B یا 2")
        self.input_sales.setFixedWidth(70)
        self.input_sales.setStyleSheet("background-color: #081024; color: #FFFFFF; padding: 4px 6px; border: 1px solid #1E3A8A; border-radius: 6px;")
        self.input_sales.textChanged.connect(self._on_manual_sales_changed)
        sales_row.addWidget(s_lbl)
        sales_row.addWidget(self.combo_sales)
        sales_row.addWidget(self.input_sales)
        s_layout.addLayout(sales_row)

        main_layout.addWidget(s_box)

        # Section 2: Purchase Sheet & Columns (اطلاعات خرید - شیت ۲ / شیلد دیگر)
        p_box = QFrame()
        p_box.setStyleSheet("background-color: #060E22; border: 1px solid #1E3A8A; border-radius: 9px; padding: 8px;")
        p_layout = QVBoxLayout(p_box)
        p_layout.setSpacing(6)

        p_header = QLabel("📦 ۲. اطلاعات خرید و بهای تمام‌شده (شیت ۲ / شیلد دیگر):")
        p_header.setStyleSheet("font-weight: bold; color: #F59E0B; font-size: 12.5px;")
        p_layout.addWidget(p_header)

        # Purchase Sheet selector
        p_sheet_row = QHBoxLayout()
        p_sheet_lbl = QLabel("انتخاب شیت خرید:")
        p_sheet_lbl.setFixedWidth(120)
        p_sheet_lbl.setStyleSheet("color: #E2E8F0; font-size: 11.5px; font-weight: bold;")
        self.combo_p_sheet = QComboBox()
        self.combo_p_sheet.setStyleSheet("background-color: #081024; color: #FFFFFF; padding: 4px 8px; border: 1px solid #1E3A8A; border-radius: 6px;")
        self.combo_p_sheet.currentTextChanged.connect(self._on_p_sheet_changed)
        p_sheet_row.addWidget(p_sheet_lbl)
        p_sheet_row.addWidget(self.combo_p_sheet)
        p_layout.addLayout(p_sheet_row)

        # Purchase amount column
        purchase_row = QHBoxLayout()
        p_lbl = QLabel("ستون مبلغ / قیمت خرید:")
        p_lbl.setFixedWidth(120)
        p_lbl.setStyleSheet("color: #CBD5E1; font-size: 11.5px;")
        self.combo_purchase = QComboBox()
        self.combo_purchase.setStyleSheet("background-color: #081024; color: #FFFFFF; padding: 4px 8px; border: 1px solid #1E3A8A; border-radius: 6px;")
        self.combo_purchase.currentTextChanged.connect(lambda: self._update_preview())
        self.input_purchase = QLineEdit()
        self.input_purchase.setPlaceholderText("مثلاً: C یا 3")
        self.input_purchase.setFixedWidth(70)
        self.input_purchase.setStyleSheet("background-color: #081024; color: #FFFFFF; padding: 4px 6px; border: 1px solid #1E3A8A; border-radius: 6px;")
        self.input_purchase.textChanged.connect(self._on_manual_purchase_changed)
        purchase_row.addWidget(p_lbl)
        purchase_row.addWidget(self.combo_purchase)
        purchase_row.addWidget(self.input_purchase)
        p_layout.addLayout(purchase_row)

        main_layout.addWidget(p_box)

        # Section: Valuation Method
        method_lbl = QLabel("🏷️ ۳. روش ارزیابی بهای تمام‌شده و انبار:")
        method_lbl.setStyleSheet("font-weight: bold; color: #10B981; font-size: 12.5px;")
        main_layout.addWidget(method_lbl)

        self.method_group = QButtonGroup(self)
        method_layout = QHBoxLayout()
        self.rb_fifo = QRadioButton("فایفو (FIFO)")
        self.rb_fifo.setChecked(True)
        self.rb_lifo = QRadioButton("لایفو (LIFO)")
        self.rb_avg = QRadioButton("میانگین موزون (Average)")

        for rb in [self.rb_fifo, self.rb_lifo, self.rb_avg]:
            rb.setStyleSheet("color: #CBD5E1; font-size: 11px;")
            self.method_group.addButton(rb)
            method_layout.addWidget(rb)
            rb.toggled.connect(lambda: self._update_preview())
        main_layout.addLayout(method_layout)

        # Section 4: Period Selection
        period_lbl = QLabel("⏱️ ۴. دوره زمانی تفکیک سود:")
        period_lbl.setStyleSheet("font-weight: bold; color: #E2E8F0; font-size: 12.5px;")
        main_layout.addWidget(period_lbl)

        self.radio_group = QButtonGroup(self)
        period_layout = QVBoxLayout()
        period_layout.setSpacing(3)

        self.rb_all = QRadioButton("📊 گزارش جامع (سود کل + تفکیک روزانه + ماهانه + سالانه)")
        self.rb_all.setChecked(True)
        self.rb_total = QRadioButton("💵 فقط سود کل (مجموع سود، میانگین حاشیه سود و پر‌سودترین بازه)")
        self.rb_daily = QRadioButton("📅 تفکیک سود روزانه (مجموع فروش، خرید و سود هر روز)")
        self.rb_monthly = QRadioButton("🗓️ تفکیک سود ماهانه (روند سودآوری ماه‌های سال)")
        self.rb_yearly = QRadioButton("📆 تفکیک سود سالانه (سهم و مجموع سود سال‌ها)")

        for rb in [self.rb_all, self.rb_total, self.rb_daily, self.rb_monthly, self.rb_yearly]:
            rb.setStyleSheet("color: #CBD5E1; font-size: 11px;")
            self.radio_group.addButton(rb)
            period_layout.addWidget(rb)
            rb.toggled.connect(lambda: self._update_preview())

        main_layout.addLayout(period_layout)

        # Output option: Create dedicated sheet
        self.chk_create_sheet = QCheckBox("📑 ایجاد شیت اختصاصی «گزارش_سود» همراه با داشبورد، جداول تفکیک‌شده و فرمول‌ها")
        self.chk_create_sheet.setChecked(True)
        self.chk_create_sheet.setStyleSheet("color: #10B981; font-weight: bold; font-size: 11.5px;")
        main_layout.addWidget(self.chk_create_sheet)

        # Live Preview Card
        self.preview_card = QFrame()
        self.preview_card.setStyleSheet("background-color: #04091A; border: 1px solid #1E3A8A; border-radius: 9px; padding: 8px;")
        card_layout = QVBoxLayout(self.preview_card)
        card_layout.setContentsMargins(8, 6, 8, 6)

        preview_title = QLabel("👁️ پیش‌نمایش نحوهٔ محاسبه سود:")
        preview_title.setStyleSheet("font-weight: bold; color: #94A3B8; font-size: 11px;")
        card_layout.addWidget(preview_title)

        self.preview_text = QLabel("در حال پردازش ستون‌ها...")
        self.preview_text.setStyleSheet("color: #F8FAFC; font-size: 11px; line-height: 1.4;")
        self.preview_text.setWordWrap(True)
        card_layout.addWidget(self.preview_text)

        main_layout.addWidget(self.preview_card)

        scroll.setWidget(content_widget)
        root_layout.addWidget(scroll, 1)

        # 3. Pinned Action Buttons Footer (ALWAYS VISIBLE AT BOTTOM)
        btn_box = QHBoxLayout()
        btn_box.setContentsMargins(0, 4, 0, 0)
        btn_box.setSpacing(12)

        self.btn_cancel = QPushButton("انصراف")
        self.btn_cancel.setStyleSheet(
            "background-color: #1E293B; color: #E2E8F0; padding: 9px 18px; border-radius: 8px; font-weight: bold; font-size: 12px;"
        )
        self.btn_cancel.clicked.connect(self.reject)

        self.btn_auto = QPushButton("⚡ محاسبه کاملاً خودکار (تشخیص هوشمند)")
        self.btn_auto.setToolTip("اجرای خودکار بدون نیاز به انتخاب دستی ستون‌ها (دقیقاً مشابه محاسبه حقوق و دستمزد)")
        self.btn_auto.setStyleSheet(
            "background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #1D4ED8, stop:1 #38BDF8); "
            "color: white; padding: 9px 18px; border-radius: 8px; font-weight: bold; font-size: 12.5px; border: 1px solid #38BDF8;"
        )
        self.btn_auto.clicked.connect(self._on_auto_submit)

        self.btn_submit = QPushButton("محاسبه با ستون‌های انتخابی 💵")
        self.btn_submit.setDefault(True)
        self.btn_submit.setAutoDefault(True)
        self.btn_submit.setStyleSheet(
            "background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #059669, stop:1 #10B981); "
            "color: white; padding: 9px 22px; border-radius: 8px; font-weight: bold; font-size: 12.5px; border: 1px solid #10B981;"
        )
        self.btn_submit.clicked.connect(self._on_submit)

        btn_box.addWidget(self.btn_cancel)
        btn_box.addWidget(self.btn_auto)
        btn_box.addWidget(self.btn_submit)
        root_layout.addLayout(btn_box)

        self._is_auto = False

    def _load_sheets_and_columns(self):
        if not self.manager or not self.manager.is_connected():
            return

        driver = self.manager.active_driver
        sheets = driver.get_sheet_names()

        self.combo_s_sheet.blockSignals(True)
        self.combo_p_sheet.blockSignals(True)

        self.combo_s_sheet.clear()
        self.combo_p_sheet.clear()

        for s in sheets:
            self.combo_s_sheet.addItem(s)
            self.combo_p_sheet.addItem(s)

        if len(sheets) >= 2:
            self.combo_s_sheet.setCurrentIndex(0)
            self.combo_p_sheet.setCurrentIndex(1)
        elif sheets:
            self.combo_s_sheet.setCurrentIndex(0)
            self.combo_p_sheet.setCurrentIndex(0)

        self.combo_s_sheet.blockSignals(False)
        self.combo_p_sheet.blockSignals(False)

        self._populate_sales_columns(self.combo_s_sheet.currentText())
        self._populate_purchase_columns(self.combo_p_sheet.currentText())

    def _populate_sales_columns(self, sheet_name: str):
        if not sheet_name or not self.manager.is_connected():
            return
        try:
            sheet = self.manager.get_sheet(sheet_name)
            min_r, min_c, max_r, max_c = sheet.get_used_range_bounds()

            best_date_idx = 0
            best_sales_idx = 0

            cust_kws = ["مشتری", "خریدار", "طرف حساب", "customer", "client", "فروشنده"]
            inv_kws = ["شماره فاکتور", "شماره سند", "فاکتور", "invoice", "doc"]
            disc_kws = ["تخفیف", "discount", "off"]
            date_kws = ["تاریخ", "date", "زمان", "time"]
            qty_kws = ["تعداد", "مقدار", "qty", "quantity"]
            sales_kws = ["مبلغ کل", "مبلغ فروش", "قیمت فروش", "قیمت واحد", "بهای واحد", "فی فروش", "فروش", "درآمد", "قیمت", "مبلغ", "فی", "price", "amount"]

            self.combo_date.blockSignals(True)
            self.combo_sales.blockSignals(True)
            self.combo_date.clear()
            self.combo_sales.clear()

            found_date = False
            found_sales = False

            for idx, c in enumerate(range(min_c, max_c + 1)):
                c_let = col_index_to_letter(c)
                hdr = str(sheet.get_value(f"{c_let}1") or "").strip()
                label = f"ستون {c_let} (شماره {c})" + (f" - {hdr}" if hdr else "")
                self.combo_date.addItem(label, c_let)
                self.combo_sales.addItem(label, c_let)

                hdr_l = hdr.lower()
                if not found_date and any(k in hdr_l for k in date_kws) and not any(k in hdr_l for k in qty_kws + disc_kws + inv_kws):
                    best_date_idx = idx
                    found_date = True
                if not found_sales and any(k in hdr_l for k in sales_kws) and not any(k in hdr_l for k in qty_kws + disc_kws + cust_kws + inv_kws):
                    best_sales_idx = idx
                    found_sales = True

            if self.combo_date.count() > 0:
                self.combo_date.setCurrentIndex(best_date_idx)
                self.combo_sales.setCurrentIndex(best_sales_idx)

            self.combo_date.blockSignals(False)
            self.combo_sales.blockSignals(False)
        except Exception:
            pass

    def _populate_purchase_columns(self, sheet_name: str):
        if not sheet_name or not self.manager.is_connected():
            return
        try:
            sheet = self.manager.get_sheet(sheet_name)
            min_r, min_c, max_r, max_c = sheet.get_used_range_bounds()

            best_purchase_idx = 0
            purchase_kws = ["بهای واحد خرید", "قیمت واحد خرید", "بهای خرید", "قیمت خرید", "مبلغ خرید", "فی خرید", "بهای تمام شده", "قیمت تمام شده", "خرید", "بها", "قیمت", "مبلغ", "فی", "cost", "purchase"]
            qty_kws = ["تعداد", "مقدار", "qty", "quantity"]

            self.combo_purchase.blockSignals(True)
            self.combo_purchase.clear()

            found_purch = False
            for idx, c in enumerate(range(min_c, max_c + 1)):
                c_let = col_index_to_letter(c)
                hdr = str(sheet.get_value(f"{c_let}1") or "").strip()
                label = f"ستون {c_let} (شماره {c})" + (f" - {hdr}" if hdr else "")
                self.combo_purchase.addItem(label, c_let)

                hdr_l = hdr.lower()
                if not found_purch and any(k in hdr_l for k in purchase_kws) and not any(k in hdr_l for k in qty_kws):
                    best_purchase_idx = idx
                    found_purch = True

            if self.combo_purchase.count() > 0:
                self.combo_purchase.setCurrentIndex(best_purchase_idx)

            self.combo_purchase.blockSignals(False)
        except Exception:
            pass

    def _on_s_sheet_changed(self, text: str):
        self._populate_sales_columns(text)
        self._update_preview()

    def _on_p_sheet_changed(self, text: str):
        self._populate_purchase_columns(text)
        self._update_preview()

    def _on_manual_date_changed(self, text: str):
        val = text.strip().upper()
        if not val: return
        if val.isdigit(): val = col_index_to_letter(int(val))
        for i in range(self.combo_date.count()):
            if self.combo_date.itemData(i) == val:
                self.combo_date.setCurrentIndex(i)
                break
        self._update_preview()

    def _on_manual_sales_changed(self, text: str):
        val = text.strip().upper()
        if not val: return
        if val.isdigit(): val = col_index_to_letter(int(val))
        for i in range(self.combo_sales.count()):
            if self.combo_sales.itemData(i) == val:
                self.combo_sales.setCurrentIndex(i)
                break
        self._update_preview()

    def _on_manual_purchase_changed(self, text: str):
        val = text.strip().upper()
        if not val: return
        if val.isdigit(): val = col_index_to_letter(int(val))
        for i in range(self.combo_purchase.count()):
            if self.combo_purchase.itemData(i) == val:
                self.combo_purchase.setCurrentIndex(i)
                break
        self._update_preview()

    def _update_preview(self):
        s_sheet = self.combo_s_sheet.currentText() or "Sheet1"
        p_sheet = self.combo_p_sheet.currentText() or "Sheet2"

        d_col = self.input_date.text().strip().upper() or self.combo_date.currentData() or "A"
        s_col = self.input_sales.text().strip().upper() or self.combo_sales.currentData() or "B"
        p_col = self.input_purchase.text().strip().upper() or self.combo_purchase.currentData() or "C"

        period_desc = "📊 گزارش جامع (سود کل + تفکیک روزانه + ماهانه + سالانه)"
        if self.rb_total.isChecked():
            period_desc = "💵 فقط سود کل و شاخص‌های مالی"
        elif self.rb_daily.isChecked():
            period_desc = "📅 تفکیک سود روزانه با جدول و فرمول تفریق"
        elif self.rb_monthly.isChecked():
            period_desc = "🗓️ تفکیک سود ماهانه روند سودآوری"
        elif self.rb_yearly.isChecked():
            period_desc = "📆 تفکیک سود سالانه"

        method_str = "فایفو (FIFO)" if self.rb_fifo.isChecked() else ("لایفو (LIFO)" if self.rb_lifo.isChecked() else "میانگین موزون (Average)")

        preview = (
            f"🛒 **فروش ({s_sheet}):** ستون فروش `{s_col}` | ستون تاریخ: `{d_col}`\n"
            f"📦 **خرید ({p_sheet}):** ستون خرید/هزینه `{p_col}`\n"
            f"🏷️ **روش محاسبه بهای تمام‌شده (COGS):** {method_str}\n"
            f"💵 **فرمول سود ناخالص:** `= فروش خالص - بهای تمام‌شده کالای فروش‌رفته ({method_str})`\n"
            f"⏱️ **دوره تفکیک گزارش:** {period_desc}\n"
            f"📑 ایجاد شیت خروجی: {'بله (شیت اختصاصی «گزارش_سود» با داشبورد و فرمول‌ها)' if self.chk_create_sheet.isChecked() else 'در شیت فعلی'}"
        )
        self.preview_text.setText(preview)

    def _on_auto_submit(self):
        self._is_auto = True
        self.accept()

    def _on_submit(self):
        self._is_auto = False
        self.accept()

    def get_action(self) -> ProfitCalculationAction:
        val_method = "fifo"
        if self.rb_lifo.isChecked():
            val_method = "lifo"
        elif self.rb_avg.isChecked():
            val_method = "average"

        period = "all"
        if self.rb_total.isChecked():
            period = "total"
        elif self.rb_daily.isChecked():
            period = "daily"
        elif self.rb_monthly.isChecked():
            period = "monthly"
        elif self.rb_yearly.isChecked():
            period = "yearly"

        period_fa = {
            "all": "روزانه، ماهانه، سالانه و کل",
            "daily": "روزانه",
            "monthly": "ماهانه",
            "yearly": "سالانه",
            "total": "کل"
        }.get(period, "روزانه، ماهانه، سالانه و کل")

        if getattr(self, "_is_auto", False):
            return ProfitCalculationAction(
                sales_sheet=None,
                purchase_sheet=None,
                date_col=None,
                sales_col=None,
                purchase_col=None,
                valuation_method=val_method,
                period=period,
                create_report_sheet=self.chk_create_sheet.isChecked(),
                description_persian=f"محاسبه هوشمند سود {period_fa} به روش {val_method} (تشخیص خودکار شیت‌ها و ستون‌ها)"
            )

        s_sheet = self.combo_s_sheet.currentText() or None
        p_sheet = self.combo_p_sheet.currentText() or None

        d_col = self.input_date.text().strip().upper() or self.combo_date.currentData() or "A"
        s_col = self.input_sales.text().strip().upper() or self.combo_sales.currentData() or "B"
        p_col = self.input_purchase.text().strip().upper() or self.combo_purchase.currentData() or "C"

        return ProfitCalculationAction(
            sales_sheet=s_sheet,
            purchase_sheet=p_sheet,
            date_col=d_col,
            sales_col=s_col,
            purchase_col=p_col,
            valuation_method=val_method,
            period=period,
            create_report_sheet=self.chk_create_sheet.isChecked(),
            description_persian=f"محاسبه سود {period_fa} با روش {val_method} (فروش {s_sheet} ستون {s_col} منهای خرید {p_sheet} ستون {p_col} بر اساس تاریخ {d_col})"
        )


class ProfitMethodSelectorDialog(QDialog):
    """Fast, modern RTL dialog for selecting FIFO, LIFO, or Weighted Average valuation method with 1 click."""

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle("🏷️ انتخاب روش محاسبه بهای تمام‌شده و سود")
        self.resize(520, 440)
        self.setMinimumSize(460, 400)
        self.setLayoutDirection(Qt.RightToLeft)
        self.selected_method = "fifo"
        self.open_advanced = False

        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 16)
        layout.setSpacing(12)

        title = QLabel("🏷️ انتخاب روش محاسبه بهای تمام‌شده و سود")
        title.setStyleSheet("font-size: 15px; font-weight: bold; color: #10B981;")
        layout.addWidget(title)

        desc = QLabel(
            "سیستم به صورت خودکار ستون‌های فروش، خرید، تعداد، فی و تخفیفات را تشخیص می‌دهد.\n"
            "لطفاً روش مورد نظر برای ارزش‌گذاری اقلام فروش‌رفته (COGS) و انبار را انتخاب کنید:"
        )
        desc.setStyleSheet("font-size: 12px; color: #CBD5E1; line-height: 1.5;")
        desc.setWordWrap(True)
        layout.addWidget(desc)

        # 1. FIFO button
        self.btn_fifo = QPushButton("🟢 ۱. روش فایفو (FIFO) - اولین وارده، اولین صادره\n   (استاندارد رسمی حسابداری - خریدهای قدیمی‌تر ابتدا فروخته می‌شوند)")
        self.btn_fifo.setStyleSheet(
            "QPushButton {"
            "  background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #064E3B, stop:1 #047857);"
            "  color: #FFFFFF; font-weight: bold; font-size: 12.5px; padding: 12px 16px; border-radius: 9px;"
            "  border: 1px solid #10B981; text-align: right;"
            "}"
            "QPushButton:hover { background: #059669; border-color: #34D399; }"
        )
        self.btn_fifo.clicked.connect(lambda: self._choose("fifo"))
        layout.addWidget(self.btn_fifo)

        # 2. LIFO button
        self.btn_lifo = QPushButton("🔵 ۲. روش لایفو (LIFO) - آخرین وارده، اولین صادره\n   (مناسب شرایط تورمی - آخرین خریدهای گران‌تر ابتدا منظور می‌شوند)")
        self.btn_lifo.setStyleSheet(
            "QPushButton {"
            "  background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #1E3A8A, stop:1 #1D4ED8);"
            "  color: #FFFFFF; font-weight: bold; font-size: 12.5px; padding: 12px 16px; border-radius: 9px;"
            "  border: 1px solid #3B82F6; text-align: right;"
            "}"
            "QPushButton:hover { background: #2563EB; border-color: #60A5FA; }"
        )
        self.btn_lifo.clicked.connect(lambda: self._choose("lifo"))
        layout.addWidget(self.btn_lifo)

        # 3. Average button
        self.btn_avg = QPushButton("🟣 ۳. روش میانگین موزون (Weighted Average)\n   (محاسبه میانگین موزون قیمت تمام خریدهای انبار)")
        self.btn_avg.setStyleSheet(
            "QPushButton {"
            "  background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #581C87, stop:1 #7E22CE);"
            "  color: #FFFFFF; font-weight: bold; font-size: 12.5px; padding: 12px 16px; border-radius: 9px;"
            "  border: 1px solid #A855F7; text-align: right;"
            "}"
            "QPushButton:hover { background: #9333EA; border-color: #C084FC; }"
        )
        self.btn_avg.clicked.connect(lambda: self._choose("average"))
        layout.addWidget(self.btn_avg)

        # 4. Compare button
        self.btn_compare = QPushButton(
            "⚖️ ۴. مقایسه تطبیقی هر سه روش (فایفو vs لایفو vs میانگین موزون)\n"
            "   (تحلیل اثر تورم بر بهای تمام‌شده، سود ناخالص، صرفه‌جویی مالیاتی و موجودی انبار)"
        )
        self.btn_compare.setStyleSheet(
            "QPushButton {"
            "  background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #0F766E, stop:1 #0284C7);"
            "  color: #FFFFFF; font-weight: bold; font-size: 12.5px; padding: 12px 16px; border-radius: 9px;"
            "  border: 1px solid #38BDF8; text-align: right;"
            "}"
            "QPushButton:hover { background: #0369A1; border-color: #7DD3FC; }"
        )
        self.btn_compare.clicked.connect(lambda: self._choose("compare"))
        layout.addWidget(self.btn_compare)

        # Footer
        footer_layout = QHBoxLayout()
        self.btn_advanced = QPushButton("⚙️ تنظیمات دستی ستون‌ها...")
        self.btn_advanced.setStyleSheet("background: transparent; color: #94A3B8; text-decoration: underline; border: none; font-size: 11.5px;")
        self.btn_advanced.clicked.connect(self._on_advanced)

        self.btn_cancel = QPushButton("انصراف")
        self.btn_cancel.setStyleSheet("background: #1E293B; color: #E2E8F0; padding: 7px 16px; border-radius: 7px; font-size: 12px;")
        self.btn_cancel.clicked.connect(self.reject)

        footer_layout.addWidget(self.btn_advanced)
        footer_layout.addStretch()
        footer_layout.addWidget(self.btn_cancel)
        layout.addLayout(footer_layout)

    def _choose(self, method: str):
        self.selected_method = method
        self.open_advanced = False
        self.accept()

    def _on_advanced(self):
        self.open_advanced = True
        self.accept()


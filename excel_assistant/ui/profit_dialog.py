"""
Interactive Profit Summary Wizard Dialog for Persian Excel Assistant.
Allows users to visually select sales sheet & columns and purchase sheet & columns
(e.g., Sales in Sheet 1 and Purchase in Sheet 2),
choose reporting periods (daily, monthly, yearly, total, or comprehensive all-in-one),
and generate styled Excel profit reports with live KPI cards and formulas.
"""
from typing import List, Optional
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QRadioButton, QButtonGroup, QLineEdit,
    QComboBox, QFrame, QCheckBox, QMessageBox
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
        self.resize(580, 720)
        self.setLayoutDirection(Qt.RightToLeft)

        self._init_ui()
        self._load_sheets_and_columns()
        self._update_preview()

    def _init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 16, 20, 16)
        main_layout.setSpacing(11)

        # 1. Header
        header_lbl = QLabel("💵 محاسبه سود روزانه، ماهانه، سالانه و کل")
        header_lbl.setStyleSheet("font-size: 16px; font-weight: bold; color: #10B981;")
        main_layout.addWidget(header_lbl)

        desc_lbl = QLabel("شیت و ستون‌های فروش (شیت ۱) و خرید (شیت ۲ یا شیلد دیگر) را جهت محاسبه سود مشخص فرمایید:")
        desc_lbl.setStyleSheet("font-size: 12px; color: #94A3B8;")
        main_layout.addWidget(desc_lbl)

        # 2. Section 1: Sales Sheet & Columns (اطلاعات فروش - شیت ۱)
        s_box = QFrame()
        s_box.setStyleSheet("background-color: #060E22; border: 1px solid #1E3A8A; border-radius: 10px; padding: 10px;")
        s_layout = QVBoxLayout(s_box)
        s_layout.setSpacing(7)

        s_header = QLabel("🛒 ۱. اطلاعات فروش (شیت ۱ / شیلد فروش):")
        s_header.setStyleSheet("font-weight: bold; color: #38BDF8; font-size: 13px;")
        s_layout.addWidget(s_header)

        # Sales Sheet selector
        s_sheet_row = QHBoxLayout()
        s_sheet_lbl = QLabel("انتخاب شیت فروش:")
        s_sheet_lbl.setFixedWidth(130)
        s_sheet_lbl.setStyleSheet("color: #E2E8F0; font-size: 12px; font-weight: bold;")
        self.combo_s_sheet = QComboBox()
        self.combo_s_sheet.setStyleSheet("background-color: #081024; color: #FFFFFF; padding: 5px 8px; border: 1px solid #1E3A8A; border-radius: 6px;")
        self.combo_s_sheet.currentTextChanged.connect(self._on_s_sheet_changed)
        s_sheet_row.addWidget(s_sheet_lbl)
        s_sheet_row.addWidget(self.combo_s_sheet)
        s_layout.addLayout(s_sheet_row)

        # Date column
        date_row = QHBoxLayout()
        d_lbl = QLabel("ستون تاریخ فاکتور:")
        d_lbl.setFixedWidth(130)
        d_lbl.setStyleSheet("color: #CBD5E1; font-size: 12px;")
        self.combo_date = QComboBox()
        self.combo_date.setStyleSheet("background-color: #081024; color: #FFFFFF; padding: 5px 8px; border: 1px solid #1E3A8A; border-radius: 6px;")
        self.combo_date.currentTextChanged.connect(lambda: self._update_preview())
        self.input_date = QLineEdit()
        self.input_date.setPlaceholderText("مثلاً: A یا 1")
        self.input_date.setFixedWidth(80)
        self.input_date.setStyleSheet("background-color: #081024; color: #FFFFFF; padding: 5px 8px; border: 1px solid #1E3A8A; border-radius: 6px;")
        self.input_date.textChanged.connect(self._on_manual_date_changed)
        date_row.addWidget(d_lbl)
        date_row.addWidget(self.combo_date)
        date_row.addWidget(self.input_date)
        s_layout.addLayout(date_row)

        # Sales amount column
        sales_row = QHBoxLayout()
        s_lbl = QLabel("ستون مبلغ / قیمت فروش:")
        s_lbl.setFixedWidth(130)
        s_lbl.setStyleSheet("color: #CBD5E1; font-size: 12px;")
        self.combo_sales = QComboBox()
        self.combo_sales.setStyleSheet("background-color: #081024; color: #FFFFFF; padding: 5px 8px; border: 1px solid #1E3A8A; border-radius: 6px;")
        self.combo_sales.currentTextChanged.connect(lambda: self._update_preview())
        self.input_sales = QLineEdit()
        self.input_sales.setPlaceholderText("مثلاً: B یا 2")
        self.input_sales.setFixedWidth(80)
        self.input_sales.setStyleSheet("background-color: #081024; color: #FFFFFF; padding: 5px 8px; border: 1px solid #1E3A8A; border-radius: 6px;")
        self.input_sales.textChanged.connect(self._on_manual_sales_changed)
        sales_row.addWidget(s_lbl)
        sales_row.addWidget(self.combo_sales)
        sales_row.addWidget(self.input_sales)
        s_layout.addLayout(sales_row)

        main_layout.addWidget(s_box)

        # 3. Section 2: Purchase Sheet & Columns (اطلاعات خرید - شیت ۲ / شیلد دیگر)
        p_box = QFrame()
        p_box.setStyleSheet("background-color: #060E22; border: 1px solid #1E3A8A; border-radius: 10px; padding: 10px;")
        p_layout = QVBoxLayout(p_box)
        p_layout.setSpacing(7)

        p_header = QLabel("📦 ۲. اطلاعات خرید و بهای تمام‌شده (شیت ۲ / شیلد دیگر):")
        p_header.setStyleSheet("font-weight: bold; color: #F59E0B; font-size: 13px;")
        p_layout.addWidget(p_header)

        # Purchase Sheet selector
        p_sheet_row = QHBoxLayout()
        p_sheet_lbl = QLabel("انتخاب شیت خرید:")
        p_sheet_lbl.setFixedWidth(130)
        p_sheet_lbl.setStyleSheet("color: #E2E8F0; font-size: 12px; font-weight: bold;")
        self.combo_p_sheet = QComboBox()
        self.combo_p_sheet.setStyleSheet("background-color: #081024; color: #FFFFFF; padding: 5px 8px; border: 1px solid #1E3A8A; border-radius: 6px;")
        self.combo_p_sheet.currentTextChanged.connect(self._on_p_sheet_changed)
        p_sheet_row.addWidget(p_sheet_lbl)
        p_sheet_row.addWidget(self.combo_p_sheet)
        p_layout.addLayout(p_sheet_row)

        # Purchase amount column
        purchase_row = QHBoxLayout()
        p_lbl = QLabel("ستون مبلغ / قیمت خرید:")
        p_lbl.setFixedWidth(130)
        p_lbl.setStyleSheet("color: #CBD5E1; font-size: 12px;")
        self.combo_purchase = QComboBox()
        self.combo_purchase.setStyleSheet("background-color: #081024; color: #FFFFFF; padding: 5px 8px; border: 1px solid #1E3A8A; border-radius: 6px;")
        self.combo_purchase.currentTextChanged.connect(lambda: self._update_preview())
        self.input_purchase = QLineEdit()
        self.input_purchase.setPlaceholderText("مثلاً: C یا 3")
        self.input_purchase.setFixedWidth(80)
        self.input_purchase.setStyleSheet("background-color: #081024; color: #FFFFFF; padding: 5px 8px; border: 1px solid #1E3A8A; border-radius: 6px;")
        self.input_purchase.textChanged.connect(self._on_manual_purchase_changed)
        purchase_row.addWidget(p_lbl)
        purchase_row.addWidget(self.combo_purchase)
        purchase_row.addWidget(self.input_purchase)
        p_layout.addLayout(purchase_row)

        main_layout.addWidget(p_box)

        # 4. Section 3: Period Selection
        period_lbl = QLabel("⏱️ ۳. دوره زمانی تفکیک سود:")
        period_lbl.setStyleSheet("font-weight: bold; color: #E2E8F0; font-size: 13px;")
        main_layout.addWidget(period_lbl)

        self.radio_group = QButtonGroup(self)
        period_layout = QVBoxLayout()
        period_layout.setSpacing(4)

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

        # 5. Output option: Create dedicated sheet
        self.chk_create_sheet = QCheckBox("📑 ایجاد شیت اختصاصی «گزارش_سود» همراه با داشبورد، جداول تفکیک‌شده و فرمول‌ها")
        self.chk_create_sheet.setChecked(True)
        self.chk_create_sheet.setStyleSheet("color: #10B981; font-weight: bold; font-size: 12px;")
        main_layout.addWidget(self.chk_create_sheet)

        # 6. Live Preview Card
        self.preview_card = QFrame()
        self.preview_card.setStyleSheet("background-color: #04091A; border: 1px solid #1E3A8A; border-radius: 10px; padding: 10px;")
        card_layout = QVBoxLayout(self.preview_card)
        card_layout.setContentsMargins(10, 6, 10, 6)

        preview_title = QLabel("👁️ پیش‌نمایش نحوهٔ محاسبه سود:")
        preview_title.setStyleSheet("font-weight: bold; color: #94A3B8; font-size: 11px;")
        card_layout.addWidget(preview_title)

        self.preview_text = QLabel("در حال پردازش ستون‌ها...")
        self.preview_text.setStyleSheet("color: #F8FAFC; font-size: 11.5px; line-height: 1.5;")
        self.preview_text.setWordWrap(True)
        card_layout.addWidget(self.preview_text)

        main_layout.addWidget(self.preview_card)

        # 7. Buttons Box
        btn_box = QHBoxLayout()
        btn_box.setSpacing(12)

        self.btn_cancel = QPushButton("انصراف")
        self.btn_cancel.setStyleSheet(
            "background-color: #1E293B; color: #E2E8F0; padding: 8px 18px; border-radius: 8px; font-weight: bold;"
        )
        self.btn_cancel.clicked.connect(self.reject)

        self.btn_submit = QPushButton("محاسبه سود 💵")
        self.btn_submit.setStyleSheet(
            "background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #059669, stop:1 #10B981); "
            "color: white; padding: 8px 24px; border-radius: 8px; font-weight: bold; font-size: 13px; border: 1px solid #10B981;"
        )
        self.btn_submit.clicked.connect(self._on_submit)

        btn_box.addWidget(self.btn_cancel)
        btn_box.addWidget(self.btn_submit)
        main_layout.addLayout(btn_box)

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
            best_sales_idx = min(1, max_c - min_c)

            date_kws = ["تاریخ", "date", "روز", "زمان", "time"]
            sales_kws = ["فروش", "مبلغ", "درآمد", "قیمت", "sales", "amount", "price", "فی"]

            self.combo_date.blockSignals(True)
            self.combo_sales.blockSignals(True)
            self.combo_date.clear()
            self.combo_sales.clear()

            for idx, c in enumerate(range(min_c, max_c + 1)):
                c_let = col_index_to_letter(c)
                hdr = str(sheet.get_value(f"{c_let}1") or "").strip()
                label = f"ستون {c_let} (شماره {c})" + (f" - {hdr}" if hdr else "")
                self.combo_date.addItem(label, c_let)
                self.combo_sales.addItem(label, c_let)

                hdr_l = hdr.lower()
                if any(k in hdr_l for k in date_kws): best_date_idx = idx
                if any(k in hdr_l for k in sales_kws): best_sales_idx = idx

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

            best_purchase_idx = min(2, max_c - min_c)
            purchase_kws = ["خرید", "هزینه", "بهای تمام شده", "قیمت تمام شده", "cost", "purchase", "قیمت", "مبلغ", "فی"]

            self.combo_purchase.blockSignals(True)
            self.combo_purchase.clear()

            for idx, c in enumerate(range(min_c, max_c + 1)):
                c_let = col_index_to_letter(c)
                hdr = str(sheet.get_value(f"{c_let}1") or "").strip()
                label = f"ستون {c_let} (شماره {c})" + (f" - {hdr}" if hdr else "")
                self.combo_purchase.addItem(label, c_let)

                hdr_l = hdr.lower()
                if any(k in hdr_l for k in purchase_kws): best_purchase_idx = idx

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

        preview = (
            f"🛒 **فروش ({s_sheet}):** ستون فروش `{s_col}` | ستون تاریخ: `{d_col}`\n"
            f"📦 **خرید ({p_sheet}):** ستون خرید/هزینه `{p_col}`\n"
            f"💵 **فرمول سود خالص:** `= ({s_sheet}!{s_col}) - ({p_sheet}!{p_col})`\n"
            f"⏱️ **دوره تفکیک گزارش:** {period_desc}\n"
            f"📑 ایجاد شیت خروجی: {'بله (شیت اختصاصی «گزارش_سود» با داشبورد و فرمول‌ها)' if self.chk_create_sheet.isChecked() else 'در شیت فعلی'}"
        )
        self.preview_text.setText(preview)

    def _on_submit(self):
        self.accept()

    def get_action(self) -> ProfitCalculationAction:
        s_sheet = self.combo_s_sheet.currentText() or None
        p_sheet = self.combo_p_sheet.currentText() or None

        d_col = self.input_date.text().strip().upper() or self.combo_date.currentData() or "A"
        s_col = self.input_sales.text().strip().upper() or self.combo_sales.currentData() or "B"
        p_col = self.input_purchase.text().strip().upper() or self.combo_purchase.currentData() or "C"

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

        return ProfitCalculationAction(
            sales_sheet=s_sheet,
            purchase_sheet=p_sheet,
            date_col=d_col,
            sales_col=s_col,
            purchase_col=p_col,
            period=period,
            create_report_sheet=self.chk_create_sheet.isChecked(),
            description_persian=f"محاسبه سود {period_fa} (فروش {s_sheet} ستون {s_col} منهای خرید {p_sheet} ستون {p_col} بر اساس تاریخ {d_col})"
        )

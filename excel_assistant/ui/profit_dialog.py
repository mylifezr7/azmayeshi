"""
Interactive Profit Summary Wizard Dialog for Persian Excel Assistant.
Allows users to visually select date column, sales and purchase/cost columns (or direct profit column),
choose reporting periods (daily, monthly, yearly, total, or comprehensive all-in-one),
and generate styled Excel profit reports with live KPI cards and formulas.
"""
from typing import List, Optional
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QRadioButton, QButtonGroup, QLineEdit,
    QComboBox, QFrame, QCheckBox, QMessageBox, QStackedWidget
)

from excel_assistant.core.utils import col_index_to_letter, col_letter_to_index
from excel_assistant.nlp.intent_models import ProfitCalculationAction
from excel_assistant.core.connection import ExcelConnectionManager


class ProfitWizardDialog(QDialog):
    """Modern RTL dialog for calculating daily, monthly, yearly, and total profit."""

    def __init__(self, manager: Optional[ExcelConnectionManager] = None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.manager = manager
        self.setWindowTitle("💵 محاسبه سود (روزانه، ماهانه، سالانه و کل)")
        self.resize(560, 660)
        self.setLayoutDirection(Qt.RightToLeft)

        self._init_ui()
        self._load_sheet_columns()
        self._update_preview()

    def _init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 18, 20, 18)
        main_layout.setSpacing(12)

        # 1. Header
        header_lbl = QLabel("💵 محاسبه سود روزانه، ماهانه، سالانه و کل")
        header_lbl.setStyleSheet("font-size: 16px; font-weight: bold; color: #10B981;")
        main_layout.addWidget(header_lbl)

        desc_lbl = QLabel("ستون تاریخ، مبالغ فروش و خرید/هزینه و دورهٔ زمانی مورد نظر را برای تحلیل سود انتخاب نمایید:")
        desc_lbl.setStyleSheet("font-size: 12px; color: #94A3B8;")
        main_layout.addWidget(desc_lbl)

        # 2. Section 1: Date Column
        date_lbl = QLabel("📅 ستون تاریخ فاکتورها / تراکنش‌ها:")
        date_lbl.setStyleSheet("font-weight: bold; color: #E2E8F0; font-size: 13px;")
        main_layout.addWidget(date_lbl)

        date_box = QHBoxLayout()
        self.combo_date = QComboBox()
        self.combo_date.setMinimumWidth(240)
        self.combo_date.setStyleSheet("background-color: #081024; color: #FFFFFF; padding: 6px 10px; border: 1px solid #1E3A8A; border-radius: 8px;")
        self.combo_date.currentTextChanged.connect(lambda: self._update_preview())

        self.input_date = QLineEdit()
        self.input_date.setPlaceholderText("شماره یا حرف (مثلاً: 1 یا A)")
        self.input_date.setStyleSheet("background-color: #081024; color: #FFFFFF; padding: 6px 10px; border: 1px solid #1E3A8A; border-radius: 8px;")
        self.input_date.textChanged.connect(self._on_manual_date_changed)

        date_box.addWidget(self.combo_date)
        date_box.addWidget(self.input_date)
        main_layout.addLayout(date_box)

        # 3. Section 2: Calculation Method & Amounts
        calc_box = QFrame()
        calc_box.setStyleSheet("background-color: #060E22; border: 1px solid #1E3A8A; border-radius: 10px; padding: 10px;")
        calc_layout = QVBoxLayout(calc_box)
        calc_layout.setSpacing(8)

        calc_title = QLabel("💳 مبالغ فروش و خرید (جهت محاسبه سود = فروش منهای خرید):")
        calc_title.setStyleSheet("font-weight: bold; color: #38BDF8; font-size: 13px;")
        calc_layout.addWidget(calc_title)

        # Row A: Sales Column
        sales_row = QHBoxLayout()
        s_lbl = QLabel("ستون مبلغ فروش:")
        s_lbl.setFixedWidth(110)
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
        calc_layout.addLayout(sales_row)

        # Row B: Purchase/Cost Column
        purchase_row = QHBoxLayout()
        p_lbl = QLabel("ستون مبلغ خرید/هزینه:")
        p_lbl.setFixedWidth(110)
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
        calc_layout.addLayout(purchase_row)

        main_layout.addWidget(calc_box)

        # 4. Section 3: Period Selection
        period_lbl = QLabel("⏱️ دوره زمانی تفکیک سود:")
        period_lbl.setStyleSheet("font-weight: bold; color: #E2E8F0; font-size: 13px;")
        main_layout.addWidget(period_lbl)

        self.radio_group = QButtonGroup(self)
        period_layout = QVBoxLayout()
        period_layout.setSpacing(5)

        self.rb_all = QRadioButton("📊 گزارش جامع (سود کل + تفکیک روزانه + ماهانه + سالانه)")
        self.rb_all.setChecked(True)
        self.rb_total = QRadioButton("💵 فقط سود کل (مجموع سود، میانگین حاشیه سود و پرسودترین قلم)")
        self.rb_daily = QRadioButton("📅 تفکیک سود روزانه (مجموع فروش، خرید و سود هر روز به همراه فرمول)")
        self.rb_monthly = QRadioButton("🗓️ تفکیک سود ماهانه (روند سودآوری ماه‌های سال)")
        self.rb_yearly = QRadioButton("📆 تفکیک سود سالانه (سهم و مجموع سود سال‌ها)")

        for rb in [self.rb_all, self.rb_total, self.rb_daily, self.rb_monthly, self.rb_yearly]:
            rb.setStyleSheet("color: #CBD5E1; font-size: 12px;")
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
        self.preview_card.setStyleSheet("background-color: #04091A; border: 1px solid #1E3A8A; border-radius: 10px; padding: 12px;")
        card_layout = QVBoxLayout(self.preview_card)
        card_layout.setContentsMargins(10, 8, 10, 8)

        preview_title = QLabel("👁️ پیش‌نمایش نحوهٔ محاسبه سود:")
        preview_title.setStyleSheet("font-weight: bold; color: #94A3B8; font-size: 11px;")
        card_layout.addWidget(preview_title)

        self.preview_text = QLabel("در حال پردازش ستون‌ها...")
        self.preview_text.setStyleSheet("color: #F8FAFC; font-size: 12px; line-height: 1.5;")
        self.preview_text.setWordWrap(True)
        card_layout.addWidget(self.preview_text)

        main_layout.addWidget(self.preview_card)

        # 7. Buttons Box
        btn_box = QHBoxLayout()
        btn_box.setSpacing(12)

        self.btn_cancel = QPushButton("انصراف")
        self.btn_cancel.setStyleSheet(
            "background-color: #1E293B; color: #E2E8F0; padding: 9px 18px; border-radius: 8px; font-weight: bold;"
        )
        self.btn_cancel.clicked.connect(self.reject)

        self.btn_submit = QPushButton("محاسبه سود 💵")
        self.btn_submit.setStyleSheet(
            "background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #059669, stop:1 #10B981); "
            "color: white; padding: 9px 24px; border-radius: 8px; font-weight: bold; font-size: 13px; border: 1px solid #10B981;"
        )
        self.btn_submit.clicked.connect(self._on_submit)

        btn_box.addWidget(self.btn_cancel)
        btn_box.addWidget(self.btn_submit)
        main_layout.addLayout(btn_box)

    def _load_sheet_columns(self):
        if not self.manager or not self.manager.is_connected():
            return

        sheet = self.manager.get_active_sheet()
        min_r, min_c, max_r, max_c = sheet.get_used_range_bounds()

        best_date_idx = 0
        best_sales_idx = min(1, max_c - min_c)
        best_purchase_idx = min(2, max_c - min_c)

        date_keywords = ["تاریخ", "date", "روز", "زمان", "time"]
        sales_keywords = ["فروش", "مبلغ", "درآمد", "قیمت", "sales", "amount", "price"]
        purchase_keywords = ["خرید", "هزینه", "بهای تمام شده", "قیمت تمام شده", "cost", "purchase"]

        items = []
        for idx, c in enumerate(range(min_c, max_c + 1)):
            col_letter = col_index_to_letter(c)
            header_val = str(sheet.get_value(f"{col_letter}1") or "").strip()
            label = f"ستون {col_letter} (شماره {c})"
            if header_val:
                label += f" - {header_val}"
            items.append((col_letter, label))

            hdr_lower = header_val.lower()
            if any(k in hdr_lower for k in date_keywords):
                best_date_idx = idx
            if any(k in hdr_lower for k in sales_keywords):
                best_sales_idx = idx
            if any(k in hdr_lower for k in purchase_keywords):
                best_purchase_idx = idx

        self.combo_date.blockSignals(True)
        self.combo_sales.blockSignals(True)
        self.combo_purchase.blockSignals(True)

        self.combo_date.clear()
        self.combo_sales.clear()
        self.combo_purchase.clear()

        for col_letter, label in items:
            self.combo_date.addItem(label, col_letter)
            self.combo_sales.addItem(label, col_letter)
            self.combo_purchase.addItem(label, col_letter)

        if items:
            self.combo_date.setCurrentIndex(best_date_idx)
            self.combo_sales.setCurrentIndex(best_sales_idx)
            self.combo_purchase.setCurrentIndex(best_purchase_idx)

        self.combo_date.blockSignals(False)
        self.combo_sales.blockSignals(False)
        self.combo_purchase.blockSignals(False)

    def _on_manual_date_changed(self, text: str):
        val = text.strip().upper()
        if not val:
            return
        if val.isdigit():
            c_idx = int(val)
            val = col_index_to_letter(c_idx)
        for i in range(self.combo_date.count()):
            if self.combo_date.itemData(i) == val:
                self.combo_date.setCurrentIndex(i)
                break
        self._update_preview()

    def _on_manual_sales_changed(self, text: str):
        val = text.strip().upper()
        if not val:
            return
        if val.isdigit():
            c_idx = int(val)
            val = col_index_to_letter(c_idx)
        for i in range(self.combo_sales.count()):
            if self.combo_sales.itemData(i) == val:
                self.combo_sales.setCurrentIndex(i)
                break
        self._update_preview()

    def _on_manual_purchase_changed(self, text: str):
        val = text.strip().upper()
        if not val:
            return
        if val.isdigit():
            c_idx = int(val)
            val = col_index_to_letter(c_idx)
        for i in range(self.combo_purchase.count()):
            if self.combo_purchase.itemData(i) == val:
                self.combo_purchase.setCurrentIndex(i)
                break
        self._update_preview()

    def _update_preview(self):
        d_col = self.input_date.text().strip().upper() or self.combo_date.currentData() or "A"
        s_col = self.input_sales.text().strip().upper() or self.combo_sales.currentData() or "B"
        p_col = self.input_purchase.text().strip().upper() or self.combo_purchase.currentData() or "C"

        period_desc = "📊 گزارش جامع (سود کل + تفکیک روزانه + ماهانه + سالانه)"
        if self.rb_total.isChecked():
            period_desc = "💵 فقط سود کل و شاخص‌های مالی"
        elif self.rb_daily.isChecked():
            period_desc = "📅 تفکیک سود روزانه با جدول و فرمول‌ها"
        elif self.rb_monthly.isChecked():
            period_desc = "🗓️ تفکیک سود ماهانه روند سودآوری"
        elif self.rb_yearly.isChecked():
            period_desc = "📆 تفکیک سود سالانه"

        preview = (
            f"📅 **ستون تاریخ:** `{d_col}`\n"
            f"💰 **ستون فروش:** `{s_col}` | 🛒 **ستون خرید/هزینه:** `{p_col}`\n"
            f"💵 **فرمول سود خالص:** `= ستون {s_col} - ستون {p_col}`\n"
            f"⏱️ **دوره تفکیک گزارش:** {period_desc}\n"
            f"📑 ایجاد شیت خروجی: {'بله (شیت اختصاصی «گزارش_سود» با داشبورد و فرمول‌ها)' if self.chk_create_sheet.isChecked() else 'در شیت فعلی'}"
        )
        self.preview_text.setText(preview)

    def _on_submit(self):
        self.accept()

    def get_action(self) -> ProfitCalculationAction:
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

        sheet_name = self.manager.get_active_sheet().name if (self.manager and self.manager.is_connected()) else None

        return ProfitCalculationAction(
            sales_sheet=sheet_name,
            purchase_sheet=sheet_name,
            date_col=d_col,
            sales_col=s_col,
            purchase_col=p_col,
            period=period,
            create_report_sheet=self.chk_create_sheet.isChecked(),
            description_persian=f"محاسبه سود {period_fa} ستون فروش {s_col} و خرید {p_col} بر اساس تاریخ {d_col}"
        )

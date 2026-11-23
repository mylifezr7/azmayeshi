"""
Interactive Sales Summary Wizard Dialog for Persian Excel Assistant.
Allows users to visually select date and sales amount columns, choose reporting periods
(daily, monthly, yearly, total, or comprehensive all-in-one), and generate styled Excel reports.
"""
from typing import List, Optional
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QRadioButton, QButtonGroup, QLineEdit,
    QComboBox, QFrame, QCheckBox, QMessageBox, QScrollArea
)

from excel_assistant.core.utils import col_index_to_letter, col_letter_to_index
from excel_assistant.nlp.intent_models import SalesSummaryAction
from excel_assistant.core.connection import ExcelConnectionManager


class SalesSummaryWizardDialog(QDialog):
    """Modern RTL dialog for calculating daily, monthly, yearly, and total sales."""

    def __init__(self, manager: Optional[ExcelConnectionManager] = None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.manager = manager
        self.setWindowTitle("💰 محاسبه و تحلیل جامع فروش (روزانه، ماهانه، سالانه و کل)")
        self.resize(540, 520)
        self.setMinimumSize(480, 380)
        self.setLayoutDirection(Qt.RightToLeft)

        self._init_ui()
        self._load_sheet_columns()
        self._update_preview()

    def _init_ui(self):
        root_layout = QVBoxLayout(self)
        root_layout.setContentsMargins(16, 12, 16, 12)
        root_layout.setSpacing(8)

        # 1. Header (pinned at top)
        header_lbl = QLabel("💰 محاسبه فروش روزانه، ماهانه، سالانه و کل")
        header_lbl.setStyleSheet("font-size: 15px; font-weight: bold; color: #38BDF8;")
        root_layout.addWidget(header_lbl)

        desc_lbl = QLabel("ستون تاریخ، ستون مبلغ فروش و دورهٔ زمانی مورد نظر را برای تحلیل هوشمند انتخاب نمایید:")
        desc_lbl.setStyleSheet("font-size: 11.5px; color: #94A3B8;")
        root_layout.addWidget(desc_lbl)

        # 2. Scrollable Body
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setStyleSheet("QScrollArea { background: transparent; border: none; }")

        content_widget = QWidget()
        content_widget.setStyleSheet("background: transparent;")
        main_layout = QVBoxLayout(content_widget)
        main_layout.setContentsMargins(2, 4, 8, 4)
        main_layout.setSpacing(10)

        # Section 1: Date Column
        date_lbl = QLabel("📅 ستون تاریخ فاکتورها / تراکنش‌ها:")
        date_lbl.setStyleSheet("font-weight: bold; color: #E2E8F0; font-size: 12.5px;")
        main_layout.addWidget(date_lbl)

        date_box = QHBoxLayout()
        self.combo_date = QComboBox()
        self.combo_date.setMinimumWidth(220)
        self.combo_date.setStyleSheet("background-color: #081024; color: #FFFFFF; padding: 5px 8px; border: 1px solid #1E3A8A; border-radius: 6px;")
        self.combo_date.currentTextChanged.connect(lambda: self._update_preview())

        self.input_date = QLineEdit()
        self.input_date.setPlaceholderText("شماره یا حرف (مثلاً: 1 یا A)")
        self.input_date.setFixedWidth(80)
        self.input_date.setStyleSheet("background-color: #081024; color: #FFFFFF; padding: 5px 8px; border: 1px solid #1E3A8A; border-radius: 6px;")
        self.input_date.textChanged.connect(self._on_manual_date_changed)

        date_box.addWidget(self.combo_date)
        date_box.addWidget(self.input_date)
        main_layout.addLayout(date_box)

        # Section 2: Sales Amount Column
        amount_lbl = QLabel("💳 ستون مبلغ / قیمت فروش:")
        amount_lbl.setStyleSheet("font-weight: bold; color: #E2E8F0; font-size: 12.5px;")
        main_layout.addWidget(amount_lbl)

        amount_box = QHBoxLayout()
        self.combo_amount = QComboBox()
        self.combo_amount.setMinimumWidth(220)
        self.combo_amount.setStyleSheet("background-color: #081024; color: #FFFFFF; padding: 5px 8px; border: 1px solid #1E3A8A; border-radius: 6px;")
        self.combo_amount.currentTextChanged.connect(lambda: self._update_preview())

        self.input_amount = QLineEdit()
        self.input_amount.setPlaceholderText("شماره یا حرف (مثلاً: 2 یا B)")
        self.input_amount.setFixedWidth(80)
        self.input_amount.setStyleSheet("background-color: #081024; color: #FFFFFF; padding: 5px 8px; border: 1px solid #1E3A8A; border-radius: 6px;")
        self.input_amount.textChanged.connect(self._on_manual_amount_changed)

        amount_box.addWidget(self.combo_amount)
        amount_box.addWidget(self.input_amount)
        main_layout.addLayout(amount_box)

        # Section 3: Period Selection
        period_lbl = QLabel("⏱️ دوره زمانی گزارش و تفکیک:")
        period_lbl.setStyleSheet("font-weight: bold; color: #E2E8F0; font-size: 12.5px;")
        main_layout.addWidget(period_lbl)

        self.radio_group = QButtonGroup(self)
        period_layout = QVBoxLayout()
        period_layout.setSpacing(4)

        self.rb_all = QRadioButton("📊 گزارش جامع (فروش کل + تفکیک روزانه + ماهانه + سالانه)")
        self.rb_all.setChecked(True)
        self.rb_total = QRadioButton("💰 فقط فروش کل (مجموع مبالغ، تعداد فاکتورها، میانگین و بالاترین فاکتور)")
        self.rb_daily = QRadioButton("📅 تفکیک فروش روزانه (مجموع فروش هر روز به همراه فاکتورها)")
        self.rb_monthly = QRadioButton("🗓️ تفکیک فروش ماهانه (روند فروش ماه‌های سال)")
        self.rb_yearly = QRadioButton("📆 تفکیک فروش سالانه (سهم و مجموع فروش سال‌ها)")

        for rb in [self.rb_all, self.rb_total, self.rb_daily, self.rb_monthly, self.rb_yearly]:
            rb.setStyleSheet("color: #CBD5E1; font-size: 11px;")
            self.radio_group.addButton(rb)
            period_layout.addWidget(rb)
            rb.toggled.connect(lambda: self._update_preview())

        main_layout.addLayout(period_layout)

        # Output option: Create dedicated sheet
        self.chk_create_sheet = QCheckBox("📑 ایجاد شیت اختصاصی «گزارش_فروش» همراه با جداول تفکیک‌شده و فرمول‌ها")
        self.chk_create_sheet.setChecked(True)
        self.chk_create_sheet.setStyleSheet("color: #38BDF8; font-weight: bold; font-size: 11.5px;")
        main_layout.addWidget(self.chk_create_sheet)

        # Live Preview Card
        self.preview_card = QFrame()
        self.preview_card.setStyleSheet("background-color: #04091A; border: 1px solid #1E3A8A; border-radius: 9px; padding: 8px;")
        card_layout = QVBoxLayout(self.preview_card)
        card_layout.setContentsMargins(8, 6, 8, 6)

        preview_title = QLabel("👁️ پیش‌نمایش نحوهٔ محاسبه:")
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
            "background-color: #1E293B; color: #E2E8F0; padding: 9px 20px; border-radius: 8px; font-weight: bold; font-size: 12.5px;"
        )
        self.btn_cancel.clicked.connect(self.reject)

        self.btn_submit = QPushButton("محاسبه و ایجاد گزارش 📈 (تایید)")
        self.btn_submit.setDefault(True)
        self.btn_submit.setAutoDefault(True)
        self.btn_submit.setStyleSheet(
            "background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #059669, stop:1 #10B981); "
            "color: white; padding: 9px 28px; border-radius: 8px; font-weight: bold; font-size: 13px; border: 1px solid #10B981;"
        )
        self.btn_submit.clicked.connect(self._on_submit)

        btn_box.addWidget(self.btn_cancel)
        btn_box.addWidget(self.btn_submit)
        root_layout.addLayout(btn_box)

    def _load_sheet_columns(self):
        if not self.manager or not self.manager.is_connected():
            return

        sheet = self.manager.get_active_sheet()
        min_r, min_c, max_r, max_c = sheet.get_used_range_bounds()

        best_date_idx = 0
        best_amount_idx = min(1, max_c - min_c)

        date_keywords = ["تاریخ", "date", "روز", "زمان", "time", "day", "month"]
        amount_keywords = ["فروش", "مبلغ", "درآمد", "قیمت", "sales", "amount", "price", "total", "جمع"]

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
            if any(k in hdr_lower for k in amount_keywords):
                best_amount_idx = idx

        self.combo_date.blockSignals(True)
        self.combo_amount.blockSignals(True)

        self.combo_date.clear()
        self.combo_amount.clear()

        for col_letter, label in items:
            self.combo_date.addItem(label, col_letter)
            self.combo_amount.addItem(label, col_letter)

        if items:
            self.combo_date.setCurrentIndex(best_date_idx)
            self.combo_amount.setCurrentIndex(best_amount_idx)

        self.combo_date.blockSignals(False)
        self.combo_amount.blockSignals(False)

    def _on_manual_date_changed(self, text: str):
        val = text.strip().upper()
        if not val:
            return
        c_letter = col_index_to_letter(int(val)) if val.isdigit() else val
        for i in range(self.combo_date.count()):
            if self.combo_date.itemData(i) == c_letter:
                self.combo_date.blockSignals(True)
                self.combo_date.setCurrentIndex(i)
                self.combo_date.blockSignals(False)
                break
        self._update_preview()

    def _on_manual_amount_changed(self, text: str):
        val = text.strip().upper()
        if not val:
            return
        c_letter = col_index_to_letter(int(val)) if val.isdigit() else val
        for i in range(self.combo_amount.count()):
            if self.combo_amount.itemData(i) == c_letter:
                self.combo_amount.blockSignals(True)
                self.combo_amount.setCurrentIndex(i)
                self.combo_amount.blockSignals(False)
                break
        self._update_preview()

    def _get_active_columns(self):
        d_man = self.input_date.text().strip().upper()
        if d_man:
            d_col = col_index_to_letter(int(d_man)) if d_man.isdigit() else d_man
        else:
            d_col = self.combo_date.currentData() or "A"

        a_man = self.input_amount.text().strip().upper()
        if a_man:
            a_col = col_index_to_letter(int(a_man)) if a_man.isdigit() else a_man
        else:
            a_col = self.combo_amount.currentData() or "B"

        return d_col, a_col

    def _get_period(self) -> str:
        if self.rb_total.isChecked():
            return "total"
        elif self.rb_daily.isChecked():
            return "daily"
        elif self.rb_monthly.isChecked():
            return "monthly"
        elif self.rb_yearly.isChecked():
            return "yearly"
        return "all"

    def _update_preview(self):
        d_col, a_col = self._get_active_columns()
        period = self._get_period()

        period_desc = {
            "all": "گزارش جامع (فروش کل، روزانه، ماهانه و سالانه)",
            "total": "فروش کل و شاخص‌های آماری",
            "daily": "تفکیک فروش روزانه",
            "monthly": "تفکیک فروش ماهانه",
            "yearly": "تفکیک فروش سالانه"
        }.get(period, "گزارش فروش")

        preview = (
            f"📌 ستون تاریخ: `{d_col}` | ستون مبلغ فروش: `{a_col}`\n"
            f"📊 دوره گزارش: **{period_desc}**\n"
            f"📑 ایجاد شیت خروجی: {'بله (شیت «گزارش_فروش» با جداول تحلیلی)' if self.chk_create_sheet.isChecked() else 'در شیت فعلی'}"
        )
        self.preview_text.setText(preview)

    def _on_submit(self):
        d_col, a_col = self._get_active_columns()
        if not d_col or not a_col:
            QMessageBox.warning(self, "خطا", "لطفاً ستون تاریخ و ستون مبلغ فروش را مشخص کنید.")
            return
        if d_col == a_col:
            QMessageBox.warning(self, "خطا", "ستون تاریخ و ستون مبلغ فروش نمی‌توانند یکسان باشند.")
            return

        self.accept()

    def get_action(self) -> SalesSummaryAction:
        d_col, a_col = self._get_active_columns()
        period = self._get_period()
        create_sheet = self.chk_create_sheet.isChecked()

        period_fa = {
            "all": "روزانه، ماهانه، سالانه و کل",
            "total": "کل",
            "daily": "روزانه",
            "monthly": "ماهانه",
            "yearly": "سالانه"
        }.get(period, "جامع")

        return SalesSummaryAction(
            date_col=d_col,
            amount_col=a_col,
            period=period,
            create_report_sheet=create_sheet,
            description_persian=f"محاسبه فروش {period_fa} ستون {a_col} بر اساس تاریخ {d_col}"
        )

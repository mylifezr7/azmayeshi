"""
Interactive Profit Wizard Dialog for Persian Excel Assistant.
Allows users to visually select sales sheet/columns and purchase sheet/columns,
and generate a comprehensive profit & loss comparison in Excel.
"""
from typing import List, Optional
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QComboBox, QFrame, QCheckBox, QMessageBox, QLineEdit,
    QRadioButton, QButtonGroup
)

from excel_assistant.core.utils import col_index_to_letter
from excel_assistant.nlp.intent_models import ProfitCalculationAction
from excel_assistant.core.connection import ExcelConnectionManager


class ProfitWizardDialog(QDialog):
    """Modern RTL dialog for calculating profit by subtracting purchase from sales (daily, monthly, yearly, total)."""

    def __init__(self, manager: Optional[ExcelConnectionManager] = None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.manager = manager
        self.setWindowTitle("💵 محاسبه سود و زیان (روزانه، ماهانه، سالانه و کل)")
        self.resize(560, 680)
        self.setLayoutDirection(Qt.RightToLeft)

        self._init_ui()
        self._load_sheets_and_columns()
        self._update_preview()

    def _init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 18, 20, 18)
        main_layout.setSpacing(12)

        # 1. Header
        header_lbl = QLabel("💵 محاسبه سود خالص (روزانه، ماهانه، سالانه و کل)")
        header_lbl.setStyleSheet("font-size: 16px; font-weight: bold; color: #38BDF8;")
        main_layout.addWidget(header_lbl)

        desc_lbl = QLabel("شیت‌ها، ستون‌ها و دوره زمانی مورد نظر را جهت محاسبه سود مشخص فرمایید:")
        desc_lbl.setStyleSheet("font-size: 12px; color: #94A3B8;")
        main_layout.addWidget(desc_lbl)

        # 2. Section 1: Sales Sheet (Sheet 1)
        s_box = QFrame()
        s_box.setStyleSheet("background-color: #060E22; border: 1px solid #1E3A8A; border-radius: 10px; padding: 10px;")
        s_layout = QVBoxLayout(s_box)

        s_title = QLabel("🛒 ۱. اطلاعات فروش (شیت ۱):")
        s_title.setStyleSheet("font-weight: bold; color: #38BDF8; font-size: 13px;")
        s_layout.addWidget(s_title)

        s_sheet_row = QHBoxLayout()
        s_sheet_row.addWidget(QLabel("شیت فروش:"))
        self.combo_s_sheet = QComboBox()
        self.combo_s_sheet.setStyleSheet("background-color: #081024; color: #FFFFFF; padding: 4px 8px; border: 1px solid #1E3A8A; border-radius: 6px;")
        self.combo_s_sheet.currentTextChanged.connect(self._on_s_sheet_changed)
        s_sheet_row.addWidget(self.combo_s_sheet)

        s_sheet_row.addWidget(QLabel("ستون تاریخ (اختیاری):"))
        self.combo_date = QComboBox()
        self.combo_date.setStyleSheet("background-color: #081024; color: #FFFFFF; padding: 4px 8px; border: 1px solid #1E3A8A; border-radius: 6px;")
        self.combo_date.currentTextChanged.connect(lambda: self._update_preview())
        s_sheet_row.addWidget(self.combo_date)
        s_layout.addLayout(s_sheet_row)

        s_cols_row = QHBoxLayout()
        s_cols_row.addWidget(QLabel("تعداد فروش:"))
        self.combo_s_qty = QComboBox()
        self.combo_s_qty.setStyleSheet("background-color: #081024; color: #FFFFFF; padding: 4px 8px; border: 1px solid #1E3A8A; border-radius: 6px;")
        self.combo_s_qty.currentTextChanged.connect(lambda: self._update_preview())
        s_cols_row.addWidget(self.combo_s_qty)

        s_cols_row.addWidget(QLabel("قیمت فروش:"))
        self.combo_s_price = QComboBox()
        self.combo_s_price.setStyleSheet("background-color: #081024; color: #FFFFFF; padding: 4px 8px; border: 1px solid #1E3A8A; border-radius: 6px;")
        self.combo_s_price.currentTextChanged.connect(lambda: self._update_preview())
        s_cols_row.addWidget(self.combo_s_price)
        s_layout.addLayout(s_cols_row)

        main_layout.addWidget(s_box)

        # 3. Section 2: Purchase Sheet (Sheet 2)
        p_box = QFrame()
        p_box.setStyleSheet("background-color: #060E22; border: 1px solid #1E3A8A; border-radius: 10px; padding: 10px;")
        p_layout = QVBoxLayout(p_box)

        p_title = QLabel("📦 ۲. اطلاعات خرید و بهای تمام‌شده (شیت ۲):")
        p_title.setStyleSheet("font-weight: bold; color: #F59E0B; font-size: 13px;")
        p_layout.addWidget(p_title)

        p_sheet_row = QHBoxLayout()
        p_sheet_row.addWidget(QLabel("شیت خرید:"))
        self.combo_p_sheet = QComboBox()
        self.combo_p_sheet.setStyleSheet("background-color: #081024; color: #FFFFFF; padding: 4px 8px; border: 1px solid #1E3A8A; border-radius: 6px;")
        self.combo_p_sheet.currentTextChanged.connect(self._on_p_sheet_changed)
        p_sheet_row.addWidget(self.combo_p_sheet)
        p_layout.addLayout(p_sheet_row)

        p_cols_row = QHBoxLayout()
        p_cols_row.addWidget(QLabel("تعداد خرید:"))
        self.combo_p_qty = QComboBox()
        self.combo_p_qty.setStyleSheet("background-color: #081024; color: #FFFFFF; padding: 4px 8px; border: 1px solid #1E3A8A; border-radius: 6px;")
        self.combo_p_qty.currentTextChanged.connect(lambda: self._update_preview())
        p_cols_row.addWidget(self.combo_p_qty)

        p_cols_row.addWidget(QLabel("قیمت خرید:"))
        self.combo_p_price = QComboBox()
        self.combo_p_price.setStyleSheet("background-color: #081024; color: #FFFFFF; padding: 4px 8px; border: 1px solid #1E3A8A; border-radius: 6px;")
        self.combo_p_price.currentTextChanged.connect(lambda: self._update_preview())
        p_cols_row.addWidget(self.combo_p_price)
        p_layout.addLayout(p_cols_row)

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
        self.rb_total = QRadioButton("💵 فقط سود کل (مجموع سود، میانگین حاشیه سود و پرسودترین قلم)")
        self.rb_daily = QRadioButton("📅 تفکیک سود روزانه (مجموع فروش، خرید و سود هر روز)")
        self.rb_monthly = QRadioButton("🗓️ تفکیک سود ماهانه (روند سودآوری ماه‌های سال)")
        self.rb_yearly = QRadioButton("📆 تفکیک سود سالانه (سهم و مجموع سود سال‌ها)")

        for rb in [self.rb_all, self.rb_total, self.rb_daily, self.rb_monthly, self.rb_yearly]:
            rb.setStyleSheet("color: #CBD5E1; font-size: 11px;")
            self.radio_group.addButton(rb)
            period_layout.addWidget(rb)
            rb.toggled.connect(lambda: self._update_preview())

        main_layout.addLayout(period_layout)

        # 5. Output option
        self.chk_create_sheet = QCheckBox("📑 ایجاد شیت اختصاصی «گزارش_سود» با داشبورد و فرمول‌های اکسل")
        self.chk_create_sheet.setChecked(True)
        self.chk_create_sheet.setStyleSheet("color: #38BDF8; font-weight: bold; font-size: 12px;")
        main_layout.addWidget(self.chk_create_sheet)

        # 6. Live Preview Card
        self.preview_card = QFrame()
        self.preview_card.setStyleSheet("background-color: #04091A; border: 1px solid #1E3A8A; border-radius: 10px; padding: 10px;")
        card_layout = QVBoxLayout(self.preview_card)
        card_layout.setContentsMargins(10, 8, 10, 8)

        preview_title = QLabel("👁️ پیش‌نمایش نحوه محاسبه:")
        preview_title.setStyleSheet("font-weight: bold; color: #94A3B8; font-size: 11px;")
        card_layout.addWidget(preview_title)

        self.preview_text = QLabel("در حال بارگذاری...")
        self.preview_text.setStyleSheet("color: #F8FAFC; font-size: 12px; line-height: 1.5;")
        self.preview_text.setWordWrap(True)
        card_layout.addWidget(self.preview_text)

        main_layout.addWidget(self.preview_card)

        # 7. Action buttons
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

        self._populate_columns_for_sheet(self.combo_s_sheet.currentText(), self.combo_s_qty, self.combo_s_price)
        self._populate_columns_for_sheet(self.combo_p_sheet.currentText(), self.combo_p_qty, self.combo_p_price)

    def _populate_columns_for_sheet(self, sheet_name: str, qty_combo: QComboBox, price_combo: QComboBox):
        if not sheet_name or not self.manager.is_connected():
            return
        try:
            sheet = self.manager.get_sheet(sheet_name)
            min_r, min_c, max_r, max_c = sheet.get_used_range_bounds()

            qty_combo.blockSignals(True)
            price_combo.blockSignals(True)
            qty_combo.clear()
            price_combo.clear()

            best_qty_idx = 0
            best_price_idx = min(1, max_c - min_c)

            qty_kws = ["تعداد", "مقدار", "qty"]
            price_kws = ["قیمت", "مبلغ", "فی", "price", "cost"]

            for idx, c in enumerate(range(min_c, max_c + 1)):
                c_let = col_index_to_letter(c)
                hdr = str(sheet.get_value(f"{c_let}1") or "").strip()
                label = f"ستون {c_let} (شماره {c})" + (f" - {hdr}" if hdr else "")
                qty_combo.addItem(label, c_let)
                price_combo.addItem(label, c_let)

                hdr_l = hdr.lower()
                if any(k in hdr_l for k in qty_kws): best_qty_idx = idx
                if any(k in hdr_l for k in price_kws): best_price_idx = idx

            if qty_combo.count() > 0:
                qty_combo.setCurrentIndex(best_qty_idx)
                price_combo.setCurrentIndex(best_price_idx)

            qty_combo.blockSignals(False)
            price_combo.blockSignals(False)

            # Also update combo_date if this is sales sheet
            if qty_combo == self.combo_s_qty:
                self.combo_date.blockSignals(True)
                self.combo_date.clear()
                self.combo_date.addItem("تشخیص خودکار تاریخ", "")
                best_date_idx = 0
                date_kws = ["تاریخ", "date", "روز", "زمان"]
                for idx, c in enumerate(range(min_c, max_c + 1)):
                    c_let = col_index_to_letter(c)
                    hdr = str(sheet.get_value(f"{c_let}1") or "").strip()
                    label = f"ستون {c_let}" + (f" - {hdr}" if hdr else "")
                    self.combo_date.addItem(label, c_let)
                    if any(k in hdr.lower() for k in date_kws):
                        best_date_idx = idx + 1
                self.combo_date.setCurrentIndex(best_date_idx)
                self.combo_date.blockSignals(False)
        except Exception:
            pass

    def _on_s_sheet_changed(self, text: str):
        self._populate_columns_for_sheet(text, self.combo_s_qty, self.combo_s_price)
        self._update_preview()

    def _on_p_sheet_changed(self, text: str):
        self._populate_columns_for_sheet(text, self.combo_p_qty, self.combo_p_price)
        self._update_preview()

    def _update_preview(self):
        s_name = self.combo_s_sheet.currentText()
        p_name = self.combo_p_sheet.currentText()
        sq = self.combo_s_qty.currentData() or "B"
        sp = self.combo_s_price.currentData() or "C"
        pq = self.combo_p_qty.currentData() or "B"
        pp = self.combo_p_price.currentData() or "C"

        period_str = "گزارش جامع (روزانه + ماهانه + سالانه + کل)"
        if self.rb_total.isChecked(): period_str = "فقط سود کل"
        elif self.rb_daily.isChecked(): period_str = "تفکیک سود روزانه"
        elif self.rb_monthly.isChecked(): period_str = "تفکیک سود ماهانه"
        elif self.rb_yearly.isChecked(): period_str = "تفکیک سود سالانه"

        date_col = self.combo_date.currentData() if self.combo_date.currentIndex() > 0 else "خودکار"

        preview = (
            f"🛒 **فروش ({s_name}):** تعداد `{sq}` × قیمت `{sp}` | تاریخ: `{date_col}`\n"
            f"📦 **خرید ({p_name}):** تعداد `{pq}` × قیمت `{pp}`\n"
            f"⏱️ **دوره محاسبه:** {period_str}\n"
            f"💵 **فرمول سود:** `(فروش) - (خرید)` با فرمول‌های اکسل و جدول داشبورد\n"
            f"📑 ایجاد شیت خروجی: {'بله («گزارش_سود»)' if self.chk_create_sheet.isChecked() else 'در شیت فعلی'}"
        )
        self.preview_text.setText(preview)

    def _on_submit(self):
        self.accept()

    def get_action(self) -> ProfitCalculationAction:
        period = "all"
        if self.rb_total.isChecked():
            period = "total"
        elif self.rb_daily.isChecked():
            period = "daily"
        elif self.rb_monthly.isChecked():
            period = "monthly"
        elif self.rb_yearly.isChecked():
            period = "yearly"

        date_col = self.combo_date.currentData() if self.combo_date.currentIndex() > 0 else None

        period_fa = {
            "all": "روزانه، ماهانه، سالانه و کل",
            "daily": "روزانه",
            "monthly": "ماهانه",
            "yearly": "سالانه",
            "total": "کل"
        }.get(period, "روزانه، ماهانه، سالانه و کل")

        return ProfitCalculationAction(
            sales_sheet=self.combo_s_sheet.currentText(),
            purchase_sheet=self.combo_p_sheet.currentText(),
            date_col=date_col,
            sales_qty_col=self.combo_s_qty.currentData(),
            sales_price_col=self.combo_s_price.currentData(),
            purchase_qty_col=self.combo_p_qty.currentData(),
            purchase_price_col=self.combo_p_price.currentData(),
            period=period,
            create_report_sheet=self.chk_create_sheet.isChecked(),
            description_persian=f"محاسبه سود {period_fa}"
        )

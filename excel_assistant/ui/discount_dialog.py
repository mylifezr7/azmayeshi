"""
Interactive Discount Wizard Dialog for Persian Excel Assistant.
Allows users to visually select the original price column, discount percentage column
(or fixed discount %), and target destination column for the discounted price.
"""
from typing import List, Optional, Tuple
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QRadioButton, QButtonGroup, QLineEdit,
    QComboBox, QFrame, QCheckBox, QMessageBox, QDoubleSpinBox
)

from excel_assistant.core.utils import col_index_to_letter, col_letter_to_index
from excel_assistant.nlp.intent_models import ApplyDiscountAction
from excel_assistant.core.connection import ExcelConnectionManager


class DiscountWizardDialog(QDialog):
    """Modern RTL dialog for applying discounts to items/prices in Excel."""

    def __init__(self, manager: Optional[ExcelConnectionManager] = None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.manager = manager
        self.setWindowTitle("🏷️ راهنمای محاسبه و اعمال تخفیف کالاها")
        self.resize(520, 560)
        self.setLayoutDirection(Qt.RightToLeft)

        self._init_ui()
        self._load_sheet_columns()
        self._update_preview()

    def _init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(18, 16, 18, 16)
        main_layout.setSpacing(14)

        # 1. Header
        header_lbl = QLabel("🏷️ محاسبه و اعمال تخفیف روی کالاها")
        header_lbl.setStyleSheet("font-size: 15px; font-weight: bold; color: #38BDF8;")
        main_layout.addWidget(header_lbl)

        desc_lbl = QLabel("ستون قیمت اولیه، ستون درصد تخفیف و ستون قیمت نهایی را مشخص کنید:")
        desc_lbl.setStyleSheet("font-size: 12px; color: #94A3B8;")
        main_layout.addWidget(desc_lbl)

        # 2. Section 1: Original Price Column
        price_lbl = QLabel("💰 ستون قیمت اصلی کالا (قبل از تخفیف):")
        price_lbl.setStyleSheet("font-weight: bold; color: #E2E8F0; font-size: 13px;")
        main_layout.addWidget(price_lbl)

        price_box = QHBoxLayout()
        self.combo_price = QComboBox()
        self.combo_price.setMinimumWidth(220)
        self.combo_price.setStyleSheet("background-color: #1E293B; color: #FFFFFF; padding: 6px 10px; border: 1px solid #475569; border-radius: 6px;")
        self.combo_price.currentTextChanged.connect(lambda: self._update_preview())

        self.input_price = QLineEdit()
        self.input_price.setPlaceholderText("شماره یا حرف (مثلاً: 2 یا B)")
        self.input_price.setStyleSheet("background-color: #1E293B; color: #FFFFFF; padding: 6px 10px; border: 1px solid #475569; border-radius: 6px;")
        self.input_price.textChanged.connect(self._on_manual_price_changed)

        price_box.addWidget(self.combo_price)
        price_box.addWidget(self.input_price)
        main_layout.addLayout(price_box)

        # 3. Section 2: Discount Percentage Selection (Column vs Fixed %)
        disc_lbl = QLabel("📊 نحوه تعیین درصد تخفیف:")
        disc_lbl.setStyleSheet("font-weight: bold; color: #E2E8F0; font-size: 13px;")
        main_layout.addWidget(disc_lbl)

        disc_frame = QFrame()
        disc_frame.setStyleSheet("background-color: #1E293B; border-radius: 8px; padding: 10px;")
        disc_layout = QVBoxLayout(disc_frame)
        disc_layout.setSpacing(10)

        # Mode A: Column-based discount
        self.radio_col = QRadioButton("خواندن درصد تخفیف از یک ستون اکسل:")
        self.radio_col.setChecked(True)
        self.radio_col.setStyleSheet("font-weight: bold; color: #F1F5F9; font-size: 12px;")
        self.radio_col.toggled.connect(self._on_discount_mode_changed)
        disc_layout.addWidget(self.radio_col)

        col_box = QHBoxLayout()
        self.combo_disc_col = QComboBox()
        self.combo_disc_col.setMinimumWidth(200)
        self.combo_disc_col.setStyleSheet("background-color: #0F172A; color: #FFFFFF; padding: 5px 8px; border: 1px solid #475569; border-radius: 6px;")
        self.combo_disc_col.currentTextChanged.connect(lambda: self._update_preview())

        self.input_disc_col = QLineEdit()
        self.input_disc_col.setPlaceholderText("شماره یا حرف (مثلاً: 3 یا C)")
        self.input_disc_col.setStyleSheet("background-color: #0F172A; color: #FFFFFF; padding: 5px 8px; border: 1px solid #475569; border-radius: 6px;")
        self.input_disc_col.textChanged.connect(self._on_manual_disc_col_changed)

        col_box.addWidget(self.combo_disc_col)
        col_box.addWidget(self.input_disc_col)
        disc_layout.addLayout(col_box)

        # Mode B: Fixed percentage
        fixed_box = QHBoxLayout()
        self.radio_fixed = QRadioButton("اعمال درصد تخفیف ثابت برای همه کالاها:")
        self.radio_fixed.setStyleSheet("font-weight: bold; color: #F1F5F9; font-size: 12px;")
        self.radio_fixed.toggled.connect(self._on_discount_mode_changed)

        self.spin_fixed = QDoubleSpinBox()
        self.spin_fixed.setRange(0.1, 100.0)
        self.spin_fixed.setValue(15.0)
        self.spin_fixed.setSuffix(" %")
        self.spin_fixed.setEnabled(False)
        self.spin_fixed.setStyleSheet("background-color: #0F172A; color: #FFFFFF; padding: 5px 8px; border: 1px solid #475569; border-radius: 6px;")
        self.spin_fixed.valueChanged.connect(lambda: self._update_preview())

        fixed_box.addWidget(self.radio_fixed)
        fixed_box.addWidget(self.spin_fixed)
        disc_layout.addLayout(fixed_box)

        main_layout.addWidget(disc_frame)

        # 4. Section 3: Target Column (Final Price After Discount)
        target_lbl = QLabel("🎯 ستون قیمت بعد از تخفیف (مقصد):")
        target_lbl.setStyleSheet("font-weight: bold; color: #E2E8F0; font-size: 13px;")
        main_layout.addWidget(target_lbl)

        target_box = QHBoxLayout()
        self.combo_target = QComboBox()
        self.combo_target.setMinimumWidth(220)
        self.combo_target.setStyleSheet("background-color: #1E293B; color: #FFFFFF; padding: 6px 10px; border: 1px solid #475569; border-radius: 6px;")
        self.combo_target.currentTextChanged.connect(lambda: self._update_preview())

        self.input_target = QLineEdit()
        self.input_target.setPlaceholderText("شماره یا حرف (مثلاً: 4 یا D)")
        self.input_target.setStyleSheet("background-color: #1E293B; color: #FFFFFF; padding: 6px 10px; border: 1px solid #475569; border-radius: 6px;")
        self.input_target.textChanged.connect(self._on_manual_target_changed)

        target_box.addWidget(self.combo_target)
        target_box.addWidget(self.input_target)
        main_layout.addLayout(target_box)

        # Optional: Discount amount column
        self.cb_amount = QCheckBox("محاسبه و درج ستون «مبلغ تخفیف (تومان/ریال)» به صورت مجزا")
        self.cb_amount.setStyleSheet("color: #E2E8F0; font-size: 12px;")
        self.cb_amount.toggled.connect(lambda: self._update_preview())
        main_layout.addWidget(self.cb_amount)

        # 5. Live Formula Preview
        self.preview_frame = QFrame()
        self.preview_frame.setStyleSheet("background-color: #132E20; border: 1px solid #107C41; border-radius: 8px; padding: 10px;")
        preview_layout = QVBoxLayout(self.preview_frame)
        preview_layout.setContentsMargins(10, 8, 10, 8)
        preview_layout.setSpacing(4)

        self.lbl_preview_title = QLabel("💡 پیش‌نمایش فرمول اکسل:")
        self.lbl_preview_title.setStyleSheet("font-weight: bold; color: #4ADE80; font-size: 12px;")
        self.lbl_preview_formula = QLabel("=")
        self.lbl_preview_formula.setStyleSheet("font-family: Consolas, monospace; font-size: 14px; font-weight: bold; color: #FFFFFF;")
        self.lbl_preview_desc = QLabel("")
        self.lbl_preview_desc.setStyleSheet("font-size: 11px; color: #A7F3D0;")

        preview_layout.addWidget(self.lbl_preview_title)
        preview_layout.addWidget(self.lbl_preview_formula)
        preview_layout.addWidget(self.lbl_preview_desc)
        main_layout.addWidget(self.preview_frame)

        # 6. Action Buttons
        btn_box = QHBoxLayout()
        btn_box.setSpacing(12)

        self.btn_execute = QPushButton("🚀 اعمال تخفیف در اکسل")
        self.btn_execute.setStyleSheet("background-color: #107C41; color: white; font-weight: bold; padding: 10px 18px; border-radius: 8px; font-size: 13px;")
        self.btn_execute.clicked.connect(self._on_execute_clicked)

        self.btn_cancel = QPushButton("❌ انصراف")
        self.btn_cancel.setStyleSheet("background-color: #334155; color: #E2E8F0; padding: 10px 18px; border-radius: 8px; font-size: 13px;")
        self.btn_cancel.clicked.connect(self.reject)

        btn_box.addWidget(self.btn_execute)
        btn_box.addWidget(self.btn_cancel)
        main_layout.addLayout(btn_box)

    def _load_sheet_columns(self):
        max_col = 10
        headers_map = {}

        if self.manager and self.manager.is_connected():
            try:
                sheet = self.manager.get_active_sheet()
                _, _, _, max_c = sheet.get_used_range_bounds()
                max_col = max(10, min(max_c + 4, 30))
                for c_i in range(1, max_col + 1):
                    c_letter = col_index_to_letter(c_i)
                    val = sheet.get_value(f"{c_letter}1")
                    if val is not None and str(val).strip():
                        headers_map[c_i] = str(val).strip()
            except Exception:
                pass

        # Populate combos
        self.combo_price.clear()
        self.combo_disc_col.clear()
        self.combo_target.clear()

        self.combo_target.addItem("خودکار (اولین ستون خالی بعد از داده‌ها)", "")

        for c_i in range(1, max_col + 1):
            c_letter = col_index_to_letter(c_i)
            hdr = headers_map.get(c_i, "")
            label = f"ستون {c_i} ({c_letter})" + (f" - «{hdr}»" if hdr else "")

            self.combo_price.addItem(label, c_letter)
            self.combo_disc_col.addItem(label, c_letter)
            self.combo_target.addItem(label, c_letter)

        # Smart defaults: if price and discount headers exist, select them
        price_idx = 1 if max_col >= 2 else 0  # Default col B
        disc_idx = 2 if max_col >= 3 else 0   # Default col C

        for i in range(self.combo_price.count()):
            txt = self.combo_price.itemText(i).lower()
            if "قیمت" in txt or "مبلغ" in txt or "price" in txt:
                price_idx = i
                break
        self.combo_price.setCurrentIndex(price_idx)

        for i in range(self.combo_disc_col.count()):
            txt = self.combo_disc_col.itemText(i).lower()
            if "تخفیف" in txt or "درصد" in txt or "discount" in txt:
                disc_idx = i
                break
        self.combo_disc_col.setCurrentIndex(disc_idx)

    def _on_discount_mode_changed(self):
        is_col = self.radio_col.isChecked()
        self.combo_disc_col.setEnabled(is_col)
        self.input_disc_col.setEnabled(is_col)
        self.spin_fixed.setEnabled(not is_col)
        self._update_preview()

    def _on_manual_price_changed(self, text: str):
        val = text.strip().upper()
        if not val:
            return
        c_let = col_index_to_letter(int(val)) if val.isdigit() else val
        for i in range(self.combo_price.count()):
            if self.combo_price.itemData(i) == c_let:
                self.combo_price.setCurrentIndex(i)
                break
        self._update_preview()

    def _on_manual_disc_col_changed(self, text: str):
        val = text.strip().upper()
        if not val:
            return
        c_let = col_index_to_letter(int(val)) if val.isdigit() else val
        for i in range(self.combo_disc_col.count()):
            if self.combo_disc_col.itemData(i) == c_let:
                self.combo_disc_col.setCurrentIndex(i)
                break
        self._update_preview()

    def _on_manual_target_changed(self, text: str):
        val = text.strip().upper()
        if not val:
            return
        c_let = col_index_to_letter(int(val)) if val.isdigit() else val
        for i in range(self.combo_target.count()):
            if self.combo_target.itemData(i) == c_let:
                self.combo_target.setCurrentIndex(i)
                break
        self._update_preview()

    def get_price_column(self) -> str:
        manual = self.input_price.text().strip().upper()
        if manual:
            return col_index_to_letter(int(manual)) if manual.isdigit() else manual
        return self.combo_price.currentData() or "B"

    def get_discount_column(self) -> Optional[str]:
        if not self.radio_col.isChecked():
            return None
        manual = self.input_disc_col.text().strip().upper()
        if manual:
            return col_index_to_letter(int(manual)) if manual.isdigit() else manual
        return self.combo_disc_col.currentData() or "C"

    def get_fixed_discount_percent(self) -> Optional[float]:
        if self.radio_fixed.isChecked():
            return self.spin_fixed.value()
        return None

    def get_target_column(self) -> Optional[str]:
        manual = self.input_target.text().strip().upper()
        if manual:
            return col_index_to_letter(int(manual)) if manual.isdigit() else manual
        data = self.combo_target.currentData()
        return data if data else None

    def _update_preview(self):
        p_col = self.get_price_column()
        d_col = self.get_discount_column()
        pct = self.get_fixed_discount_percent()
        target = self.get_target_column() or "خودکار"

        if d_col:
            formula = f"=IF({d_col}2>1, {p_col}2*(1-{d_col}2/100), {p_col}2*(1-{d_col}2))"
            desc = f"محاسبه قیمت با تخفیف: قیمت اولیه ستون {p_col} منهای درصد تخفیف ستون {d_col} -> ستون {target}"
        elif pct is not None:
            formula = f"={p_col}2*(1 - {pct}/100)"
            desc = f"محاسبه قیمت با تخفیف {pct}% روی ستون {p_col} -> ستون {target}"
        else:
            formula = f"={p_col}2"
            desc = f"ستون {p_col}"

        self.lbl_preview_formula.setText(formula)
        self.lbl_preview_desc.setText(desc)

    def _on_execute_clicked(self):
        p_col = self.get_price_column()
        d_col = self.get_discount_column()
        if self.radio_col.isChecked() and p_col == d_col:
            QMessageBox.warning(self, "خطای انتخاب ستون", "ستون قیمت اصلی و ستون درصد تخفیف نمی‌توانند یکسان باشند.")
            return
        self.accept()

    def get_action(self) -> ApplyDiscountAction:
        p_col = self.get_price_column()
        d_col = self.get_discount_column()
        pct = self.get_fixed_discount_percent()
        target = self.get_target_column()

        da_col = None
        if self.cb_amount.isChecked():
            # If target given, place discount amount column next to target, else auto
            if target:
                da_idx = col_letter_to_index(target) + 1
                da_col = col_index_to_letter(da_idx)

        d_str = f"ستون {d_col}" if d_col else f"{pct}%"
        return ApplyDiscountAction(
            price_col=p_col,
            discount_col=d_col,
            discount_percent=pct,
            target_col=target,
            discount_amount_col=da_col,
            description_persian=f"محاسبه تخفیف ستون {p_col} با درصد {d_str} و درج در ستون {target or 'جدید'}"
        )

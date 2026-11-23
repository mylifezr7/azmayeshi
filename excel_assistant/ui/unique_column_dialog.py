"""
Interactive Unique Column Wizard Dialog for Persian Excel Assistant.
Allows users to select a column (e.g. column 12 / L for product code or national ID),
and apply Excel Data Validation to strictly prevent duplicate entries with custom Persian error alert.
"""
from typing import List, Optional
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QLineEdit, QComboBox, QFrame, QCheckBox, QMessageBox
)

from excel_assistant.core.utils import col_index_to_letter, col_letter_to_index
from excel_assistant.nlp.intent_models import PreventDuplicatesAction
from excel_assistant.core.connection import ExcelConnectionManager


class UniqueColumnWizardDialog(QDialog):
    """Modern RTL dialog for locking a column to prevent duplicate values (e.g. National ID / Product Code)."""

    def __init__(self, manager: Optional[ExcelConnectionManager] = None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.manager = manager
        self.setWindowTitle("🔒 قفل یکتاسازی ستون و جلوگیری از ورود داده تکراری")
        self.resize(520, 540)
        self.setLayoutDirection(Qt.RightToLeft)

        self._init_ui()
        self._load_sheet_columns()
        self._update_preview()

    def _init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(18, 16, 18, 16)
        main_layout.setSpacing(14)

        # 1. Header
        header_lbl = QLabel("🔒 جلوگیری از ورود دادهٔ تکراری در ستون (کد محصول / کد ملی)")
        header_lbl.setStyleSheet("font-size: 14px; font-weight: bold; color: #38BDF8;")
        main_layout.addWidget(header_lbl)

        desc_lbl = QLabel(
            "با فعال‌سازی این قابلیت، اکسل به طور خودکار ورود هرگونه مقدار تکراری در ستون انتخابی\n"
            "را مسدود کرده و با نمایش پیام اخطار فارسی، مانع از ثبت کد تکراری می‌شود."
        )
        desc_lbl.setStyleSheet("font-size: 12px; color: #94A3B8; line-height: 1.5;")
        main_layout.addWidget(desc_lbl)

        # 2. Section 1: Target Column
        col_lbl = QLabel("📌 ستون مورد نظر برای یکتاسازی:")
        col_lbl.setStyleSheet("font-weight: bold; color: #E2E8F0; font-size: 13px;")
        main_layout.addWidget(col_lbl)

        col_box = QHBoxLayout()
        self.combo_col = QComboBox()
        self.combo_col.setMinimumWidth(220)
        self.combo_col.setStyleSheet("background-color: #1E293B; color: #FFFFFF; padding: 6px 10px; border: 1px solid #475569; border-radius: 6px;")
        self.combo_col.currentTextChanged.connect(lambda: self._update_preview())

        self.input_col = QLineEdit()
        self.input_col.setPlaceholderText("شماره یا حرف ستون (مثلاً: 12 یا L)")
        self.input_col.setStyleSheet("background-color: #1E293B; color: #FFFFFF; padding: 6px 10px; border: 1px solid #475569; border-radius: 6px;")
        self.input_col.textChanged.connect(self._on_manual_col_changed)

        col_box.addWidget(self.combo_col)
        col_box.addWidget(self.input_col)
        main_layout.addLayout(col_box)

        # 3. Section 2: Error Alert Configuration
        alert_lbl = QLabel("🛑 تنظیمات پیام خطای اکسل در صورت ورود تکراری:")
        alert_lbl.setStyleSheet("font-weight: bold; color: #E2E8F0; font-size: 13px;")
        main_layout.addWidget(alert_lbl)

        alert_frame = QFrame()
        alert_frame.setStyleSheet("background-color: #1E293B; border-radius: 8px; padding: 10px;")
        alert_layout = QVBoxLayout(alert_frame)
        alert_layout.setSpacing(8)

        lbl_t = QLabel("عنوان پیام خطای اکسل:")
        lbl_t.setStyleSheet("color: #94A3B8; font-size: 11px;")
        self.input_title = QLineEdit("خطای ورود داده تکراری")
        self.input_title.setStyleSheet("background-color: #0F172A; color: #FFFFFF; padding: 5px 8px; border: 1px solid #475569; border-radius: 6px;")

        lbl_m = QLabel("متن اخطار مانع‌شونده (Stop Alert):")
        lbl_m.setStyleSheet("color: #94A3B8; font-size: 11px;")
        self.input_msg = QLineEdit("این کد قبلاً در این ستون ثبت شده است و تکراری می‌باشد. لطفاً کد یکتای دیگری وارد کنید.")
        self.input_msg.setStyleSheet("background-color: #0F172A; color: #FFFFFF; padding: 5px 8px; border: 1px solid #475569; border-radius: 6px;")

        alert_layout.addWidget(lbl_t)
        alert_layout.addWidget(self.input_title)
        alert_layout.addWidget(lbl_m)
        alert_layout.addWidget(self.input_msg)
        main_layout.addWidget(alert_frame)

        # 4. Section 3: Highlight existing duplicates checkbox
        self.cb_highlight = QCheckBox("بررسی و هایلایت داده‌های تکراری فعلی در صورت وجود")
        self.cb_highlight.setChecked(True)
        self.cb_highlight.setStyleSheet("color: #E2E8F0; font-size: 12px; font-weight: 500;")
        main_layout.addWidget(self.cb_highlight)

        # 5. Live Preview
        self.preview_frame = QFrame()
        self.preview_frame.setStyleSheet("background-color: #132E20; border: 1px solid #107C41; border-radius: 8px; padding: 10px;")
        preview_layout = QVBoxLayout(self.preview_frame)
        preview_layout.setContentsMargins(10, 8, 10, 8)
        preview_layout.setSpacing(4)

        self.lbl_preview_title = QLabel("💡 فرمول اعتبارسنجی اکسل (Data Validation):")
        self.lbl_preview_title.setStyleSheet("font-weight: bold; color: #4ADE80; font-size: 12px;")
        self.lbl_preview_formula = QLabel("")
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

        self.btn_execute = QPushButton("🚀 فعال‌سازی قفل یکتاسازی در اکسل")
        self.btn_execute.setStyleSheet("background-color: #107C41; color: white; font-weight: bold; padding: 10px 18px; border-radius: 8px; font-size: 13px;")
        self.btn_execute.clicked.connect(self._on_execute_clicked)

        self.btn_cancel = QPushButton("❌ انصراف")
        self.btn_cancel.setStyleSheet("background-color: #334155; color: #E2E8F0; padding: 10px 18px; border-radius: 8px; font-size: 13px;")
        self.btn_cancel.clicked.connect(self.reject)

        btn_box.addWidget(self.btn_execute)
        btn_box.addWidget(self.btn_cancel)
        main_layout.addLayout(btn_box)

    def _load_sheet_columns(self):
        max_col = 15
        headers_map = {}

        if self.manager and self.manager.is_connected():
            try:
                sheet = self.manager.get_active_sheet()
                _, _, _, max_c = sheet.get_used_range_bounds()
                max_col = max(15, min(max_c + 4, 30))
                for c_i in range(1, max_col + 1):
                    c_letter = col_index_to_letter(c_i)
                    val = sheet.get_value(f"{c_letter}1")
                    if val is not None and str(val).strip():
                        headers_map[c_i] = str(val).strip()
            except Exception:
                pass

        self.combo_col.clear()
        selected_idx = 0
        for c_i in range(1, max_col + 1):
            c_letter = col_index_to_letter(c_i)
            hdr = headers_map.get(c_i, "")
            label = f"ستون {c_i} ({c_letter})" + (f" - «{hdr}»" if hdr else "")
            self.combo_col.addItem(label, c_letter)

            # Smart pick: prioritize column with code/id/national
            hdr_low = hdr.lower()
            if any(k in hdr_low for k in ["کد", "شناسه", "ملی", "code", "id", "product", "محصول"]):
                selected_idx = c_i - 1

        self.combo_col.setCurrentIndex(selected_idx)

    def _on_manual_col_changed(self, text: str):
        val = text.strip().upper()
        if not val:
            return
        c_let = col_index_to_letter(int(val)) if val.isdigit() else val
        for i in range(self.combo_col.count()):
            if self.combo_col.itemData(i) == c_let:
                self.combo_col.setCurrentIndex(i)
                break
        self._update_preview()

    def get_target_column(self) -> str:
        manual = self.input_col.text().strip().upper()
        if manual:
            return col_index_to_letter(int(manual)) if manual.isdigit() else manual
        return self.combo_col.currentData() or "A"

    def _update_preview(self):
        col = self.get_target_column()
        formula = f"=COUNTIF({col}:{col}, {col}1)<=1"
        self.lbl_preview_formula.setText(formula)
        self.lbl_preview_desc.setText(f"قفل ستون {col}: اکسل اجازه ورود هیچ مقدار تکراری را نخواهد داد.")

    def _on_execute_clicked(self):
        self.accept()

    def get_action(self) -> PreventDuplicatesAction:
        col = self.get_target_column()
        title = self.input_title.text().strip() or "خطای ورود داده تکراری"
        msg = self.input_msg.text().strip() or "این مقدار قبلاً در این ستون ثبت شده است و تکراری می‌باشد."
        hl = self.cb_highlight.isChecked()

        return PreventDuplicatesAction(
            target_col=col,
            error_title=title,
            error_message=msg,
            highlight_existing=hl,
            description_persian=f"قفل یکتاسازی ستون {col} و جلوگیری از ورود دادهٔ تکراری در اکسل"
        )

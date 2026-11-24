"""
Interactive Unique Column Wizard Dialog for Persian Excel Assistant.
Provides a clear choice between:
1. Applying uniqueness lock on a column (قفل یکتایی)
2. Canceling/removing uniqueness lock from a column (لغو قفل یکتایی)
"""
from typing import List, Optional, Union
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QLineEdit, QComboBox, QFrame, QCheckBox,
    QRadioButton, QButtonGroup, QMessageBox
)

from excel_assistant.core.utils import col_index_to_letter, col_letter_to_index
from excel_assistant.nlp.intent_models import PreventDuplicatesAction, RemoveUniqueValidationAction
from excel_assistant.core.connection import ExcelConnectionManager


class UniqueColumnWizardDialog(QDialog):
    """Modern RTL dialog providing dedicated sections for applying or canceling column uniqueness lock."""

    def __init__(self, manager: Optional[ExcelConnectionManager] = None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.manager = manager
        self.mode = "lock"  # "lock" or "unlock"
        self.setWindowTitle("🔒 مدیریت یکتاسازی ستون اکسل (اعمال یا لغو قفل)")
        self.resize(540, 600)
        self.setLayoutDirection(Qt.RightToLeft)

        self._init_ui()
        self._load_sheet_columns()
        self._update_preview()

    def _init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(18, 16, 18, 16)
        main_layout.setSpacing(12)

        # 1. Header
        header_lbl = QLabel("🔒 مدیریت یکتاسازی ستون‌ها (کد محصول / کد ملی)")
        header_lbl.setStyleSheet("font-size: 14px; font-weight: bold; color: #38BDF8;")
        main_layout.addWidget(header_lbl)

        desc_lbl = QLabel(
            "عملیات مورد نظر خود را انتخاب کنید؛ می‌توانید روی یک ستون قفل یکتایی اعمال کنید\n"
            "یا قفل یکتایی قبلی را از ستون لغو و حذف نمایید."
        )
        desc_lbl.setStyleSheet("font-size: 12px; color: #94A3B8; line-height: 1.5;")
        main_layout.addWidget(desc_lbl)

        # 2. Mode Selector: Apply Lock vs Cancel Lock
        mode_frame = QFrame()
        mode_frame.setStyleSheet("background-color: #0F172A; border: 1px solid #334155; border-radius: 10px; padding: 6px;")
        mode_layout = QHBoxLayout(mode_frame)
        mode_layout.setSpacing(10)
        mode_layout.setContentsMargins(6, 6, 6, 6)

        self.btn_group = QButtonGroup(self)

        self.rb_lock = QRadioButton("🔒 اعمال قفل یکتایی بر ستون")
        self.rb_lock.setChecked(True)
        self.rb_lock.setStyleSheet("""
            QRadioButton {
                background-color: #1E293B; color: #E2E8F0; font-weight: bold;
                padding: 8px 14px; border-radius: 8px; border: 1px solid #475569; font-size: 12px;
            }
            QRadioButton:checked {
                background-color: #14532D; border: 1px solid #22C55E; color: #86EFAC;
            }
        """)

        self.rb_unlock = QRadioButton("🔓 لغو قفل یکتایی ستون")
        self.rb_unlock.setStyleSheet("""
            QRadioButton {
                background-color: #1E293B; color: #E2E8F0; font-weight: bold;
                padding: 8px 14px; border-radius: 8px; border: 1px solid #475569; font-size: 12px;
            }
            QRadioButton:checked {
                background-color: #7F1D1D; border: 1px solid #EF4444; color: #FECACA;
            }
        """)

        self.btn_group.addButton(self.rb_lock)
        self.btn_group.addButton(self.rb_unlock)

        self.rb_lock.toggled.connect(self._on_mode_toggled)

        mode_layout.addWidget(self.rb_lock)
        mode_layout.addWidget(self.rb_unlock)
        main_layout.addWidget(mode_frame)

        # 3. Target Column Section
        self.lbl_col_section = QLabel("📌 ستون مورد نظر برای اعمال قفل یکتایی:")
        self.lbl_col_section.setStyleSheet("font-weight: bold; color: #E2E8F0; font-size: 13px;")
        main_layout.addWidget(self.lbl_col_section)

        col_box = QHBoxLayout()
        self.combo_col = QComboBox()
        self.combo_col.setMinimumWidth(240)
        self.combo_col.setStyleSheet("background-color: #1E293B; color: #FFFFFF; padding: 6px 10px; border: 1px solid #475569; border-radius: 6px;")
        self.combo_col.currentTextChanged.connect(lambda: self._update_preview())

        self.input_col = QLineEdit()
        self.input_col.setPlaceholderText("حرف یا شماره ستون (مثلاً: L یا 12)")
        self.input_col.setStyleSheet("background-color: #1E293B; color: #FFFFFF; padding: 6px 10px; border: 1px solid #475569; border-radius: 6px;")
        self.input_col.textChanged.connect(self._on_manual_col_changed)

        col_box.addWidget(self.combo_col)
        col_box.addWidget(self.input_col)
        main_layout.addLayout(col_box)

        # 3.5 Status Indicator Badge
        self.status_frame = QFrame()
        self.status_frame.setStyleSheet("background-color: #1E293B; border-radius: 8px; padding: 8px 12px;")
        status_layout = QHBoxLayout(self.status_frame)
        status_layout.setContentsMargins(8, 6, 8, 6)
        status_layout.setSpacing(8)

        self.lbl_status_icon = QLabel("ℹ️")
        self.lbl_status_icon.setStyleSheet("font-size: 16px;")
        self.lbl_status_text = QLabel("وضعیت ستون: در حال بررسی...")
        self.lbl_status_text.setStyleSheet("font-size: 12px; font-weight: bold; color: #E2E8F0;")

        status_layout.addWidget(self.lbl_status_icon)
        status_layout.addWidget(self.lbl_status_text, stretch=1)
        main_layout.addWidget(self.status_frame)

        # 4. Error Alert Settings (visible in lock mode)
        self.lbl_alert_section = QLabel("🛑 تنظیمات پیام اخطار مانع‌شونده اکسل:")
        self.lbl_alert_section.setStyleSheet("font-weight: bold; color: #E2E8F0; font-size: 13px;")
        main_layout.addWidget(self.lbl_alert_section)

        self.alert_frame = QFrame()
        self.alert_frame.setStyleSheet("background-color: #1E293B; border-radius: 8px; padding: 10px;")
        alert_layout = QVBoxLayout(self.alert_frame)
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
        main_layout.addWidget(self.alert_frame)

        # 4.5 Unlock Info Frame (visible in unlock mode)
        self.unlock_info_frame = QFrame()
        self.unlock_info_frame.setStyleSheet("background-color: #311B1B; border: 1px solid #7F1D1D; border-radius: 8px; padding: 10px;")
        self.unlock_info_frame.setVisible(False)
        unlock_info_layout = QVBoxLayout(self.unlock_info_frame)
        unlock_info_lbl = QLabel(
            "🔓 **لغو قفل یکتایی:**\n"
            "با اجرای این گزینه، محدودیت اعتبارسنجی داده (Data Validation) از این ستون برداشته می‌شود\n"
            "و کاربران از این پس می‌توانند مقادیر دلخواه (حتی تکراری) را بدون اخطار در ستون ثبت نمایند."
        )
        unlock_info_lbl.setStyleSheet("color: #FECACA; font-size: 12px; line-height: 1.5;")
        unlock_info_layout.addWidget(unlock_info_lbl)
        main_layout.addWidget(self.unlock_info_frame)

        # 5. Checkbox (highlight / clear highlight)
        self.cb_highlight = QCheckBox("بررسی و مدیریت هایلایت رنگی داده‌های تکراری")
        self.cb_highlight.setChecked(True)
        self.cb_highlight.setStyleSheet("color: #E2E8F0; font-size: 12px; font-weight: 500;")
        main_layout.addWidget(self.cb_highlight)

        # 6. Live Preview
        self.preview_frame = QFrame()
        self.preview_frame.setStyleSheet("background-color: #132E20; border: 1px solid #107C41; border-radius: 8px; padding: 10px;")
        preview_layout = QVBoxLayout(self.preview_frame)
        preview_layout.setContentsMargins(10, 8, 10, 8)
        preview_layout.setSpacing(4)

        self.lbl_preview_title = QLabel("💡 پیش‌نمایش فرمول اعتبارسنجی:")
        self.lbl_preview_title.setStyleSheet("font-weight: bold; color: #4ADE80; font-size: 12px;")
        self.lbl_preview_formula = QLabel("")
        self.lbl_preview_formula.setStyleSheet("font-family: Consolas, monospace; font-size: 14px; font-weight: bold; color: #FFFFFF;")
        self.lbl_preview_desc = QLabel("")
        self.lbl_preview_desc.setStyleSheet("font-size: 11px; color: #A7F3D0;")

        preview_layout.addWidget(self.lbl_preview_title)
        preview_layout.addWidget(self.lbl_preview_formula)
        preview_layout.addWidget(self.lbl_preview_desc)
        main_layout.addWidget(self.preview_frame)

        # 7. Action Buttons
        btn_box = QHBoxLayout()
        btn_box.setSpacing(10)

        self.btn_execute = QPushButton("🔒 اعمال قفل یکتایی در اکسل")
        self.btn_execute.setStyleSheet("background-color: #107C41; color: white; font-weight: bold; padding: 10px 20px; border-radius: 8px; font-size: 13px;")
        self.btn_execute.clicked.connect(self._on_execute_clicked)

        self.btn_cancel = QPushButton("❌ بستن")
        self.btn_cancel.setStyleSheet("background-color: #334155; color: #E2E8F0; padding: 10px 18px; border-radius: 8px; font-size: 13px;")
        self.btn_cancel.clicked.connect(self.reject)

        btn_box.addWidget(self.btn_execute, stretch=2)
        btn_box.addWidget(self.btn_cancel, stretch=1)
        main_layout.addLayout(btn_box)

    def _on_mode_toggled(self, is_lock_checked: bool):
        if is_lock_checked:
            self.mode = "lock"
            self.lbl_col_section.setText("📌 ستون مورد نظر برای اعمال قفل یکتایی:")
            self.lbl_alert_section.setVisible(True)
            self.alert_frame.setVisible(True)
            self.unlock_info_frame.setVisible(False)
            self.cb_highlight.setText("بررسی و هایلایت داده‌های تکراری فعلی")
            self.btn_execute.setText("🔒 اعمال قفل یکتایی در اکسل")
            self.btn_execute.setStyleSheet("background-color: #107C41; color: white; font-weight: bold; padding: 10px 20px; border-radius: 8px; font-size: 13px;")
            self.preview_frame.setStyleSheet("background-color: #132E20; border: 1px solid #107C41; border-radius: 8px; padding: 10px;")
            self.lbl_preview_title.setText("💡 پیش‌نمایش فرمول اعتبارسنجی (Data Validation):")
            self.lbl_preview_title.setStyleSheet("font-weight: bold; color: #4ADE80; font-size: 12px;")
            self.lbl_preview_desc.setStyleSheet("font-size: 11px; color: #A7F3D0;")
        else:
            self.mode = "unlock"
            self.lbl_col_section.setText("📌 ستون مورد نظر برای لغو قفل یکتایی:")
            self.lbl_alert_section.setVisible(False)
            self.alert_frame.setVisible(False)
            self.unlock_info_frame.setVisible(True)
            self.cb_highlight.setText("پاک‌سازی رنگ‌های اخطار تکراری از سلول‌های این ستون")
            self.btn_execute.setText("🔓 لغو و حذف قفل یکتایی از ستون")
            self.btn_execute.setStyleSheet("background-color: #DC2626; color: white; font-weight: bold; padding: 10px 20px; border-radius: 8px; font-size: 13px;")
            self.preview_frame.setStyleSheet("background-color: #3B1219; border: 1px solid #DC2626; border-radius: 8px; padding: 10px;")
            self.lbl_preview_title.setText("💡 عملیات لغو قفل ستون:")
            self.lbl_preview_title.setStyleSheet("font-weight: bold; color: #FCA5A5; font-size: 12px;")
            self.lbl_preview_desc.setStyleSheet("font-size: 11px; color: #FECACA;")

        self._update_preview()

    def _load_sheet_columns(self):
        max_col = 15
        headers_map = {}
        locked_cols = set()

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
                    if sheet.has_unique_validation(c_letter):
                        locked_cols.add(c_letter)
            except Exception:
                pass

        self.combo_col.clear()
        selected_idx = 0
        for c_i in range(1, max_col + 1):
            c_letter = col_index_to_letter(c_i)
            hdr = headers_map.get(c_i, "")
            is_locked = c_letter in locked_cols
            lock_badge = " 🔒 [دارای قفل]" if is_locked else ""
            label = f"ستون {c_i} ({c_letter}){lock_badge}" + (f" - «{hdr}»" if hdr else "")
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

        is_locked = False
        if self.manager and self.manager.is_connected():
            try:
                sheet = self.manager.get_active_sheet()
                is_locked = sheet.has_unique_validation(col)
            except Exception:
                is_locked = False

        if self.mode == "lock":
            formula = f"=COUNTIF({col}:{col}, {col}1)<=1"
            self.lbl_preview_formula.setText(formula)

            if is_locked:
                self.lbl_status_icon.setText("🔒")
                self.lbl_status_text.setText(f"وضعیت ستون {col}: قفل یکتایی هم‌اکنون روی این ستون فعال است.")
                self.lbl_status_text.setStyleSheet("font-size: 12px; font-weight: bold; color: #FCA5A5;")
                self.status_frame.setStyleSheet("background-color: #450A0A; border: 1px solid #DC2626; border-radius: 8px; padding: 8px 12px;")
                self.lbl_preview_desc.setText(f"اعمال مجدد، تنظیمات متن اخطار ستون {col} را به‌روزرسانی می‌کند.")
            else:
                self.lbl_status_icon.setText("🔓")
                self.lbl_status_text.setText(f"وضعیت ستون {col}: این ستون آزاد است (بدون قفل یکتایی).")
                self.lbl_status_text.setStyleSheet("font-size: 12px; font-weight: bold; color: #86EFAC;")
                self.status_frame.setStyleSheet("background-color: #0F291E; border: 1px solid #16A34A; border-radius: 8px; padding: 8px 12px;")
                self.lbl_preview_desc.setText(f"با فعال‌سازی، اکسل مانع ورود هرگونه مقدار تکراری در ستون {col} خواهد شد.")
        else:
            # Unlock mode
            self.lbl_preview_formula.setText(f"حذف کامل محدودیت Data Validation از ستون {col}")

            if is_locked:
                self.lbl_status_icon.setText("🔒")
                self.lbl_status_text.setText(f"وضعیت ستون {col}: قفل یکتایی فعال است و با زدن دکمه زیر لغو خواهد شد.")
                self.lbl_status_text.setStyleSheet("font-size: 12px; font-weight: bold; color: #FCA5A5;")
                self.status_frame.setStyleSheet("background-color: #450A0A; border: 1px solid #DC2626; border-radius: 8px; padding: 8px 12px;")
                self.lbl_preview_desc.setText(f"قاعده اعتبارسنجی حذف شده و ورود مقادیر در ستون {col} آزاد می‌شود.")
            else:
                self.lbl_status_icon.setText("ℹ️")
                self.lbl_status_text.setText(f"وضعیت ستون {col}: این ستون در حال حاضر قفل یکتایی ندارد.")
                self.lbl_status_text.setStyleSheet("font-size: 12px; font-weight: bold; color: #94A3B8;")
                self.status_frame.setStyleSheet("background-color: #1E293B; border: 1px solid #475569; border-radius: 8px; padding: 8px 12px;")
                self.lbl_preview_desc.setText(f"ستون {col} همین الان هم آزاد است، اما اجرای این گزینه هرگونه قاعده مرتبط را پاک می‌کند.")

    def _on_execute_clicked(self):
        col = self.get_target_column()
        if self.mode == "unlock":
            confirm = QMessageBox.question(
                self,
                "تأیید لغو قفل یکتایی",
                f"آیا مطمئن هستید که می‌خواهید قفل یکتایی را از ستون {col} لغو کنید؟\n\n"
                f"با لغو قفل، محدودیت اعتبارسنجی داده (Data Validation) از اکسل حذف شده و "
                f"امکان ورود مقادیر تکراری در ستون {col} برقرار خواهد شد.",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes
            )
            if confirm != QMessageBox.Yes:
                return
        self.accept()

    def get_action(self) -> Union[PreventDuplicatesAction, RemoveUniqueValidationAction]:
        col = self.get_target_column()
        if self.mode == "unlock":
            return RemoveUniqueValidationAction(
                target_col=col,
                clear_highlight=self.cb_highlight.isChecked(),
                description_persian=f"برداشتن قفل یکتایی ستون {col}"
            )
        else:
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

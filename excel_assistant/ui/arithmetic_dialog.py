"""
Interactive Arithmetic Wizard Dialog for Persian Excel Assistant.
Allows users to visually select operation (Add, Subtract, Multiply, Divide),
source columns (e.g. 2, 4, 5), and target destination column (e.g. 8 / H).
"""
from typing import List, Optional, Tuple
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QRadioButton, QButtonGroup, QLineEdit,
    QComboBox, QFrame, QScrollArea, QCheckBox, QMessageBox
)

from excel_assistant.core.utils import col_index_to_letter, col_letter_to_index
from excel_assistant.nlp.intent_models import ColumnArithmeticAction
from excel_assistant.core.connection import ExcelConnectionManager


class ArithmeticWizardDialog(QDialog):
    """Modern RTL dialog for 4 basic arithmetic operations on Excel columns."""

    def __init__(self, manager: Optional[ExcelConnectionManager] = None,
                 initial_op: str = "add", parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.manager = manager
        self.initial_op = initial_op
        self.setWindowTitle("🧮 راهنمای ۴ عمل اصلی روی ستون‌ها")
        self.resize(520, 580)
        self.setLayoutDirection(Qt.RightToLeft)

        self._init_ui()
        self._load_sheet_columns()
        self._update_preview()

    def _init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(18, 16, 18, 16)
        main_layout.setSpacing(14)

        # 1. Header
        header_lbl = QLabel("🧮 محاسبه ۴ عمل اصلی (جمع، تفریق، ضرب، تقسیم)")
        header_lbl.setStyleSheet("font-size: 15px; font-weight: bold; color: #38BDF8;")
        main_layout.addWidget(header_lbl)

        desc_lbl = QLabel("نوع عملیات، ستون‌های مبدأ و ستون مقصد را مشخص کنید:")
        desc_lbl.setStyleSheet("font-size: 12px; color: #94A3B8;")
        main_layout.addWidget(desc_lbl)

        # 2. Operation Selection
        op_frame = QFrame()
        op_frame.setStyleSheet("background-color: #1E293B; border-radius: 8px; padding: 6px;")
        op_layout = QHBoxLayout(op_frame)
        op_layout.setSpacing(12)

        self.op_group = QButtonGroup(self)

        self.radio_add = QRadioButton("➕ جمع")
        self.radio_sub = QRadioButton("➖ تفریق")
        self.radio_mul = QRadioButton("✖️ ضرب")
        self.radio_div = QRadioButton("➗ تقسیم")

        for r in (self.radio_add, self.radio_sub, self.radio_mul, self.radio_div):
            r.setStyleSheet("font-weight: bold; font-size: 13px; color: #F1F5F9;")
            op_layout.addWidget(r)

        self.op_group.addButton(self.radio_add, 1)
        self.op_group.addButton(self.radio_sub, 2)
        self.op_group.addButton(self.radio_mul, 3)
        self.op_group.addButton(self.radio_div, 4)

        if self.initial_op == "multiply":
            self.radio_mul.setChecked(True)
        elif self.initial_op == "subtract":
            self.radio_sub.setChecked(True)
        elif self.initial_op == "divide":
            self.radio_div.setChecked(True)
        else:
            self.radio_add.setChecked(True)

        self.op_group.idToggled.connect(lambda: self._update_preview())
        main_layout.addWidget(op_frame)

        # 3. Source Columns Section
        cols_lbl = QLabel("📌 ستون‌های مورد استفاده (ستون‌های مبدأ):")
        cols_lbl.setStyleSheet("font-weight: bold; color: #E2E8F0; font-size: 13px;")
        main_layout.addWidget(cols_lbl)

        # Scroll area with checkboxes for available columns
        self.scroll_area = QScrollArea()
        self.scroll_area.setMaximumHeight(160)
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setStyleSheet("background-color: #0F172A; border: 1px solid #334155; border-radius: 6px;")

        self.cols_container = QWidget()
        self.cols_layout = QVBoxLayout(self.cols_container)
        self.cols_layout.setContentsMargins(10, 8, 10, 8)
        self.cols_layout.setSpacing(6)
        self.scroll_area.setWidget(self.cols_container)
        main_layout.addWidget(self.scroll_area)

        # Manual input for source columns
        manual_box = QHBoxLayout()
        manual_lbl = QLabel("یا وارد کردن دستی:")
        manual_lbl.setStyleSheet("color: #94A3B8; font-size: 12px;")
        self.input_cols = QLineEdit()
        self.input_cols.setPlaceholderText("مثلاً: 2, 4, 5 یا B, D, E")
        self.input_cols.setStyleSheet("background-color: #1E293B; color: #FFFFFF; padding: 6px 10px; border: 1px solid #475569; border-radius: 6px;")
        self.input_cols.textChanged.connect(self._on_manual_cols_changed)
        manual_box.addWidget(manual_lbl)
        manual_box.addWidget(self.input_cols)
        main_layout.addLayout(manual_box)

        # 4. Target Column Section
        target_lbl = QLabel("🎯 ستون مقصد (نتیجه در کدام ستون قرار گیرد؟):")
        target_lbl.setStyleSheet("font-weight: bold; color: #E2E8F0; font-size: 13px;")
        main_layout.addWidget(target_lbl)

        target_box = QHBoxLayout()
        self.combo_target = QComboBox()
        self.combo_target.setMinimumWidth(200)
        self.combo_target.setStyleSheet("background-color: #1E293B; color: #FFFFFF; padding: 6px 10px; border: 1px solid #475569; border-radius: 6px;")
        self.combo_target.currentTextChanged.connect(lambda: self._update_preview())

        self.input_target = QLineEdit()
        self.input_target.setPlaceholderText("شماره یا حرف ستون (مثلاً: 8 یا H)")
        self.input_target.setStyleSheet("background-color: #1E293B; color: #FFFFFF; padding: 6px 10px; border: 1px solid #475569; border-radius: 6px;")
        self.input_target.textChanged.connect(self._on_manual_target_changed)

        target_box.addWidget(self.combo_target)
        target_box.addWidget(self.input_target)
        main_layout.addLayout(target_box)

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

        self.btn_execute = QPushButton("🚀 محاسبه و اعمال در اکسل")
        self.btn_execute.setStyleSheet("background-color: #107C41; color: white; font-weight: bold; padding: 10px 18px; border-radius: 8px; font-size: 13px;")
        self.btn_execute.clicked.connect(self._on_execute_clicked)

        self.btn_cancel = QPushButton("❌ انصراف")
        self.btn_cancel.setStyleSheet("background-color: #334155; color: #E2E8F0; padding: 10px 18px; border-radius: 8px; font-size: 13px;")
        self.btn_cancel.clicked.connect(self.reject)

        btn_box.addWidget(self.btn_execute)
        btn_box.addWidget(self.btn_cancel)
        main_layout.addLayout(btn_box)

        self.checkboxes: List[Tuple[int, str, QCheckBox]] = []
        self._updating_controls = False

    def _load_sheet_columns(self):
        """Loads available sheet columns and headers into checkboxes and dropdowns."""
        self.checkboxes.clear()
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

        # Populate source column checkboxes
        for c_i in range(1, max_col + 1):
            c_letter = col_index_to_letter(c_i)
            hdr = headers_map.get(c_i, "")
            label_text = f"ستون {c_i} ({c_letter})" + (f" - «{hdr}»" if hdr else "")
            cb = QCheckBox(label_text)
            cb.setStyleSheet("color: #E2E8F0; font-size: 12px;")
            cb.toggled.connect(self._on_checkbox_toggled)
            self.cols_layout.addWidget(cb)
            self.checkboxes.append((c_i, c_letter, cb))

        self.cols_layout.addStretch()

        # Populate target column combo
        self.combo_target.clear()
        self.combo_target.addItem("خودکار (اولین ستون خالی بعد از داده‌ها)", "")
        for c_i in range(1, max_col + 5):
            c_letter = col_index_to_letter(c_i)
            hdr = headers_map.get(c_i, "")
            lbl = f"ستون {c_i} ({c_letter})" + (f" - «{hdr}»" if hdr else "")
            self.combo_target.addItem(lbl, c_letter)

        # Pre-check columns 2, 4, 5 if available, or 2 and 3
        if len(self.checkboxes) >= 5:
            self._updating_controls = True
            self.checkboxes[1][2].setChecked(True)  # Col 2
            self.checkboxes[2][2].setChecked(True)  # Col 3
            self._updating_controls = False
            self.input_cols.setText("2, 3")

    def _on_checkbox_toggled(self):
        if self._updating_controls:
            return
        self._updating_controls = True
        checked_cols = [str(c_idx) for c_idx, _, cb in self.checkboxes if cb.isChecked()]
        self.input_cols.setText(", ".join(checked_cols))
        self._updating_controls = False
        self._update_preview()

    def _on_manual_cols_changed(self, text: str):
        if self._updating_controls:
            return
        self._updating_controls = True
        # Parse tokens
        import re
        tokens = [t.strip().upper() for t in re.split(r'[,،\s]+', text) if t.strip()]
        for c_idx, c_let, cb in self.checkboxes:
            is_match = (str(c_idx) in tokens or c_let in tokens)
            cb.setChecked(is_match)
        self._updating_controls = False
        self._update_preview()

    def _on_manual_target_changed(self, text: str):
        val = text.strip().upper()
        if not val:
            return
        # If number
        if val.isdigit():
            c_let = col_index_to_letter(int(val))
        else:
            c_let = val

        for i in range(self.combo_target.count()):
            if self.combo_target.itemData(i) == c_let:
                self.combo_target.setCurrentIndex(i)
                break
        self._update_preview()

    def get_selected_operation(self) -> str:
        if self.radio_mul.isChecked():
            return "multiply"
        elif self.radio_sub.isChecked():
            return "subtract"
        elif self.radio_div.isChecked():
            return "divide"
        return "add"

    def get_selected_columns(self) -> List[str]:
        # Prefer checked checkboxes
        cols = [c_let for _, c_let, cb in self.checkboxes if cb.isChecked()]
        if not cols:
            import re
            tokens = [t.strip().upper() for t in re.split(r'[,،\s]+', self.input_cols.text()) if t.strip()]
            for t in tokens:
                if t.isdigit():
                    cols.append(col_index_to_letter(int(t)))
                elif t.isalpha():
                    cols.append(t)
        return cols

    def get_target_column(self) -> Optional[str]:
        manual = self.input_target.text().strip().upper()
        if manual:
            if manual.isdigit():
                return col_index_to_letter(int(manual))
            return manual
        combo_data = self.combo_target.currentData()
        return combo_data if combo_data else None

    def _update_preview(self):
        op = self.get_selected_operation()
        cols = self.get_selected_columns()
        target = self.get_target_column() or "خودکار"

        op_sym = {"add": "+", "subtract": "-", "multiply": "*", "divide": "/"}.get(op, "+")
        op_fa = {"add": "جمع", "subtract": "تفریق", "multiply": "ضرب", "divide": "تقسیم"}.get(op, "محاسبه")

        if len(cols) >= 2:
            formula = f"={op_sym.join(f'{c}2' for c in cols)}"
            desc = f"عملیات {op_fa} ستون‌های {', '.join(cols)} و درج نتیجه در ستون {target}"
            self.lbl_preview_formula.setText(formula)
            self.lbl_preview_desc.setText(desc)
            self.preview_frame.setStyleSheet("background-color: #132E20; border: 1px solid #107C41; border-radius: 8px; padding: 10px;")
        else:
            self.lbl_preview_formula.setText("⚠️ لطفاً حداقل ۲ ستون برای محاسبه انتخاب کنید")
            self.lbl_preview_desc.setText("مثلاً ستون‌های ۲، ۴ و ۵")
            self.preview_frame.setStyleSheet("background-color: #2D1A1A; border: 1px solid #EF4444; border-radius: 8px; padding: 10px;")

    def _on_execute_clicked(self):
        cols = self.get_selected_columns()
        if len(cols) < 2:
            QMessageBox.warning(self, "انتخاب ستون‌ها", "لطفاً حداقل دو ستون مبدأ برای انجام محاسبه انتخاب نمایید.")
            return
        self.accept()

    def get_action(self) -> ColumnArithmeticAction:
        op = self.get_selected_operation()
        cols = self.get_selected_columns()
        target = self.get_target_column()
        op_fa = {"add": "جمع", "subtract": "تفریق", "multiply": "ضرب", "divide": "تقسیم"}.get(op, "محاسبه")

        return ColumnArithmeticAction(
            operation=op,
            col1=cols[0] if len(cols) > 0 else "A",
            col2=cols[1] if len(cols) > 1 else "B",
            columns=cols,
            target_col=target,
            as_formula=True,
            description_persian=f"{op_fa} ستون‌های {'، '.join(cols)}{' و درج نتیجه در ستون ' + target if target else ''}"
        )

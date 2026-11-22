"""
Interactive Date Insertion Wizard Dialog for Persian Excel Assistant.
Allows users to visually select target column, row range (e.g. 1 to 30),
and date (day, month, year) with fixed or sequential increment options.
"""
import datetime
from typing import List, Optional, Tuple
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QRadioButton, QButtonGroup, QLineEdit,
    QComboBox, QFrame, QSpinBox, QMessageBox
)

from excel_assistant.core.utils import col_index_to_letter, col_letter_to_index
from excel_assistant.nlp.intent_models import InsertDateAction
from excel_assistant.core.connection import ExcelConnectionManager


SHAMSI_MONTH_NAMES = [
    "فروردین", "اردیبهشت", "خرداد", "تیر", "مرداد", "شهریور",
    "مهر", "آبان", "آذر", "دی", "بهمن", "اسفند"
]


def get_current_shamsi_date():
    """Converts Gregorian date to Shamsi (Jalali) date."""
    now = datetime.date.today()
    gy, gm, gd = now.year, now.month, now.day
    g_d_m = [0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334]
    gy2 = (gy - 1600) if (gm > 2) else (gy - 1601)
    days = (365 * gy2) + ((gy2 + 3) // 4) - ((gy2 + 99) // 100) + ((gy2 + 399) // 400) - 80 + gd + g_d_m[gm - 1]
    jy = 979 + (33 * (days // 12053))
    days %= 12053
    jy += 4 * (days // 1461)
    days %= 1461
    if days > 365:
        jy += (days - 1) // 365
        days = (days - 1) % 365
    if days < 186:
        jm = 1 + (days // 31)
        jd = 1 + (days % 31)
    else:
        days -= 186
        jm = 7 + (days // 30)
        jd = 1 + (days % 30)
    return jy, jm, jd


class DateWizardDialog(QDialog):
    """Modern RTL dialog for inserting dates into Excel ranges."""

    def __init__(self, manager: Optional[ExcelConnectionManager] = None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.manager = manager
        self.setWindowTitle("📅 راهنمای درج تاریخ در اکسل")
        self.resize(520, 560)
        self.setLayoutDirection(Qt.RightToLeft)

        self._init_ui()
        self._load_sheet_columns()
        self._set_today_date()
        self._update_preview()

    def _init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(18, 16, 18, 16)
        main_layout.setSpacing(14)

        # 1. Header
        header_lbl = QLabel("📅 درج خودکار تاریخ در ستون و ردیف‌های اکسل")
        header_lbl.setStyleSheet("font-size: 15px; font-weight: bold; color: #38BDF8;")
        main_layout.addWidget(header_lbl)

        desc_lbl = QLabel("ستون، بازه سطرها (مثلاً ۱ تا ۳۰) و روز، ماه و سال مورد نظر را مشخص کنید:")
        desc_lbl.setStyleSheet("font-size: 12px; color: #94A3B8;")
        main_layout.addWidget(desc_lbl)

        # 2. Section 1: Target Column
        col_lbl = QLabel("📌 ستون مقصد (تاریخ در کدام ستون درج شود؟):")
        col_lbl.setStyleSheet("font-weight: bold; color: #E2E8F0; font-size: 13px;")
        main_layout.addWidget(col_lbl)

        col_box = QHBoxLayout()
        self.combo_col = QComboBox()
        self.combo_col.setMinimumWidth(220)
        self.combo_col.setStyleSheet("background-color: #1E293B; color: #FFFFFF; padding: 6px 10px; border: 1px solid #475569; border-radius: 6px;")
        self.combo_col.currentTextChanged.connect(lambda: self._update_preview())

        self.input_col = QLineEdit()
        self.input_col.setPlaceholderText("شماره یا حرف ستون (مثلاً: 1 یا A)")
        self.input_col.setStyleSheet("background-color: #1E293B; color: #FFFFFF; padding: 6px 10px; border: 1px solid #475569; border-radius: 6px;")
        self.input_col.textChanged.connect(self._on_manual_col_changed)

        col_box.addWidget(self.combo_col)
        col_box.addWidget(self.input_col)
        main_layout.addLayout(col_box)

        # 3. Section 2: Row Range (Start and End Row)
        range_lbl = QLabel("🔢 محدوده ردیف‌ها:")
        range_lbl.setStyleSheet("font-weight: bold; color: #E2E8F0; font-size: 13px;")
        main_layout.addWidget(range_lbl)

        range_frame = QFrame()
        range_frame.setStyleSheet("background-color: #1E293B; border-radius: 8px; padding: 8px;")
        range_layout = QHBoxLayout(range_frame)
        range_layout.setSpacing(12)

        lbl_from = QLabel("از سطر:")
        lbl_from.setStyleSheet("color: #E2E8F0; font-size: 12px;")
        self.spin_start_row = QSpinBox()
        self.spin_start_row.setRange(1, 100000)
        self.spin_start_row.setValue(1)
        self.spin_start_row.setStyleSheet("background-color: #0F172A; color: #FFFFFF; padding: 5px 8px; border: 1px solid #475569; border-radius: 6px;")
        self.spin_start_row.valueChanged.connect(lambda: self._update_preview())

        lbl_to = QLabel("تا سطر:")
        lbl_to.setStyleSheet("color: #E2E8F0; font-size: 12px;")
        self.spin_end_row = QSpinBox()
        self.spin_end_row.setRange(1, 100000)
        self.spin_end_row.setValue(30)
        self.spin_end_row.setStyleSheet("background-color: #0F172A; color: #FFFFFF; padding: 5px 8px; border: 1px solid #475569; border-radius: 6px;")
        self.spin_end_row.valueChanged.connect(lambda: self._update_preview())

        range_layout.addWidget(lbl_from)
        range_layout.addWidget(self.spin_start_row)
        range_layout.addWidget(lbl_to)
        range_layout.addWidget(self.spin_end_row)
        range_layout.addStretch()
        main_layout.addWidget(range_frame)

        # 4. Section 3: Date Picker (Day, Month, Year)
        date_lbl = QLabel("📅 انتخاب روز، ماه و سال:")
        date_lbl.setStyleSheet("font-weight: bold; color: #E2E8F0; font-size: 13px;")
        main_layout.addWidget(date_lbl)

        date_frame = QFrame()
        date_frame.setStyleSheet("background-color: #1E293B; border-radius: 8px; padding: 10px;")
        date_layout = QHBoxLayout(date_frame)
        date_layout.setSpacing(10)

        # Day
        lbl_day = QLabel("روز:")
        lbl_day.setStyleSheet("color: #E2E8F0; font-size: 12px;")
        self.spin_day = QSpinBox()
        self.spin_day.setRange(1, 31)
        self.spin_day.setValue(1)
        self.spin_day.setStyleSheet("background-color: #0F172A; color: #FFFFFF; padding: 5px 8px; border: 1px solid #475569; border-radius: 6px;")
        self.spin_day.valueChanged.connect(lambda: self._update_preview())

        # Month
        lbl_month = QLabel("ماه:")
        lbl_month.setStyleSheet("color: #E2E8F0; font-size: 12px;")
        self.combo_month = QComboBox()
        self.combo_month.addItems(SHAMSI_MONTH_NAMES)
        self.combo_month.setStyleSheet("background-color: #0F172A; color: #FFFFFF; padding: 5px 8px; border: 1px solid #475569; border-radius: 6px;")
        self.combo_month.currentIndexChanged.connect(lambda: self._update_preview())

        # Year
        lbl_year = QLabel("سال:")
        lbl_year.setStyleSheet("color: #E2E8F0; font-size: 12px;")
        self.spin_year = QSpinBox()
        self.spin_year.setRange(1380, 1430)
        self.spin_year.setValue(1403)
        self.spin_year.setStyleSheet("background-color: #0F172A; color: #FFFFFF; padding: 5px 8px; border: 1px solid #475569; border-radius: 6px;")
        self.spin_year.valueChanged.connect(lambda: self._update_preview())

        # Quick Today button
        self.btn_today = QPushButton("🗓️ امروز")
        self.btn_today.setStyleSheet("background-color: #0284C7; color: white; padding: 5px 12px; border-radius: 6px; font-weight: bold;")
        self.btn_today.clicked.connect(self._set_today_date)

        date_layout.addWidget(lbl_day)
        date_layout.addWidget(self.spin_day)
        date_layout.addWidget(lbl_month)
        date_layout.addWidget(self.combo_month)
        date_layout.addWidget(lbl_year)
        date_layout.addWidget(self.spin_year)
        date_layout.addWidget(self.btn_today)
        main_layout.addWidget(date_frame)

        # 5. Section 4: Fill Mode (Fixed vs Sequential)
        mode_lbl = QLabel("⚙️ نحوه درج در ردیف‌ها:")
        mode_lbl.setStyleSheet("font-weight: bold; color: #E2E8F0; font-size: 13px;")
        main_layout.addWidget(mode_lbl)

        mode_frame = QFrame()
        mode_frame.setStyleSheet("background-color: #1E293B; border-radius: 8px; padding: 6px;")
        mode_layout = QHBoxLayout(mode_frame)
        mode_layout.setSpacing(14)

        self.radio_fixed = QRadioButton("📌 تاریخ ثابت (یکسان در تمام ردیف‌ها)")
        self.radio_fixed.setChecked(True)
        self.radio_fixed.setStyleSheet("font-weight: bold; color: #F1F5F9; font-size: 12px;")
        self.radio_fixed.toggled.connect(lambda: self._update_preview())

        self.radio_seq = QRadioButton("📈 تاریخ متوالی و روزشمار (افزایش روزانه)")
        self.radio_seq.setStyleSheet("font-weight: bold; color: #F1F5F9; font-size: 12px;")
        self.radio_seq.toggled.connect(lambda: self._update_preview())

        mode_layout.addWidget(self.radio_fixed)
        mode_layout.addWidget(self.radio_seq)
        main_layout.addWidget(mode_frame)

        # 6. Live Preview
        self.preview_frame = QFrame()
        self.preview_frame.setStyleSheet("background-color: #132E20; border: 1px solid #107C41; border-radius: 8px; padding: 10px;")
        preview_layout = QVBoxLayout(self.preview_frame)
        preview_layout.setContentsMargins(10, 8, 10, 8)
        preview_layout.setSpacing(4)

        self.lbl_preview_title = QLabel("💡 پیش‌نمایش درج تاریخ:")
        self.lbl_preview_title.setStyleSheet("font-weight: bold; color: #4ADE80; font-size: 12px;")
        self.lbl_preview_text = QLabel("")
        self.lbl_preview_text.setStyleSheet("font-family: Consolas, monospace; font-size: 14px; font-weight: bold; color: #FFFFFF;")
        self.lbl_preview_desc = QLabel("")
        self.lbl_preview_desc.setStyleSheet("font-size: 11px; color: #A7F3D0;")

        preview_layout.addWidget(self.lbl_preview_title)
        preview_layout.addWidget(self.lbl_preview_text)
        preview_layout.addWidget(self.lbl_preview_desc)
        main_layout.addWidget(self.preview_frame)

        # 7. Action Buttons
        btn_box = QHBoxLayout()
        btn_box.setSpacing(12)

        self.btn_execute = QPushButton("🚀 درج تاریخ در اکسل")
        self.btn_execute.setStyleSheet("background-color: #107C41; color: white; font-weight: bold; padding: 10px 18px; border-radius: 8px; font-size: 13px;")
        self.btn_execute.clicked.connect(self._on_execute_clicked)

        self.btn_cancel = QPushButton("❌ انصراف")
        self.btn_cancel.setStyleSheet("background-color: #334155; color: #E2E8F0; padding: 10px 18px; border-radius: 8px; font-size: 13px;")
        self.btn_cancel.clicked.connect(self.reject)

        btn_box.addWidget(self.btn_execute)
        btn_box.addWidget(self.btn_cancel)
        main_layout.addLayout(btn_box)

    def _set_today_date(self):
        jy, jm, jd = get_current_shamsi_date()
        self.spin_year.setValue(jy)
        self.combo_month.setCurrentIndex(max(0, min(jm - 1, 11)))
        self.spin_day.setValue(jd)
        self._update_preview()

    def _load_sheet_columns(self):
        max_col = 10
        headers_map = {}

        if self.manager and self.manager.is_connected():
            try:
                sheet = self.manager.get_active_sheet()
                _, _, max_r, max_c = sheet.get_used_range_bounds()
                max_col = max(10, min(max_c + 4, 30))
                if max_r > 1:
                    self.spin_end_row.setValue(max_r)
                for c_i in range(1, max_col + 1):
                    c_letter = col_index_to_letter(c_i)
                    val = sheet.get_value(f"{c_letter}1")
                    if val is not None and str(val).strip():
                        headers_map[c_i] = str(val).strip()
            except Exception:
                pass

        self.combo_col.clear()
        date_col_idx = 0
        for c_i in range(1, max_col + 1):
            c_letter = col_index_to_letter(c_i)
            hdr = headers_map.get(c_i, "")
            label = f"ستون {c_i} ({c_letter})" + (f" - «{hdr}»" if hdr else "")
            self.combo_col.addItem(label, c_letter)
            if "تاریخ" in hdr.lower() or "date" in hdr.lower():
                date_col_idx = c_i - 1

        self.combo_col.setCurrentIndex(date_col_idx)

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
        start_r = self.spin_start_row.value()
        end_r = max(start_r, self.spin_end_row.value())
        count = end_r - start_r + 1

        y = self.spin_year.value()
        m = self.combo_month.currentIndex() + 1
        d = self.spin_day.value()

        date_sample = f"{y:04d}/{m:02d}/{d:02d}"
        range_sample = f"{col}{start_r}:{col}{end_r}"

        mode_str = "متوالی و افزایشی" if self.radio_seq.isChecked() else "ثابت"
        self.lbl_preview_text.setText(f"{date_sample}  ({mode_str})")
        self.lbl_preview_desc.setText(f"درج در محدوده {range_sample} (مجموعاً {count} سلول)")

    def _on_execute_clicked(self):
        if self.spin_start_row.value() > self.spin_end_row.value():
            QMessageBox.warning(self, "خطای محدوده", "ردیف شروع نمی‌تواند بزرگتر از ردیف پایان باشد.")
            return
        self.accept()

    def get_action(self) -> InsertDateAction:
        col = self.get_target_column()
        start_r = self.spin_start_row.value()
        end_r = self.spin_end_row.value()
        y = self.spin_year.value()
        m = self.combo_month.currentIndex() + 1
        d = self.spin_day.value()
        mode = "sequential" if self.radio_seq.isChecked() else "fixed"

        return InsertDateAction(
            target_col=col,
            start_row=start_r,
            end_row=end_r,
            year=y,
            month=m,
            day=d,
            mode=mode,
            description_persian=f"درج تاریخ {y:04d}/{m:02d}/{d:02d} در ستون {col} از سطر {start_r} تا {end_r}"
        )

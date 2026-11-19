"""
Header and Status Bar Widget for Excel Assistant.
Displays live Excel 2024 connection status, workbook name, and active sheet selector.
"""
from typing import List, Optional
from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QWidget, QHBoxLayout, QLabel, QComboBox, QPushButton, QFrame
)


class ExcelStatusBar(QFrame):
    """Top bar displaying Excel connection status and active workbook/sheet."""

    reconnect_requested = Signal()
    sheet_switched = Signal(str)

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setObjectName("headerPanel")
        self.setLayoutDirection(Qt.RightToLeft)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(12, 8, 12, 8)
        layout.setSpacing(16)

        # Connection status indicator
        self.status_icon = QLabel("🟢")
        self.status_text = QLabel("در حال اتصال...")
        self.status_text.setStyleSheet("font-weight: bold; color: #10B981;")
        layout.addWidget(self.status_icon)
        layout.addWidget(self.status_text)

        # Divider
        divider1 = QLabel("|")
        divider1.setStyleSheet("color: #475569;")
        layout.addWidget(divider1)

        # Workbook name label
        self.wb_label = QLabel("📁 فایل: -")
        self.wb_label.setStyleSheet("color: #E2E8F0; font-weight: 500;")
        layout.addWidget(self.wb_label)

        # Active Sheet selector
        sheet_lbl = QLabel("📑 شیت:")
        sheet_lbl.setStyleSheet("color: #94A3B8;")
        layout.addWidget(sheet_lbl)

        self.sheet_combo = QComboBox()
        self.sheet_combo.setMinimumWidth(120)
        self.sheet_combo.currentTextChanged.connect(self._on_sheet_combo_changed)
        layout.addWidget(self.sheet_combo)

        layout.addStretch()

        # Reconnect button
        self.btn_reconnect = QPushButton("🔄 اتصال مجدد")
        self.btn_reconnect.clicked.connect(self.reconnect_requested.emit)
        layout.addWidget(self.btn_reconnect)

    def set_connected(self, is_connected: bool, app_desc: str = "Microsoft Excel 2024"):
        if is_connected:
            self.status_icon.setText("🟢")
            self.status_text.setText(f"متصل به {app_desc}")
            self.status_text.setStyleSheet("font-weight: bold; color: #10B981;")
        else:
            self.status_icon.setText("🔴")
            self.status_text.setText("اتصال به اکسل برقرار نیست")
            self.status_text.setStyleSheet("font-weight: bold; color: #EF4444;")

    def update_workbook_info(self, wb_name: str, sheets: List[str], active_sheet: str):
        self.wb_label.setText(f"📁 فایل: {wb_name}")
        self.sheet_combo.blockSignals(True)
        self.sheet_combo.clear()
        self.sheet_combo.addItems(sheets)
        if active_sheet in sheets:
            self.sheet_combo.setCurrentText(active_sheet)
        self.sheet_combo.blockSignals(False)

    def _on_sheet_combo_changed(self, text: str):
        if text:
            self.sheet_switched.emit(text)

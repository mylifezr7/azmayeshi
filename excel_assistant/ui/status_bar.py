"""
Header and Status Bar Widget for Excel Assistant.
Displays live Excel 2024 connection status, workbook selector, file browser, and active sheet selector.
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
    workbook_switched = Signal(str)
    open_file_requested = Signal()

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setObjectName("headerPanel")
        self.setLayoutDirection(Qt.RightToLeft)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(12, 8, 12, 8)
        layout.setSpacing(12)

        # Connection status indicator
        self.status_icon = QLabel("🟢")
        self.status_text = QLabel("در حال اتصال...")
        self.status_text.setStyleSheet("font-weight: bold; color: #10B981; border: none;")
        layout.addWidget(self.status_icon)
        layout.addWidget(self.status_text)

        # Divider
        divider1 = QLabel("|")
        divider1.setStyleSheet("color: #475569; border: none;")
        layout.addWidget(divider1)

        # Open File Button
        self.btn_open_file = QPushButton("📂 باز کردن فایل اکسل...")
        self.btn_open_file.setStyleSheet("background-color: #2563EB; color: white; font-weight: bold; padding: 5px 12px; border-radius: 6px;")
        self.btn_open_file.clicked.connect(self.open_file_requested.emit)
        layout.addWidget(self.btn_open_file)

        # Workbook selector dropdown
        wb_lbl = QLabel("📁 فایل:")
        wb_lbl.setStyleSheet("color: #94A3B8; border: none;")
        layout.addWidget(wb_lbl)

        self.wb_combo = QComboBox()
        self.wb_combo.setMinimumWidth(140)
        self.wb_combo.currentTextChanged.connect(self._on_wb_combo_changed)
        layout.addWidget(self.wb_combo)

        # Active Sheet selector
        sheet_lbl = QLabel("📑 شیت:")
        sheet_lbl.setStyleSheet("color: #94A3B8; border: none;")
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
            self.status_text.setStyleSheet("font-weight: bold; color: #10B981; border: none;")
        else:
            self.status_icon.setText("🔴")
            self.status_text.setText("اتصال به اکسل برقرار نیست")
            self.status_text.setStyleSheet("font-weight: bold; color: #EF4444; border: none;")

    def update_workbook_info(self, wb_name: str, sheets: List[str], active_sheet: str, open_workbooks: Optional[List[str]] = None):
        # Update workbooks dropdown
        wbs = open_workbooks or [wb_name]
        if wb_name not in wbs:
            wbs.append(wb_name)
        self.wb_combo.blockSignals(True)
        self.wb_combo.clear()
        self.wb_combo.addItems(wbs)
        self.wb_combo.setCurrentText(wb_name)
        self.wb_combo.blockSignals(False)

        # Update sheets dropdown
        self.sheet_combo.blockSignals(True)
        self.sheet_combo.clear()
        self.sheet_combo.addItems(sheets)
        if active_sheet in sheets:
            self.sheet_combo.setCurrentText(active_sheet)
        self.sheet_combo.blockSignals(False)

    def _on_sheet_combo_changed(self, text: str):
        if text:
            self.sheet_switched.emit(text)

    def _on_wb_combo_changed(self, text: str):
        if text:
            self.workbook_switched.emit(text)

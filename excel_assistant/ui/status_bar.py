"""
Header and Status Bar Widget for Excel Assistant.
Displays user badge, live Excel 2024 connection status, workbook selector, file browser, and active sheet selector.
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
        layout.setContentsMargins(14, 8, 14, 8)
        layout.setSpacing(12)

        # 1. User Profile Badge (Rightmost in RTL)
        self.user_badge = QLabel("M  MIYUM")
        self.user_badge.setObjectName("userBadge")
        self.user_badge.setStyleSheet(
            "background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #08132B, stop:1 #0E224F); "
            "color: #38BDF8; font-weight: bold; font-size: 13px; "
            "padding: 5px 14px; border: 1.5px solid #2563EB; border-radius: 14px;"
        )
        layout.addWidget(self.user_badge)

        # 2. Connection Status Pill
        self.status_container = QFrame()
        self.status_container.setObjectName("statusContainer")
        self.status_container.setStyleSheet(
            "background-color: #081024; border: 1px solid #1E293B; border-radius: 14px; padding: 2px 10px;"
        )
        pill_layout = QHBoxLayout(self.status_container)
        pill_layout.setContentsMargins(6, 2, 6, 2)
        pill_layout.setSpacing(6)

        self.status_text = QLabel("در حال اتصال به اکسل...")
        self.status_text.setStyleSheet("font-weight: bold; color: #10B981; border: none; font-size: 12px;")
        pill_layout.addWidget(self.status_text)

        self.status_icon = QLabel("🟢")
        self.status_icon.setStyleSheet("border: none; font-size: 11px;")
        pill_layout.addWidget(self.status_icon)

        layout.addWidget(self.status_container)

        # 3. Open File Button (glowing blue gradient)
        self.btn_open_file = QPushButton("📂 باز کردن فایل اکسل...")
        self.btn_open_file.setObjectName("openFileButton")
        self.btn_open_file.setStyleSheet(
            "QPushButton#openFileButton {"
            "  background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #1D4ED8, stop:1 #3B82F6);"
            "  border: 1px solid #60A5FA; border-radius: 10px; color: #FFFFFF;"
            "  font-weight: bold; padding: 6px 14px; font-size: 13px;"
            "}"
            "QPushButton#openFileButton:hover {"
            "  background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #2563EB, stop:1 #60A5FA);"
            "  border-color: #93C5FD;"
            "}"
            "QPushButton#openFileButton:pressed { background: #1E40AF; }"
        )
        self.btn_open_file.clicked.connect(self.open_file_requested.emit)
        layout.addWidget(self.btn_open_file)

        # 4. Workbook selector dropdown
        wb_lbl = QLabel("📁 فایل:")
        wb_lbl.setStyleSheet("color: #94A3B8; border: none; font-size: 12px;")
        layout.addWidget(wb_lbl)

        self.wb_combo = QComboBox()
        self.wb_combo.setMinimumWidth(140)
        self.wb_combo.setStyleSheet(
            "background-color: #081024; border: 1px solid #1E3A8A; border-radius: 8px; padding: 5px 10px; color: #E2E8F0;"
        )
        self.wb_combo.currentTextChanged.connect(self._on_wb_combo_changed)
        layout.addWidget(self.wb_combo)

        # 5. Active Sheet selector dropdown
        sheet_lbl = QLabel("📑 شیت:")
        sheet_lbl.setStyleSheet("color: #94A3B8; border: none; font-size: 12px;")
        layout.addWidget(sheet_lbl)

        self.sheet_combo = QComboBox()
        self.sheet_combo.setMinimumWidth(120)
        self.sheet_combo.setStyleSheet(
            "background-color: #081024; border: 1px solid #1E3A8A; border-radius: 8px; padding: 5px 10px; color: #E2E8F0;"
        )
        self.sheet_combo.currentTextChanged.connect(self._on_sheet_combo_changed)
        layout.addWidget(self.sheet_combo)

        layout.addStretch()

        # 6. Reconnect button
        self.btn_reconnect = QPushButton("📶 اتصال مجدد")
        self.btn_reconnect.setObjectName("reconnectButton")
        self.btn_reconnect.setStyleSheet(
            "QPushButton#reconnectButton {"
            "  background-color: #081024; border: 1px solid #1E3A8A; border-radius: 8px;"
            "  color: #93C5FD; font-weight: bold; padding: 6px 12px; font-size: 12px;"
            "}"
            "QPushButton#reconnectButton:hover {"
            "  background-color: #1E3A8A; color: #FFFFFF;"
            "}"
        )
        self.btn_reconnect.clicked.connect(self.reconnect_requested.emit)
        layout.addWidget(self.btn_reconnect)

    def set_connected(self, is_connected: bool, app_desc: str = "Microsoft Excel 2024 (COM)"):
        if is_connected:
            self.status_icon.setText("🟢")
            self.status_text.setText(f"{app_desc}")
            self.status_text.setStyleSheet("font-weight: bold; color: #10B981; border: none; font-size: 12px;")
            self.status_container.setStyleSheet(
                "background-color: #081024; border: 1px solid #1E293B; border-radius: 14px; padding: 2px 10px;"
            )
        else:
            self.status_icon.setText("🔴")
            self.status_text.setText("عدم اتصال به اکسل")
            self.status_text.setStyleSheet("font-weight: bold; color: #EF4444; border: none; font-size: 12px;")
            self.status_container.setStyleSheet(
                "background-color: #180808; border: 1px solid #7F1D1D; border-radius: 14px; padding: 2px 10px;"
            )

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

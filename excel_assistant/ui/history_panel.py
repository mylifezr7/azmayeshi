"""
Command History and Quick Actions Sidebar Panel.
Features vibrant distinct primary gradient buttons and modern dark glassmorphism.
"""
from typing import List, Optional
from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QListWidget, QPushButton, QFrame
)


class HistoryPanel(QFrame):
    """Sidebar for command history and quick actions."""

    command_selected = Signal(str)
    save_requested = Signal()
    undo_requested = Signal()
    analyze_requested = Signal()
    arithmetic_requested = Signal()
    discount_requested = Signal()
    date_requested = Signal()
    unique_col_requested = Signal()
    sales_requested = Signal()
    profit_requested = Signal()

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setObjectName("sidebarPanel")
        self.setMinimumWidth(230)
        self.setMaximumWidth(290)
        self.setLayoutDirection(Qt.RightToLeft)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        # Title
        title_lbl = QLabel("📜 تاریخچه دستورات")
        title_lbl.setStyleSheet("font-weight: bold; color: #CBD5E1; font-size: 13px; margin-bottom: 2px;")
        layout.addWidget(title_lbl)

        # History list
        self.list_widget = QListWidget()
        self.list_widget.setObjectName("historyList")
        self.list_widget.itemDoubleClicked.connect(self._on_item_double_clicked)
        layout.addWidget(self.list_widget)

        # 0. Save button: Vibrant Teal / Cyan-Emerald
        self.btn_save = QPushButton("💾 ذخیره فایل (Save)  [Ctrl+S]")
        self.btn_save.setObjectName("saveButton")
        self.btn_save.setProperty("class", "sidebarActionBtn")
        self.btn_save.setToolTip("ذخیره تغییرات در فایل اکسل: Ctrl+S")
        self.btn_save.setStyleSheet(
            "QPushButton#saveButton {"
            "  background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #0F766E, stop:1 #14B8A6);"
            "  color: #FFFFFF; font-weight: bold; font-size: 13px; padding: 10px 14px;"
            "  border: 1px solid #2DD4BF; border-radius: 10px;"
            "}"
            "QPushButton#saveButton:hover {"
            "  background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #115E59, stop:1 #2DD4BF);"
            "  border: 1px solid #5EEAD4;"
            "}"
            "QPushButton#saveButton:pressed { background: #134E4A; }"
        )
        self.btn_save.clicked.connect(self.save_requested.emit)
        layout.addWidget(self.btn_save)

        # 1. Undo button: Warm Vibrant Orange / Amber
        self.btn_undo = QPushButton("↩️ لغو آخرین تغییر (Undo)")
        self.btn_undo.setObjectName("undoButton")
        self.btn_undo.setProperty("class", "sidebarActionBtn")
        self.btn_undo.setStyleSheet(
            "QPushButton#undoButton {"
            "  background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #EA580C, stop:1 #FB923C);"
            "  color: #FFFFFF; font-weight: bold; font-size: 13px; padding: 10px 14px;"
            "  border: 1px solid #F97316; border-radius: 10px;"
            "}"
            "QPushButton#undoButton:hover {"
            "  background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #F97316, stop:1 #FDBA74);"
            "  border: 1px solid #FDBA74;"
            "}"
            "QPushButton#undoButton:pressed { background: #C2410C; }"
        )
        self.btn_undo.clicked.connect(self.undo_requested.emit)
        layout.addWidget(self.btn_undo)

        # 2. Arithmetic operations button: Rich Vibrant Emerald Green
        self.btn_arithmetic = QPushButton("🧮 ۴ عمل اصلی ستون‌ها...")
        self.btn_arithmetic.setObjectName("arithmeticButton")
        self.btn_arithmetic.setProperty("class", "sidebarActionBtn")
        self.btn_arithmetic.setStyleSheet(
            "QPushButton#arithmeticButton {"
            "  background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #047857, stop:1 #10B981);"
            "  color: #FFFFFF; font-weight: bold; font-size: 13px; padding: 10px 14px;"
            "  border: 1px solid #10B981; border-radius: 10px;"
            "}"
            "QPushButton#arithmeticButton:hover {"
            "  background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #059669, stop:1 #34D399);"
            "  border: 1px solid #34D399;"
            "}"
            "QPushButton#arithmeticButton:pressed { background: #064E3B; }"
        )
        self.btn_arithmetic.clicked.connect(self.arithmetic_requested.emit)
        layout.addWidget(self.btn_arithmetic)

        # 3. Product discount button: Warm Vibrant Mango / Amber Orange
        self.btn_discount = QPushButton("🏷️ محاسبه تخفیف کالاها...")
        self.btn_discount.setObjectName("discountButton")
        self.btn_discount.setProperty("class", "sidebarActionBtn")
        self.btn_discount.setStyleSheet(
            "QPushButton#discountButton {"
            "  background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #B45309, stop:1 #F59E0B);"
            "  color: #FFFFFF; font-weight: bold; font-size: 13px; padding: 10px 14px;"
            "  border: 1px solid #F59E0B; border-radius: 10px;"
            "}"
            "QPushButton#discountButton:hover {"
            "  background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #D97706, stop:1 #FBBF24);"
            "  border: 1px solid #FBBF24;"
            "}"
            "QPushButton#discountButton:pressed { background: #78350F; }"
        )
        self.btn_discount.clicked.connect(self.discount_requested.emit)
        layout.addWidget(self.btn_discount)

        layout.addStretch()

    def add_command(self, cmd: str):
        if cmd.strip():
            self.list_widget.insertItem(0, cmd.strip())

    def _on_item_double_clicked(self, item):
        self.command_selected.emit(item.text())

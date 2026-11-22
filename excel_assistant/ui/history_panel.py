"""
Command History and Quick Actions Sidebar Panel.
"""
from typing import List, Optional
from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QListWidget, QPushButton, QFrame, QHBoxLayout
)


class HistoryPanel(QFrame):
    """Sidebar for command history and quick actions."""

    command_selected = Signal(str)
    undo_requested = Signal()
    analyze_requested = Signal()
    arithmetic_requested = Signal()
    discount_requested = Signal()
    date_requested = Signal()

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setMinimumWidth(220)
        self.setMaximumWidth(280)
        self.setLayoutDirection(Qt.RightToLeft)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        # Title
        title_lbl = QLabel("📜 تاریخچه دستورات")
        title_lbl.setStyleSheet("font-weight: bold; color: #94A3B8; font-size: 13px;")
        layout.addWidget(title_lbl)

        # History list
        self.list_widget = QListWidget()
        self.list_widget.setObjectName("historyList")
        self.list_widget.itemDoubleClicked.connect(self._on_item_double_clicked)
        layout.addWidget(self.list_widget)

        # Undo button
        self.btn_undo = QPushButton("↩️ لغو آخرین تغییر (Undo)")
        self.btn_undo.setObjectName("undoButton")
        self.btn_undo.clicked.connect(self.undo_requested.emit)
        layout.addWidget(self.btn_undo)

        # Arithmetic button
        self.btn_arithmetic = QPushButton("🧮 ۴ عمل اصلی ستون‌ها...")
        self.btn_arithmetic.setStyleSheet("background-color: #107C41; color: white; font-weight: bold;")
        self.btn_arithmetic.clicked.connect(self.arithmetic_requested.emit)
        layout.addWidget(self.btn_arithmetic)

        # Discount button
        self.btn_discount = QPushButton("🏷️ محاسبه تخفیف کالاها...")
        self.btn_discount.setStyleSheet("background-color: #D97706; color: white; font-weight: bold;")
        self.btn_discount.clicked.connect(self.discount_requested.emit)
        layout.addWidget(self.btn_discount)

        # Date button
        self.btn_date = QPushButton("📅 درج تاریخ در ستون...")
        self.btn_date.setStyleSheet("background-color: #0284C7; color: white; font-weight: bold;")
        self.btn_date.clicked.connect(self.date_requested.emit)
        layout.addWidget(self.btn_date)

        # Analyze button
        self.btn_analyze = QPushButton("📊 تحلیل هوشمند شیت")
        self.btn_analyze.setStyleSheet("background-color: #475569; color: white;")
        self.btn_analyze.clicked.connect(self.analyze_requested.emit)
        layout.addWidget(self.btn_analyze)

    def add_command(self, cmd: str):
        if cmd.strip():
            self.list_widget.insertItem(0, cmd.strip())

    def _on_item_double_clicked(self, item):
        self.command_selected.emit(item.text())

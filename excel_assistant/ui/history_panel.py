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

        # 2. Unique column lock button: Vibrant Royal Purple / Violet
        self.btn_unique = QPushButton("🔒 قفل یکتایی ستون اکسل...")
        self.btn_unique.setObjectName("uniqueColButton")
        self.btn_unique.setProperty("class", "sidebarActionBtn")
        self.btn_unique.setStyleSheet(
            "QPushButton#uniqueColButton {"
            "  background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #6D28D9, stop:1 #8B5CF6);"
            "  color: #FFFFFF; font-weight: bold; font-size: 13px; padding: 10px 14px;"
            "  border: 1px solid #8B5CF6; border-radius: 10px;"
            "}"
            "QPushButton#uniqueColButton:hover {"
            "  background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #7C3AED, stop:1 #A78BFA);"
            "  border: 1px solid #A78BFA;"
            "}"
            "QPushButton#uniqueColButton:pressed { background: #5B21B6; }"
        )
        self.btn_unique.clicked.connect(self.unique_col_requested.emit)
        layout.addWidget(self.btn_unique)

        # 3. Arithmetic operations button: Rich Vibrant Emerald Green
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

        # 4. Product discount button: Warm Vibrant Mango / Amber Orange
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

        # 5. Insert Date button: Vibrant Sky Blue / Azure
        self.btn_date = QPushButton("📅 درج تاریخ در شیت...")
        self.btn_date.setObjectName("dateButton")
        self.btn_date.setProperty("class", "sidebarActionBtn")
        self.btn_date.setStyleSheet(
            "QPushButton#dateButton {"
            "  background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #0284C7, stop:1 #38BDF8);"
            "  color: #FFFFFF; font-weight: bold; font-size: 13px; padding: 10px 14px;"
            "  border: 1px solid #38BDF8; border-radius: 10px;"
            "}"
            "QPushButton#dateButton:hover {"
            "  background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #0369A1, stop:1 #7DD3FC);"
            "  border: 1px solid #7DD3FC;"
            "}"
            "QPushButton#dateButton:pressed { background: #075985; }"
        )
        self.btn_date.clicked.connect(self.date_requested.emit)
        layout.addWidget(self.btn_date)

        # 6. Smart Analysis button: Deep Sapphire / Royal Blue
        self.btn_analyze = QPushButton("📊 تحلیل هوشمند شیت")
        self.btn_analyze.setObjectName("analyzeButton")
        self.btn_analyze.setProperty("class", "sidebarActionBtn")
        self.btn_analyze.setStyleSheet(
            "QPushButton#analyzeButton {"
            "  background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #1E3A8A, stop:1 #2563EB);"
            "  color: #FFFFFF; font-weight: bold; font-size: 13px; padding: 10px 14px;"
            "  border: 1px solid #3B82F6; border-radius: 10px;"
            "}"
            "QPushButton#analyzeButton:hover {"
            "  background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #1D4ED8, stop:1 #60A5FA);"
            "  border: 1px solid #60A5FA;"
            "}"
            "QPushButton#analyzeButton:pressed { background: #172554; }"
        )
        self.btn_analyze.clicked.connect(self.analyze_requested.emit)
        layout.addWidget(self.btn_analyze)

        # 7. Sales Calculation & Dashboard button: Radiant Ruby / Rose Coral
        self.btn_sales = QPushButton("💰 محاسبه فروش (روزانه/ماهانه/سالانه/کل)...  [Ctrl+Shift+S]")
        self.btn_sales.setObjectName("salesButton")
        self.btn_sales.setProperty("class", "sidebarActionBtn")
        self.btn_sales.setToolTip("کلید میانبر پنجره تعاملی: Ctrl+Shift+S")
        self.btn_sales.setStyleSheet(
            "QPushButton#salesButton {"
            "  background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #BE123C, stop:1 #F43F5E);"
            "  color: #FFFFFF; font-weight: bold; font-size: 13px; padding: 10px 14px;"
            "  border: 1px solid #F43F5E; border-radius: 10px;"
            "}"
            "QPushButton#salesButton:hover {"
            "  background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #E11D48, stop:1 #FB7185);"
            "  border: 1px solid #FB7185;"
            "}"
            "QPushButton#salesButton:pressed { background: #9F1239; }"
        )
        self.btn_sales.clicked.connect(self.sales_requested.emit)
        layout.addWidget(self.btn_sales)

        # 8. Profit Calculation button: Vibrant Emerald / Bright Green
        self.btn_profit = QPushButton("💵 محاسبه سود (روزانه/ماهانه/سالانه/کل)...  [Ctrl+P]")
        self.btn_profit.setObjectName("profitButton")
        self.btn_profit.setProperty("class", "sidebarActionBtn")
        self.btn_profit.setToolTip("کلید میانبر پنجره تعاملی: Ctrl+P")
        self.btn_profit.setStyleSheet(
            "QPushButton#profitButton {"
            "  background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #059669, stop:1 #10B981);"
            "  color: #FFFFFF; font-weight: bold; font-size: 13px; padding: 10px 14px;"
            "  border: 1px solid #10B981; border-radius: 10px;"
            "}"
            "QPushButton#profitButton:hover {"
            "  background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #10B981, stop:1 #34D399);"
            "  border: 1px solid #34D399;"
            "}"
            "QPushButton#profitButton:pressed { background: #064E3B; }"
        )
        self.btn_profit.clicked.connect(self.profit_requested.emit)
        layout.addWidget(self.btn_profit)

    def add_command(self, cmd: str):
        if cmd.strip():
            self.list_widget.insertItem(0, cmd.strip())

    def _on_item_double_clicked(self, item):
        self.command_selected.emit(item.text())

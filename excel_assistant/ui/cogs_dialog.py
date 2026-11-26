"""
Cost of Goods Sold (COGS) Valuation Method Selector Dialog for MUYUM Excel Assistant.
Allows users to visually choose between FIFO, LIFO, Weighted Average, Comparison, or Periodic Statement.
"""
from typing import Optional
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QFrame
)


class CogsMethodSelectorDialog(QDialog):
    """Fast, modern RTL dialog for selecting FIFO, LIFO, Weighted Average, or Periodic COGS method."""

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle("🏷️ انتخاب روش محاسبه بهای تمام شده (COGS)")
        self.resize(540, 480)
        self.setMinimumSize(480, 420)
        self.setLayoutDirection(Qt.RightToLeft)
        self.selected_method = "fifo"

        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 16)
        layout.setSpacing(12)

        title = QLabel("🏷️ انتخاب روش محاسبه بهای تمام شده کالای فروش رفته (COGS)")
        title.setStyleSheet("font-size: 15px; font-weight: bold; color: #10B981;")
        layout.addWidget(title)

        desc = QLabel(
            "سیستم به صورت خودکار مقادیر خرید، کرایه، برگشت و موجودی را تشخیص می‌دهد.\n"
            "لطفاً روش مورد نظر برای بهای تمام شده را انتخاب کنید (لایفو، فایفو، میانگین یا فرمول جامع):"
        )
        desc.setStyleSheet("font-size: 12px; color: #CBD5E1; line-height: 1.5;")
        desc.setWordWrap(True)
        layout.addWidget(desc)

        # 1. FIFO button
        self.btn_fifo = QPushButton(
            "🟢 ۱. روش فایفو (FIFO) - اولین وارده، اولین صادره\n"
            "   (استاندارد رسمی حسابداری - خریدهای قبلی ابتدا منظور می‌شوند و بهای تمام شده کمتر است)"
        )
        self.btn_fifo.setStyleSheet(
            "QPushButton {"
            "  background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #064E3B, stop:1 #047857);"
            "  color: #FFFFFF; font-weight: bold; font-size: 12.5px; padding: 12px 16px; border-radius: 9px;"
            "  border: 1px solid #10B981; text-align: right;"
            "}"
            "QPushButton:hover { background: #059669; border-color: #34D399; }"
        )
        self.btn_fifo.clicked.connect(lambda: self._choose("fifo"))
        layout.addWidget(self.btn_fifo)

        # 2. LIFO button
        self.btn_lifo = QPushButton(
            "🔵 ۲. روش لایفو (LIFO) - آخرین وارده، اولین صادره\n"
            "   (مناسب شرایط تورمی - خریدهای جدیدتر و گران‌تر منظور شده و موجب صرفه‌جویی مالیاتی می‌شود)"
        )
        self.btn_lifo.setStyleSheet(
            "QPushButton {"
            "  background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #1E3A8A, stop:1 #1D4ED8);"
            "  color: #FFFFFF; font-weight: bold; font-size: 12.5px; padding: 12px 16px; border-radius: 9px;"
            "  border: 1px solid #3B82F6; text-align: right;"
            "}"
            "QPushButton:hover { background: #2563EB; border-color: #60A5FA; }"
        )
        self.btn_lifo.clicked.connect(lambda: self._choose("lifo"))
        layout.addWidget(self.btn_lifo)

        # 3. Weighted Average button
        self.btn_avg = QPushButton(
            "🟣 ۳. روش میانگین موزون (Weighted Average)\n"
            "   (محاسبه میانگین موزون بهای تمام شده و تعدیل نوسانات قیمت خرید)"
        )
        self.btn_avg.setStyleSheet(
            "QPushButton {"
            "  background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #581C87, stop:1 #7E22CE);"
            "  color: #FFFFFF; font-weight: bold; font-size: 12.5px; padding: 12px 16px; border-radius: 9px;"
            "  border: 1px solid #A855F7; text-align: right;"
            "}"
            "QPushButton:hover { background: #9333EA; border-color: #C084FC; }"
        )
        self.btn_avg.clicked.connect(lambda: self._choose("average"))
        layout.addWidget(self.btn_avg)

        # 4. Compare button
        self.btn_compare = QPushButton(
            "⚖️ ۴. مقایسه تطبیقی هر سه روش (فایفو vs لایفو vs میانگین موزون)\n"
            "   (تحلیل اثر تورم بر بهای تمام‌شده، سود ناخالص، صرفه‌جویی مالیاتی و موجودی پایان دوره)"
        )
        self.btn_compare.setStyleSheet(
            "QPushButton {"
            "  background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #0F766E, stop:1 #0284C7);"
            "  color: #FFFFFF; font-weight: bold; font-size: 12.5px; padding: 12px 16px; border-radius: 9px;"
            "  border: 1px solid #38BDF8; text-align: right;"
            "}"
            "QPushButton:hover { background: #0369A1; border-color: #7DD3FC; }"
        )
        self.btn_compare.clicked.connect(lambda: self._choose("compare"))
        layout.addWidget(self.btn_compare)

        # 5. Periodic Standard Statement button
        self.btn_periodic = QPushButton(
            "🏢 ۵. صورت استاندارد ادواری حسابداری (فرمول جامع)\n"
            "   (موجودی اول دوره + خرید خالص - موجودی پایان دوره = بهای تمام شده)"
        )
        self.btn_periodic.setStyleSheet(
            "QPushButton {"
            "  background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #92400E, stop:1 #D97706);"
            "  color: #FFFFFF; font-weight: bold; font-size: 12.5px; padding: 12px 16px; border-radius: 9px;"
            "  border: 1px solid #F59E0B; text-align: right;"
            "}"
            "QPushButton:hover { background: #B45309; border-color: #FCD34D; }"
        )
        self.btn_periodic.clicked.connect(lambda: self._choose("periodic"))
        layout.addWidget(self.btn_periodic)

        # Footer
        footer_layout = QHBoxLayout()
        self.btn_cancel = QPushButton("انصراف")
        self.btn_cancel.setStyleSheet("background: #1E293B; color: #E2E8F0; padding: 7px 18px; border-radius: 7px; font-size: 12px;")
        self.btn_cancel.clicked.connect(self.reject)

        footer_layout.addStretch()
        footer_layout.addWidget(self.btn_cancel)
        layout.addLayout(footer_layout)

    def _choose(self, method: str):
        self.selected_method = method
        self.accept()

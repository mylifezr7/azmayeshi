"""
Tests for UI Theme, Wallpaper Resources, and Vibrant Shortcut Buttons.
"""
import os
import pytest
from pathlib import Path
from excel_assistant.ui.style import QSS_STYLE, BACKGROUND_IMAGE_PATH, RESOURCES_DIR
from excel_assistant.ui.web_server import HTML_TEMPLATE


def test_theme_background_resource_exists():
    """Verify that background wallpaper resource exists and has valid content."""
    assert os.path.isdir(RESOURCES_DIR)
    bg_file = RESOURCES_DIR / "background.jpg"
    assert bg_file.exists()
    assert bg_file.stat().st_size > 10000  # Wallpaper must be populated


def test_qss_style_contains_all_vibrant_shortcuts():
    """Verify that desktop QSS stylesheet contains vibrant styles for all 6 sidebar buttons."""
    # 1. Undo
    assert "QPushButton#undoButton" in QSS_STYLE
    assert "#EA580C" in QSS_STYLE or "#F97316" in QSS_STYLE

    # 2. Unique lock
    assert "QPushButton#uniqueColButton" in QSS_STYLE
    assert "#6D28D9" in QSS_STYLE or "#8B5CF6" in QSS_STYLE

    # 3. Arithmetic operations
    assert "QPushButton#arithmeticButton" in QSS_STYLE
    assert "#047857" in QSS_STYLE or "#10B981" in QSS_STYLE

    # 4. Product discount
    assert "QPushButton#discountButton" in QSS_STYLE
    assert "#B45309" in QSS_STYLE or "#F59E0B" in QSS_STYLE

    # 5. Insert date
    assert "QPushButton#dateButton" in QSS_STYLE
    assert "#0284C7" in QSS_STYLE or "#38BDF8" in QSS_STYLE

    # 6. Smart analysis
    assert "QPushButton#analyzeButton" in QSS_STYLE
    assert "#1E3A8A" in QSS_STYLE or "#2563EB" in QSS_STYLE

    # 7. Sales summary
    assert "QPushButton#salesButton" in QSS_STYLE
    assert "#BE123C" in QSS_STYLE or "#F43F5E" in QSS_STYLE

    # 8. Profit calculation
    assert "QPushButton#profitButton" in QSS_STYLE
    assert "#059669" in QSS_STYLE or "#10B981" in QSS_STYLE

    # Background image in chat scroll area
    assert "QScrollArea#chatScrollArea" in QSS_STYLE
    assert "background.jpg" in QSS_STYLE


def test_web_template_theme_elements():
    """Verify that Web UI template contains the cosmic theme, MIYUM badge, and colorful shortcuts."""
    assert "M MIYUM" in HTML_TEMPLATE
    assert "btn-undo" in HTML_TEMPLATE
    assert "btn-unique" in HTML_TEMPLATE
    assert "btn-arithmetic" in HTML_TEMPLATE
    assert "btn-discount" in HTML_TEMPLATE
    assert "btn-date" in HTML_TEMPLATE
    assert "btn-analyze" in HTML_TEMPLATE
    assert "btn-sales" in HTML_TEMPLATE
    assert "btn-profit" in HTML_TEMPLATE
    assert "/resources/background.jpg" in HTML_TEMPLATE


def test_quick_actions_dropdown_and_scroll_style():
    """Verify that quick action bar styles include dropdown combo and scroll area."""
    assert "QComboBox#quickCommandsCombo" in QSS_STYLE
    assert "QComboBox#quickCategoryCombo" in QSS_STYLE
    assert "QComboBox#quickActionCombo" in QSS_STYLE
    assert "QPushButton#runQuickActionButton" in QSS_STYLE
    assert "QScrollArea#chipsScrollArea" in QSS_STYLE
    assert "QPushButton.quickChip" in QSS_STYLE
    assert "chip-select" in HTML_TEMPLATE


def test_five_categories_cascading_dropdowns():
    """Verify that Web UI template has the 5 categories cascading selector and options."""
    assert "id=\"catSelect\"" in HTML_TEMPLATE
    assert "id=\"actSelect\"" in HTML_TEMPLATE
    assert "FIVE_CATEGORIES" in HTML_TEMPLATE
    assert "⚡ ۱. قابلیت‌های ضروری" in HTML_TEMPLATE
    assert "💼 ۲. قابلیت‌های حسابداری" in HTML_TEMPLATE
    assert "🛒 ۳. قابلیت‌های فروش و انبارداری" in HTML_TEMPLATE
    assert "🤖 ۴. قابلیت‌های هوشمند و کاربردی" in HTML_TEMPLATE
    assert "⚙️ ۵. امکانات مدیریتی خود برنامه" in HTML_TEMPLATE
    assert "runSelectedAction" in HTML_TEMPLATE


def test_dropup_and_enlarged_controls():
    """Verify that action controls are enlarged, dropup combobox is used, and chips bar is removed."""
    # 1. Enlarged QSS styles
    assert "min-height: 42px;" in QSS_STYLE
    assert "font-size: 14px;" in QSS_STYLE
    assert "QComboBox#quickCategoryCombo" in QSS_STYLE
    assert "QComboBox#quickActionCombo" in QSS_STYLE
    assert "QPushButton#runQuickActionButton" in QSS_STYLE

    # 2. Main window uses DropUpComboBox and removed chips_scroll
    main_window_src = (Path(__file__).parent.parent / "excel_assistant" / "ui" / "main_window.py").read_text(encoding="utf-8")
    assert "class DropUpComboBox(QComboBox):" in main_window_src
    assert "self.category_combo = DropUpComboBox()" in main_window_src
    assert "self.action_combo = DropUpComboBox()" in main_window_src
    assert "chips_scroll = QScrollArea()" not in main_window_src
    assert "HorizontalWheelFilter(chips_scroll)" not in main_window_src

    # 3. Web UI uses action-run-btn and removed bottom chip buttons from chips-bar
    assert "class=\"action-run-btn\"" in HTML_TEMPLATE
    assert "min-height: 44px" in HTML_TEMPLATE
    assert "<button class=\"chip\" onclick=\"sendCommand('مغایرت‌گیری" not in HTML_TEMPLATE


def test_sidebar_and_categories_reorganization():
    """Verify that uniqueness lock, date, analysis, profit, and sales were moved to categories and accounting."""
    main_window_src = (Path(__file__).parent.parent / "excel_assistant" / "ui" / "main_window.py").read_text(encoding="utf-8")
    history_panel_src = (Path(__file__).parent.parent / "excel_assistant" / "ui" / "history_panel.py").read_text(encoding="utf-8")

    # In HistoryPanel, the 5 action buttons are no longer added to sidebar layout
    assert "layout.addWidget(self.btn_unique)" not in history_panel_src
    assert "layout.addWidget(self.btn_date)" not in history_panel_src
    assert "layout.addWidget(self.btn_analyze)" not in history_panel_src
    assert "layout.addWidget(self.btn_sales)" not in history_panel_src
    assert "layout.addWidget(self.btn_profit)" not in history_panel_src

    # Accounting category contains profit and sales calculation
    assert "محاسبه سود" in main_window_src
    assert "محاسبه فروش" in main_window_src

    # Web UI template also includes profit and sales in accounting
    assert 'label: "💵 محاسبه سود (روزانه/ماهانه/سالانه/کل)...", cmd: "محاسبه سود"' in HTML_TEMPLATE
    assert 'label: "💰 محاسبه فروش (روزانه/ماهانه/سالانه/کل)...", cmd: "محاسبه فروش"' in HTML_TEMPLATE




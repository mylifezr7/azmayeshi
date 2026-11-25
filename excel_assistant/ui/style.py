"""
Deep Midnight Cosmic Glassmorphism RTL Persian Styling for Excel AI Assistant.
Provides high-vibrancy shortcut buttons, glowing blue accents, and dark cosmic aesthetics.
"""
from pathlib import Path

# Path to background wallpaper image
RESOURCES_DIR = Path(__file__).parent / "resources"
BACKGROUND_IMAGE_PATH = (RESOURCES_DIR / "background.jpg").as_posix()

QSS_STYLE = f"""
/* Main Window & Base Widgets */
QMainWindow, QDialog, QWidget {{
    background-color: #031412;
    color: #F8FAFC;
    font-family: 'B Nazanin', 'Sahel', 'Vazirmatn', 'IRANSans', 'Tahoma', 'Segoe UI', serif;
    font-size: 13px;
}}

/* Top Header Panel */
QFrame#headerPanel {{
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #062822, stop:0.5 #0A221E, stop:1 #031412);
    border-bottom: 2.5px solid #D97706;
    padding: 8px 16px;
}}

QLabel#userBadge {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #78350F, stop:0.5 #B45309, stop:1 #D97706);
    color: #FEF3C7;
    font-weight: bold;
    font-size: 13px;
    padding: 4px 14px;
    border: 1.5px solid #F59E0B;
    border-radius: 14px;
}}

QLabel#excelStatusPill {{
    background-color: #061E1A;
    border: 1px solid #059669;
    border-radius: 14px;
    padding: 4px 12px;
    color: #6EE7B7;
    font-size: 12px;
    font-weight: 500;
}}

QPushButton#openFileButton {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #0F766E, stop:1 #14B8A6);
    border: 1px solid #2DD4BF;
    border-radius: 10px;
    color: #FFFFFF;
    font-weight: bold;
    padding: 6px 14px;
    font-size: 13px;
}}

QPushButton#openFileButton:hover {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #115E59, stop:1 #2DD4BF);
    border-color: #5EEAD4;
}}

QPushButton#openFileButton:pressed {{
    background: #134E4A;
}}

/* Combo Boxes */
QComboBox {{
    background-color: #061E1A;
    border: 1px solid #0F766E;
    border-radius: 8px;
    padding: 5px 10px;
    color: #E2E8F0;
    font-size: 12px;
}}

QComboBox:hover {{
    border-color: #2DD4BF;
}}

QComboBox::drop-down {{
    border: none;
    width: 20px;
}}

/* Reconnect Button */
QPushButton#reconnectButton {{
    background-color: #061E1A;
    border: 1px solid #0F766E;
    border-radius: 8px;
    color: #5EEAD4;
    font-weight: bold;
    padding: 6px 12px;
}}

QPushButton#reconnectButton:hover {{
    background-color: #115E59;
    color: #FFFFFF;
}}

/* Sidebar Panel */
QFrame#sidebarPanel {{
    background-color: #051A18;
    border: 1.5px solid #0F3A34;
    border-radius: 12px;
    padding: 10px;
}}

/* Sidebar History List */
QListWidget#historyList {{
    background-color: #020E0D;
    border: 1px solid #0F3A34;
    border-radius: 10px;
    color: #E2E8F0;
    padding: 6px;
    font-size: 12px;
}}

QListWidget#historyList::item {{
    padding: 8px;
    border-radius: 6px;
    margin-bottom: 4px;
}}

QListWidget#historyList::item:hover {{
    background-color: #0B2C27;
    color: #6EE7B7;
}}

QListWidget#historyList::item:selected {{
    background-color: #047857;
    color: #FFFFFF;
}}
    margin-bottom: 4px;
}}

QListWidget#historyList::item:hover {{
    background-color: #0B193D;
    color: #60A5FA;
}}

QListWidget#historyList::item:selected {{
    background-color: #1D4ED8;
    color: #FFFFFF;
}}

/* Sidebar Action Buttons with High Vibrancy Primary Colors */
QPushButton.sidebarActionBtn {{
    border-radius: 10px;
    color: #FFFFFF;
    font-weight: bold;
    font-size: 13px;
    padding: 10px 14px;
    text-align: center;
}}

/* 0. Save: Vibrant Teal / Emerald */
QPushButton#saveButton {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #0F766E, stop:1 #14B8A6);
    border: 1px solid #2DD4BF;
    border-radius: 10px;
    color: #FFFFFF;
    font-weight: bold;
    font-size: 13px;
    padding: 10px 14px;
}}

QPushButton#saveButton:hover {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #115E59, stop:1 #2DD4BF);
    border: 1px solid #5EEAD4;
}}

QPushButton#saveButton:pressed {{
    background: #134E4A;
}}

/* 1. Undo: Warm Vibrant Amber/Orange */
QPushButton#undoButton {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #EA580C, stop:1 #FB923C);
    border: 1px solid #F97316;
    border-radius: 10px;
    color: #FFFFFF;
    font-weight: bold;
    font-size: 13px;
    padding: 10px 14px;
}}

QPushButton#undoButton:hover {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #F97316, stop:1 #FDBA74);
    border: 1px solid #FDBA74;
}}

QPushButton#undoButton:pressed {{
    background: #C2410C;
}}

/* 2. Unique Column Lock: Vibrant Royal Purple */
QPushButton#uniqueColButton {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #6D28D9, stop:1 #8B5CF6);
    border: 1px solid #8B5CF6;
    border-radius: 10px;
    color: #FFFFFF;
    font-weight: bold;
    font-size: 13px;
    padding: 10px 14px;
}}

QPushButton#uniqueColButton:hover {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #7C3AED, stop:1 #A78BFA);
    border: 1px solid #A78BFA;
}}

QPushButton#uniqueColButton:pressed {{
    background: #5B21B6;
}}

/* 3. 4 Arithmetic Operations: Rich Vibrant Emerald Green */
QPushButton#arithmeticButton {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #047857, stop:1 #10B981);
    border: 1px solid #10B981;
    border-radius: 10px;
    color: #FFFFFF;
    font-weight: bold;
    font-size: 13px;
    padding: 10px 14px;
}}

QPushButton#arithmeticButton:hover {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #059669, stop:1 #34D399);
    border: 1px solid #34D399;
}}

QPushButton#arithmeticButton:pressed {{
    background: #064E3B;
}}

/* 4. Product Discount: Warm Vibrant Mango/Amber */
QPushButton#discountButton {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #B45309, stop:1 #F59E0B);
    border: 1px solid #F59E0B;
    border-radius: 10px;
    color: #FFFFFF;
    font-weight: bold;
    font-size: 13px;
    padding: 10px 14px;
}}

QPushButton#discountButton:hover {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #D97706, stop:1 #FBBF24);
    border: 1px solid #FBBF24;
}}

QPushButton#discountButton:pressed {{
    background: #78350F;
}}

/* 5. Insert Date: Vibrant Sky Blue / Azure */
QPushButton#dateButton {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #0284C7, stop:1 #38BDF8);
    border: 1px solid #38BDF8;
    border-radius: 10px;
    color: #FFFFFF;
    font-weight: bold;
    font-size: 13px;
    padding: 10px 14px;
}}

QPushButton#dateButton:hover {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #0369A1, stop:1 #7DD3FC);
    border: 1px solid #7DD3FC;
}}

QPushButton#dateButton:pressed {{
    background: #075985;
}}

/* 6. Smart Analysis: Deep Sapphire / Royal Blue */
QPushButton#analyzeButton {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #1E3A8A, stop:1 #2563EB);
    border: 1px solid #3B82F6;
    border-radius: 10px;
    color: #FFFFFF;
    font-weight: bold;
    font-size: 13px;
    padding: 10px 14px;
}}

QPushButton#analyzeButton:hover {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #1D4ED8, stop:1 #60A5FA);
    border: 1px solid #60A5FA;
}}

QPushButton#analyzeButton:pressed {{
    background: #172554;
}}

/* 7. Sales Calculation: Radiant Ruby / Rose Coral */
QPushButton#salesButton {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #BE123C, stop:1 #F43F5E);
    border: 1px solid #F43F5E;
    border-radius: 10px;
    color: #FFFFFF;
    font-weight: bold;
    font-size: 13px;
    padding: 10px 14px;
}}

QPushButton#salesButton:hover {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #E11D48, stop:1 #FB7185);
    border: 1px solid #FB7185;
}}

QPushButton#salesButton:pressed {{
    background: #9F1239;
}}

/* 8. Profit Calculation: Vibrant Emerald Green */
QPushButton#profitButton {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #059669, stop:1 #10B981);
    border: 1px solid #10B981;
    border-radius: 10px;
    color: #FFFFFF;
    font-weight: bold;
    font-size: 13px;
    padding: 10px 14px;
}}

QPushButton#profitButton:hover {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #10B981, stop:1 #34D399);
    border: 1px solid #34D399;
}}

QPushButton#profitButton:pressed {{
    background: #064E3B;
}}

/* Chat Scroll Area */
QScrollArea#chatScrollArea {{
    background-color: #020617;
    background-image: url({BACKGROUND_IMAGE_PATH});
    background-position: center;
    background-repeat: no-repeat;
    border: none;
}}

QScrollArea#chatScrollArea > QWidget > QWidget {{
    background: transparent;
}}

/* Quick Action Chips & Cascading 5-Category Dropdowns (Enlarged & Prominent) */
QComboBox#quickCommandsCombo, QComboBox#quickCategoryCombo, QComboBox#quickActionCombo {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #062822, stop:1 #0B362F);
    border: 2px solid #D97706;
    border-radius: 12px;
    padding: 8px 18px;
    color: #FEF3C7;
    font-size: 14px;
    font-weight: bold;
    min-height: 42px;
}}

QComboBox#quickCommandsCombo:hover, QComboBox#quickCategoryCombo:hover, QComboBox#quickActionCombo:hover {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #083830, stop:1 #0F483E);
    border-color: #F59E0B;
    color: #FFFFFF;
}}

QComboBox#quickCommandsCombo::drop-down, QComboBox#quickCategoryCombo::drop-down, QComboBox#quickActionCombo::drop-down {{
    border: none;
    width: 28px;
}}

QComboBox#quickCommandsCombo QAbstractItemView, QComboBox#quickCategoryCombo QAbstractItemView, QComboBox#quickActionCombo QAbstractItemView {{
    background-color: #031412;
    border: 2px solid #D97706;
    border-radius: 10px;
    color: #E2E8F0;
    selection-background-color: #047857;
    selection-color: #FFFFFF;
    padding: 6px;
    font-size: 13px;
    font-weight: 500;
}}

QPushButton#runQuickActionButton {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #059669, stop:1 #10B981);
    color: #FFFFFF;
    font-size: 14px;
    font-weight: bold;
    border-radius: 12px;
    padding: 8px 24px;
    border: none;
    min-height: 42px;
}}

QPushButton#runQuickActionButton:hover {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #047857, stop:1 #059669);
    border: 1px solid #34D399;
}}

QScrollArea#chipsScrollArea {{
    background: transparent;
    border: none;
}}

QScrollArea#chipsScrollArea QScrollBar:horizontal {{
    height: 4px;
    background: rgba(8, 16, 38, 0.4);
    border-radius: 2px;
}}

QScrollArea#chipsScrollArea QScrollBar::handle:horizontal {{
    background: #0F766E;
    border-radius: 2px;
    min-width: 20px;
}}

QScrollArea#chipsScrollArea QScrollBar::handle:horizontal:hover {{
    background: #2DD4BF;
}}

QScrollArea#chipsScrollArea QScrollBar::add-line:horizontal,
QScrollArea#chipsScrollArea QScrollBar::sub-line:horizontal {{
    width: 0px;
    height: 0px;
}}

QPushButton.quickChip {{
    background-color: rgba(6, 32, 28, 0.85);
    border: 1px solid #0F766E;
    border-radius: 14px;
    padding: 5px 14px;
    font-size: 12px;
    color: #A7F3D0;
    font-weight: 500;
    white-space: nowrap;
}}

QPushButton.quickChip:hover {{
    background-color: #047857;
    color: #FFFFFF;
    border-color: #34D399;
}}

QPushButton.quickChip:pressed {{
    background-color: #064E3B;
}}

/* Chat Input Bar */
QFrame#inputPanel {{
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #062822, stop:1 #031412);
    border-top: 2px solid #D97706;
    padding: 10px 16px;
}}

QLineEdit#chatInput {{
    background-color: #041715;
    border: 1.5px solid #0D9488;
    border-radius: 10px;
    padding: 10px 14px;
    color: #FFFFFF;
    font-size: 13px;
    selection-background-color: #047857;
}}

QLineEdit#chatInput:focus {{
    border: 1.5px solid #F59E0B;
}}

/* Send Button */
QPushButton#sendButton {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #047857, stop:1 #10B981);
    border: 1px solid #34D399;
    color: #FFFFFF;
    border-radius: 10px;
    padding: 10px 22px;
    font-size: 13px;
    font-weight: bold;
}}

QPushButton#sendButton:hover {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #059669, stop:1 #34D399);
    border-color: #6EE7B7;
}}

QPushButton#sendButton:pressed {{
    background: #064E3B;
}}

/* Stop Button */
QPushButton#stopButton {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #991B1B, stop:1 #DC2626);
    border: 1px solid #EF4444;
    color: #FFFFFF;
    border-radius: 10px;
    padding: 10px 18px;
    font-size: 13px;
    font-weight: bold;
}}

QPushButton#stopButton:hover {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #B91C1C, stop:1 #EF4444);
}}

/* General Push Buttons */
QPushButton {{
    background-color: #0E1D40;
    border: 1px solid #1E3A8A;
    border-radius: 8px;
    color: #E2E8F0;
    padding: 8px 16px;
    font-weight: bold;
}}

QPushButton:hover {{
    background-color: #1D4ED8;
    border-color: #38BDF8;
    color: #FFFFFF;
}}

QPushButton:pressed {{
    background-color: #0B193D;
}}

/* Scrollbars */
QScrollBar:vertical {{
    border: none;
    background: #04091A;
    width: 8px;
    margin: 0px;
}}

QScrollBar::handle:vertical {{
    background: #1E3A8A;
    min-height: 20px;
    border-radius: 4px;
}}

QScrollBar::handle:vertical:hover {{
    background: #38BDF8;
}}

QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0px;
}}
"""

"""
Modern RTL Persian Styling for Excel AI Assistant.
Provides clean Excel-themed Dark & Fluent styles.
"""

QSS_STYLE = """
/* Main Window & Core Widgets */
QMainWindow, QDialog, QWidget {
    background-color: #1E1E24;
    color: #E2E8F0;
    font-family: 'Segoe UI', 'Vazirmatn', 'Tahoma', sans-serif;
    font-size: 13px;
}

/* Header & Panels */
QFrame#headerPanel {
    background-color: #262730;
    border-bottom: 2px solid #107C41;
    padding: 8px 16px;
}

QFrame#statusCard {
    background-color: #2D303E;
    border-radius: 8px;
    padding: 6px 12px;
}

/* Chat Scroll Area */
QScrollArea {
    background-color: #1A1A20;
    border: none;
}

QScrollArea > QWidget > QWidget {
    background-color: #1A1A20;
}

/* Chat Input Bar */
QFrame#inputPanel {
    background-color: #262730;
    border-top: 1px solid #333842;
    padding: 10px 16px;
}

QLineEdit#chatInput {
    background-color: #1E1E24;
    border: 1px solid #3F4451;
    border-radius: 8px;
    padding: 10px 14px;
    color: #FFFFFF;
    font-size: 14px;
    selection-background-color: #107C41;
}

QLineEdit#chatInput:focus {
    border: 1px solid #107C41;
}

/* Push Buttons */
QPushButton {
    background-color: #2D303E;
    border: 1px solid #3F4451;
    border-radius: 6px;
    color: #E2E8F0;
    padding: 8px 16px;
    font-weight: bold;
}

QPushButton:hover {
    background-color: #383C4D;
    border-color: #52596B;
}

QPushButton:pressed {
    background-color: #252834;
}

QPushButton#sendButton {
    background-color: #107C41;
    border: 1px solid #0E6B38;
    color: #FFFFFF;
    border-radius: 8px;
    padding: 10px 20px;
    font-size: 14px;
}

QPushButton#sendButton:hover {
    background-color: #148A49;
}

QPushButton#stopButton {
    background-color: #D32F2F;
    border: 1px solid #B71C1C;
    color: #FFFFFF;
    border-radius: 8px;
    padding: 10px 16px;
}

QPushButton#stopButton:hover {
    background-color: #E53935;
}

QPushButton#undoButton {
    background-color: #D97706;
    border: 1px solid #B45309;
    color: #FFFFFF;
}

QPushButton#undoButton:hover {
    background-color: #F59E0B;
}

/* Quick Action Chips */
QPushButton.quickChip {
    background-color: #2A2E3D;
    border: 1px solid #3D4255;
    border-radius: 14px;
    padding: 5px 12px;
    font-size: 12px;
    color: #94A3B8;
}

QPushButton.quickChip:hover {
    background-color: #107C41;
    color: #FFFFFF;
    border-color: #107C41;
}

/* Combo Box */
QComboBox {
    background-color: #2D303E;
    border: 1px solid #3F4451;
    border-radius: 6px;
    padding: 4px 10px;
    color: #E2E8F0;
}

QComboBox::drop-down {
    border: none;
    width: 20px;
}

/* Sidebar History List */
QListWidget#historyList {
    background-color: #1A1A20;
    border: 1px solid #2B2F3D;
    border-radius: 8px;
    color: #CBD5E1;
    padding: 6px;
}

QListWidget#historyList::item {
    padding: 8px;
    border-radius: 6px;
    margin-bottom: 4px;
}

QListWidget#historyList::item:hover {
    background-color: #262A37;
    color: #FFFFFF;
}

QListWidget#historyList::item:selected {
    background-color: #107C41;
    color: #FFFFFF;
}

/* Scrollbars */
QScrollBar:vertical {
    border: none;
    background: #1A1A20;
    width: 8px;
    margin: 0px;
}

QScrollBar::handle:vertical {
    background: #3B4050;
    min-height: 20px;
    border-radius: 4px;
}

QScrollBar::handle:vertical:hover {
    background: #107C41;
}

QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}
"""

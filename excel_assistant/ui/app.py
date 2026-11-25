"""
Application launcher for PySide6 GUI.
"""
from pathlib import Path
import sys
from typing import Optional

try:
    from PySide6.QtCore import Qt
    from PySide6.QtGui import QFont, QFontDatabase
    from PySide6.QtWidgets import QApplication
except ImportError:
    Qt = None
    QFont = QFontDatabase = QApplication = None

from excel_assistant.core.connection import ExcelConnectionManager
from excel_assistant.ui.main_window import MainWindow

RESOURCES_DIR = Path(__file__).parent / "resources"


def setup_application_font(app):
    """Loads bundled classic Persian fonts and sets global typography."""
    if not app:
        return
    try:
        font_files = [
            RESOURCES_DIR / "Vazirmatn-Regular.ttf",
            RESOURCES_DIR / "Vazirmatn-Bold.ttf"
        ]
        loaded_family = None
        for f in font_files:
            if f.exists():
                font_id = QFontDatabase.addApplicationFont(str(f))
                if font_id >= 0:
                    families = QFontDatabase.applicationFontFamilies(font_id)
                    if families and not loaded_family:
                        loaded_family = families[0]

        app_font = QFont()
        families = []
        if loaded_family:
            families.append(loaded_family)
        families.extend(["Vazirmatn", "B Nazanin", "Sahel", "IRANSans", "Tahoma", "Segoe UI"])
        app_font.setFamilies(families)
        app_font.setPointSize(10)
        app.setFont(app_font)
    except Exception as e:
        print(f"Font initialization note: {e}")


def run_gui(workbook_path: Optional[str] = None):
    """Initializes and runs the PySide6 modern GUI."""
    if QApplication is None:
        print("❌ پکیج PySide6 روی سیستم شما نصب نیست. لطفاً دستور زیر را اجرا کنید:\npip install PySide6")
        return 1

    app = QApplication.instance() or QApplication(sys.argv)
    app.setLayoutDirection(Qt.RightToLeft)
    setup_application_font(app)

    manager = ExcelConnectionManager()
    if workbook_path:
        manager.connect(workbook_path)

    window = MainWindow(manager=manager)
    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(run_gui())


"""
Application launcher for PySide6 GUI.
"""
import sys
from typing import Optional
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication

from excel_assistant.core.connection import ExcelConnectionManager
from excel_assistant.ui.main_window import MainWindow


def run_gui(workbook_path: Optional[str] = None):
    """Initializes and runs the PySide6 modern GUI."""
    app = QApplication.instance() or QApplication(sys.argv)
    app.setLayoutDirection(Qt.RightToLeft)

    manager = ExcelConnectionManager()
    if workbook_path:
        manager.connect(workbook_path)

    window = MainWindow(manager=manager)
    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(run_gui())


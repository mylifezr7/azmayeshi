"""
Excel Connection Manager for Excel AI Assistant.
Manages connection lifecycle, multi-file support, and automatic driver selection
(Microsoft Excel 2024 via xlwings/COM with seamless fallback for cross-platform/headless).
"""
import os
import sys
from typing import Dict, List, Optional, Type

from excel_assistant.core.interfaces import IExcelDriver, ISheet, WorkbookInfo
from excel_assistant.core.openpyxl_driver import OpenpyxlDriver
from excel_assistant.core.xlwings_driver import XlwingsDriver, XLWINGS_AVAILABLE, check_xlwings
from excel_assistant.core.exceptions import (
    ExcelConnectionError, WorkbookNotFoundError, SheetNotFoundError
)


class ExcelConnectionManager:
    """Central manager for Excel sessions and workbooks."""

    def __init__(self, preferred_driver: Optional[str] = None):
        """
        :param preferred_driver: 'xlwings', 'openpyxl', or None for auto-detect.
        """
        self.preferred_driver = preferred_driver
        self.active_driver: Optional[IExcelDriver] = None
        self.driver_type: Optional[str] = None
        self.last_driver_error: Optional[str] = None

    def detect_best_driver(self) -> str:
        """Determines whether live Microsoft Excel 2024 (xlwings) or Openpyxl should be used."""
        if self.preferred_driver:
            return self.preferred_driver

        check_xlwings()
        from excel_assistant.core.xlwings_driver import XLWINGS_AVAILABLE

        if sys.platform == 'win32' and XLWINGS_AVAILABLE:
            return 'xlwings'

        # Also check if xlwings apps are currently alive
        if XLWINGS_AVAILABLE:
            try:
                import xlwings as xw
                if len(xw.apps) > 0:
                    return 'xlwings'
            except Exception:
                pass

        return 'openpyxl'

    def connect(self, workbook_path_or_name: Optional[str] = None, force_driver: Optional[str] = None) -> WorkbookInfo:
        """
        Connects to Microsoft Excel 2024 or opens/creates workbook.
        """
        target_driver_name = force_driver or self.detect_best_driver()
        last_error = None
        self.last_driver_error = None

        if target_driver_name == 'xlwings':
            try:
                driver = XlwingsDriver()
                info = driver.connect(workbook_path_or_name)
                self.active_driver = driver
                self.driver_type = 'xlwings'
                self.last_driver_error = None
                return info
            except Exception as e:
                last_error = e
                self.last_driver_error = str(e)
                # Fallback to openpyxl if xlwings fails (e.g. running in container or Excel not installed)
                if not force_driver:
                    target_driver_name = 'openpyxl'
                else:
                    raise ExcelConnectionError(f"Could not connect to live Excel 2024: {e}")

        if target_driver_name == 'openpyxl':
            try:
                driver = OpenpyxlDriver(file_path=workbook_path_or_name)
                info = driver.connect(workbook_path_or_name)
                self.active_driver = driver
                self.driver_type = 'openpyxl'
                return info
            except Exception as e:
                raise ExcelConnectionError(f"Failed to open Excel workbook with openpyxl: {e}") from last_error

        raise ExcelConnectionError(f"Unknown driver type: {target_driver_name}")

    def is_connected(self) -> bool:
        return self.active_driver is not None and self.active_driver.is_connected()

    def open_workbook(self, file_path: str) -> WorkbookInfo:
        """Opens an existing Excel workbook file and makes it active."""
        if not self.active_driver:
            return self.connect(file_path)
        info = self.active_driver.open_workbook(file_path)
        return info

    def switch_workbook(self, workbook_name: str) -> WorkbookInfo:
        """Switches active focus to another open workbook."""
        if not self.active_driver:
            return self.connect(workbook_name)
        info = self.active_driver.switch_workbook(workbook_name)
        return info

    def get_workbook_info(self) -> WorkbookInfo:
        if not self.is_connected():
            raise WorkbookNotFoundError("No Excel workbook currently connected.")
        return self.active_driver.get_workbook_info()

    def get_active_sheet(self) -> ISheet:
        if not self.is_connected():
            raise WorkbookNotFoundError("No Excel workbook currently connected.")
        return self.active_driver.get_active_sheet()

    def get_sheet(self, sheet_name: str) -> ISheet:
        if not self.is_connected():
            raise WorkbookNotFoundError("No Excel workbook currently connected.")
        return self.active_driver.get_sheet(sheet_name)

    def activate_sheet(self, sheet_name: str) -> ISheet:
        if not self.is_connected():
            raise WorkbookNotFoundError("No Excel workbook currently connected.")
        return self.active_driver.activate_sheet(sheet_name)

    def list_open_workbooks(self) -> List[str]:
        if not self.active_driver:
            return []
        return self.active_driver.list_open_workbooks()

    def list_sheets(self) -> List[str]:
        if not self.is_connected():
            return []
        return self.active_driver.get_sheet_names()

    def save(self, path: Optional[str] = None) -> str:
        if not self.is_connected():
            raise WorkbookNotFoundError("No Excel workbook currently connected.")
        return self.active_driver.save(path)

    def close(self) -> None:
        if self.active_driver:
            self.active_driver.close()
            self.active_driver = None
            self.driver_type = None

    def get_status_persian(self) -> str:
        """Returns human-readable Persian connection status."""
        if not self.is_connected():
            return "❌ اتصال به اکسل برقرار نیست."
        info = self.get_workbook_info()
        driver_desc = "اکسل ویندوز (Excel 2024 Live COM)" if self.driver_type == 'xlwings' else "موتور مستقل اکسل (OpenPyXL Engine)"
        return (
            f"✅ وضعیت: متصل به {driver_desc}\n"
            f"📁 نام فایل: {info.name}\n"
            f"📑 شیت فعال: {info.active_sheet}\n"
            f"📋 تعداد شیت‌ها: {len(info.sheets)} ({', '.join(info.sheets)})"
        )

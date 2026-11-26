"""
Tests for Phase 1: Excel Connection and Core Driver Operations.
"""
import os
import tempfile
import pytest

from excel_assistant.core.connection import ExcelConnectionManager
from excel_assistant.core.openpyxl_driver import OpenpyxlDriver
from excel_assistant.core.exceptions import SheetNotFoundError, WorkbookNotFoundError
from excel_assistant.core.utils import (
    col_letter_to_index, col_index_to_letter, parse_range_address, parse_cell_address
)


def test_column_letter_index_conversion():
    assert col_letter_to_index("A") == 1
    assert col_letter_to_index("B") == 2
    assert col_letter_to_index("Z") == 26
    assert col_letter_to_index("AA") == 27
    assert col_letter_to_index("AB") == 28

    assert col_index_to_letter(1) == "A"
    assert col_index_to_letter(2) == "B"
    assert col_index_to_letter(26) == "Z"
    assert col_index_to_letter(27) == "AA"
    assert col_index_to_letter(28) == "AB"


def test_parse_cell_and_range():
    assert parse_cell_address("A1") == (1, 1)
    assert parse_cell_address("C15") == (15, 3)
    assert parse_cell_address("$B$4") == (4, 2)

    assert parse_range_address("A1:C5") == (1, 1, 5, 3)
    assert parse_range_address("B2:B10") == (2, 2, 10, 2)
    assert parse_range_address("D4") == (4, 4, 4, 4)


def test_openpyxl_driver_lifecycle():
    with tempfile.TemporaryDirectory() as tmpdir:
        test_file = os.path.join(tmpdir, "test_excel.xlsx")
        driver = OpenpyxlDriver()
        driver.connect(test_file)

        assert driver.is_connected()
        info = driver.get_workbook_info()
        assert info.name == "test_excel.xlsx"
        assert len(info.sheets) >= 1

        sheet = driver.get_active_sheet()
        # Write values
        sheet.set_value("A1", "Product")
        sheet.set_value("B1", "Price")
        sheet.set_value("C1", "Qty")
        sheet.set_value("A2", "Laptop")
        sheet.set_value("B2", 1500)
        sheet.set_value("C2", 3)

        # Read values
        assert sheet.get_value("A1") == "Product"
        assert sheet.get_value("B2") == 1500
        assert sheet.get_value("C2") == 3

        # Write and read formula
        sheet.set_formula("D1", "Total")
        sheet.set_formula("D2", "=B2*C2")
        assert sheet.get_formula("D2") == "=B2*C2"

        # Bounds check
        bounds = sheet.get_used_range_bounds()
        assert bounds[0] == 1  # min row
        assert bounds[1] == 1  # min col
        assert bounds[2] >= 2  # max row
        assert bounds[3] >= 4  # max col

        # Colors & Fonts
        sheet.set_cell_color("A1:D1", "#FFFF00")
        sheet.set_font("A1:D1", bold=True, size=12)

        # Create new sheet & switch
        driver.create_sheet("Sales")
        assert "Sales" in driver.get_sheet_names()
        sales_sheet = driver.activate_sheet("Sales")
        assert sales_sheet.name == "Sales"

        # Save and close
        driver.save()
        assert os.path.exists(test_file)
        driver.close()
        assert not driver.is_connected()


def test_connection_manager():
    with tempfile.TemporaryDirectory() as tmpdir:
        test_file = os.path.join(tmpdir, "manager_test.xlsx")
        manager = ExcelConnectionManager(preferred_driver="openpyxl")
        info = manager.connect(test_file)

        assert manager.is_connected()
        assert info.name == "manager_test.xlsx"
        assert manager.driver_type == "openpyxl"

        sheet = manager.get_active_sheet()
        sheet.set_value("A1", 100)
        assert sheet.get_value("A1") == 100

        status = manager.get_status_persian()
        assert "وضعیت: متصل" in status
        assert "manager_test.xlsx" in status

        manager.save()
        manager.close()
        assert not manager.is_connected()


def test_sheet_not_found():
    manager = ExcelConnectionManager(preferred_driver="openpyxl")
    manager.connect()
    with pytest.raises(SheetNotFoundError):
        manager.get_sheet("NonExistentSheetXYZ")
    manager.close()


def test_to_float_safe_persian_and_formats():
    from excel_assistant.core.utils import to_float_safe
    assert to_float_safe(100) == 100.0
    assert to_float_safe("۲۳,۰۰۰ تومان") == 23000.0
    assert to_float_safe("۵۵٬۲۰۰") == 55200.0
    assert to_float_safe("3,500,000 ریال") == 3500000.0
    assert to_float_safe("۱۲٫۵") == 12.5
    assert to_float_safe("۲۰٪") == 20.0
    assert to_float_safe(None) == 0.0
    assert to_float_safe("invalid", default=99.0) == 99.0


def test_openpyxl_csv_support():
    with tempfile.TemporaryDirectory() as tmpdir:
        csv_file = os.path.join(tmpdir, "data.csv")
        with open(csv_file, "w", encoding="utf-8") as f:
            f.write("کالا,تعداد,قیمت\nسیب,10,2000\nپرتقال,5,3000\n")

        driver = OpenpyxlDriver()
        driver.connect(csv_file)
        assert driver.is_connected()
        sheet = driver.get_active_sheet()
        assert sheet.get_value("A1") == "کالا"
        assert sheet.get_value("B2") == "10"
        driver.close()


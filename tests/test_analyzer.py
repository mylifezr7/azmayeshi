"""
Tests for Phase 2: Sheet & Workbook Structure Analysis, Type Detection, and Suggestions.
"""
import datetime
import pytest

from excel_assistant.core.openpyxl_driver import OpenpyxlDriver
from excel_assistant.analyzer.type_detector import detect_cell_type, analyze_column
from excel_assistant.analyzer.inspector import SheetInspector, WorkbookInspector
from excel_assistant.analyzer.formula_checker import check_sheet_formulas_and_errors
from excel_assistant.analyzer.suggestions import generate_smart_suggestions


def test_detect_cell_type():
    assert detect_cell_type(100) == "integer"
    assert detect_cell_type(45.5) == "float"
    assert detect_cell_type("Hello") == "text"
    assert detect_cell_type("50%") == "percentage"
    assert detect_cell_type("$1,200") == "currency"
    assert detect_cell_type("1000 تومان") == "currency"
    assert detect_cell_type(datetime.date(2025, 1, 1)) == "date"
    assert detect_cell_type("2025-05-12") == "date"
    assert detect_cell_type(True) == "boolean"
    assert detect_cell_type(None) == "empty"
    assert detect_cell_type("#DIV/0!") == "error"


def test_analyze_column():
    summary = analyze_column("A", 1, "قیمت", [100, 200, 300, 100, None])
    assert summary.column_letter == "A"
    assert summary.header_name == "قیمت"
    assert summary.detected_type in ("numeric", "integer")
    assert summary.null_count == 1
    assert summary.duplicate_count == 1
    assert summary.min_val == 100
    assert summary.max_val == 300


def test_sheet_inspection_and_anomalies():
    driver = OpenpyxlDriver()
    driver.connect()
    sheet = driver.get_active_sheet()

    # Setup test sheet
    sheet.set_value("A1", "نام")
    sheet.set_value("B1", "تعداد")
    sheet.set_value("C1", "قیمت")
    sheet.set_value("D1", "خطا")

    # Row 2
    sheet.set_value("A2", "موبایل")
    sheet.set_value("B2", 5)
    sheet.set_value("C2", 1000)
    sheet.set_value("D2", 10)

    # Row 3 (Duplicate row)
    sheet.set_value("A3", "موبایل")
    sheet.set_value("B3", 5)
    sheet.set_value("C3", 1000)
    sheet.set_value("D3", 10)

    # Row 4 (Empty row - left empty)

    # Row 5 (With error)
    sheet.set_value("A5", "تبلت")
    sheet.set_value("B5", 2)
    sheet.set_value("C5", 2500)
    sheet.set_value("D5", "#DIV/0!")

    analysis = SheetInspector.inspect_sheet(sheet)

    assert analysis.sheet_name == sheet.name
    assert len(analysis.headers) == 4
    assert analysis.headers[0][1] == "نام"
    assert analysis.headers[1][1] == "تعداد"
    assert analysis.headers[2][1] == "قیمت"

    # Empty rows check
    assert 4 in analysis.empty_row_indices

    # Duplicate rows check
    assert 3 in analysis.duplicate_row_indices

    # Formula/Error check
    assert "D5" in analysis.formula_report.error_cells

    # Suggestions check
    suggestions = generate_smart_suggestions(analysis)
    assert len(suggestions) > 0
    suggestions_text = " ".join(suggestions)
    assert "ردیف خالی" in suggestions_text or "ردیف‌های خالی" in suggestions_text
    assert "تکراری" in suggestions_text

    # Workbook report formatting
    wb_analysis = WorkbookInspector.inspect(driver)
    report = WorkbookInspector.format_persian_summary(wb_analysis)
    assert "گزارش تحلیل هوشمند" in report
    assert "ستون" in report

    driver.close()

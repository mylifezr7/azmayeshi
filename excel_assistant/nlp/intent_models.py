"""
Pydantic action schemas for Excel AI Assistant operations.
"""
from typing import Any, Dict, List, Literal, Optional, Union
from pydantic import BaseModel, Field


class BaseActionModel(BaseModel):
    action_type: str
    description_persian: str = ""
    is_destructive: bool = False
    affected_range: str = ""


class ColumnArithmeticAction(BaseActionModel):
    action_type: Literal["column_arithmetic"] = "column_arithmetic"
    operation: Literal["multiply", "add", "subtract", "divide"]  # ضرب، جمع، تفریق، تقسیم
    col1: Optional[str] = "A"  # e.g. "A" or header name
    col2: Optional[str] = "B"  # e.g. "B"
    columns: Optional[List[str]] = None  # multi-column operands e.g. ["2", "4", "5"] or ["B", "D", "E"]
    target_col: Optional[str] = None  # e.g. "C" or "H"
    as_formula: bool = True  # generate Excel formula like =A2*B2 or compute values


class ColumnStatsAction(BaseActionModel):
    action_type: Literal["column_stats"] = "column_stats"
    stat_type: Literal["sum", "average", "median", "max", "min", "count", "percentage"]  # میانگین، مجموع، میانه، بیشترین، کمترین، تعداد، درصد
    source_col: str  # e.g. "D"
    target_cell: Optional[str] = None  # cell to write result into, or just return result in chat


class HighlightConditionAction(BaseActionModel):
    action_type: Literal["highlight_condition"] = "highlight_condition"
    target_col: Optional[str] = None
    target_range: Optional[str] = None
    condition: Literal["gt", "gte", "lt", "lte", "eq", "neq", "between", "duplicates"]  # بیشتر از، کمتر از، مساوی، تکراری
    value: Optional[Union[float, int, str]] = None
    value2: Optional[Union[float, int, str]] = None
    color_hex: str = "#90EE90"  # default light green


class HighlightDuplicatesAction(BaseActionModel):
    action_type: Literal["highlight_duplicates"] = "highlight_duplicates"
    target_col: Optional[str] = None
    target_range: Optional[str] = None
    color_hex: str = "#FFFF00"  # default yellow
    text_only: bool = True  # Only highlight textual/typographic duplicates (skip pure numbers)
    distinct_colors: bool = True  # Assign distinct color to each duplicate value group


class DeleteEmptyRowsAction(BaseActionModel):
    action_type: Literal["delete_empty_rows"] = "delete_empty_rows"
    is_destructive: bool = True


class DeleteDuplicatesAction(BaseActionModel):
    action_type: Literal["delete_duplicates"] = "delete_duplicates"
    target_col: Optional[str] = None
    is_destructive: bool = True


class DeleteRowsOrColsAction(BaseActionModel):
    action_type: Literal["delete_rows_or_cols"] = "delete_rows_or_cols"
    target_type: Literal["row", "column"]
    targets: List[Union[int, str]]
    is_destructive: bool = True


class SortDataAction(BaseActionModel):
    action_type: Literal["sort_data"] = "sort_data"
    key_col: str
    ascending: bool = False  # False = descending (بیشترین به کمترین)


class AutoFitColumnsAction(BaseActionModel):
    action_type: Literal["autofit_columns"] = "autofit_columns"
    columns: Optional[List[str]] = None


class CreateTableAction(BaseActionModel):
    action_type: Literal["create_table"] = "create_table"
    range_str: Optional[str] = None
    table_name: Optional[str] = None


class FormatCellsAction(BaseActionModel):
    action_type: Literal["format_cells"] = "format_cells"
    target_range: str
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    font_size: Optional[float] = None
    font_name: Optional[str] = None
    bg_color: Optional[str] = None
    font_color: Optional[str] = None


class ApplyDiscountAction(BaseActionModel):
    action_type: Literal["apply_discount"] = "apply_discount"
    price_col: str  # Original price column e.g. "B" or "2"
    discount_col: Optional[str] = None  # Discount percentage column e.g. "C" or "3"
    discount_percent: Optional[float] = None  # Fixed discount percentage e.g. 15.0
    target_col: Optional[str] = None  # Final price after discount column e.g. "D" or "4"
    discount_amount_col: Optional[str] = None  # Optional discount amount column e.g. "E"


class InsertDateAction(BaseActionModel):
    action_type: Literal["insert_date"] = "insert_date"
    target_col: str = "A"  # e.g. "A" or "1"
    start_row: int = 1     # e.g. 1
    end_row: int = 30      # e.g. 30
    year: int = 1403      # e.g. 1403
    month: int = 8        # e.g. 8
    day: int = 29         # e.g. 29
    mode: Literal["fixed", "sequential"] = "fixed"
    add_header: bool = False


class ApplyFormulaAction(BaseActionModel):
    action_type: Literal["apply_formula"] = "apply_formula"
    target_cell_or_range: str
    formula_template: str  # e.g. "=SUM(B2:B20)" or "=A{row}*B{row}"


class AnalyzeSheetAction(BaseActionModel):
    action_type: Literal["analyze_sheet"] = "analyze_sheet"
    sheet_name: Optional[str] = None


class UndoLastAction(BaseActionModel):
    action_type: Literal["undo_last_action"] = "undo_last_action"


ExcelAction = Union[
    ColumnArithmeticAction,
    ColumnStatsAction,
    ApplyDiscountAction,
    InsertDateAction,
    HighlightConditionAction,
    HighlightDuplicatesAction,
    DeleteEmptyRowsAction,
    DeleteDuplicatesAction,
    DeleteRowsOrColsAction,
    SortDataAction,
    AutoFitColumnsAction,
    CreateTableAction,
    FormatCellsAction,
    ApplyFormulaAction,
    AnalyzeSheetAction,
    UndoLastAction,
]

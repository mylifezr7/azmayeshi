"""
Abstract interfaces and data structures for Excel connection and operations.
"""
from abc import ABC, abstractmethod
from typing import Any, Dict, List, Optional, Tuple, Union
from dataclasses import dataclass, field


@dataclass
class WorkbookInfo:
    name: str
    path: Optional[str] = None
    is_saved: bool = True
    sheets: List[str] = field(default_factory=list)
    active_sheet: str = ""
    app_version: str = "Excel 2024"


@dataclass
class SheetStructure:
    name: str
    used_range: str
    min_row: int
    min_col: int
    max_row: int
    max_col: int
    headers: List[str] = field(default_factory=list)
    column_types: Dict[str, str] = field(default_factory=dict)
    has_formulas: bool = False
    formula_count: int = 0
    empty_row_count: int = 0
    row_count: int = 0
    column_count: int = 0


class ISheet(ABC):
    """Abstract representation of an Excel worksheet."""

    @property
    @abstractmethod
    def name(self) -> str:
        pass

    @abstractmethod
    def get_used_range_bounds(self) -> Tuple[int, int, int, int]:
        """Returns (min_row, min_col, max_row, max_col), 1-indexed."""
        pass

    @abstractmethod
    def get_value(self, cell_or_range: str) -> Any:
        """Read value(s) from a cell or range address like 'A1' or 'A1:C10'."""
        pass

    @abstractmethod
    def set_value(self, cell_or_range: str, value: Any) -> None:
        """Write value(s) into a cell or range."""
        pass

    @abstractmethod
    def get_formula(self, cell_or_range: str) -> Any:
        """Get formula string from a cell or matrix of formulas from a range."""
        pass

    @abstractmethod
    def set_formula(self, cell_or_range: str, formula: str) -> None:
        """Write an Excel formula (e.g. '=A2*B2') into cell or range."""
        pass

    @abstractmethod
    def get_column_values(self, col: Union[int, str], include_header: bool = False) -> List[Any]:
        """Get all values in a column within the used range."""
        pass

    @abstractmethod
    def set_cell_color(self, cell_or_range: str, color_hex: str) -> None:
        """Set background color using hex color code like '#FFFF00' or '#28A745'."""
        pass

    @abstractmethod
    def set_font(self, cell_or_range: str, name: Optional[str] = None, size: Optional[float] = None,
                 bold: Optional[bool] = None, italic: Optional[bool] = None, color_hex: Optional[str] = None) -> None:
        """Configure font properties for range."""
        pass

    @abstractmethod
    def auto_fit_columns(self, cols: Optional[List[Union[int, str]]] = None) -> None:
        """Autofit width of specified columns or all used columns."""
        pass

    @abstractmethod
    def delete_rows(self, row_indices: List[int]) -> int:
        """Delete specific rows by 1-based index (descending order). Returns count deleted."""
        pass

    @abstractmethod
    def insert_rows(self, row_index: int, count: int = 1) -> None:
        """Insert row(s) before row_index."""
        pass

    @abstractmethod
    def delete_cols(self, col_indices: List[Union[int, str]]) -> int:
        """Delete column(s) by index or letter. Returns count deleted."""
        pass

    @abstractmethod
    def sort_range(self, range_str: str, key_col: Union[int, str], ascending: bool = True, has_headers: bool = True) -> None:
        """Sort data in range by a key column."""
        pass

    @abstractmethod
    def create_table(self, range_str: str, name: Optional[str] = None, style: str = "TableStyleMedium9") -> str:
        """Convert range to Excel Table/ListObject."""
        pass


class IExcelDriver(ABC):
    """Abstract interface for Excel application automation drivers."""

    @abstractmethod
    def connect(self, workbook_name_or_path: Optional[str] = None) -> WorkbookInfo:
        """Connect to running Excel instance or open a workbook."""
        pass

    @abstractmethod
    def is_connected(self) -> bool:
        """Check if connection is active."""
        pass

    @abstractmethod
    def list_open_workbooks(self) -> List[str]:
        """List names of currently open workbooks."""
        pass

    @abstractmethod
    def get_workbook_info(self) -> WorkbookInfo:
        """Get metadata for the currently active workbook."""
        pass

    @abstractmethod
    def get_sheet_names(self) -> List[str]:
        """List all sheet names in current workbook."""
        pass

    @abstractmethod
    def get_active_sheet(self) -> ISheet:
        """Get ISheet wrapper for active worksheet."""
        pass

    @abstractmethod
    def get_sheet(self, name: str) -> ISheet:
        """Get ISheet wrapper for worksheet by name."""
        pass

    @abstractmethod
    def activate_sheet(self, name: str) -> ISheet:
        """Activate and return the specified sheet."""
        pass

    @abstractmethod
    def create_sheet(self, name: str) -> ISheet:
        """Create new sheet with given name."""
        pass

    @abstractmethod
    def save(self, path: Optional[str] = None) -> str:
        """Save workbook to path or current file. Returns saved path."""
        pass

    @abstractmethod
    def close(self) -> None:
        """Close current connection/workbook cleanly."""
        pass

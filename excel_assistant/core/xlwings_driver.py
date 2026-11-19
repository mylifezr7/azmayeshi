"""
Live Microsoft Excel 2024 Automation Driver using xlwings and Windows COM.
Directly interfaces with running Microsoft Excel instances on Windows.
"""
import os
import sys
from typing import Any, Dict, List, Optional, Tuple, Union

try:
    import xlwings as xw
    XLWINGS_AVAILABLE = True
except ImportError:
    xw = None
    XLWINGS_AVAILABLE = False

from excel_assistant.core.interfaces import ISheet, IExcelDriver, WorkbookInfo
from excel_assistant.core.exceptions import (
    ExcelNotRunningError, ExcelConnectionError, SheetNotFoundError,
    WorkbookNotFoundError, ActionExecutionError
)
from excel_assistant.core.utils import (
    col_letter_to_index, col_index_to_letter, parse_range_address,
    parse_cell_address, hex_to_rgb, normalize_hex_color, bounds_to_range_str
)


class XlwingsSheet(ISheet):
    """Implementation of ISheet wrapping a live xlwings Sheet."""

    def __init__(self, xw_sheet):
        self._sheet = xw_sheet

    @property
    def name(self) -> str:
        return self._sheet.name

    def get_used_range_bounds(self) -> Tuple[int, int, int, int]:
        try:
            used = self._sheet.used_range
            min_row = used.row
            min_col = used.column
            max_row = min_row + used.rows.count - 1
            max_col = min_col + used.columns.count - 1
            return min_row, min_col, max_row, max_col
        except Exception:
            return 1, 1, 1, 1

    def get_value(self, cell_or_range: str) -> Any:
        try:
            val = self._sheet.range(cell_or_range).value
            return val
        except Exception as e:
            raise ActionExecutionError(f"Failed to read from range '{cell_or_range}' in live Excel: {e}")

    def set_value(self, cell_or_range: str, value: Any) -> None:
        try:
            self._sheet.range(cell_or_range).value = value
        except Exception as e:
            raise ActionExecutionError(f"Failed to write to range '{cell_or_range}' in live Excel: {e}")

    def get_formula(self, cell_or_range: str) -> Any:
        try:
            return self._sheet.range(cell_or_range).formula
        except Exception as e:
            raise ActionExecutionError(f"Failed to get formula from '{cell_or_range}': {e}")

    def set_formula(self, cell_or_range: str, formula: str) -> None:
        try:
            if not formula.startswith('='):
                formula = '=' + formula
            self._sheet.range(cell_or_range).formula = formula
        except Exception as e:
            raise ActionExecutionError(f"Failed to set formula on '{cell_or_range}': {e}")

    def get_column_values(self, col: Union[int, str], include_header: bool = False) -> List[Any]:
        col_idx = col if isinstance(col, int) else col_letter_to_index(col)
        min_r, _, max_r, _ = self.get_used_range_bounds()
        start_r = min_r if include_header else min_r + 1
        if start_r > max_r:
            return []
        col_letter = col_index_to_letter(col_idx)
        rng = f"{col_letter}{start_r}:{col_letter}{max_r}"
        vals = self._sheet.range(rng).value
        if not isinstance(vals, list):
            return [vals] if vals is not None else []
        return vals

    def set_cell_color(self, cell_or_range: str, color_hex: str) -> None:
        try:
            rgb = hex_to_rgb(color_hex)
            self._sheet.range(cell_or_range).color = rgb
        except Exception as e:
            raise ActionExecutionError(f"Failed to color range '{cell_or_range}': {e}")

    def set_font(self, cell_or_range: str, name: Optional[str] = None, size: Optional[float] = None,
                 bold: Optional[bool] = None, italic: Optional[bool] = None, color_hex: Optional[str] = None) -> None:
        try:
            rng = self._sheet.range(cell_or_range)
            if name:
                rng.font.name = name
            if size is not None:
                rng.font.size = size
            if bold is not None:
                rng.font.bold = bold
            if italic is not None:
                rng.font.italic = italic
            if color_hex:
                rng.font.color = hex_to_rgb(color_hex)
        except Exception as e:
            raise ActionExecutionError(f"Failed to set font on '{cell_or_range}': {e}")

    def auto_fit_columns(self, cols: Optional[List[Union[int, str]]] = None) -> None:
        try:
            if cols:
                for c in cols:
                    letter = c if isinstance(c, str) and not c.isdigit() else col_index_to_letter(int(c))
                    self._sheet.range(f"{letter}:{letter}").columns.autofit()
            else:
                self._sheet.used_range.columns.autofit()
        except Exception as e:
            raise ActionExecutionError(f"Failed to autofit columns: {e}")

    def delete_rows(self, row_indices: List[int]) -> int:
        if not row_indices:
            return 0
        sorted_rows = sorted(set(row_indices), reverse=True)
        count = 0
        for r in sorted_rows:
            try:
                self._sheet.range(f"{r}:{r}").delete(shift='up')
                count += 1
            except Exception as e:
                raise ActionExecutionError(f"Failed to delete row {r}: {e}")
        return count

    def insert_rows(self, row_index: int, count: int = 1) -> None:
        try:
            for _ in range(count):
                self._sheet.range(f"{row_index}:{row_index}").insert(shift='down')
        except Exception as e:
            raise ActionExecutionError(f"Failed to insert row at {row_index}: {e}")

    def delete_cols(self, col_indices: List[Union[int, str]]) -> int:
        if not col_indices:
            return 0
        indices = [c if isinstance(c, int) else col_letter_to_index(c) for c in col_indices]
        sorted_cols = sorted(set(indices), reverse=True)
        count = 0
        for c in sorted_cols:
            letter = col_index_to_letter(c)
            try:
                self._sheet.range(f"{letter}:{letter}").delete(shift='left')
                count += 1
            except Exception as e:
                raise ActionExecutionError(f"Failed to delete column {letter}: {e}")
        return count

    def sort_range(self, range_str: str, key_col: Union[int, str], ascending: bool = True, has_headers: bool = True) -> None:
        try:
            # Check if COM api available on Windows
            key_letter = key_col if isinstance(key_col, str) and not key_col.isdigit() else col_index_to_letter(int(key_col))
            r1, c1, r2, c2 = parse_range_address(range_str)
            key_idx = (key_col if isinstance(key_col, int) else col_letter_to_index(key_col)) - c1
            start_r = (r1 + 1) if has_headers else r1
            if start_r >= r2:
                return

            # Read in-memory values and sort
            target_rng = f"{col_index_to_letter(c1)}{start_r}:{col_index_to_letter(c2)}{r2}"
            data = self._sheet.range(target_rng).value
            if not isinstance(data, list):
                return
            # Normalize 1D to 2D
            if data and not isinstance(data[0], list):
                data = [[x] for x in data]

            def sort_key(item):
                val = item[key_idx] if key_idx < len(item) else None
                if val is None:
                    return (1, "")
                if isinstance(val, (int, float)):
                    return (0, float(val))
                return (0, str(val).lower())

            data.sort(key=sort_key, reverse=not ascending)
            self._sheet.range(target_rng).value = data
        except Exception as e:
            raise ActionExecutionError(f"Failed to sort range {range_str}: {e}")

    def create_table(self, range_str: str, name: Optional[str] = None, style: str = "TableStyleMedium9") -> str:
        try:
            rng = self._sheet.range(range_str)
            table_name = name or f"Table_{self.name}_{len(self._sheet.tables) + 1}"
            tbl = self._sheet.tables.add(source=rng, name=table_name)
            return table_name
        except Exception as e:
            raise ActionExecutionError(f"Failed to create table on {range_str}: {e}")


class XlwingsDriver(IExcelDriver):
    """Driver for connecting to Microsoft Excel 2024 via xlwings / COM."""

    def __init__(self):
        if not XLWINGS_AVAILABLE:
            raise ExcelConnectionError("xlwings is not installed in the environment.")
        self._app: Optional[xw.App] = None
        self._book: Optional[xw.Book] = None

    @classmethod
    def is_excel_installed(cls) -> bool:
        """Check if Excel COM can be reached."""
        if not XLWINGS_AVAILABLE:
            return False
        try:
            apps = xw.apps
            return True
        except Exception:
            return False

    def connect(self, workbook_name_or_path: Optional[str] = None) -> WorkbookInfo:
        if not XLWINGS_AVAILABLE:
            raise ExcelConnectionError("xlwings is required for Microsoft Excel 2024 live connection.")

        try:
            # 1. Connect to active instance of Excel or start one
            if len(xw.apps) > 0:
                self._app = xw.apps.active
            else:
                self._app = xw.App(visible=True, add_book=False)

            # 2. Find or open target workbook
            if workbook_name_or_path:
                # Check if already open by name
                found = False
                base_target = os.path.basename(workbook_name_or_path).lower()
                for b in self._app.books:
                    if b.name.lower() == base_target:
                        self._book = b
                        found = True
                        break
                if not found:
                    if os.path.exists(workbook_name_or_path):
                        self._book = self._app.books.open(workbook_name_or_path)
                    else:
                        self._book = self._app.books.add()
                        self._book.save(workbook_name_or_path)
            else:
                # Connect to currently active workbook if open
                if len(self._app.books) > 0:
                    self._book = self._app.books.active
                else:
                    self._book = self._app.books.add()

            return self.get_workbook_info()
        except Exception as e:
            raise ExcelConnectionError(f"Failed to connect to Microsoft Excel 2024: {e}")

    def is_connected(self) -> bool:
        if not self._app or not self._book:
            return False
        try:
            _ = self._book.name
            return True
        except Exception:
            return False

    def list_open_workbooks(self) -> List[str]:
        if not self._app:
            return []
        try:
            return [b.name for b in self._app.books]
        except Exception:
            return []

    def get_workbook_info(self) -> WorkbookInfo:
        if not self.is_connected():
            raise WorkbookNotFoundError("No active workbook connected in Microsoft Excel.")
        name = self._book.name
        try:
            path = self._book.fullname
        except Exception:
            path = None

        sheets = [s.name for s in self._book.sheets]
        active = self._book.sheets.active.name if self._book.sheets else ""

        # Determine Excel version
        version_str = "Microsoft Excel 2024"
        try:
            if hasattr(self._app, 'api') and hasattr(self._app.api, 'Version'):
                v = self._app.api.Version
                version_str = f"Microsoft Excel (COM {v})"
        except Exception:
            pass

        return WorkbookInfo(
            name=name,
            path=path,
            is_saved=True,
            sheets=sheets,
            active_sheet=active,
            app_version=version_str
        )

    def get_sheet_names(self) -> List[str]:
        if not self.is_connected():
            raise WorkbookNotFoundError("No active workbook connected.")
        return [s.name for s in self._book.sheets]

    def get_active_sheet(self) -> ISheet:
        if not self.is_connected():
            raise WorkbookNotFoundError("No active workbook connected.")
        return XlwingsSheet(self._book.sheets.active)

    def get_sheet(self, name: str) -> ISheet:
        if not self.is_connected():
            raise WorkbookNotFoundError("No active workbook connected.")
        for s in self._book.sheets:
            if s.name == name:
                return XlwingsSheet(s)
        raise SheetNotFoundError(f"Sheet '{name}' not found in active workbook.")

    def activate_sheet(self, name: str) -> ISheet:
        sheet = self.get_sheet(name)
        self._book.sheets[name].activate()
        return sheet

    def create_sheet(self, name: str) -> ISheet:
        if not self.is_connected():
            raise WorkbookNotFoundError("No active workbook connected.")
        s = self._book.sheets.add(name=name, after=self._book.sheets[-1])
        return XlwingsSheet(s)

    def save(self, path: Optional[str] = None) -> str:
        if not self.is_connected():
            raise WorkbookNotFoundError("No active workbook connected.")
        if path:
            self._book.save(path)
            return path
        else:
            self._book.save()
            return self._book.fullname

    def close(self) -> None:
        if self._book:
            try:
                self._book.close()
            except Exception:
                pass
            self._book = None

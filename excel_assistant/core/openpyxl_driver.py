"""
OpenPyXL Driver for Excel AI Assistant.
Provides high-fidelity, cross-platform Excel workbook automation matching ISheet & IExcelDriver.
"""
import os
import shutil
from typing import Any, Dict, List, Optional, Tuple, Union

import openpyxl
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.worksheet.table import Table, TableStyleInfo

from excel_assistant.core.interfaces import ISheet, IExcelDriver, WorkbookInfo
from excel_assistant.core.exceptions import (
    SheetNotFoundError, WorkbookNotFoundError, RangeInvalidError, ActionExecutionError
)
from excel_assistant.core.utils import (
    col_letter_to_index, col_index_to_letter, parse_range_address,
    parse_cell_address, normalize_hex_color, bounds_to_range_str
)


class OpenpyxlSheet(ISheet):
    """Implementation of ISheet wrapping an openpyxl Worksheet."""

    def __init__(self, worksheet: openpyxl.worksheet.worksheet.Worksheet):
        self._ws = worksheet

    @property
    def name(self) -> str:
        return self._ws.title

    def get_used_range_bounds(self) -> Tuple[int, int, int, int]:
        min_row = self._ws.min_row or 1
        min_col = self._ws.min_column or 1
        max_row = self._ws.max_row or 1
        max_col = self._ws.max_column or 1
        # If worksheet is completely empty
        if max_row == 1 and max_col == 1 and self._ws.cell(row=1, column=1).value is None:
            return 1, 1, 1, 1
        return min_row, min_col, max_row, max_col

    def get_value(self, cell_or_range: str) -> Any:
        try:
            r1, c1, r2, c2 = parse_range_address(cell_or_range)
            # Clip if open-ended
            ub_r1, ub_c1, ub_r2, ub_c2 = self.get_used_range_bounds()
            r1 = max(1, r1)
            c1 = max(1, c1)
            r2 = min(r2, max(ub_r2, r1))
            c2 = min(c2, max(ub_c2, c1))

            if r1 == r2 and c1 == c2:
                return self._ws.cell(row=r1, column=c1).value

            rows = []
            for r in range(r1, r2 + 1):
                row_vals = []
                for c in range(c1, c2 + 1):
                    row_vals.append(self._ws.cell(row=r, column=c).value)
                rows.append(row_vals)
            return rows
        except Exception as e:
            raise ActionExecutionError(f"Failed to read value from {cell_or_range}: {e}")

    def set_value(self, cell_or_range: str, value: Any) -> None:
        try:
            r1, c1, r2, c2 = parse_range_address(cell_or_range)
            if r1 == r2 and c1 == c2:
                self._ws.cell(row=r1, column=c1).value = value
            elif isinstance(value, (list, tuple)):
                for r_offset, row_data in enumerate(value):
                    current_r = r1 + r_offset
                    if current_r > r2:
                        break
                    if isinstance(row_data, (list, tuple)):
                        for c_offset, val in enumerate(row_data):
                            current_c = c1 + c_offset
                            if current_c > c2:
                                break
                            self._ws.cell(row=current_r, column=current_c).value = val
                    else:
                        self._ws.cell(row=current_r, column=c1).value = row_data
            else:
                # Set same scalar value across entire range
                for r in range(r1, r2 + 1):
                    for c in range(c1, c2 + 1):
                        self._ws.cell(row=r, column=c).value = value
        except Exception as e:
            raise ActionExecutionError(f"Failed to set value on {cell_or_range}: {e}")

    def get_formula(self, cell_or_range: str) -> Any:
        try:
            r1, c1, r2, c2 = parse_range_address(cell_or_range)
            if r1 == r2 and c1 == c2:
                val = self._ws.cell(row=r1, column=c1).value
                return str(val) if isinstance(val, str) and val.startswith('=') else None

            res = []
            for r in range(r1, r2 + 1):
                row_f = []
                for c in range(c1, c2 + 1):
                    val = self._ws.cell(row=r, column=c).value
                    row_f.append(str(val) if isinstance(val, str) and val.startswith('=') else None)
                res.append(row_f)
            return res
        except Exception as e:
            raise ActionExecutionError(f"Failed to get formula from {cell_or_range}: {e}")

    def set_formula(self, cell_or_range: str, formula: str) -> None:
        if not formula.startswith('='):
            formula = '=' + formula
        r1, c1, r2, c2 = parse_range_address(cell_or_range)
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                self._ws.cell(row=r, column=c).value = formula

    def get_column_values(self, col: Union[int, str], include_header: bool = False) -> List[Any]:
        col_idx = col if isinstance(col, int) else col_letter_to_index(col)
        min_r, _, max_r, _ = self.get_used_range_bounds()
        start_r = min_r if include_header else min_r + 1
        if start_r > max_r:
            return []
        values = []
        for r in range(start_r, max_r + 1):
            values.append(self._ws.cell(row=r, column=col_idx).value)
        return values

    def set_cell_color(self, cell_or_range: str, color_hex: str) -> None:
        try:
            hex_clean = normalize_hex_color(color_hex)
            fill = PatternFill(start_color=f"FF{hex_clean}", end_color=f"FF{hex_clean}", fill_type="solid")
            r1, c1, r2, c2 = parse_range_address(cell_or_range)
            _, _, ub_r2, ub_c2 = self.get_used_range_bounds()
            r2 = min(r2, max(ub_r2, r1))
            c2 = min(c2, max(ub_c2, c1))
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    self._ws.cell(row=r, column=c).fill = fill
        except Exception as e:
            raise ActionExecutionError(f"Failed to set cell color on {cell_or_range}: {e}")

    def set_font(self, cell_or_range: str, name: Optional[str] = None, size: Optional[float] = None,
                 bold: Optional[bool] = None, italic: Optional[bool] = None, color_hex: Optional[str] = None) -> None:
        try:
            r1, c1, r2, c2 = parse_range_address(cell_or_range)
            _, _, ub_r2, ub_c2 = self.get_used_range_bounds()
            r2 = min(r2, max(ub_r2, r1))
            c2 = min(c2, max(ub_c2, c1))
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    cell = self._ws.cell(row=r, column=c)
                    current_font = cell.font or Font()
                    cell.font = Font(
                        name=name or current_font.name,
                        size=size or current_font.size,
                        bold=bold if bold is not None else current_font.bold,
                        italic=italic if italic is not None else current_font.italic,
                        color=f"FF{normalize_hex_color(color_hex)}" if color_hex else current_font.color
                    )
        except Exception as e:
            raise ActionExecutionError(f"Failed to set font on {cell_or_range}: {e}")

    def auto_fit_columns(self, cols: Optional[List[Union[int, str]]] = None) -> None:
        min_r, min_c, max_r, max_c = self.get_used_range_bounds()
        if cols:
            target_indices = [c if isinstance(c, int) else col_letter_to_index(c) for c in cols]
        else:
            target_indices = list(range(min_c, max_c + 1))

        for col_idx in target_indices:
            col_letter = get_column_letter(col_idx)
            max_len = 0
            for r in range(min_r, max_r + 1):
                val = self._ws.cell(row=r, column=col_idx).value
                if val is not None:
                    max_len = max(max_len, len(str(val)))
            self._ws.column_dimensions[col_letter].width = max(max_len + 4, 10)

    def delete_rows(self, row_indices: List[int]) -> int:
        if not row_indices:
            return 0
        sorted_rows = sorted(set(row_indices), reverse=True)
        count = 0
        for r in sorted_rows:
            self._ws.delete_rows(r, 1)
            count += 1
        return count

    def insert_rows(self, row_index: int, count: int = 1) -> None:
        self._ws.insert_rows(row_index, amount=count)

    def delete_cols(self, col_indices: List[Union[int, str]]) -> int:
        if not col_indices:
            return 0
        indices = [c if isinstance(c, int) else col_letter_to_index(c) for c in col_indices]
        sorted_cols = sorted(set(indices), reverse=True)
        count = 0
        for c in sorted_cols:
            self._ws.delete_cols(c, 1)
            count += 1
        return count

    def sort_range(self, range_str: str, key_col: Union[int, str], ascending: bool = True, has_headers: bool = True) -> None:
        r1, c1, r2, c2 = parse_range_address(range_str)
        key_idx = (key_col if isinstance(key_col, int) else col_letter_to_index(key_col)) - c1
        start_r = (r1 + 1) if has_headers else r1
        if start_r >= r2:
            return

        # Read row data and cell styles
        row_cells = []
        for r in range(start_r, r2 + 1):
            row_vals = [self._ws.cell(row=r, column=c).value for c in range(c1, c2 + 1)]
            row_cells.append(row_vals)

        def sort_key(item):
            val = item[key_idx] if key_idx < len(item) else None
            if val is None:
                return (1, "")
            if isinstance(val, (int, float)):
                return (0, float(val))
            return (0, str(val).lower())

        row_cells.sort(key=sort_key, reverse=not ascending)

        # Write sorted rows back
        for r_offset, row_vals in enumerate(row_cells):
            curr_r = start_r + r_offset
            for c_offset, val in enumerate(row_vals):
                self._ws.cell(row=curr_r, column=c1 + c_offset).value = val

    def create_table(self, range_str: str, name: Optional[str] = None, style: str = "TableStyleMedium9") -> str:
        table_name = name or f"Table_{self.name}_{len(self._ws.tables) + 1}"
        table = Table(displayName=table_name, ref=range_str)
        table.tableStyleInfo = TableStyleInfo(name=style, showFirstColumn=False, showLastColumn=False,
                                              showRowStripes=True, showColumnStripes=False)
        self._ws.add_table(table)
        return table_name


class OpenpyxlDriver(IExcelDriver):
    """Driver implementation using openpyxl for standalone/testing Excel operations."""

    def __init__(self, file_path: Optional[str] = None):
        self.file_path = file_path
        self._wb: Optional[openpyxl.Workbook] = None
        self._active_sheet_name: Optional[str] = None
        if file_path and os.path.exists(file_path):
            self.connect(file_path)

    def connect(self, workbook_name_or_path: Optional[str] = None) -> WorkbookInfo:
        target_path = workbook_name_or_path or self.file_path
        if target_path and os.path.exists(target_path):
            self._wb = openpyxl.load_workbook(target_path, data_only=False)
            self.file_path = target_path
        else:
            self._wb = openpyxl.Workbook()
            if target_path:
                self.file_path = target_path
                self._wb.save(target_path)
            else:
                self.file_path = "Book1.xlsx"

        self._active_sheet_name = self._wb.active.title
        return self.get_workbook_info()

    def is_connected(self) -> bool:
        return self._wb is not None

    def list_open_workbooks(self) -> List[str]:
        if self.file_path:
            return [os.path.basename(self.file_path)]
        return ["Book1.xlsx"] if self._wb else []

    def get_workbook_info(self) -> WorkbookInfo:
        if not self._wb:
            raise WorkbookNotFoundError("No workbook is open.")
        name = os.path.basename(self.file_path) if self.file_path else "Book1.xlsx"
        sheets = self.get_sheet_names()
        active = self._active_sheet_name or (sheets[0] if sheets else "")
        return WorkbookInfo(
            name=name,
            path=self.file_path,
            is_saved=True,
            sheets=sheets,
            active_sheet=active,
            app_version="Excel 2024 (Standalone Engine)"
        )

    def get_sheet_names(self) -> List[str]:
        if not self._wb:
            raise WorkbookNotFoundError("No workbook is open.")
        return list(self._wb.sheetnames)

    def get_active_sheet(self) -> ISheet:
        if not self._wb:
            raise WorkbookNotFoundError("No workbook is open.")
        if self._active_sheet_name and self._active_sheet_name in self._wb.sheetnames:
            ws = self._wb[self._active_sheet_name]
        else:
            ws = self._wb.active
            self._active_sheet_name = ws.title
        return OpenpyxlSheet(ws)

    def get_sheet(self, name: str) -> ISheet:
        if not self._wb:
            raise WorkbookNotFoundError("No workbook is open.")
        if name not in self._wb.sheetnames:
            raise SheetNotFoundError(f"Sheet '{name}' not found in workbook.")
        return OpenpyxlSheet(self._wb[name])

    def activate_sheet(self, name: str) -> ISheet:
        if not self._wb:
            raise WorkbookNotFoundError("No workbook is open.")
        if name not in self._wb.sheetnames:
            raise SheetNotFoundError(f"Sheet '{name}' not found in workbook.")
        self._active_sheet_name = name
        self._wb.active = self._wb[name]
        return OpenpyxlSheet(self._wb[name])

    def create_sheet(self, name: str) -> ISheet:
        if not self._wb:
            raise WorkbookNotFoundError("No workbook is open.")
        ws = self._wb.create_sheet(title=name)
        self._active_sheet_name = name
        return OpenpyxlSheet(ws)

    def save(self, path: Optional[str] = None) -> str:
        if not self._wb:
            raise WorkbookNotFoundError("No workbook is open.")
        target = path or self.file_path or "Book1.xlsx"
        self._wb.save(target)
        self.file_path = target
        return target

    def close(self) -> None:
        if self._wb:
            try:
                self._wb.close()
            except Exception:
                pass
            self._wb = None

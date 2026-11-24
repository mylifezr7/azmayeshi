"""
Backup and Snapshot Engine for Excel Workbooks.
Provides atomic snapshots for undo functionality and safety backups before large mutations.
"""
import copy
import os
import shutil
import tempfile
import time
from typing import Any, Dict, List, Optional

from excel_assistant.core.interfaces import ISheet, IExcelDriver


class SheetSnapshot:
    """In-memory snapshot of a worksheet's values, formulas, and used range."""

    def __init__(self, sheet: ISheet):
        self.sheet_name = sheet.name
        self.bounds = sheet.get_used_range_bounds()
        min_r, min_c, max_r, max_c = self.bounds
        # Fast single batch read
        self.matrix = sheet.get_matrix(min_r, min_c, max_r, max_c)
        self.formula_matrix = sheet.get_formula_matrix(min_r, min_c, max_r, max_c)

    def get_cell(self, r: int, c: int) -> tuple[Any, Optional[str]]:
        """Returns (value, formula) for cell at 1-based (r, c)."""
        min_r, min_c, max_r, max_c = self.bounds
        if min_r <= r <= max_r and min_c <= c <= max_c:
            r_idx = r - min_r
            c_idx = c - min_c
            val = None
            if r_idx < len(self.matrix):
                row = self.matrix[r_idx]
                val = row[c_idx] if c_idx < len(row) else None
            formula = None
            if r_idx < len(self.formula_matrix):
                f_row = self.formula_matrix[r_idx]
                formula = f_row[c_idx] if c_idx < len(f_row) else None
            return val, formula
        return None, None

    def restore_to(self, sheet: ISheet) -> None:
        """Restores values and formulas into the sheet without wiping unrelated manual cells."""
        from excel_assistant.core.utils import col_index_to_letter

        # Restore from snapshotted matrix without clearing unrelated external user cells
        orig_min_r, orig_min_c, orig_max_r, orig_max_c = self.bounds
        for r_offset, row in enumerate(self.matrix):
            r = orig_min_r + r_offset
            f_row = self.formula_matrix[r_offset] if r_offset < len(self.formula_matrix) else []
            for c_offset, val in enumerate(row):
                c = orig_min_c + c_offset
                addr = f"{col_index_to_letter(c)}{r}"
                formula = f_row[c_offset] if c_offset < len(f_row) else None
                if formula:
                    sheet.set_formula(addr, formula)
                else:
                    sheet.set_value(addr, val)


class BackupEngine:
    """Manages file-based and in-memory snapshots for undo and disaster recovery."""

    def __init__(self, backup_dir: Optional[str] = None):
        self.backup_dir = backup_dir or os.path.join(tempfile.gettempdir(), "excel_ai_backups")
        os.makedirs(self.backup_dir, exist_ok=True)

    def create_file_backup(self, driver: IExcelDriver) -> Optional[str]:
        """Creates an on-disk snapshot of the current workbook."""
        try:
            info = driver.get_workbook_info()
            if info.path and os.path.exists(info.path):
                timestamp = time.strftime("%Y%m%d_%H%M%S")
                base = os.path.basename(info.path)
                backup_name = f"backup_{timestamp}_{base}"
                dest = os.path.join(self.backup_dir, backup_name)
                shutil.copy2(info.path, dest)
                return dest
            else:
                # Save to a temporary backup file
                timestamp = time.strftime("%Y%m%d_%H%M%S")
                dest = os.path.join(self.backup_dir, f"backup_{timestamp}_{info.name}")
                driver.save(dest)
                return dest
        except Exception:
            return None

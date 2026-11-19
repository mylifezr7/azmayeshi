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
        self.cell_values: Dict[Tuple[int, int], Any] = {}
        self.cell_formulas: Dict[Tuple[int, int], Any] = {}

        for r in range(min_r, max_r + 1):
            for c in range(min_c, max_c + 1):
                from excel_assistant.core.utils import col_index_to_letter
                addr = f"{col_index_to_letter(c)}{r}"
                val = sheet.get_value(addr)
                formula = sheet.get_formula(addr)
                self.cell_values[(r, c)] = copy.deepcopy(val)
                self.cell_formulas[(r, c)] = formula

    def restore_to(self, sheet: ISheet) -> None:
        """Restores values and formulas into the sheet."""
        # Clear currently used range values
        curr_bounds = sheet.get_used_range_bounds()
        min_r, min_c, max_r, max_c = curr_bounds
        from excel_assistant.core.utils import col_index_to_letter

        for r in range(min_r, max_r + 1):
            for c in range(min_c, max_c + 1):
                addr = f"{col_index_to_letter(c)}{r}"
                sheet.set_value(addr, None)

        # Restore snapshotted cells
        for (r, c), formula in self.cell_formulas.items():
            addr = f"{col_index_to_letter(c)}{r}"
            if formula:
                sheet.set_formula(addr, formula)
            else:
                val = self.cell_values.get((r, c))
                if val is not None:
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

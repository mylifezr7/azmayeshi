"""
Undo and Transaction Manager for Excel AI Assistant.
Maintains granular history of cell deltas and snapshots, providing safe rollbacks
that never wipe out unrelated manual data entered by the user.
"""
from dataclasses import dataclass
from typing import Any, List, Optional
from excel_assistant.core.interfaces import ISheet
from excel_assistant.safety.backup_engine import SheetSnapshot


@dataclass
class CellChange:
    row: int
    col: int
    addr: str
    old_value: Any
    old_formula: Optional[str]
    new_value: Any
    new_formula: Optional[str]


@dataclass
class ActionUndoRecord:
    sheet_name: str
    changes: List[CellChange]
    snapshot_before: Optional[SheetSnapshot] = None


class UndoManager:
    """Manages granular undo history per worksheet."""

    def __init__(self, max_history: int = 20):
        self.max_history = max_history
        self._history: List[ActionUndoRecord] = []
        self._pending_snapshot: Optional[SheetSnapshot] = None

    def begin_action(self, sheet: ISheet) -> None:
        """Takes a pre-action snapshot before an action begins."""
        try:
            self._pending_snapshot = SheetSnapshot(sheet)
        except Exception:
            self._pending_snapshot = None

    def commit_action(self, sheet: ISheet) -> None:
        """Compares post-action state with pre-action state and records exact cell deltas."""
        if not self._pending_snapshot:
            return
        try:
            snapshot_after = SheetSnapshot(sheet)
            changes = self._compute_delta(self._pending_snapshot, snapshot_after)
            if self._history and self._history[-1].snapshot_before is self._pending_snapshot:
                # Update record previously added by save_state
                self._history[-1].changes = changes
            else:
                record = ActionUndoRecord(
                    sheet_name=sheet.name,
                    changes=changes,
                    snapshot_before=self._pending_snapshot
                )
                self._history.append(record)
                if len(self._history) > self.max_history:
                    self._history.pop(0)
        except Exception:
            pass
        finally:
            self._pending_snapshot = None

    def discard_pending(self) -> None:
        """Discards pending pre-action snapshot if action was aborted or failed."""
        self._pending_snapshot = None

    def save_state(self, sheet: ISheet) -> None:
        """Takes a snapshot of sheet state before an action (backward compatible)."""
        try:
            snapshot = SheetSnapshot(sheet)
            self._pending_snapshot = snapshot
            record = ActionUndoRecord(
                sheet_name=sheet.name,
                changes=[],
                snapshot_before=snapshot
            )
            self._history.append(record)
            if len(self._history) > self.max_history:
                self._history.pop(0)
        except Exception:
            pass

    def _compute_delta(self, before: SheetSnapshot, after: SheetSnapshot) -> List[CellChange]:
        from excel_assistant.core.utils import col_index_to_letter

        changes: List[CellChange] = []
        b_min_r, b_min_c, b_max_r, b_max_c = before.bounds
        a_min_r, a_min_c, a_max_r, a_max_c = after.bounds

        min_r = min(b_min_r, a_min_r)
        min_c = min(b_min_c, a_min_c)
        max_r = max(b_max_r, a_max_r)
        max_c = max(b_max_c, a_max_c)

        for r in range(min_r, max_r + 1):
            for c in range(min_c, max_c + 1):
                val_b, f_b = before.get_cell(r, c)
                val_a, f_a = after.get_cell(r, c)

                if val_b != val_a or f_b != f_a:
                    addr = f"{col_index_to_letter(c)}{r}"
                    changes.append(CellChange(
                        row=r,
                        col=c,
                        addr=addr,
                        old_value=val_b,
                        old_formula=f_b,
                        new_value=val_a,
                        new_formula=f_a
                    ))
        return changes

    def can_undo(self) -> bool:
        return len(self._history) > 0

    def undo(self, sheet: ISheet) -> bool:
        """Restores only the cells that were changed by the last action, preserving manual edits elsewhere."""
        if not self._history:
            return False
        record = self._history.pop()
        if record.changes:
            for change in record.changes:
                addr = change.addr
                if change.old_formula:
                    sheet.set_formula(addr, change.old_formula)
                else:
                    sheet.set_value(addr, change.old_value)
            return True
        elif record.snapshot_before:
            record.snapshot_before.restore_to(sheet)
            return True
        return False

    def clear(self) -> None:
        self._history.clear()
        self._pending_snapshot = None

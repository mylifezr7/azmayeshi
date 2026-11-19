"""
Undo and Transaction Manager for Excel AI Assistant.
Maintains history of snapshots and provides rollbacks on demand.
"""
from typing import List, Optional
from excel_assistant.core.interfaces import ISheet
from excel_assistant.safety.backup_engine import SheetSnapshot


class UndoManager:
    """Manages undo history per worksheet."""

    def __init__(self, max_history: int = 20):
        self.max_history = max_history
        self._history: List[SheetSnapshot] = []

    def save_state(self, sheet: ISheet) -> None:
        """Takes a snapshot of sheet state before an action."""
        try:
            snapshot = SheetSnapshot(sheet)
            self._history.append(snapshot)
            if len(self._history) > self.max_history:
                self._history.pop(0)
        except Exception:
            pass

    def can_undo(self) -> bool:
        return len(self._history) > 0

    def undo(self, sheet: ISheet) -> bool:
        """Restores the last state."""
        if not self._history:
            return False
        snapshot = self._history.pop()
        snapshot.restore_to(sheet)
        return True

    def clear(self) -> None:
        self._history.clear()

"""
Security and Safety Layer for Excel Assistant.
Validates boundaries, classifies action risks, and manages user confirmations for destructive operations.
"""
from enum import Enum
from typing import Optional, Tuple

from excel_assistant.nlp.intent_models import BaseActionModel, ExcelAction
from excel_assistant.core.exceptions import SafetyConfirmationRequiredError


class RiskLevel(Enum):
    LOW = "low"          # Read, stats, coloring, font change
    MEDIUM = "medium"    # Writing formulas, adding columns, autofit, sorting
    HIGH = "high"        # Deleting rows, deleting columns, deleting duplicates, clearing


class SecurityPolicy:
    """Enforces safety rules before mutations occur."""

    @classmethod
    def evaluate_risk(cls, action: ExcelAction) -> RiskLevel:
        if getattr(action, "is_destructive", False):
            return RiskLevel.HIGH
        action_type = action.action_type
        if action_type in ("delete_empty_rows", "delete_duplicates", "delete_rows_or_cols"):
            return RiskLevel.HIGH
        if action_type in ("column_arithmetic", "sort_data", "apply_formula"):
            return RiskLevel.MEDIUM
        return RiskLevel.LOW

    @classmethod
    def check_permission(cls, action: ExcelAction, confirmed: bool = False) -> None:
        """
        Raises SafetyConfirmationRequiredError if the action is destructive and not explicitly confirmed.
        """
        risk = cls.evaluate_risk(action)
        if risk == RiskLevel.HIGH and not confirmed:
            desc = getattr(action, "description_persian", action.action_type)
            raise SafetyConfirmationRequiredError(
                action_name=action.action_type,
                details=f"عملیات «{desc}» غیرقابل بازگشت مستقیم روی برخی داده‌هاست و نیازمند تأیید شماست.",
                affected_range=getattr(action, "affected_range", "")
            )

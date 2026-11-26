"""
Tests for Command Search Engine and Search UI Integration.
"""
import pytest
from excel_assistant.ui.command_search import search_commands, COMMAND_SEARCH_REGISTRY, normalize_persian_search_term
from excel_assistant.nlp.parser import PersianIntentParser
from excel_assistant.nlp.intent_models import ProfitCalculationAction, AttendanceAction


def test_search_profit_annual():
    """Verify searching 'سود' returns 'محاسبه سود سالانه' at the very top."""
    results = search_commands("سود", limit=10)
    assert len(results) > 0
    # Top result should be محاسبه سود سالانه
    top_title = results[0]["title"]
    top_cmd = results[0]["cmd"]
    assert "سود سالانه" in top_title
    assert top_cmd == "محاسبه سود سالانه"

    # Verify parser recognizes the command
    parser = PersianIntentParser()
    action = parser.parse(top_cmd)
    assert isinstance(action, ProfitCalculationAction)
    assert action.period == "yearly"
    assert action.valuation_method == "fifo"


def test_search_profit_explicit_query():
    """Verify searching 'سود سالانه' matches 'محاسبه سود سالانه'."""
    results = search_commands("سود سالانه", limit=5)
    assert len(results) > 0
    assert results[0]["cmd"] == "محاسبه سود سالانه"


def test_search_attendance_timesheet():
    """Verify searching 'کارکرد' or 'ورود و خروج' surfaces attendance timesheet calculation."""
    for q in ["کارکرد", "ورود و خروج", "حضور و غیاب", "ساعت کاری"]:
        res = search_commands(q, limit=5)
        assert len(res) > 0
        assert res[0]["cmd"] == "محاسبه ساعت کاری"

        parser = PersianIntentParser()
        action = parser.parse(res[0]["cmd"])
        assert isinstance(action, AttendanceAction)


def test_search_sales():
    """Verify searching 'فروش' surfaces sales calculation."""
    res = search_commands("فروش", limit=5)
    assert len(res) > 0
    titles = [r["title"] for r in res]
    assert any("فروش" in t for t in titles)
    assert any(r["cmd"] == "محاسبه فروش" for r in res)


def test_search_filter_and_sort():
    """Verify searching 'فیلتر' and 'مرتب سازی' surfaces table operations."""
    res_filter = search_commands("فیلتر", limit=5)
    assert any("فیلتر" in r["title"] for r in res_filter)

    res_sort = search_commands("مرتب سازی", limit=5)
    assert any("مرتب" in r["title"] for r in res_sort)


def test_search_lock_unique():
    """Verify searching 'قفل' or 'یکتا' surfaces unique column locking."""
    res_lock = search_commands("قفل", limit=5)
    assert any("قفل یکتاسازی" in r["title"] for r in res_lock)

    res_unique = search_commands("یکتا", limit=5)
    assert any("قفل یکتاسازی" in r["title"] for r in res_unique)


def test_search_empty_and_spaces():
    """Verify empty or whitespace queries return empty results."""
    assert search_commands("") == []
    assert search_commands("   ") == []
    assert search_commands(None) == []


def test_command_registry_integrity():
    """Verify all items in the command registry have valid categories, commands, and keywords."""
    valid_categories = {"ضروری", "حسابداری", "فروش و انبارداری", "هوشمند", "عمومی"}
    for item in COMMAND_SEARCH_REGISTRY:
        assert "title" in item and len(item["title"]) > 0
        assert "cmd" in item and len(item["cmd"]) > 0
        assert "category" in item and item["category"] in valid_categories
        assert "keywords" in item and isinstance(item["keywords"], list)
        assert len(item["keywords"]) > 0

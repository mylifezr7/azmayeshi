"""
Interactive Persian CLI for Excel AI Assistant.
Allows running the assistant in terminal environments or headless servers.
"""
import sys
from typing import Optional

from excel_assistant.core.connection import ExcelConnectionManager
from excel_assistant.analyzer.inspector import SheetInspector
from excel_assistant.nlp.parser import PersianIntentParser
from excel_assistant.actions.base import ActionContext
from excel_assistant.actions.executor import ActionExecutor
from excel_assistant.actions.undo_manager import UndoManager
from excel_assistant.core.exceptions import AmbiguousCommandError


def run_cli(workbook_path: Optional[str] = None):
    print("=" * 60)
    print("  🟢 دستیار هوشمند Microsoft Excel (Excel AI Assistant)")
    print("=" * 60)

    manager = ExcelConnectionManager()
    try:
        info = manager.connect(workbook_path)
        print(manager.get_status_persian())
    except Exception as e:
        print(f"❌ خطا در اتصال به اکسل: {e}")
        return

    parser = PersianIntentParser()
    undo_mgr = UndoManager()
    executor = ActionExecutor(undo_mgr)

    print("\n💡 راهنما: دستورات خود را به زبان فارسی بنویسید (مثلاً: «ستون A را در ستون B ضرب کن»)")
    print("برای خروج، عبارت «خروج» یا «exit» را وارد کنید.\n")

    while True:
        try:
            cmd = input("👤 دستور شما > ").strip()
            if not cmd:
                continue
            if cmd.lower() in ("خروج", "exit", "quit", "q"):
                print("خداحافظ! 👋")
                break

            sheet = manager.get_active_sheet()
            sa = SheetInspector.inspect_sheet(sheet)

            try:
                action = parser.parse(cmd, sheet_analysis=sa)
            except AmbiguousCommandError as ace:
                print(f"⚠️ {ace}")
                if ace.suggested_options:
                    print(ace.clarification_prompt)
                    for opt in ace.suggested_options:
                        print(f"  • {opt}")
                continue

            ctx = ActionContext(manager=manager, sheet=sheet, sheet_analysis=sa, confirmed=False)
            result = executor.execute(action, ctx)

            if result.requires_confirmation:
                print(f"\n⚠️ {result.confirmation_prompt}")
                ans = input("آیا مایل به انجام عملیات هستید؟ (بله / خیر): ").strip().lower()
                if ans in ("بله", "yes", "y", "آره"):
                    ctx.confirmed = True
                    result = executor.execute(action, ctx)
                else:
                    print("عملیات لغو شد.")
                    continue

            print("\n" + result.message_persian + "\n")

        except (KeyboardInterrupt, EOFError):
            print("\nخروج.")
            break
        except Exception as e:
            print(f"❌ خطای غیرمنتظره: {e}")


if __name__ == "__main__":
    run_cli()

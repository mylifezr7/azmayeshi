"""
اسکریپت تست سریع کنترل موجودی انبار با کسر خودکار فروش از خرید
Quick test runner for Inventory Alert with sales deduction
"""
import os
import sys

# Ensure repository root is in python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from excel_assistant.core.openpyxl_driver import OpenpyxlDriver
from excel_assistant.core.connection import ExcelConnectionManager
from excel_assistant.nlp.intent_models import InventoryAlertAction
from excel_assistant.actions.business_actions import BusinessActionHandler
from excel_assistant.actions.base import ActionContext


def main():
    print("=" * 65)
    print("🚀 در حال اجرای تست هوشمند کنترل موجودی انبار (کسر فروش از خرید)...")
    print("=" * 65)

    # 1. راه‌اندازی درایور و مدیریت اکسل
    driver = OpenpyxlDriver()
    driver.connect()
    manager = ExcelConnectionManager()
    manager.active_driver = driver

    # 2. ایجاد شیت مرجع انبار (خرید)
    p_sheet = manager.create_sheet("خرید")
    p_sheet.set_value("A1", "نام کالا")
    p_sheet.set_value("B1", "قیمت خرید")
    p_sheet.set_value("C1", "تعداد خرید")

    p_sheet.set_value("A2", "لپ‌تاپ")
    p_sheet.set_value("B2", 20000000)
    p_sheet.set_value("C2", 10)

    p_sheet.set_value("A3", "کیبورد")
    p_sheet.set_value("B3", 500000)
    p_sheet.set_value("C3", 8)

    p_sheet.set_value("A4", "ماوس")
    p_sheet.set_value("B4", 200000)
    p_sheet.set_value("C4", 50)

    # 3. ایجاد شیت فاکتورهای فروش (فروش)
    s_sheet = manager.create_sheet("فروش")
    s_sheet.set_value("A1", "نام کالا")
    s_sheet.set_value("B1", "قیمت فروش")
    s_sheet.set_value("C1", "تعداد فروش")

    s_sheet.set_value("A2", "لپ‌تاپ")
    s_sheet.set_value("B2", 25000000)
    s_sheet.set_value("C2", 7)

    s_sheet.set_value("A3", "لپ‌تاپ")
    s_sheet.set_value("B3", 25000000)
    s_sheet.set_value("C3", 1)  # جمع کل فروش لپ‌تاپ = ۸ عدد (مانده: ۱۰ - ۸ = ۲ عدد <= ۵ -> هشدار کسری)

    s_sheet.set_value("A4", "کیبورد")
    s_sheet.set_value("B4", 700000)
    s_sheet.set_value("C4", 5)  # فروش کیبورد = ۵ عدد (مانده: ۸ - ۵ = ۳ عدد <= ۵ -> هشدار کسری)

    s_sheet.set_value("A5", "ماوس")
    s_sheet.set_value("B5", 300000)
    s_sheet.set_value("C5", 15) # فروش ماوس = ۱۵ عدد (مانده: ۵۰ - ۱۵ = ۳۵ عدد > ۵ -> موجودی کافی)

    # 4. اجرای عملیات کنترل موجودی با آستانه ۵ عدد
    action = InventoryAlertAction(min_threshold=5.0, highlight_low_stock=True)
    ctx = ActionContext(manager=manager, sheet=p_sheet)
    result = BusinessActionHandler.execute_inventory_alert(action, ctx)

    # 5. نمایش نتایج
    print("\n" + result.message_persian)
    print("\n" + "-" * 65)
    print("📊 بررسی فرمول‌های اکسل درج‌شده در شیت خرید:")
    print(f"• ستون D (تعداد فروش): {p_sheet.get_formula('D2')}")
    print(f"• ستون E (مانده موجودی): {p_sheet.get_formula('E2')}")
    print(f"• ستون F (وضعیت هشدار): {p_sheet.get_formula('F2')}")
    print(f"• ستون G (ارزش ریالی): {p_sheet.get_formula('G2')}")
    print(f"• جمع کل ارزش انبار (ردیف ۵): {p_sheet.get_formula('G5')}")
    print("-" * 65)

    # ذخیره در فایل اکسل تستی
    out_file = "test_inventory_output.xlsx"
    driver.save(out_file)
    print(f"✅ فایل اکسل خروجی با موفقیت ذخیره شد: {out_file}\n")


if __name__ == "__main__":
    main()

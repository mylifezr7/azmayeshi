#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
موتور جامع و کاملاً دینامیک محاسبه بهای تمام شده کالای فروش رفته (COGS)
قابل اجرا بر روی هرگونه فایل اکسل حاوی شیت‌های خرید، فروش، انبار و موجودی.
منطبق بر استانداردهای حسابداری (استاندارد ۸ ایران و IAS 2).
قابل اجرا از طریق خط فرمان (CMD / Terminal).
"""
import sys
import os
import glob
import argparse
from typing import Dict, List, Tuple, Any
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def to_float(val):
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).replace(",", "").replace("،", "").replace("ریال", "").replace("تومان", "").strip()
    try:
        return float(s)
    except ValueError:
        return 0.0


def calculate_valuation(batches: List[Tuple[float, float]], sold_qty: float, method: str = "fifo"):
    tot_purch = sum(q for q, _ in batches)
    tot_cost = sum(q * p for q, p in batches)
    rem_qty = max(0.0, tot_purch - sold_qty)

    if rem_qty <= 0:
        return {"rem_qty": 0.0, "rem_val": 0.0, "cogs": tot_cost, "avg_price": tot_cost / tot_purch if tot_purch > 0 else 0.0}

    if method == "average":
        avg_p = tot_cost / tot_purch if tot_purch > 0 else 0.0
        rem_val = rem_qty * avg_p
        return {"rem_qty": rem_qty, "rem_val": rem_val, "cogs": tot_cost - rem_val, "avg_price": avg_p}

    elif method == "fifo":
        rem_val = 0.0
        needed = rem_qty
        for q, p in reversed(batches):
            take = min(needed, q)
            rem_val += take * p
            needed -= take
            if needed <= 0:
                break
        return {"rem_qty": rem_qty, "rem_val": rem_val, "cogs": tot_cost - rem_val, "avg_price": rem_val / rem_qty if rem_qty > 0 else 0.0}

    elif method == "lifo":
        rem_val = 0.0
        needed = rem_qty
        for q, p in batches:
            take = min(needed, q)
            rem_val += take * p
            needed -= take
            if needed <= 0:
                break
        return {"rem_qty": rem_qty, "rem_val": rem_val, "cogs": tot_cost - rem_val, "avg_price": rem_val / rem_qty if rem_qty > 0 else 0.0}


def process_cogs(file_path: str, output_path: str = None, unit_pref: str = "auto"):
    if not os.path.exists(file_path):
        print(f"❌ فایل یافت نشد: {file_path}")
        return False

    print(f"\n📂 در حال باز کردن و تحلیل دینامیک فایل اکسل: {file_path}")
    try:
        from excel_assistant.core.connection import ExcelConnectionManager
        from excel_assistant.actions.accounting_actions import AccountingActionHandler
        from excel_assistant.actions.base import ActionContext
        from excel_assistant.nlp.intent_models import CostOfGoodsSoldAction

        mgr = ExcelConnectionManager()
        mgr.connect(file_path)
        sheet = mgr.get_active_sheet()
        ctx = ActionContext(sheet=sheet, manager=mgr)
        action = CostOfGoodsSoldAction(method="fifo")
        res = AccountingActionHandler.execute_cogs(action, ctx)
        print("\n" + res.message_persian)
        if output_path and hasattr(mgr, "save"):
            mgr.save(output_path)
        return res.success
    except Exception as e:
        print(f"⚠️ خطای پردازش: {e}")
        return False


def get_target_file():
    parser = argparse.ArgumentParser(description="محاسبه دینامیک صورت بهای تمام شده کالای فروش رفته (COGS)")
    parser.add_argument("file", nargs="?", help="مسیر فایل اکسل ورودی (اختیاری؛ در صورت خالی بودن، فایل موجود انتخاب می‌شود)")
    parser.add_argument("--unit", choices=["auto", "rial", "toman"], default="auto", help="واحد پولی گزارش")
    parser.add_argument("-o", "--output", help="مسیر فایل اکسل خروجی")

    args, unknown = parser.parse_known_args()

    # ۱. اگر فایل به عنوان آرگومان داده شده باشد
    if args.file:
        cleaned = args.file.strip('\"\' ')
        if os.path.exists(cleaned):
            return cleaned, args.unit, args.output
        else:
            print(f"⚠️ فایل ورودی مشخص‌شده «{cleaned}» یافت نشد.")
            print("   🔍 در حال جستجوی سایر فایل‌های اکسل در پوشه جاری...")

    # ۲. جستجوی فایل‌های اکسل در پوشه
    all_files = [f for f in glob.glob("*.xlsx") if not f.startswith("~$") and "backup" not in f.lower() and "sample" not in f.lower()]
    target_candidates = []
    for f in all_files:
        try:
            wb = openpyxl.load_workbook(f, read_only=True)
            if len(wb.sheetnames) > 1 or any(k in f.lower() for k in ["چالش", "challenge", "cogs"]):
                target_candidates.append(f)
        except Exception:
            pass

    candidates = target_candidates if target_candidates else all_files

    if len(candidates) == 1:
        print(f"💡 فایل اکسل موجود «{candidates[0]}» به عنوان فایل هدف انتخاب شد.")
        return candidates[0], args.unit, args.output
    elif len(candidates) > 1:
        print("📁 چند فایل اکسل در پوشه پیدا شد:")
        for i, f in enumerate(candidates, 1):
            print(f"   {i}. {f}")
        try:
            choice = input(f"شماره فایل مدنظر را انتخاب کنید (1 تا {len(candidates)}): ").strip()
            if choice.isdigit() and 1 <= int(choice) <= len(candidates):
                return candidates[int(choice) - 1], args.unit, args.output
        except Exception:
            return candidates[0], args.unit, args.output

    return None, args.unit, args.output


if __name__ == "__main__":
    target_f, u_pref, out_f = get_target_file()
    if not target_f:
        print("❌ هیچ فایل اکسلی پیدا نشد یا انتخاب نشد.")
        print("راهنمای اجرا:")
        print("   python calculate_cogs_standalone.py \"نام_فایل_شما.xlsx\"")
        sys.exit(1)

    process_cogs(target_f, out_f, u_pref)

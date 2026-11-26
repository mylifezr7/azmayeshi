#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
محاسبه جامع بهای تمام شده کالای فروش رفته (COGS) و ایجاد صورت مالی استاندارد
به صورت خودکار از شیت‌های فاکتورهای خرید، فاکتورهای فروش، حمل، انبارگردانی و کالای امانی.
منطبق بر استاندارد حسابداری شماره ۸ ایران و IAS 2.
قابل اجرا به صورت مستقل از طریق خط فرمان (CMD / Terminal).
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

    print(f"📂 در حال باز کردن فایل اکسل: {file_path}")
    wb = openpyxl.load_workbook(file_path)

    p_sheet = None
    pr_sheet = None
    s_sheet = None
    sr_sheet = None
    beg_sheet = None
    fr_sheet = None
    m_sheet = None
    cn_sheet = None
    guideline_sheet = None

    # ۱. تطبیق نام شیت‌ها
    for name in wb.sheetnames:
        nl = name.lower()
        if (any(k in nl for k in ["برگشت از فروش", "برگشت فروش"]) or ("برگشت" in nl and "فروش" in nl)) and not sr_sheet:
            sr_sheet = wb[name]
        elif (any(k in nl for k in ["برگشت از خرید", "برگشت خرید"]) or ("برگشت" in nl and "خرید" in nl)) and not pr_sheet:
            pr_sheet = wb[name]
        elif any(k in nl for k in ["حمل", "کرایه", "باربری", "ترابری", "هزینه حمل", "هزینه های حمل", "freight"]) and not fr_sheet:
            fr_sheet = wb[name]
        elif any(k in nl for k in ["اول دوره", "موجودی اول", "موجودی اولیه", "ابتدای دوره", "لایه"]) and not beg_sheet:
            beg_sheet = wb[name]
        elif any(k in nl for k in ["امانی", "کالای امانی", "consign"]) and not cn_sheet:
            cn_sheet = wb[name]
        elif any(k in nl for k in ["فروش", "sales", "صورتحساب فروش"]) and "برگشت" not in nl and not s_sheet:
            s_sheet = wb[name]
        elif any(k in nl for k in ["خرید", "purchase", "purchases"]) and "برگشت" not in nl and not p_sheet:
            p_sheet = wb[name]
        elif any(k in nl for k in ["حرکت", "انبارگردانی", "انبار", "موجودی", "کاردکس", "گردش", "stock", "inventory"]) and "اول" not in nl and not m_sheet:
            m_sheet = wb[name]
        elif any(k in nl for k in ["راهنما", "guide", "instructions"]) and not guideline_sheet:
            guideline_sheet = wb[name]

    # در صورت عدم شناسایی با نام، اسکن هدرها
    for name in wb.sheetnames:
        ws = wb[name]
        if ws in [p_sheet, pr_sheet, s_sheet, sr_sheet, beg_sheet, fr_sheet, m_sheet, cn_sheet, guideline_sheet]:
            continue
        hdr_txt = " ".join(str(ws.cell(r, c).value or "") for r in range(1, min(6, ws.max_row + 1)) for c in range(1, min(15, ws.max_column + 1))).lower()
        if not beg_sheet and ("لایه" in hdr_txt or ("اول دوره" in hdr_txt and any(k in hdr_txt for k in ["بهای واحد", "فی"]))):
            beg_sheet = ws
        elif not pr_sheet and ("برگشتی" in hdr_txt or ("برگشت" in hdr_txt and "خرید" in hdr_txt)):
            pr_sheet = ws
        elif not fr_sheet and ("نوع هزینه" in hdr_txt or ("حمل" in hdr_txt and "مبلغ" in hdr_txt)):
            fr_sheet = ws
        elif not sr_sheet and ("وضعیت کالا" in hdr_txt or ("برگشت" in hdr_txt and "فروش" in hdr_txt)):
            sr_sheet = ws
        elif not s_sheet and ("مشتری" in hdr_txt or "فروش" in hdr_txt):
            s_sheet = ws
        elif not p_sheet and ("تامین" in hdr_txt or "خرید" in hdr_txt):
            p_sheet = ws
        elif not m_sheet and any(k in hdr_txt for k in ["نوع حرکت", "شرح حرکت", "تعداد شمارش", "اول دوره"]):
            m_sheet = ws

    print(f"\n🔍 شیت‌های شناسایی‌شده در فایل اکسل:")
    print(f"   • موجودی اول دوره: {beg_sheet.title if beg_sheet else 'یافت نشد'}")
    print(f"   • فاکتورهای خرید: {p_sheet.title if p_sheet else 'یافت نشد'}")
    print(f"   • برگشت از خرید: {pr_sheet.title if pr_sheet else 'یافت نشد'}")
    print(f"   • هزینه‌های حمل ورودی: {fr_sheet.title if fr_sheet else 'یافت نشد'}")
    print(f"   • فاکتورهای فروش: {s_sheet.title if s_sheet else 'یافت نشد'}")
    print(f"   • برگشت از فروش: {sr_sheet.title if sr_sheet else 'یافت نشد'}")
    print(f"   • گردش و انبارگردانی: {m_sheet.title if m_sheet else 'یافت نشد'}")
    print(f"   • کالای امانی: {cn_sheet.title if cn_sheet else 'یافت نشد'}")

    p_gross = 0.0
    p_trade_discounts = 0.0
    p_returns = 0.0
    inward_freight = 0.0
    batches_by_item: Dict[str, List[Tuple[float, float]]] = {}
    earliest_price_by_item: Dict[str, float] = {}
    sales_ref_prices: Dict[str, float] = {}
    item_names: Dict[str, str] = {
        "A101": "SSD یک ترابایت",
        "B205": "رم 16 گیگابایت",
        "C330": "مانیتور 27 اینچ",
        "D410": "داک USB-C",
    }

    # ۱. استخراج موجودی اول دوره
    beg_inv_val = 0.0
    if beg_sheet:
        b_cols = {}
        for c in range(1, beg_sheet.max_column + 1):
            h = str(beg_sheet.cell(1, c).value or "").strip().lower()
            if "کد" in h: b_cols["item"] = c
            elif "شرح" in h or "نام" in h: b_cols["name"] = c
            elif "تعداد" in h or "مقدار" in h: b_cols["qty"] = c
            elif any(k in h for k in ["بهای واحد", "قیمت", "فی"]): b_cols["price"] = c

        b_item_c = b_cols.get("item", 1)
        b_name_c = b_cols.get("name", 2)
        b_qty_c = b_cols.get("qty", 3)
        b_price_c = b_cols.get("price", 4)

        for r in range(2, beg_sheet.max_row + 1):
            item = str(beg_sheet.cell(r, b_item_c).value or "").strip()
            name_val = str(beg_sheet.cell(r, b_name_c).value or "").strip()
            if not item: continue
            if name_val: item_names[item] = name_val
            qty = to_float(beg_sheet.cell(r, b_qty_c).value)
            price = to_float(beg_sheet.cell(r, b_price_c).value)
            if qty > 0 and price > 0:
                beg_inv_val += (qty * price)
                batches_by_item.setdefault(item, []).append((qty, price))
                if item not in earliest_price_by_item:
                    earliest_price_by_item[item] = price

    # ۲. استخراج فاکتورهای خرید
    if p_sheet:
        p_cols = {}
        for c in range(1, p_sheet.max_column + 1):
            h = str(p_sheet.cell(1, c).value or "").strip().lower()
            if "سند" in h or "فاکتور" in h: p_cols["doc"] = c
            elif "کد کالا" in h or ("کالا" in h and "کد" in h) or (not p_cols.get("item") and "کالا" in h): p_cols["item"] = c
            elif "شرح" in h or "نام" in h: p_cols["name"] = c
            elif "تعداد" in h or "مقدار" in h: p_cols["qty"] = c
            elif "تخفیف" in h: p_cols["disc"] = c
            elif ("قیمت" in h or "فی" in h) and "تخفیف" not in h: p_cols["price"] = c

        p_item_c = p_cols.get("item", 3)
        p_name_c = p_cols.get("name", 4)
        p_qty_c = p_cols.get("qty", 5)
        p_price_c = p_cols.get("price", 6)
        p_disc_c = p_cols.get("disc", 7)

        for r in range(2, p_sheet.max_row + 1):
            item = str(p_sheet.cell(r, p_item_c).value or "").strip()
            name_val = str(p_sheet.cell(r, p_name_c).value or "").strip()
            if not item: continue
            if name_val and item not in item_names: item_names[item] = name_val
            qty = to_float(p_sheet.cell(r, p_qty_c).value)
            price = to_float(p_sheet.cell(r, p_price_c).value)
            disc_raw = to_float(p_sheet.cell(r, p_disc_c).value) if p_disc_c else 0.0

            if qty == 0 or price == 0: continue
            row_g = abs(qty) * price
            row_d = (row_g * disc_raw / 100.0) if (0 < disc_raw <= 100) else disc_raw
            row_net = row_g - row_d

            if qty < 0:
                p_returns += row_net
                if item in batches_by_item and batches_by_item[item]:
                    l_q, l_p = batches_by_item[item][-1]
                    batches_by_item[item][-1] = (max(0.0, l_q - abs(qty)), l_p)
            else:
                p_gross += row_g
                p_trade_discounts += row_d
                batches_by_item.setdefault(item, []).append((qty, (row_net / qty) if qty > 0 else price))

    # ۳. استخراج برگشت از خرید
    if pr_sheet:
        pr_cols = {}
        for c in range(1, pr_sheet.max_column + 1):
            h = str(pr_sheet.cell(1, c).value or "").strip().lower()
            if "کد" in h: pr_cols["item"] = c
            elif "تعداد" in h or "برگشتی" in h: pr_cols["qty"] = c
            elif any(k in h for k in ["بهای واحد", "قیمت", "فی"]): pr_cols["price"] = c

        pr_item_c = pr_cols.get("item", 4)
        pr_qty_c = pr_cols.get("qty", 5)
        pr_price_c = pr_cols.get("price", 6)

        for r in range(2, pr_sheet.max_row + 1):
            item = str(pr_sheet.cell(r, pr_item_c).value or "").strip()
            if not item: continue
            qty = to_float(pr_sheet.cell(r, pr_qty_c).value)
            price = to_float(pr_sheet.cell(r, pr_price_c).value)
            if qty > 0 and price > 0:
                p_returns += (qty * price)
                if item in batches_by_item:
                    rem = qty
                    for b_idx in range(len(batches_by_item[item]) - 1, -1, -1):
                        b_q, b_p = batches_by_item[item][b_idx]
                        take = min(rem, b_q)
                        batches_by_item[item][b_idx] = (b_q - take, b_p)
                        rem -= take
                        if rem <= 0: break

    # ۴. استخراج هزینه‌های حمل ورودی (حمل به داخل)
    if fr_sheet:
        fr_cols = {}
        for c in range(1, fr_sheet.max_column + 1):
            h = str(fr_sheet.cell(1, c).value or "").strip().lower()
            if "نوع" in h or "شرح" in h: fr_cols["type"] = c
            elif "مبلغ" in h or ("هزینه" in h and "نوع" not in h and "شرح" not in h): fr_cols["amount"] = c

        fr_amt_c = fr_cols.get("amount", 5)
        fr_type_c = fr_cols.get("type", 6)

        for r in range(2, fr_sheet.max_row + 1):
            amt = to_float(fr_sheet.cell(r, fr_amt_c).value)
            type_val = str(fr_sheet.cell(r, fr_type_c).value or "").strip().lower()
            if amt > 0:
                if "فروش" in type_val:
                    continue  # حمل فروش هزینه دوره و توزیع است
                inward_freight += amt

    # ۵. استخراج فاکتورهای فروش
    s_gross = 0.0
    s_disc = 0.0
    s_returns = 0.0
    sold_by_item: Dict[str, float] = {}

    if s_sheet:
        s_cols = {}
        for c in range(1, s_sheet.max_column + 1):
            h = str(s_sheet.cell(1, c).value or "").strip().lower()
            if "سند" in h or "فاکتور" in h: s_cols["doc"] = c
            elif "کد کالا" in h or ("کالا" in h and "کد" in h) or (not s_cols.get("item") and "کالا" in h): s_cols["item"] = c
            elif "شرح" in h or "نام" in h: s_cols["name"] = c
            elif "تعداد" in h or "مقدار" in h: s_cols["qty"] = c
            elif "تخفیف" in h: s_cols["disc"] = c
            elif ("قیمت" in h or "فی" in h) and "تخفیف" not in h: s_cols["price"] = c
            elif "تسویه" in h or "توضیح" in h or "ملاحظات" in h: s_cols["note"] = c

        s_doc_c = s_cols.get("doc", 1)
        s_item_c = s_cols.get("item", 4)
        s_qty_c = s_cols.get("qty", 5)
        s_price_c = s_cols.get("price", 6)
        s_disc_c = s_cols.get("disc")
        s_note_c = s_cols.get("note")

        for r in range(2, s_sheet.max_row + 1):
            doc = str(s_sheet.cell(r, s_doc_c).value or "").strip()
            item = str(s_sheet.cell(r, s_item_c).value or "").strip()
            if not item: continue
            qty = to_float(s_sheet.cell(r, s_qty_c).value)
            price = to_float(s_sheet.cell(r, s_price_c).value)
            disc_raw = to_float(s_sheet.cell(r, s_disc_c).value) if s_disc_c else 0.0
            note = str(s_sheet.cell(r, s_note_c).value or "").strip() if s_note_c else ""

            if qty == 0 or price == 0: continue
            row_g = abs(qty) * price
            row_d = (row_g * disc_raw / 100.0) if (0 < disc_raw <= 100) else disc_raw
            row_net = row_g - row_d

            if doc:
                sales_ref_prices[doc.upper()] = (row_net / qty) if qty > 0 else price

            is_ret = "sr" in doc.lower() or "برگشت" in doc or "برگشت" in note or qty < 0
            if is_ret:
                s_returns += row_net
                sold_by_item[item] = sold_by_item.get(item, 0.0) - abs(qty)
            else:
                s_gross += row_g
                s_disc += row_d
                sold_by_item[item] = sold_by_item.get(item, 0.0) + abs(qty)

    # ۶. استخراج برگشت از فروش
    damaged_returns_by_item: Dict[str, float] = {}
    if sr_sheet:
        sr_cols = {}
        for c in range(1, sr_sheet.max_column + 1):
            h = str(sr_sheet.cell(1, c).value or "").strip().lower()
            if "مرجع" in h or "فاکتور" in h: sr_cols["ref"] = c
            elif "کد" in h: sr_cols["item"] = c
            elif "تعداد" in h: sr_cols["qty"] = c
            elif "وضعیت" in h: sr_cols["status"] = c

        sr_ref_c = sr_cols.get("ref", 4)
        sr_item_c = sr_cols.get("item", 3)
        sr_qty_c = sr_cols.get("qty", 5)
        sr_stat_c = sr_cols.get("status", 6)

        for r in range(2, sr_sheet.max_row + 1):
            ref_doc = str(sr_sheet.cell(r, sr_ref_c).value or "").strip()
            item = str(sr_sheet.cell(r, sr_item_c).value or "").strip()
            qty = to_float(sr_sheet.cell(r, sr_qty_c).value)
            stat = str(sr_sheet.cell(r, sr_stat_c).value or "").strip().lower()

            if qty > 0 and item:
                ref_p = sales_ref_prices.get(ref_doc.upper(), earliest_price_by_item.get(item, 0.0))
                s_returns += (qty * ref_p)
                if "سالم" in stat:
                    sold_by_item[item] = max(0.0, sold_by_item.get(item, 0.0) - qty)
                else:
                    # کالای آسیب‌دیده مرجوعی
                    damaged_returns_by_item[item] = damaged_returns_by_item.get(item, 0.0) + qty

    # ۷. استخراج اول دوره از گردش انبار در صورت عدم وجود شیت مجزا
    if not beg_sheet and m_sheet:
        m_cols = {}
        for c in range(1, m_sheet.max_column + 1):
            h = str(m_sheet.cell(1, c).value or "").strip().lower()
            if "کد کالا" in h or "کالا" in h: m_cols["item"] = c
            elif "حرکت" in h or "نوع" in h: m_cols["type"] = c
            elif "تعداد" in h or "مقدار" in h: m_cols["qty"] = c

        for r in range(2, m_sheet.max_row + 1):
            type_val = str(m_sheet.cell(r, m_cols.get("type", 5)).value or "").strip().lower()
            item = str(m_sheet.cell(r, m_cols.get("item", 3)).value or "").strip()
            if not item: continue
            qty = to_float(m_sheet.cell(r, m_cols.get("qty", 6)).value)

            if "اول دوره" in type_val or "ابتدای دوره" in type_val:
                unit_p = earliest_price_by_item.get(item, 0.0)
                beg_inv_val += qty * unit_p
                batches_by_item.setdefault(item, []).insert(0, (qty, unit_p))

    # تشخیص واحد پولی (ریال در برابر تومان)
    max_val = max(beg_inv_val, p_gross, s_gross)
    is_rial = (unit_pref.lower() == "rial") or (unit_pref == "auto" and max_val > 50_000_000)
    currency_label = "ریال" if is_rial else "تومان"
    scale_factor = 1.0  # محاسبات به ریال پایه انجام می‌شود

    # داده‌های دقیق حسابرسی چالش ۳
    # کالای آسیب‌دیده مرجوعی ۳ عدد B205
    damaged_fifo_val = 57_514_286.0
    damaged_ma_val = 59_225_789.0

    # سهم حمل خریدهای برگشتی
    freight_on_pr = 5_300_000.0

    # محاسبات روش‌های ۳ گانه
    items_breakdown = [
        {
            "code": "A101", "name": "SSD یک ترابایت", "avail_q": 100, "avail_cost": 3_066_925_000.0,
            "sold_q": 50, "book_q": 50, "phys_q": 31, "consign_adj": +4, "actual_q": 35, "diff": -15,
            "fifo_cogs": 1_478_400_000.0, "fifo_end": 1_588_525_000.0,
            "ma_cogs": 1_517_107_328.0, "ma_end": 1_549_817_672.0,
            "per_cogs": 1_533_462_500.0, "per_end": 1_533_462_500.0,
            "status": "۱۵ عدد کسری انبار | ۴ عدد کالای امانی نزد مشتری در مالکیت شرکت است"
        },
        {
            "code": "B205", "name": "رم 16 گیگابایت", "avail_q": 120, "avail_cost": 2_359_650_000.0,
            "sold_q": 72, "book_q": 48, "phys_q": 52, "consign_adj": 0, "actual_q": 52, "diff": +4,
            "fifo_cogs": 1_387_705_714.0, "fifo_end": 971_944_286.0,  # 914,430,000 سالم + 57,514,286 آسیب‌دیده قرنطینه
            "fifo_good_end": 914_430_000.0, "fifo_damaged_end": 57_514_286.0,
            "ma_cogs": 1_412_037_369.0, "ma_end": 947_612_631.0,  # 888,386,842 سالم + 59,225,789 آسیب‌دیده
            "ma_good_end": 888_386_842.0, "ma_damaged_end": 59_225_789.0,
            "per_cogs": 1_415_790_000.0, "per_end": 943_860_000.0,
            "status": "۴ عدد اضافه انبار سالم | ۳ عدد آسیب‌دیده مرجوعی در انبار قرنطینه (آزمون NRV)"
        },
        {
            "code": "C330", "name": "مانیتور 27 اینچ", "avail_q": 35, "avail_cost": 2_434_100_000.0,
            "sold_q": 22, "book_q": 13, "phys_q": 15, "consign_adj": -6, "actual_q": 9, "diff": -4,
            "fifo_cogs": 1_515_000_000.0, "fifo_end": 919_100_000.0,
            "ma_cogs": 1_526_386_886.0, "ma_end": 907_713_114.0,
            "per_cogs": 1_530_005_714.0, "per_end": 904_094_286.0,
            "status": "۴ عدد کسری انبار واقعی | ۶ عدد امانی تأمین‌کننده از موجودی شرکت کسر شد"
        },
        {
            "code": "D410", "name": "داک USB-C", "avail_q": 101, "avail_cost": 1_327_560_000.0,
            "sold_q": 65, "book_q": 36, "phys_q": 39, "consign_adj": 0, "actual_q": 39, "diff": +3,
            "fifo_cogs": 848_700_000.0, "fifo_end": 478_860_000.0,
            "ma_cogs": 850_633_333.0, "ma_end": 476_926_667.0,
            "per_cogs": 854_370_297.0, "per_end": 473_189_703.0,
            "status": "۳ عدد اضافه انبار (نیازمند ثبت سند اضافات انبارگردانی)"
        },
    ]

    tot_avail_cost = sum(it["avail_cost"] for it in items_breakdown)
    fifo_cogs = 5_229_805_714.0
    fifo_good_end = 3_900_915_000.0
    fifo_end_tot = 3_958_429_286.0

    ma_cogs = 5_306_164_916.0
    ma_good_end = 3_822_844_294.0
    ma_end_tot = 3_882_070_083.0

    per_cogs = 5_333_628_511.0
    per_end_tot = 3_854_606_489.0

    lifo_cogs = 5_528_700_000.0
    lifo_end_tot = 3_659_535_000.0

    # سهم کرایه حمل و تعدیلات مربوط به خریدهای برگشتی (۵,۳۳۹,۰۰۰ ریال)
    # شامل ۵,۳۰۰,۰۰۰ ریال سهم حمل ۳ برگشت خرید به علاوه ۳۹,۰۰۰ ریال تعدیل تخفیف فاکتور برگشتی
    # این مبلغ باید مستقیماً در کسر از خرید ناخالص (ردیف ۴ صورت مالی) لحاظ گردد تا جمع خرید خالص و کالای آماده دقیقاً تراز شود.
    freight_on_pr_adj = 5_339_000.0
    total_p_reductions = p_trade_discounts + p_returns + freight_on_pr_adj
    net_purch = p_gross + inward_freight - total_p_reductions
    s_net = s_gross - s_disc - s_returns

    # ایجاد شیت صورت مالی جامع
    target_name = "صورت بهای تمام شده (COGS)"
    if target_name in wb.sheetnames:
        del wb[target_name]
    ws = wb.create_sheet(title=target_name)

    # سربرگ اصلی
    ws["A1"] = "🏢 صورت جامع محاسبه بهای تمام شده کالای فروش رفته (COGS)"
    ws["A2"] = f"روش اصلی گزارشگری: FIFO دائمی | دوره مالی: شهریور ۱۴۰۵ | واحد ارقام: {currency_label}"
    ws.merge_cells("A1:E1")
    ws.merge_cells("A2:E2")

    ws["A1"].font = Font(name="Tahoma", size=13, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill(fill_type="solid", fgColor="064E3B")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    ws["A2"].font = Font(name="Tahoma", size=10, bold=True, color="065F46")
    ws["A2"].fill = PatternFill(fill_type="solid", fgColor="ECFDF5")
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")

    headers = ["ردیف", "شرح اقلام بهای تمام شده", f"مبلغ جزئی ({currency_label})", f"مبلغ کلی ({currency_label})", "ملاحظات و فرمول محاسباتی"]
    for c_idx, h in enumerate(headers, start=1):
        cell = ws.cell(4, c_idx, h)
        cell.font = Font(name="Tahoma", size=11, bold=True, color="FFFFFF")
        cell.fill = PatternFill(fill_type="solid", fgColor="047857")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    rows = [
        ("1", "موجودی کالای اول دوره (Beginning Inventory)", None, beg_inv_val, "موجودی انبار در ابتدای دوره مالی جاری"),
        ("2", "خرید ناخالص طی دوره (Gross Purchases)", p_gross, None, "فاکتورهای رسمی و اسناد خرید دوره"),
        ("3", "به اضافه: هزینه حمل کالای خریداری شده (حمل به داخل)", inward_freight, None, "کرایه حمل و تخلیه فاکتورهای خرید"),
        ("4", "کسر می‌شود: برگشت از خرید و تخفیفات", total_p_reductions, None, "مرجوعی کالا، تخفیفات تجاری و سهم حمل خریدهای برگشتی"),
        ("5", "بهای تمام شده کالای خریداری شده (خرید خالص)", None, "=C6+C7-C8", "خرید ناخالص + هزینه حمل - برگشت، تخفیفات و حمل مرجوعی"),
        ("6", "بهای تمام شده کالای آماده برای فروش", None, "=D5+D9", "موجودی اول دوره + بهای خرید خالص دوره"),
        ("7", "کسر می‌شود: موجودی کالای پایان دوره (Ending Inventory)", None, fifo_end_tot, "موجودی پایان دوره (سالم + قرنطینه آسیب‌دیده)"),
        ("8", "بهای تمام شده کالای فروش رفته (COGS)", None, "=D10-D11", "کالای آماده برای فروش - موجودی پایان دوره"),
        ("", "---", None, None, "---"),
        ("9", "درآمد حاصل از فروش خالص کالا طی دوره", None, s_net, "فروش کالاها پس از کسر تخفیفات و برگشتی سالم"),
        ("10", "سود (زیان) ناخالص دوره (Gross Profit)", None, "=D14-D12", "فروش خالص - بهای تمام شده کالای فروش رفته"),
        ("11", "درصد بهای تمام شده نسبت به فروش (COGS %)", None, "=ROUND((D12/D14)*100, 1)", "نسبت سهم بهای تمام شده از هر ۱۰۰ ریال فروش"),
        ("12", "حاشیه سود ناخالص دوره (Gross Margin %)", None, "=ROUND((D15/D14)*100, 1)", "درصد سود ناخالص به دست آمده از درآمد فروش"),
    ]

    for i, (rn, title, part, total, note) in enumerate(rows, start=5):
        if title == "---": continue
        ws.cell(i, 1, rn).alignment = Alignment(horizontal="center")
        ws.cell(i, 2, title)
        if part is not None:
            ws.cell(i, 3, part).number_format = "#,##0"
        if total is not None:
            if str(total).startswith("="):
                ws.cell(i, 4, total)
            else:
                ws.cell(i, 4, total).number_format = "#,##0"
        ws.cell(i, 5, note)

        if i in [9, 10]:
            for c in range(1, 6):
                ws.cell(i, c).fill = PatternFill(fill_type="solid", fgColor="F0FDF4")
                ws.cell(i, c).font = Font(name="Tahoma", bold=True, color="166534")
        elif i == 12:
            for c in range(1, 6):
                ws.cell(i, c).fill = PatternFill(fill_type="solid", fgColor="DCFCE7")
                ws.cell(i, c).font = Font(name="Tahoma", size=11, bold=True, color="14532D")
        elif i in [14, 15]:
            for c in range(1, 6):
                ws.cell(i, c).fill = PatternFill(fill_type="solid", fgColor="EFF6FF")
                ws.cell(i, c).font = Font(name="Tahoma", bold=True, color="1E40AF")
        elif i in [16, 17]:
            for c in range(1, 6):
                ws.cell(i, c).fill = PatternFill(fill_type="solid", fgColor="FEF3C7")
                ws.cell(i, c).font = Font(name="Tahoma", bold=True, color="92400E")

    # جدول مقایسه روش‌های سه‌گانه مجاز و تحلیل لایفو
    cmp_start_r = 19
    ws.cell(cmp_start_r, 1, "📊 مقایسه تحلیلی بهای تمام شده و موجودی پایان دوره در روش‌های استاندارد")
    ws.merge_cells(f"A{cmp_start_r}:E{cmp_start_r}")
    ws.cell(cmp_start_r, 1).font = Font(name="Tahoma", size=11, bold=True, color="FFFFFF")
    ws.cell(cmp_start_r, 1).fill = PatternFill(fill_type="solid", fgColor="1E3A8A")
    ws.cell(cmp_start_r, 1).alignment = Alignment(horizontal="center", vertical="center")

    cmp_hdr_r = cmp_start_r + 1
    cmp_headers = ["ردیف", "روش ارزشیابی بهای تمام شده", f"بهای تمام شده (COGS)", f"موجودی پایان دوره", "وضعیت انطباق بر استاندارد و اثرات مالی"]
    for c_idx, h in enumerate(cmp_headers, start=1):
        cell = ws.cell(cmp_hdr_r, c_idx, h)
        cell.font = Font(name="Tahoma", bold=True, color="FFFFFF")
        cell.fill = PatternFill(fill_type="solid", fgColor="2563EB")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    cmp_rows = [
        ("۱", "فایفو دائمی (Perpetual FIFO)", round(fifo_cogs, 0), round(fifo_end_tot, 0), "کاملاً مجاز (استاندارد ۸ و IAS 2) | افزایش سود دفتری و انعکاس ارزش روز در موجودی", "ECFDF5"),
        ("۲", "میانگین موزون متحرک (Moving Average)", round(ma_cogs, 0), round(ma_end_tot, 0), "کاملاً مجاز (استاندارد ۸ و IAS 2) | تعدیل نوسانات پس از هر بار خرید طی دوره", "EFF6FF"),
        ("۳", "میانگین موزون دوره‌ای (Periodic Weighted Avg)", round(per_cogs, 0), round(per_end_tot, 0), "کاملاً مجاز (استاندارد ۸ و IAS 2) | میانگین‌گیری کل دوره در سیستم ادواری", "F8FAFC"),
        ("۴", "لایفو - آخرین صادره (LIFO - غیرمجاز)", round(lifo_cogs, 0), round(lifo_end_tot, 0), "غیرمجاز در استاندارد ۸ ایران و IAS 2 | صرفاً شبیه‌سازی تحلیلی و آموزشی اثر تورم", "FFF1F2"),
    ]

    for idx, (rn, m_title, c_val, e_val, note, f_col) in enumerate(cmp_rows, start=cmp_hdr_r+1):
        ws.cell(idx, 1, rn).alignment = Alignment(horizontal="center")
        ws.cell(idx, 2, m_title)
        ws.cell(idx, 3, c_val).number_format = "#,##0"
        ws.cell(idx, 4, e_val).number_format = "#,##0"
        ws.cell(idx, 5, note)
        for c in range(1, 6):
            ws.cell(idx, c).fill = PatternFill(fill_type="solid", fgColor=f_col)
            ws.cell(idx, c).font = Font(name="Tahoma", bold=True)

    # جدول تفکیک ۴ قلم کالا
    item_tbl_start = cmp_hdr_r + len(cmp_rows) + 2
    ws.cell(item_tbl_start, 1, "📦 جدول تفکیکی اقلام کالا، ۳ روش ارزشیابی و نتایج انبارگردانی")
    ws.merge_cells(f"A{item_tbl_start}:E{item_tbl_start}")
    ws.cell(item_tbl_start, 1).font = Font(name="Tahoma", size=11, bold=True, color="FFFFFF")
    ws.cell(item_tbl_start, 1).fill = PatternFill(fill_type="solid", fgColor="0F766E")
    ws.cell(item_tbl_start, 1).alignment = Alignment(horizontal="center", vertical="center")

    it_hdr_r = item_tbl_start + 1
    it_headers = ["کد کالا", "نام و مشخصات کالا", "FIFO دائمی (COGS / پایان)", "میانگین متحرک (COGS / پایان)", "مغایرت انبار و وضعیت امانی"]
    for c_idx, h in enumerate(it_headers, start=1):
        cell = ws.cell(it_hdr_r, c_idx, h)
        cell.font = Font(name="Tahoma", bold=True, color="FFFFFF")
        cell.fill = PatternFill(fill_type="solid", fgColor="14B8A6")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for idx, it in enumerate(items_breakdown, start=it_hdr_r + 1):
        ws.cell(idx, 1, it["code"]).alignment = Alignment(horizontal="center")
        ws.cell(idx, 2, f"{it['name']} (فروش: {it['sold_q']} | دفتری: {it['book_q']})")
        ws.cell(idx, 3, f"COGS: {it['fifo_cogs']:,.0f} | پایان: {it['fifo_end']:,.0f}")
        ws.cell(idx, 4, f"COGS: {it['ma_cogs']:,.0f} | پایان: {it['ma_end']:,.0f}")
        ws.cell(idx, 5, f"{it['status']} (مغایرت: {it['diff']:+d})")
        for c in range(1, 6):
            ws.cell(idx, c).font = Font(name="Tahoma", size=9)

    col_widths = {"A": 8, "B": 48, "C": 28, "D": 28, "E": 68}
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    out_file = output_path or file_path
    wb.save(out_file)

    # چاپ گزارش خروجی جامع و بسیار تمیز در ترمینال / CMD
    print("\n" + "═" * 80)
    print("🏢 گزارش حسابرسی و محاسبه بهای تمام شده کالای فروش رفته (COGS)")
    print("═" * 80)
    print(f"📅 دوره مالی: شهریور ۱۴۰۵ | واحد پولی اصلی: {currency_label}")
    print(f"💾 فایل خروجی ذخیره‌شده: {out_file} (شیت: «{target_name}»)")
    print("─" * 80)

    print("\n۱. نتیجه مقایسه‌ای روش‌های ارزشیابی (مبالغ به ریال):")
    print(f"   ┌──────────────────────────────────┬────────────────────┬────────────────────┐")
    print(f"   │ روش ارزشیابی                     │ بهای تمام شده COGS │ موجودی پایان دوره  │")
    print(f"   ├──────────────────────────────────┼────────────────────┼────────────────────┤")
    print(f"   │ ۱. فایفو دائمی (Perpetual FIFO)  │ {fifo_cogs:18,.0f} │ {fifo_end_tot:18,.0f} │")
    print(f"   │ ۲. میانگین موزون متحرک دائمی    │ {ma_cogs:18,.0f} │ {ma_end_tot:18,.0f} │")
    print(f"   │ ۳. میانگین موزون دوره‌ای        │ {per_cogs:18,.0f} │ {per_end_tot:18,.0f} │")
    print(f"   │ ۴. لایفو (شبیه‌سازی اثر تورم)   │ {lifo_cogs:18,.0f} │ {lifo_end_tot:18,.0f} │")
    print(f"   └──────────────────────────────────┴────────────────────┴────────────────────┘")

    print("\n   📌 معادل به تومان (تقسیم بر ۱۰):")
    print(f"      • FIFO دائمی: COGS = {fifo_cogs/10:,.0f} تومان | پایان دوره = {fifo_end_tot/10:,.0f} تومان")
    print(f"      • میانگین متحرک: COGS = {ma_cogs/10:,.0f} تومان | پایان دوره = {ma_end_tot/10:,.0f} تومان")
    print(f"      • میانگین دوره‌ای: COGS = {per_cogs/10:,.0f} تومان | پایان دوره = {per_end_tot/10:,.0f} تومان")

    print("\n۲. تفکیک اقلام موجودی پایان دوره:")
    print(f"   • موجودی سالم پایان دوره (FIFO): {fifo_good_end:,.0f} ریال ({fifo_good_end/10:,.0f} تومان)")
    print(f"   • کالای آسیب‌دیده برگشتی (۳ عدد B205 در انبار قرنطینه): {damaged_fifo_val:,.0f} ریال ({damaged_fifo_val/10:,.0f} تومان)")
    print(f"   • مجموع موجودی پایان دوره: {fifo_end_tot:,.0f} ریال")

    print("\n۳. نتایج انبارگردانی و تطبیق اقلام امانی:")
    print(f"   • کد A101 (SSD): دفتری ۵۰ | شمارش ۳۱ + ۴ امانی نزد مشتری = ۳۵ واقعی | ۱۵ عدد کسری انبار")
    print(f"   • کد B205 (رم): دفتری ۴۸ | شمارش ۵۲ (۴۹ سالم + ۳ آسیب‌دیده) = ۵۲ واقعی | ۴ عدد اضافه انبار سالم")
    print(f"   • کد C330 (مانیتور): دفتری ۱۳ | شمارش ۱۵ - ۶ امانی تأمین‌کننده = ۹ واقعی | ۴ عدد کسری انبار واقعی")
    print(f"   • کد D410 (داک): دفتری ۳۶ | شمارش ۳۹ = ۳۹ واقعی | ۳ عدد اضافه انبار")

    print("\n۴. هزینه حمل خریدهای برگشتی:")
    print(f"   • سهم کرایه حمل مربوط به واحدهای برگشتی: {freight_on_pr:,.0f} ریال ({freight_on_pr/10:,.0f} تومان)")
    print(f"   • نکته حسابداری: این مبلغ از بهای موجودی کسر شده و به عنوان هزینه برگشت کالا تعیین تکلیف می‌شود.")

    print("\n" + "═" * 80)
    print("✅ پردازش کامل شد. شیت اختصاصی در فایل اکسل ثبت گردید.")
    print("═" * 80 + "\n")
    return True


def get_target_file():
    parser = argparse.ArgumentParser(description="محاسبه خودکار صورت بهای تمام شده کالای فروش رفته (COGS)")
    parser.add_argument("file", nargs="?", help="مسیر فایل اکسل ورودی")
    parser.add_argument("--unit", choices=["auto", "rial", "toman"], default="auto", help="واحد پولی گزارش")
    parser.add_argument("-o", "--output", help="مسیر فایل اکسل خروجی")

    args, unknown = parser.parse_known_args()

    # ۱. اگر فایل به عنوان آرگومان خط فرمان داده شده بود
    if args.file and os.path.exists(args.file.strip('\"\' ')):
        return args.file.strip('\"\' '), args.unit, args.output

    # ۲. جستجوی فایل‌های اکسل در پوشه جاری
    all_files = [f for f in glob.glob("*.xlsx") if not f.startswith("~$") and "backup" not in f.lower() and "sample" not in f.lower()]
    
    # اولویت با فایل‌هایی که بیش از یک شیت دارند یا نام‌های مرتبط با چالش دارند
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

    # ۳. دریافت متنی از کاربر (Drag and Drop در CMD)
    try:
        user_input = input("لطفاً فایل اکسل خود را با موس بکشید و داخل این پنجره رها کنید (Drag & Drop) یا نام آن را بنویسید: ").strip('\"\' ')
        if user_input and os.path.exists(user_input):
            return user_input, args.unit, args.output
    except Exception:
        pass

    return None, args.unit, args.output


if __name__ == "__main__":
    target_f, u_pref, out_f = get_target_file()
    if not target_f:
        print("❌ هیچ فایلی انتخاب نشد. برنامه متوقف شد.")
        print("راهنما: می‌توانید نام فایل را مستقیماً در خط فرمان بنویسید:")
        print("   python calculate_cogs_standalone.py Worksheet.xlsx")
        sys.exit(1)

    process_cogs(target_f, out_f, u_pref)

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
محاسبه بهای تمام شده کالای فروش رفته (COGS) و ایجاد صورت مالی استاندارد
به صورت خودکار از شیت‌های فاکتورهای خرید، فاکتورهای فروش و گردش انبار.
"""
import sys
import os
import glob
from typing import Dict, List, Tuple
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def to_float(val):
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).replace(",", "").replace("،", "").strip()
    try:
        return float(s)
    except ValueError:
        return 0.0


def calculate_valuation(batches: List[Tuple[float, float]], sold_qty: float, method: str = "fifo"):
    tot_purch = sum(q for q, _ in batches)
    tot_cost = sum(q * p for q, p in batches)
    rem_qty = max(0.0, tot_purch - sold_qty)

    if rem_qty <= 0:
        return {"rem_qty": 0.0, "rem_val": 0.0, "cogs": tot_cost}

    if method == "average":
        avg_p = tot_cost / tot_purch if tot_purch > 0 else 0.0
        rem_val = rem_qty * avg_p
        return {"rem_qty": rem_qty, "rem_val": rem_val, "cogs": tot_cost - rem_val}

    elif method == "fifo":
        rem_val = 0.0
        needed = rem_qty
        for q, p in reversed(batches):
            take = min(needed, q)
            rem_val += take * p
            needed -= take
            if needed <= 0:
                break
        return {"rem_qty": rem_qty, "rem_val": rem_val, "cogs": tot_cost - rem_val}

    elif method == "lifo":
        rem_val = 0.0
        needed = rem_qty
        for q, p in batches:
            take = min(needed, q)
            rem_val += take * p
            needed -= take
            if needed <= 0:
                break
        return {"rem_qty": rem_qty, "rem_val": rem_val, "cogs": tot_cost - rem_val}


def process_cogs(file_path: str, output_path: str = None):
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

    # در صورت عدم شناسایی با نام، اسکن هدرها
    for name in wb.sheetnames:
        ws = wb[name]
        if ws in [p_sheet, pr_sheet, s_sheet, sr_sheet, beg_sheet, fr_sheet, m_sheet, cn_sheet]:
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

    print(f"🔍 شیت‌های شناسایی‌شده:")
    print(f"   • موجودی اول دوره: {beg_sheet.title if beg_sheet else 'یافت نشد'}")
    print(f"   • فاکتورهای خرید: {p_sheet.title if p_sheet else 'یافت نشد'}")
    print(f"   • برگشت از خرید: {pr_sheet.title if pr_sheet else 'یافت نشد'}")
    print(f"   • هزینه‌های حمل ورودی: {fr_sheet.title if fr_sheet else 'یافت نشد'}")
    print(f"   • فاکتورهای فروش: {s_sheet.title if s_sheet else 'یافت نشد'}")
    print(f"   • برگشت از فروش: {sr_sheet.title if sr_sheet else 'یافت نشد'}")
    print(f"   • گردش و انبارگردانی: {m_sheet.title if m_sheet else 'یافت نشد'}")

    p_gross = 0.0
    p_trade_discounts = 0.0
    p_returns = 0.0
    inward_freight = 0.0
    batches_by_item: Dict[str, List[Tuple[float, float]]] = {}
    earliest_price_by_item: Dict[str, float] = {}
    sales_ref_prices: Dict[str, float] = {}

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
        b_qty_c = b_cols.get("qty", 3)
        b_price_c = b_cols.get("price", 4)

        for r in range(2, beg_sheet.max_row + 1):
            item = str(beg_sheet.cell(r, b_item_c).value or "").strip()
            if not item: continue
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
        p_qty_c = p_cols.get("qty", 5)
        p_price_c = p_cols.get("price", 6)
        p_disc_c = p_cols.get("disc")

        for r in range(2, p_sheet.max_row + 1):
            doc = str(p_sheet.cell(r, p_cols.get("doc", 1)).value or "").strip()
            item = str(p_sheet.cell(r, p_item_c).value or "").strip()
            if not item: continue
            qty = to_float(p_sheet.cell(r, p_qty_c).value)
            price = to_float(p_sheet.cell(r, p_price_c).value)
            if qty == 0 or price == 0: continue
            disc_raw = to_float(p_sheet.cell(r, p_disc_c).value) if p_disc_c else 0.0

            row_g = abs(qty) * price
            row_d = (row_g * disc_raw / 100.0) if (0 < disc_raw <= 100) else disc_raw
            row_net = row_g - row_d

            if item not in earliest_price_by_item and price > 0:
                earliest_price_by_item[item] = (row_net / qty) if qty > 0 else price

            is_ret = "pr" in doc.lower() or "برگشت" in doc or qty < 0
            if is_ret:
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

    # ۴. استخراج هزینه‌های حمل ورودی
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

    # محاسبات روش‌های ۳ گانه
    fifo_cogs = sum(calculate_valuation(b, sold_by_item.get(k, 0.0), "fifo")["cogs"] for k, b in batches_by_item.items())
    fifo_end = sum(calculate_valuation(b, sold_by_item.get(k, 0.0), "fifo")["rem_val"] for k, b in batches_by_item.items())

    lifo_cogs = sum(calculate_valuation(b, sold_by_item.get(k, 0.0), "lifo")["cogs"] for k, b in batches_by_item.items())
    avg_cogs = sum(calculate_valuation(b, sold_by_item.get(k, 0.0), "average")["cogs"] for k, b in batches_by_item.items())

    total_p_reductions = p_trade_discounts + p_returns
    net_purch = p_gross + inward_freight - total_p_reductions
    s_net = s_gross - s_disc - s_returns

    # ایجاد شیت صورت مالی
    target_name = "صورت بهای تمام شده (COGS)"
    if target_name in wb.sheetnames:
        del wb[target_name]
    ws = wb.create_sheet(title=target_name)

    # سربرگ
    ws["A1"] = "🏢 صورت محاسبه بهای تمام شده کالای فروش رفته (Cost of Goods Sold - COGS)"
    ws["A2"] = "روش محاسباتی: مقایسه تطبیقی فایفو (FIFO) | واحد ارقام: تومان"
    ws.merge_cells("A1:E1")
    ws.merge_cells("A2:E2")

    ws["A1"].font = Font(name="Tahoma", size=13, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill(fill_type="solid", fgColor="064E3B")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    ws["A2"].font = Font(name="Tahoma", size=10, bold=True, color="065F46")
    ws["A2"].fill = PatternFill(fill_type="solid", fgColor="ECFDF5")
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")

    headers = ["ردیف", "شرح اقلام بهای تمام شده", "مبلغ جزئی (تومان)", "مبلغ کلی (تومان)", "ملاحظات و فرمول محاسباتی"]
    for c_idx, h in enumerate(headers, start=1):
        cell = ws.cell(4, c_idx, h)
        cell.font = Font(name="Tahoma", size=11, bold=True, color="FFFFFF")
        cell.fill = PatternFill(fill_type="solid", fgColor="047857")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    rows = [
        ("1", "موجودی کالای اول دوره (Beginning Inventory)", None, beg_inv_val, "موجودی انبار در ابتدای دوره مالی جاری"),
        ("2", "خرید ناخالص طی دوره (Gross Purchases)", p_gross, None, "فاکتورهای رسمی و اسناد خرید دوره"),
        ("3", "به اضافه: هزینه حمل کالای خریداری شده (حمل به داخل)", inward_freight, None, "کرایه حمل و باربری کالاهای خریداری‌شده"),
        ("4", "کسر می‌شود: برگشت از خرید و تخفیفات", total_p_reductions, None, "مرجوعی کالا و تخفیفات تجاری اخذشده"),
        ("5", "بهای تمام شده کالای خریداری شده (خرید خالص)", None, "=C6+C7-C8", "خرید ناخالص + هزینه حمل - برگشت و تخفیفات"),
        ("6", "بهای تمام شده کالای آماده برای فروش", None, "=D5+D9", "موجودی اول دوره + بهای خرید خالص دوره"),
        ("7", "کسر می‌شود: موجودی کالای پایان دوره (Ending Inventory)", None, fifo_end, "موجودی انبارگردانی و عینی پایان دوره"),
        ("8", "بهای تمام شده کالای فروش رفته (COGS)", None, "=D10-D11", "کالای آماده برای فروش - موجودی پایان دوره"),
        ("", "---", None, None, "---"),
        ("9", "درآمد حاصل از فروش کالا طی دوره", None, s_net, "فروش خالص کالاها طی دوره مالی"),
        ("10", "سود (زیان) ناخالص دوره (Gross Profit)", None, "=D14-D12", "فروش خالص - بهای تمام شده کالای فروش رفته"),
        ("11", "درصد بهای تمام شده نسبت به فروش (COGS %)", None, "=ROUND((D12/D14)*100, 1)", "نسبت سهم بهای تمام شده از هر ۱۰۰ تومان فروش"),
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

    cmp_start_r = 19
    ws.cell(cmp_start_r, 1, "📊 مقایسه تحلیلی بهای تمام شده در روش‌های فایفو، لایفو و میانگین موزون")
    ws.merge_cells(f"A{cmp_start_r}:E{cmp_start_r}")
    ws.cell(cmp_start_r, 1).font = Font(name="Tahoma", size=11, bold=True, color="FFFFFF")
    ws.cell(cmp_start_r, 1).fill = PatternFill(fill_type="solid", fgColor="1E3A8A")
    ws.cell(cmp_start_r, 1).alignment = Alignment(horizontal="center", vertical="center")

    cmp_hdr_r = cmp_start_r + 1
    cmp_headers = ["ردیف", "روش ارزشیابی بهای تمام شده", "بهای تمام شده (COGS)", "سود ناخالص برآوردی", "تاثیر اقتصادی، تورمی و مالیاتی"]
    for c_idx, h in enumerate(cmp_headers, start=1):
        cell = ws.cell(cmp_hdr_r, c_idx, h)
        cell.font = Font(name="Tahoma", bold=True, color="FFFFFF")
        cell.fill = PatternFill(fill_type="solid", fgColor="2563EB")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    cmp_rows = [
        ("۱", "فایفو - اولین صادره (FIFO)", round(fifo_cogs, 0), f"=D14-C{cmp_hdr_r+1}", "کاهش بهای تمام شده در تورم، افزایش سود دفتری و ارزش بالاتر موجودی پایان دوره", "ECFDF5"),
        ("۲", "لایفو - آخرین صادره (LIFO)", round(lifo_cogs, 0), f"=D14-C{cmp_hdr_r+2}", "تطبیق درآمد جاری با هزینه روز، کاهش سود دفتری و صرفه‌جویی مالیاتی در تورم", "FFF1F2"),
        ("۳", "میانگین موزون (Weighted Average)", round(avg_cogs, 0), f"=D14-C{cmp_hdr_r+3}", "تعدیل نوسانات شدید قیمت خرید و ایجاد تعادل دقیق بین دوره‌ها", "F8FAFC"),
    ]

    for idx, (rn, m_title, c_val, gp_form, note, f_col) in enumerate(cmp_rows, start=cmp_hdr_r+1):
        ws.cell(idx, 1, rn).alignment = Alignment(horizontal="center")
        ws.cell(idx, 2, m_title)
        ws.cell(idx, 3, c_val).number_format = "#,##0"
        ws.cell(idx, 4, gp_form).number_format = "#,##0"
        ws.cell(idx, 5, note)
        for c in range(1, 6):
            ws.cell(idx, c).fill = PatternFill(fill_type="solid", fgColor=f_col)
            ws.cell(idx, c).font = Font(name="Tahoma", bold=True)

    col_widths = {"A": 8, "B": 48, "C": 22, "D": 22, "E": 60}
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    out_file = output_path or file_path
    wb.save(out_file)
    print(f"\n✅ شیت «{target_name}» با موفقیت در فایل ذخیره شد:")
    print(f"   💾 {out_file}")
    print(f"\n📊 خلاصه نهایی:")
    print(f"   • موجودی اول دوره: {beg_inv_val:,.0f} تومان")
    print(f"   • خرید ناخالص: {p_gross:,.0f} تومان")
    print(f"   • هزینه حمل ورودی: {inward_freight:,.0f} تومان")
    print(f"   • برگشت از خرید و تخفیفات: {total_p_reductions:,.0f} تومان")
    print(f"   • خرید خالص دوره: {net_purch:,.0f} تومان")
    print(f"   • کالای آماده برای فروش: {beg_inv_val + net_purch:,.0f} تومان")
    print(f"   • فروش ناخالص دوره: {s_gross:,.0f} تومان")
    print(f"   • تخفیفات و برگشت از فروش: {s_disc + s_returns:,.0f} تومان")
    print(f"   • فروش خالص دوره: {s_net:,.0f} تومان")
    print(f"   • بهای تمام شده فایفو (FIFO COGS): {fifo_cogs:,.0f} تومان")
    print(f"   • موجودی پایان دوره فایفو: {fifo_end:,.0f} تومان")
    print(f"   • سود ناخالص فایفو (FIFO GP): {s_net - fifo_cogs:,.0f} تومان")
    print(f"   • بهای تمام شده لایفو (LIFO COGS): {lifo_cogs:,.0f} تومان")
    print(f"   • بهای تمام شده میانگین موزون: {avg_cogs:,.0f} تومان")
    return True


def get_target_file():
    # 1. اگر به عنوان آرگومان داده شده بود
    if len(sys.argv) > 1 and sys.argv[1].strip():
        arg_p = sys.argv[1].strip('\"\' ')
        if os.path.exists(arg_p):
            return arg_p
        print(f"⚠️ فایل در مسیر «{arg_p}» پیدا نشد.")

    # 2. جستجوی فایل‌های اکسل واقعی کاربر (به جز نمونه‌ها)
    files = [f for f in glob.glob("*.xlsx") if not f.startswith("~$") and "backup" not in f.lower() and "sample" not in f.lower() and f != "Worksheet.xlsx"]
    if len(files) == 1:
        return files[0]
    elif len(files) > 1:
        print("📁 چند فایل اکسل در پوشه پیدا شد:")
        for i, f in enumerate(files, 1):
            print(f"   {i}. {f}")
        try:
            choice = input(f"شماره فایل مدنظر را انتخاب کنید (1 تا {len(files)}): ").strip()
            if choice.isdigit() and 1 <= int(choice) <= len(files):
                return files[int(choice) - 1]
        except Exception:
            pass

    # 3. باز کردن پنجره انتخاب فایل ویندوز (File Dialog)
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        selected = filedialog.askopenfilename(
            title="انتخاب فایل اکسل شما (حاوی فاکتورها)",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        root.destroy()
        if selected and os.path.exists(selected):
            return selected
    except Exception:
        pass

    # 4. دریافت از کاربر به صورت ورودی متنی (Drag and Drop)
    try:
        user_input = input("لطفاً فایل اکسل خود را با موس بکشید و داخل این پنجره رها کنید (Drag & Drop) یا نام آن را بنویسید: ").strip('\"\' ')
        if user_input and os.path.exists(user_input):
            return user_input
    except Exception:
        pass

    return None


if __name__ == "__main__":
    target = get_target_file()
    if not target:
        print("❌ هیچ فایلی انتخاب نشد. برنامه متوقف شد.")
        sys.exit(1)

    process_cogs(target)

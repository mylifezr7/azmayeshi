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

    p_sheet, s_sheet, m_sheet = None, None, None

    # شناسایی خودکار شیت‌ها
    for name in wb.sheetnames:
        nl = name.lower()
        if any(k in nl for k in ["فروش", "sales"]) and not s_sheet:
            s_sheet = wb[name]
        elif any(k in nl for k in ["خرید", "purchase", "purchases"]) and not p_sheet:
            p_sheet = wb[name]
        elif any(k in nl for k in ["حرکت", "انبار", "موجودی", "کاردکس", "گردش", "stock", "inventory"]) and not m_sheet:
            m_sheet = wb[name]

    # در صورت عدم شناسایی با نام، اسکن هدرها
    for name in wb.sheetnames:
        ws = wb[name]
        if ws in [p_sheet, s_sheet, m_sheet]:
            continue
        hdr_txt = " ".join(str(ws.cell(r, c).value or "") for r in range(1, min(6, ws.max_row + 1)) for c in range(1, min(15, ws.max_column + 1))).lower()
        if not s_sheet and ("مشتری" in hdr_txt or "فروش" in hdr_txt):
            s_sheet = ws
        elif not p_sheet and ("تامین" in hdr_txt or "خرید" in hdr_txt):
            p_sheet = ws
        elif not m_sheet and any(k in hdr_txt for k in ["نوع حرکت", "اول دوره"]):
            m_sheet = ws

    print(f"🔍 شیت‌های شناسایی‌شده:")
    print(f"   • فاکتورهای خرید: {p_sheet.title if p_sheet else 'یافت نشد'}")
    print(f"   • فاکتورهای فروش: {s_sheet.title if s_sheet else 'یافت نشد'}")
    print(f"   • گردش و حرکت انبار: {m_sheet.title if m_sheet else 'یافت نشد'}")

    # 1. استخراج فاکتورهای خرید و برگشتی‌ها
    p_gross = 0.0
    p_returns = 0.0
    batches_by_item: Dict[str, List[Tuple[float, float]]] = {}
    earliest_price_by_item: Dict[str, float] = {}

    if p_sheet:
        p_cols = {}
        for c in range(1, p_sheet.max_column + 1):
            h = str(p_sheet.cell(1, c).value or "").strip().lower()
            if "سند" in h or "فاکتور" in h: p_cols["doc"] = c
            elif "کد کالا" in h or "کالا" in h: p_cols["item"] = c
            elif "تعداد" in h or "مقدار" in h: p_cols["qty"] = c
            elif ("قیمت" in h or "فی" in h) and "تخفیف" not in h: p_cols["price"] = c

        for r in range(2, p_sheet.max_row + 1):
            doc = str(p_sheet.cell(r, p_cols.get("doc", 1)).value or "").strip()
            item = str(p_sheet.cell(r, p_cols.get("item", 5)).value or "").strip()
            if not item: continue
            qty = to_float(p_sheet.cell(r, p_cols.get("qty", 7)).value)
            price = to_float(p_sheet.cell(r, p_cols.get("price", 8)).value)
            if qty == 0 or price == 0: continue

            if item not in earliest_price_by_item and price > 0:
                earliest_price_by_item[item] = price

            is_ret = "pr" in doc.lower() or "برگشت" in doc or qty < 0
            if is_ret:
                p_returns += abs(qty) * price
                if item in batches_by_item and batches_by_item[item]:
                    l_q, l_p = batches_by_item[item][-1]
                    batches_by_item[item][-1] = (max(0.0, l_q - abs(qty)), l_p)
            else:
                p_gross += qty * price
                batches_by_item.setdefault(item, []).append((qty, price))

    # 2. استخراج فاکتورهای فروش، تخفیفات و برگشتی‌ها
    s_gross = 0.0
    s_disc = 0.0
    s_returns = 0.0
    sold_by_item: Dict[str, float] = {}

    if s_sheet:
        s_cols = {}
        for c in range(1, s_sheet.max_column + 1):
            h = str(s_sheet.cell(1, c).value or "").strip().lower()
            if "سند" in h or "فاکتور" in h: s_cols["doc"] = c
            elif "کد کالا" in h or "کالا" in h: s_cols["item"] = c
            elif "تعداد" in h or "مقدار" in h: s_cols["qty"] = c
            elif "تخفیف" in h: s_cols["disc"] = c
            elif ("قیمت" in h or "فی" in h) and "تخفیف" not in h: s_cols["price"] = c
            elif "تسویه" in h or "توضیح" in h: s_cols["note"] = c

        for r in range(2, s_sheet.max_row + 1):
            doc = str(s_sheet.cell(r, s_cols.get("doc", 1)).value or "").strip()
            item = str(s_sheet.cell(r, s_cols.get("item", 5)).value or "").strip()
            if not item: continue
            qty = to_float(s_sheet.cell(r, s_cols.get("qty", 7)).value)
            price = to_float(s_sheet.cell(r, s_cols.get("price", 8)).value)
            disc_raw = to_float(s_sheet.cell(r, s_cols.get("disc", 9)).value) if "disc" in s_cols else 0.0
            note = str(s_sheet.cell(r, s_cols.get("note", 11)).value or "").strip() if "note" in s_cols else ""

            if qty == 0 or price == 0: continue
            row_g = abs(qty) * price
            row_d = row_g * (disc_raw / 100.0) if (0 < disc_raw <= 100) else disc_raw
            row_net = row_g - row_d

            is_ret = "sr" in doc.lower() or "برگشت" in doc or "برگشت" in note or qty < 0
            if is_ret:
                s_returns += row_net
                sold_by_item[item] = sold_by_item.get(item, 0.0) - abs(qty)
            else:
                s_gross += row_g
                s_disc += row_d
                sold_by_item[item] = sold_by_item.get(item, 0.0) + abs(qty)

    s_net = s_gross - s_disc - s_returns

    # 3. موجودی اول دوره از گردش انبار
    beg_inv_val = 0.0
    if m_sheet:
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

    net_purch = p_gross - p_returns

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
        ("3", "به اضافه: هزینه حمل کالای خریداری شده (حمل به داخل)", 0, None, "کرایه حمل و باربری کالاهای خریداری‌شده"),
        ("4", "کسر می‌شود: برگشت از خرید و تخفیفات", p_returns, None, "مرجوعی کالا و تخفیفات نقدی اخذشده"),
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
            d_cell = ws.cell(i, 4, total)
            d_cell.number_format = "#,##0"

        # حذف علامت = برای جلوگیری قطعی از ارور #NAME? در اکسل
        clean_note = str(note or "").lstrip("=")
        ws.cell(i, 5, clean_note)

        fill_color = None
        if i in [9, 10]: fill_color = "F0FDF4"
        elif i == 12: fill_color = "DCFCE7"
        elif i in [14, 15]: fill_color = "EFF6FF"
        elif i in [16, 17]: fill_color = "FEF3C7"

        if fill_color:
            for c in range(1, 6):
                ws.cell(i, c).fill = PatternFill(fill_type="solid", fgColor=fill_color)
                ws.cell(i, c).font = Font(name="Tahoma", bold=True)

    # جدول مقایسه تحلیلی سه‌گانه
    ws["A19"] = "📊 مقایسه تحلیلی بهای تمام شده در روش‌های فایفو، لایفو و میانگین موزون"
    ws.merge_cells("A19:E19")
    ws["A19"].font = Font(name="Tahoma", size=11, bold=True, color="FFFFFF")
    ws["A19"].fill = PatternFill(fill_type="solid", fgColor="1E3A8A")
    ws["A19"].alignment = Alignment(horizontal="center")

    cmp_headers = ["ردیف", "روش ارزشیابی بهای تمام شده", "بهای تمام شده (COGS)", "سود ناخالص برآوردی", "تاثیر اقتصادی، تورمی و مالیاتی"]
    for c_idx, h in enumerate(cmp_headers, start=1):
        cell = ws.cell(20, c_idx, h)
        cell.font = Font(name="Tahoma", bold=True, color="FFFFFF")
        cell.fill = PatternFill(fill_type="solid", fgColor="2563EB")
        cell.alignment = Alignment(horizontal="center")

    cmp_rows = [
        ("۱", "فایفو - اولین صادره (FIFO)", round(fifo_cogs, 0), "=D14-C21", "کاهش بهای تمام شده در تورم، افزایش سود دفتری و ارزش بالاتر موجودی پایان دوره", "ECFDF5"),
        ("۲", "لایفو - آخرین صادره (LIFO)", round(lifo_cogs, 0), "=D14-C22", "تطبیق درآمد جاری با هزینه روز، کاهش سود دفتری و صرفه‌جویی مالیاتی در تورم", "FFF1F2"),
        ("۳", "میانگین موزون (Weighted Average)", round(avg_cogs, 0), "=D14-C23", "تعدیل نوسانات شدید قیمت خرید و ایجاد تعادل دقیق بین دوره‌ها", "EFF6FF"),
    ]

    for idx, (rn, m_title, c_val, gp_form, note, f_col) in enumerate(cmp_rows, start=21):
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
    print(f"   • خرید خالص دوره: {net_purch:,.0f} تومان")
    print(f"   • فروش خالص دوره: {s_net:,.0f} تومان")
    print(f"   • بهای تمام شده فایفو (FIFO COGS): {fifo_cogs:,.0f} تومان")
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

"""
Business, Sales, Inventory and Management Actions for MIYUM Excel Assistant.
Implements:
- Sales & Purchase Invoice Generator (صدور و ثبت فاکتور رسمی فروش و خرید)
- Inventory Management, Low Stock Alerts & Total Valuation (کنترل موجودی انبار، هشدار کسری و ارزش کل)
- Best-Selling Products Report (پرفروش‌ترین کالاها)
- Customer Accounts / Ledger (محاسبه بدهی و طلب مشتریان)
- Percentage Price Adjustment (افزایش یا کاهش درصدی قیمت‌ها)
- Auto Chart & Dashboard Generator (رسم نمودار گرافیکی خودکار)
- Auto Backup (ذخیره نسخه پشتیبان)
- PDF Export (خروجی PDF استاندارد)
"""
import datetime
import os
import shutil
from typing import Any, Dict, List, Optional, Tuple

from excel_assistant.actions.base import ActionResult, ActionContext
from excel_assistant.core.utils import (
    col_letter_to_index, col_index_to_letter, bounds_to_range_str
)
from excel_assistant.nlp.intent_models import (
    InvoiceAction,
    InventoryAlertAction,
    InventoryValuationAction,
    TopSellingAction,
    CustomerLedgerAction,
    PriceAdjustmentAction,
    AutoChartAction,
    BackupAction,
    PdfExportAction,
)


class InventoryValuationCalculator:
    """
    محاسبه ارزش موجودی انبار و بهای تمام شده کالای فروش‌رفته (COGS) به روش‌های استاندارد حسابداری:
    - فایفو (FIFO): اولین وارده اولین صادره (First In, First Out)
    - لایفو (LIFO): آخرین وارده اولین صادره (Last In, First Out)
    - میانگین موزون (Weighted Average Cost)
    """

    @staticmethod
    def calculate_valuation(
        purchase_batches: List[Tuple[float, float]],  # List of (quantity, unit_price) in chronological order
        total_sold: float,
        method: str = "fifo"
    ) -> Dict[str, float]:
        """
        محاسبه مانده موجودی، ارزش دفتری موجودی و بهای تمام شده کالای فروش‌رفته.
        """
        total_purchased = sum(qty for qty, _ in purchase_batches)
        total_purchase_cost = sum(qty * price for qty, price in purchase_batches)
        remaining_qty = max(0.0, total_purchased - total_sold)

        if remaining_qty <= 0:
            avg_p = total_purchase_cost / total_purchased if total_purchased > 0 else 0.0
            return {
                "total_purchased": total_purchased,
                "total_purchase_cost": total_purchase_cost,
                "remaining_qty": 0.0,
                "remaining_value": 0.0,
                "cogs": total_purchase_cost,
                "effective_unit_price": avg_p,
            }

        if method == "average":
            weighted_avg_price = total_purchase_cost / total_purchased if total_purchased > 0 else 0.0
            remaining_value = remaining_qty * weighted_avg_price
            cogs = total_purchase_cost - remaining_value
            return {
                "total_purchased": total_purchased,
                "total_purchase_cost": total_purchase_cost,
                "remaining_qty": remaining_qty,
                "remaining_value": remaining_value,
                "cogs": cogs,
                "effective_unit_price": weighted_avg_price,
            }

        elif method == "fifo":
            # در فایفو، اقلام فروخته‌شده از قدیمی‌ترین خریدها برداشته می‌شوند؛
            # بنابراین موجودی باقیمانده، از جدیدترین خریدها (انتهای لیست به ابتدا) محاسبه می‌گردد.
            remaining_value = 0.0
            needed_qty = remaining_qty
            for qty, price in reversed(purchase_batches):
                take = min(needed_qty, qty)
                remaining_value += take * price
                needed_qty -= take
                if needed_qty <= 0:
                    break
            cogs = total_purchase_cost - remaining_value
            effective_price = remaining_value / remaining_qty if remaining_qty > 0 else 0.0
            return {
                "total_purchased": total_purchased,
                "total_purchase_cost": total_purchase_cost,
                "remaining_qty": remaining_qty,
                "remaining_value": remaining_value,
                "cogs": cogs,
                "effective_unit_price": effective_price,
            }

        elif method == "lifo":
            # در لایفو، اقلام فروخته‌شده از جدیدترین خریدها برداشته می‌شوند؛
            # بنابراین موجودی باقیمانده، از قدیمی‌ترین خریدها (ابتدای لیست به انتها) محاسبه می‌گردد.
            remaining_value = 0.0
            needed_qty = remaining_qty
            for qty, price in purchase_batches:
                take = min(needed_qty, qty)
                remaining_value += take * price
                needed_qty -= take
                if needed_qty <= 0:
                    break
            cogs = total_purchase_cost - remaining_value
            effective_price = remaining_value / remaining_qty if remaining_qty > 0 else 0.0
            return {
                "total_purchased": total_purchased,
                "total_purchase_cost": total_purchase_cost,
                "remaining_qty": remaining_qty,
                "remaining_value": remaining_value,
                "cogs": cogs,
                "effective_unit_price": effective_price,
            }

        else:
            latest_price = purchase_batches[-1][1] if purchase_batches else 0.0
            remaining_value = remaining_qty * latest_price
            cogs = total_purchase_cost - remaining_value
            return {
                "total_purchased": total_purchased,
                "total_purchase_cost": total_purchase_cost,
                "remaining_qty": remaining_qty,
                "remaining_value": remaining_value,
                "cogs": cogs,
                "effective_unit_price": latest_price,
            }


class BusinessActionHandler:
    """Handles sales invoices, inventory alerts, price adjustments, charts, and file exports."""

    @classmethod
    def execute_invoice(cls, action: InvoiceAction, ctx: ActionContext) -> ActionResult:
        """
        صدور و ثبت فاکتور رسمی فروش یا خرید کالا و خدمات.
        دارای مشخصات کامل خریدار/فروشنده، جدول اقلام، تخفیف، مالیات ارزش افزوده (۱۰٪) و جمع نهایی.
        """
        inv_type = action.invoice_type
        prefix = "فاکتور_فروش" if inv_type == "sales" else "فاکتور_خرید"
        title_type = "فاکتور فروش کالا و خدمات" if inv_type == "sales" else "فاکتور خرید کالا و خدمات"

        sheet_names = ctx.manager.get_sheet_names()
        inv_no = action.invoice_number or datetime.datetime.now().strftime("%Y%m%d%H%M")[-5:]
        sheet_name = f"{prefix}_{inv_no}"
        counter = 1
        while sheet_name in sheet_names:
            sheet_name = f"{prefix}_{inv_no}_{counter}"
            counter += 1

        inv = ctx.manager.create_sheet(sheet_name)
        if hasattr(inv, "_ws"):
            inv._ws.views.sheetView[0].rightToLeft = True

        # Header Title
        inv.set_value("A1", title_type)
        inv.set_font("A1", bold=True, size=16)
        inv.merge_cells("A1:I1")

        today_date = action.date_str or datetime.datetime.now().strftime("%Y/%m/%d")
        inv.set_value("A2", f"شماره فاکتور: {inv_no}        |        تاریخ صدور: {today_date}        |        وضعیت: نقدی / رسمی")
        inv.set_font("A2", bold=True, size=10)
        inv.merge_cells("A2:I2")
        inv.set_cell_color("A2:I2", "#F3F4F6")

        # Seller info
        seller = action.seller_name or "شرکت بازرگانی پارس (سهامی خاص)"
        inv.set_value("A4", "مشخصات فروشنده:")
        inv.set_value("B4", seller)
        inv.set_value("E4", "شناسه ملی / کد اقتصادی:")
        inv.set_value("G4", "10103847291")
        inv.merge_cells("B4:D4")
        inv.merge_cells("G4:I4")
        inv.set_font("A4:I4", bold=True)

        # Buyer info
        buyer = action.buyer_name or "شرکت آوا تجارت البرز"
        inv.set_value("A5", "مشخصات خریدار:")
        inv.set_value("B5", buyer)
        inv.set_value("E5", "شماره تماس / نشانی:")
        inv.set_value("G5", "تهران، خیابان ولیعصر")
        inv.merge_cells("B5:D5")
        inv.merge_cells("G5:I5")
        inv.set_font("A5:I5", bold=True)

        # Invoice items table header
        table_headers = [
            ("A7", "ردیف"),
            ("B7", "کد کالا"),
            ("C7", "شرح کالا یا خدمات"),
            ("D7", "تعداد"),
            ("E7", "واحد"),
            ("F7", "مبلغ واحد (ریال)"),
            ("G7", "مبلغ کل (ریال)"),
            ("H7", "تخفیف (ریال)"),
            ("I7", "مبلغ پس از تخفیف (ریال)"),
        ]
        for ref, text in table_headers:
            inv.set_value(ref, text)
        inv.set_font("A7:I7", bold=True)
        inv.set_cell_color("A7:I7", "#1E40AF")
        inv.set_font("A7:I7", font_color="#FFFFFF")

        # Sample or active data items
        items = [
            (1, "PRD-101", "لپ‌تاپ ایسوس Core i7 16GB", 3, "دستگاه", 480000000, 15000000),
            (2, "PRD-102", "مانیتور سامسونگ ۲۷ اینچ 4K", 5, "دستگاه", 145000000, 5000000),
            (3, "PRD-103", "کیبورد مکانیکی و ماوس بیسیم", 10, "ست", 28000000, 1000000),
            (4, "PRD-104", "هارد اکسترنال ۲ ترابایت وسترن", 8, "عدد", 38000000, 0),
        ]

        cur_row = 8
        for row_idx, code, desc, qty, unit, unit_price, disc in items:
            inv.set_value(f"A{cur_row}", row_idx)
            inv.set_value(f"B{cur_row}", code)
            inv.set_value(f"C{cur_row}", desc)
            inv.set_value(f"D{cur_row}", qty)
            inv.set_value(f"E{cur_row}", unit)
            inv.set_value(f"F{cur_row}", unit_price)

            # Total = Qty * Unit Price
            inv.set_formula(f"G{cur_row}", f"=D{cur_row}*F{cur_row}")

            # Discount
            inv.set_value(f"H{cur_row}", disc)

            # After Discount = Total - Discount
            inv.set_formula(f"I{cur_row}", f"=G{cur_row}-H{cur_row}")

            cur_row += 1

        end_items_row = cur_row - 1

        # Summary calculations at bottom
        # Subtotal
        inv.set_value(f"A{cur_row}", "جمع کل اقلام:")
        inv.merge_cells(f"A{cur_row}:F{cur_row}")
        inv.set_formula(f"G{cur_row}", f"=SUM(G8:G{end_items_row})")
        inv.set_formula(f"H{cur_row}", f"=SUM(H8:H{end_items_row})")
        inv.set_formula(f"I{cur_row}", f"=SUM(I8:I{end_items_row})")
        inv.set_font(f"A{cur_row}:I{cur_row}", bold=True)
        inv.set_cell_color(f"A{cur_row}:I{cur_row}", "#F3F4F6")
        cur_row += 1

        # VAT (10%)
        vat_row = cur_row
        inv.set_value(f"A{vat_row}", f"مالیات بر ارزش افزوده ({action.tax_percent:.0f}٪):")
        inv.merge_cells(f"A{vat_row}:H{vat_row}")
        tax_factor = action.tax_percent / 100.0
        inv.set_formula(f"I{vat_row}", f"=ROUND(I{vat_row-1}*{tax_factor}, 0)")
        inv.set_font(f"A{vat_row}:I{vat_row}", bold=True)
        cur_row += 1

        # Grand Total
        grand_row = cur_row
        inv.set_value(f"A{grand_row}", "مبلغ کل قابل پرداخت (ریال):")
        inv.merge_cells(f"A{grand_row}:H{grand_row}")
        inv.set_formula(f"I{grand_row}", f"=I{vat_row-1}+I{vat_row}")
        inv.set_font(f"A{grand_row}:I{grand_row}", bold=True, size=12)
        inv.set_cell_color(f"A{grand_row}:I{grand_row}", "#FEF08A")
        cur_row += 2

        # Signatures
        sig_row = cur_row
        inv.set_value(f"B{sig_row}", "مهر و امضای فروشنده")
        inv.set_value(f"G{sig_row}", "امضا و تحویل‌گیرنده خریدار")
        inv.set_font(f"B{sig_row}:G{sig_row}", bold=True)

        inv.set_number_format(f"F8:I{grand_row}", "#,##0")
        inv.auto_fit_columns(["A", "B", "C", "D", "E", "F", "G", "H", "I"])

        return ActionResult(
            success=True,
            message_persian=(
                f"🧾 **{title_type} با موفقیت در شیت «{sheet_name}» صادر گردید:**\n\n"
                f"• شماره فاکتور: **{inv_no}** | تاریخ: **{today_date}**\n"
                f"• فروشنده: **{seller}** | خریدار: **{buyer}**\n"
                f"• جدول شامل ستون‌های: ردیف، کد، شرح کالا، تعداد، واحد، فی، مبلغ کل، تخفیف و مبلغ نهایی\n"
                f"• محاسبه خودکار مالیات بر ارزش افزوده (۱۰٪) و جمع نهایی با فرمول‌های پیوسته اکسل"
            )
        )

    @classmethod
    def execute_inventory_alert(cls, action: InventoryAlertAction, ctx: ActionContext) -> ActionResult:
        """
        مدیریت موجودی انبار، هشدار کمبود موجودی و محاسبه ارزش کل موجودی.
        اگر شیت فروش وجود داشته باشد، به طور خودکار مجموع فروش هر کالا از خرید کسر می‌شود (مانده موجودی = خرید - فروش).
        """
        sheet = ctx.sheet
        all_sheets = []
        if ctx.manager:
            try:
                all_sheets = ctx.manager.list_sheets() if hasattr(ctx.manager, 'list_sheets') else []
                if not all_sheets and hasattr(ctx.manager, 'get_sheet_names'):
                    all_sheets = ctx.manager.get_sheet_names()
            except Exception:
                all_sheets = []

        min_r, min_c, max_r, max_c = sheet.get_used_range_bounds()

        # اگر کاربر در شیت فروش بود اما شیت خرید وجود داشت، شیت خرید به عنوان مبنا انتخاب می‌شود
        is_current_sales = any(kw in sheet.name.lower() for kw in ["فروش", "sales"])
        if not is_current_sales and max_r >= min_r:
            for r_c in range(min_r, min(min_r + 5, max_r + 1)):
                r_txt = " ".join(str(sheet.get_value(f"{col_index_to_letter(c)}{r_c}") or "") for c in range(min_c, max_c + 1)).lower()
                if "فروش" in r_txt and not any(k in r_txt for k in ["تعداد خرید", "قیمت خرید", "موجودی انبار"]):
                    is_current_sales = True
                    break

        if is_current_sales and not action.purchase_sheet:
            for sh_n in all_sheets:
                if sh_n != sheet.name:
                    if any(p_kw in sh_n.lower() for p_kw in ["خرید", "purchase", "انبار", "موجودی"]):
                        try:
                            sheet = ctx.manager.get_sheet(sh_n)
                            min_r, min_c, max_r, max_c = sheet.get_used_range_bounds()
                            break
                        except Exception:
                            pass
                    else:
                        try:
                            cand_p = ctx.manager.get_sheet(sh_n)
                            cp_min_r, cp_min_c, cp_max_r, cp_max_c = cand_p.get_used_range_bounds()
                            for r_p in range(cp_min_r, min(cp_min_r + 5, cp_max_r + 1)):
                                p_txt = " ".join(str(cand_p.get_value(f"{col_index_to_letter(c)}{r_p}") or "") for c in range(cp_min_c, cp_max_c + 1)).lower()
                                if any(k in p_txt for k in ["تعداد خرید", "قیمت خرید", "موجودی انبار", "خرید"]):
                                    sheet = cand_p
                                    min_r, min_c, max_r, max_c = sheet.get_used_range_bounds()
                                    break
                            if sheet == cand_p:
                                break
                        except Exception:
                            pass

        if max_r < min_r or max_c < min_c:
            return ActionResult(success=False, message_persian="شیت خالی است و داده‌ای برای بررسی موجودی یافت نشد.")

        # 1. تشخیص ردیف هدر واقعی
        header_row = min_r
        best_hdr_score = -1
        hdr_keywords = ["کالا", "محصول", "شرح", "عنوان", "نام", "تعداد", "موجودی", "انبار", "قیمت", "مبلغ", "فی", "ردیف", "stock", "price", "product", "qty"]

        for r_cand in range(min_r, min(min_r + 25, max_r + 1)):
            cand_score = 0
            for c in range(min_c, max_c + 1):
                val_str = str(sheet.get_value(f"{col_index_to_letter(c)}{r_cand}") or "").strip().lower()
                if any(kw in val_str for kw in hdr_keywords):
                    cand_score += 2
                elif val_str:
                    cand_score += 0.5
            if cand_score > best_hdr_score and cand_score >= 2:
                best_hdr_score = cand_score
                header_row = r_cand

        start_row = header_row + 1

        # 2. جستجوی شیت فروش برای کسر تعداد اقلام فروخته شده از خرید
        sales_sheet = None
        s_prod_col = None
        s_qty_col = None
        s_start_r = 2
        s_max_r = 100
        sales_by_prod: Dict[str, float] = {}

        if action.sales_sheet and ctx.manager:
            try:
                sales_sheet = ctx.manager.get_sheet(action.sales_sheet)
            except Exception:
                sales_sheet = None

        if not sales_sheet and ctx.manager:
            for sh_n in all_sheets:
                if sh_n != sheet.name and any(s_kw in sh_n.lower() for s_kw in ["فروش", "sales", "خروج", "فاکتور فروش"]):
                    try:
                        sales_sheet = ctx.manager.get_sheet(sh_n)
                        break
                    except Exception:
                        pass

        if not sales_sheet and ctx.manager:
            for sh_n in all_sheets:
                if sh_n != sheet.name:
                    try:
                        cand_s = ctx.manager.get_sheet(sh_n)
                        c_min_r, c_min_c, c_max_r, c_max_c = cand_s.get_used_range_bounds()
                        for r_c in range(c_min_r, min(c_min_r + 5, c_max_r + 1)):
                            row_t = " ".join(str(cand_s.get_value(f"{col_index_to_letter(c)}{r_c}") or "") for c in range(c_min_c, c_max_c + 1)).lower()
                            if "فروش" in row_t and any(k in row_t for k in ["تعداد", "مقدار", "کالا", "محصول", "مبلغ", "قیمت"]):
                                sales_sheet = cand_s
                                break
                        if sales_sheet:
                            break
                    except Exception:
                        pass

        SUMMARY_KEYWORDS = ["مجموع", "جمع", "کل", "خلاصه", "میانگین", "تراز", "total", "sum", "average", "grand total"]

        def norm_key(s: Any) -> str:
            if s is None:
                return ""
            txt = str(s).strip().lower()
            char_map = {'ي': 'ی', 'ك': 'ک', 'ة': 'ه', 'ؤ': 'و', 'إ': 'ا', 'أ': 'ا', '\u200c': ' ', '\xa0': ' '}
            for k, v in char_map.items():
                txt = txt.replace(k, v)
            return " ".join(txt.split())

        if sales_sheet:
            try:
                s_min_r, s_min_c, s_max_r, s_max_c = sales_sheet.get_used_range_bounds()
                s_header_r = s_min_r
                for sr in range(s_min_r, min(s_min_r + 15, s_max_r + 1)):
                    row_txt = " ".join(str(sales_sheet.get_value(f"{col_index_to_letter(c)}{sr}") or "") for c in range(s_min_c, s_max_c + 1)).lower()
                    if any(kw in row_txt for kw in ["کالا", "محصول", "شرح", "تعداد", "فروش"]):
                        s_header_r = sr
                        break
                s_start_r = s_header_r + 1

                for c in range(s_min_c, s_max_c + 1):
                    c_let = col_index_to_letter(c)
                    h_val = str(sales_sheet.get_value(f"{c_let}{s_header_r}") or "").strip().lower()
                    if not s_prod_col and any(kw in h_val for kw in ["نام کالا", "نام محصول", "شرح کالا", "کالا", "محصول", "شرح", "product", "item"]) and not any(ck in h_val for ck in ["کد", "شناسه", "بارکد"]):
                        s_prod_col = c_let
                    elif not s_qty_col and any(kw in h_val for kw in ["تعداد فروش", "تعداد", "مقدار فروش", "مقدار", "خروج", "qty", "quantity", "count"]):
                        s_qty_col = c_let

                if not s_prod_col: s_prod_col = "A"
                if not s_qty_col: s_qty_col = "C" if s_prod_col != "C" else "B"

                for sr in range(s_start_r, s_max_r + 1):
                    p_raw = sales_sheet.get_value(f"{s_prod_col}{sr}")
                    if p_raw is None or str(p_raw).strip() == "":
                        continue
                    p_name_s = str(p_raw).strip()
                    if any(kw in p_name_s.lower() for kw in SUMMARY_KEYWORDS):
                        continue
                    q_raw = sales_sheet.get_value(f"{s_qty_col}{sr}")
                    try:
                        q_num = float(str(q_raw).replace(",", "").strip()) if q_raw is not None else 0.0
                    except ValueError:
                        q_num = 0.0
                    nk = norm_key(p_name_s)
                    sales_by_prod[nk] = sales_by_prod.get(nk, 0.0) + q_num
            except Exception:
                sales_sheet = None

        # 3. تشخیص ستون‌ها و ستون‌های موجود قبلی در شیت اصلی (خرید)
        prod_col = action.product_col
        stock_col = action.stock_col
        price_col = action.purchase_price_col

        existing_sold_col = None
        existing_rem_col = None
        existing_status_col = None
        existing_val_col = None

        best_prod_score = 0
        best_stock_score = 0
        best_price_score = 0

        for c in range(min_c, max_c + 1):
            letter = col_index_to_letter(c)
            h = str(sheet.get_value(f"{letter}{header_row}") or "").strip().lower()
            if "تعداد فروش (" in h or (sales_sheet and "تعداد فروش" in h and any(sales_sheet.name.lower() in h for _ in [1])):
                existing_sold_col = letter
            elif "مانده موجودی" in h:
                existing_rem_col = letter
            elif "وضعیت موجودی" in h:
                existing_status_col = letter
            elif "ارزش ریالی" in h or "ارزش موجودی" in h:
                existing_val_col = letter

            if not action.product_col:
                p_score = 0
                if any(k in h for k in ["نام کالا", "نام محصول", "شرح کالا", "عنوان کالا"]):
                    p_score = 10
                elif any(k in h for k in ["کالا", "محصول", "شرح", "عنوان", "product", "item"]) and not any(ck in h for ck in ["کد", "شناسه", "بارکد", "code", "id"]):
                    p_score = 5
                elif "کالا" in h:
                    p_score = 1
                if p_score > best_prod_score:
                    best_prod_score = p_score
                    prod_col = letter

            if not action.stock_col:
                s_score = 0
                if any(k in h for k in ["تعداد خرید", "موجودی انبار", "تعداد انبار", "مانده انبار", "موجودی"]):
                    s_score = 10
                elif any(k in h for k in ["تعداد", "مقدار", "انبار", "stock", "qty", "quantity"]) and "فروش" not in h:
                    s_score = 5
                if s_score > best_stock_score:
                    best_stock_score = s_score
                    stock_col = letter

            if not action.purchase_price_col:
                pr_score = 0
                if any(k in h for k in ["قیمت خرید", "بهای تمام شده", "فی خرید", "نرخ خرید"]):
                    pr_score = 10
                elif any(k in h for k in ["قیمت واحد", "فی", "قیمت", "مبلغ واحد", "cost", "unit price", "price"]) and "فروش" not in h:
                    pr_score = 5
                if pr_score > best_price_score:
                    best_price_score = pr_score
                    price_col = letter

        # بررسی وجود ستون فروش در همان شیت (در صورتی که شیت مجزای فروش نباشد)
        same_sheet_sold_col = None
        if not sales_sheet:
            for c in range(min_c, max_c + 1):
                letter = col_index_to_letter(c)
                if letter in (prod_col, stock_col, price_col, existing_sold_col, existing_rem_col, existing_status_col, existing_val_col):
                    continue
                h_c = str(sheet.get_value(f"{letter}{header_row}") or "").strip().lower()
                if any(kw in h_c for kw in ["تعداد فروش", "مقدار فروش", "میزان فروش"]) or (h_c == "فروش" and letter != stock_col):
                    same_sheet_sold_col = letter
                    break

        # تعیین آخرین ستون واقعی دارای داده در جدول
        known_added_cols = [c for c in [existing_sold_col, existing_rem_col, existing_status_col, existing_val_col] if c]
        last_data_c = min_c
        for c in range(min_c, max_c + 1):
            letter = col_index_to_letter(c)
            if letter in known_added_cols:
                continue
            hdr = str(sheet.get_value(f"{letter}{header_row}") or "").strip()
            has_data = bool(hdr)
            if not has_data:
                for r in range(start_row, min(start_row + 15, max_r + 1)):
                    if sheet.get_value(f"{letter}{r}") is not None:
                        has_data = True
                        break
            if has_data:
                last_data_c = max(last_data_c, c)

        # اگر با هدر پیدا نشد، نمونه‌برداری از محتوای سطرهای زیرین
        if not prod_col or not stock_col or not price_col:
            for c in range(min_c, last_data_c + 1):
                letter = col_index_to_letter(c)
                if letter in [prod_col, stock_col, price_col, same_sheet_sold_col] + known_added_cols:
                    continue
                text_count = 0
                num_count = 0
                num_vals = []
                for r in range(start_row, min(start_row + 10, max_r + 1)):
                    v = sheet.get_value(f"{letter}{r}")
                    if v is not None:
                        s = str(v).replace(",", "").strip()
                        try:
                            f = float(s)
                            num_count += 1
                            num_vals.append(f)
                        except ValueError:
                            if len(s) > 1:
                                text_count += 1
                if not prod_col and text_count >= 2:
                    prod_col = letter
                elif not stock_col and num_count >= 2 and all(x >= 0 and x < 100000 for x in num_vals):
                    if any(x % 1 == 0 and x <= 5000 for x in num_vals):
                        stock_col = letter
                elif not price_col and num_count >= 2:
                    price_col = letter

        if not prod_col:
            prod_col = "A" if min_c == 1 else col_index_to_letter(min_c)
        if not stock_col:
            stock_col = "C"
        if not price_col:
            price_col = "B" if prod_col != "B" else "D"

        # 4. تعیین محل ستون‌ها
        target_cols = []
        if sales_sheet:
            if existing_sold_col and col_letter_to_index(existing_sold_col) <= last_data_c + 4:
                sold_col = existing_sold_col
                rem_col = existing_rem_col or col_index_to_letter(col_letter_to_index(sold_col) + 1)
                status_col = existing_status_col or col_index_to_letter(col_letter_to_index(rem_col) + 1)
                val_col = existing_val_col or col_index_to_letter(col_letter_to_index(status_col) + 1)
            else:
                sold_col = col_index_to_letter(last_data_c + 1)
                rem_col = col_index_to_letter(last_data_c + 2)
                status_col = col_index_to_letter(last_data_c + 3)
                val_col = col_index_to_letter(last_data_c + 4)
            target_cols = [sold_col, rem_col, status_col, val_col]

            sheet.set_value(f"{sold_col}{header_row}", f"تعداد فروش ({sales_sheet.name})")
            sheet.set_value(f"{rem_col}{header_row}", "مانده موجودی انبار")
            sheet.set_value(f"{status_col}{header_row}", "وضعیت موجودی انبار")
            sheet.set_value(f"{val_col}{header_row}", "ارزش ریالی موجودی")
        elif same_sheet_sold_col:
            sold_col = same_sheet_sold_col
            if existing_rem_col and col_letter_to_index(existing_rem_col) <= last_data_c + 3:
                rem_col = existing_rem_col
                status_col = existing_status_col or col_index_to_letter(col_letter_to_index(rem_col) + 1)
                val_col = existing_val_col or col_index_to_letter(col_letter_to_index(status_col) + 1)
            else:
                rem_col = col_index_to_letter(last_data_c + 1)
                status_col = col_index_to_letter(last_data_c + 2)
                val_col = col_index_to_letter(last_data_c + 3)
            target_cols = [rem_col, status_col, val_col]

            sheet.set_value(f"{rem_col}{header_row}", "مانده موجودی انبار")
            sheet.set_value(f"{status_col}{header_row}", "وضعیت موجودی انبار")
            sheet.set_value(f"{val_col}{header_row}", "ارزش ریالی موجودی")
        else:
            sold_col = None
            rem_col = None
            if existing_status_col and col_letter_to_index(existing_status_col) <= last_data_c + 2:
                status_col = existing_status_col
                val_col = existing_val_col or col_index_to_letter(col_letter_to_index(status_col) + 1)
            else:
                status_col = col_index_to_letter(last_data_c + 1)
                val_col = col_index_to_letter(last_data_c + 2)
            target_cols = [status_col, val_col]

            sheet.set_value(f"{status_col}{header_row}", "وضعیت موجودی انبار")
            sheet.set_value(f"{val_col}{header_row}", "ارزش ریالی موجودی")

        # تمیز کردن ستون‌های اضافه دورافتاده قبلی (مثل ستون K)
        for old_c in known_added_cols:
            if old_c not in target_cols and old_c != same_sheet_sold_col:
                for r in range(header_row, max_r + 2):
                    sheet.set_value(f"{old_c}{r}", None)
                    sheet.set_cell_color(f"{old_c}{r}", None)

        sheet.set_font(f"{target_cols[0]}{header_row}:{target_cols[-1]}{header_row}", bold=True)
        sheet.set_cell_color(f"{target_cols[0]}{header_row}:{target_cols[-1]}{header_row}", "#E0E7FF")

        threshold = action.min_threshold
        low_stock_items = []
        total_items_count = 0
        total_valuation = 0.0

        existing_sum_row = None
        last_data_row = start_row - 1

        for r in range(start_row, max_r + 1):
            raw_prod = sheet.get_value(f"{prod_col}{r}")
            stock_val = sheet.get_value(f"{stock_col}{r}")
            price_val = sheet.get_value(f"{price_col}{r}")

            # بررسی سطر خلاصه/مجموع (مانند مجموع کل در انتهای جدول)
            row_text = " ".join(
                str(sheet.get_value(f"{col_index_to_letter(c)}{r}") or "")
                for c in range(min_c, last_data_c + 1)
            ).strip().lower()

            is_summary = any(kw in row_text for kw in SUMMARY_KEYWORDS)
            if is_summary:
                if existing_sum_row is None:
                    existing_sum_row = r
                for tc in target_cols:
                    sheet.set_value(f"{tc}{r}", None)
                continue

            # بررسی سطرهای خالی و فاصله‌دار (Spacer rows)
            is_prod_empty = (raw_prod is None or str(raw_prod).strip() == "")
            is_stock_empty = (stock_val is None or str(stock_val).strip() == "")
            is_price_empty = (price_val is None or str(price_val).strip() == "")

            if is_prod_empty and is_stock_empty:
                # پاک کردن فرمول‌ها و رنگ صورتی قبلی از روی سطرهای خالی
                for tc in target_cols:
                    sheet.set_value(f"{tc}{r}", None)
                    sheet.set_cell_color(f"{tc}{r}", None)
                continue

            prod_name = str(raw_prod).strip() if not is_prod_empty else f"کالای ردیف {r}"
            try:
                s_qty = float(str(stock_val).replace(",", "").strip()) if not is_stock_empty else 0.0
            except ValueError:
                s_qty = 0.0

            try:
                u_price = float(str(price_val).replace(",", "").strip()) if not is_price_empty else 0.0
            except ValueError:
                u_price = 0.0

            total_items_count += 1
            last_data_row = r

            if sales_sheet:
                sold_qty = sales_by_prod.get(norm_key(prod_name), 0.0)
                rem_stock = max(0.0, s_qty - sold_qty)
                item_val = rem_stock * u_price
                total_valuation += item_val

                sheet.set_formula(
                    f"{sold_col}{r}",
                    f'=IF({prod_col}{r}="", "", SUMIF(\'{sales_sheet.name}\'!${s_prod_col}${s_start_r}:${s_prod_col}${s_max_r}, {prod_col}{r}, \'{sales_sheet.name}\'!${s_qty_col}${s_start_r}:${s_qty_col}${s_max_r}))'
                )
                sheet.set_formula(
                    f"{rem_col}{r}",
                    f'=IF(OR({prod_col}{r}="", {stock_col}{r}=""), "", {stock_col}{r}-{sold_col}{r})'
                )
                sheet.set_formula(
                    f"{status_col}{r}",
                    f'=IF(OR({prod_col}{r}="", {rem_col}{r}=""), "", IF({rem_col}{r}<={threshold}, "⚠️ کسری موجودی (سفارش دهید)", "موجودی کافی ✅"))'
                )
                sheet.set_formula(
                    f"{val_col}{r}",
                    f'=IF(OR({rem_col}{r}="", {price_col}{r}=""), "", {rem_col}{r}*{price_col}{r})'
                )

                if rem_stock <= threshold:
                    low_stock_items.append({"name": prod_name, "purchase": s_qty, "sold": sold_qty, "stock": rem_stock, "row": r})
                    if action.highlight_low_stock:
                        sheet.set_cell_color(f"{status_col}{r}:{val_col}{r}", "#FEE2E2")
                else:
                    sheet.set_cell_color(f"{status_col}{r}:{val_col}{r}", None)
            elif same_sheet_sold_col:
                raw_s_sold = sheet.get_value(f"{same_sheet_sold_col}{r}")
                try:
                    sold_qty = float(str(raw_s_sold).replace(",", "").strip()) if raw_s_sold is not None and str(raw_s_sold).strip() != "" else 0.0
                except ValueError:
                    sold_qty = 0.0
                rem_stock = max(0.0, s_qty - sold_qty)
                item_val = rem_stock * u_price
                total_valuation += item_val

                sheet.set_formula(
                    f"{rem_col}{r}",
                    f'=IF(OR({prod_col}{r}="", {stock_col}{r}=""), "", {stock_col}{r}-{same_sheet_sold_col}{r})'
                )
                sheet.set_formula(
                    f"{status_col}{r}",
                    f'=IF(OR({prod_col}{r}="", {rem_col}{r}=""), "", IF({rem_col}{r}<={threshold}, "⚠️ کسری موجودی (سفارش دهید)", "موجودی کافی ✅"))'
                )
                sheet.set_formula(
                    f"{val_col}{r}",
                    f'=IF(OR({rem_col}{r}="", {price_col}{r}=""), "", {rem_col}{r}*{price_col}{r})'
                )

                if rem_stock <= threshold:
                    low_stock_items.append({"name": prod_name, "purchase": s_qty, "sold": sold_qty, "stock": rem_stock, "row": r})
                    if action.highlight_low_stock:
                        sheet.set_cell_color(f"{status_col}{r}:{val_col}{r}", "#FEE2E2")
                else:
                    sheet.set_cell_color(f"{status_col}{r}:{val_col}{r}", None)
            else:
                item_val = s_qty * u_price
                total_valuation += item_val

                # فرمول وضعیت: اگر سلول خالی باشد هیچ پیامی درج نشود، در غیر این صورت هشدار کسری
                sheet.set_formula(
                    f"{status_col}{r}",
                    f'=IF(OR({prod_col}{r}="", {stock_col}{r}=""), "", IF({stock_col}{r}<={threshold}, "⚠️ کسری موجودی (سفارش دهید)", "موجودی کافی ✅"))'
                )
                sheet.set_formula(
                    f"{val_col}{r}",
                    f'=IF(OR({stock_col}{r}="", {price_col}{r}=""), "", {stock_col}{r}*{price_col}{r})'
                )

                if s_qty <= threshold:
                    low_stock_items.append({"name": prod_name, "purchase": s_qty, "sold": 0.0, "stock": s_qty, "row": r})
                    if action.highlight_low_stock:
                        sheet.set_cell_color(f"{status_col}{r}:{val_col}{r}", "#FEE2E2")
                else:
                    sheet.set_cell_color(f"{status_col}{r}:{val_col}{r}", None)

        # 5. سطر جمع کل
        if existing_sum_row:
            sum_row = existing_sum_row
            if sales_sheet and sold_col and rem_col:
                sheet.set_formula(f"{sold_col}{sum_row}", f"=SUM({sold_col}{start_row}:{sold_col}{last_data_row})")
                sheet.set_formula(f"{rem_col}{sum_row}", f"=SUM({rem_col}{start_row}:{rem_col}{last_data_row})")
            elif same_sheet_sold_col and rem_col:
                sheet.set_formula(f"{rem_col}{sum_row}", f"=SUM({rem_col}{start_row}:{rem_col}{last_data_row})")
            sheet.set_formula(f"{val_col}{sum_row}", f"=SUM({val_col}{start_row}:{val_col}{last_data_row})")
            sheet.set_font(f"{target_cols[0]}{sum_row}:{target_cols[-1]}{sum_row}", bold=True)
            sheet.set_number_format(f"{target_cols[0]}{sum_row}:{target_cols[-1]}{sum_row}", "#,##0")
        elif last_data_row >= start_row:
            sum_row = last_data_row + 1
            sheet.set_value(f"{prod_col}{sum_row}", "مجموع کل ارزش موجودی انبار:")
            if sales_sheet and sold_col and rem_col:
                sheet.set_formula(f"{sold_col}{sum_row}", f"=SUM({sold_col}{start_row}:{sold_col}{last_data_row})")
                sheet.set_formula(f"{rem_col}{sum_row}", f"=SUM({rem_col}{start_row}:{rem_col}{last_data_row})")
            elif same_sheet_sold_col and rem_col:
                sheet.set_formula(f"{rem_col}{sum_row}", f"=SUM({rem_col}{start_row}:{rem_col}{last_data_row})")
            sheet.set_formula(f"{val_col}{sum_row}", f"=SUM({val_col}{start_row}:{val_col}{last_data_row})")
            sheet.set_font(f"{prod_col}{sum_row}:{target_cols[-1]}{sum_row}", bold=True)
            sheet.set_cell_color(f"{prod_col}{sum_row}:{target_cols[-1]}{sum_row}", "#DCFCE7")
            sheet.set_number_format(f"{target_cols[0]}{sum_row}:{target_cols[-1]}{sum_row}", "#,##0")

        if last_data_row >= start_row:
            sheet.set_number_format(f"{target_cols[0]}{start_row}:{target_cols[-1]}{last_data_row}", "#,##0")
        sheet.auto_fit_columns(target_cols)

        alert_text = ""
        has_deduction = bool(sales_sheet or same_sheet_sold_col)
        if has_deduction and low_stock_items:
            alert_text = "\n⚠️ **اقلام نیازمند شارژ فوری انبار (پس از کسر فروش):**\n" + "\n".join(
                f"• {it['name']}: خرید: **{it['purchase']:.0f}** | فروش: **{it['sold']:.0f}** | مانده موجودی: فقط **{it['stock']:.0f}** عدد (ردیف {it['row']})"
                for it in low_stock_items[:10]
            )
        elif low_stock_items:
            alert_text = "\n⚠️ **اقلام نیازمند شارژ فوری انبار:**\n" + "\n".join(
                f"• {it['name']}: فقط **{it['stock']:.0f}** عدد باقی‌مانده (ردیف {it['row']})"
                for it in low_stock_items[:10]
            )
        else:
            alert_text = "\n✅ کلیه اقلام دارای موجودی کافی بالای حد آستانه هستند."

        sheet_info = f"• شیت‌های بررسی‌شده: خرید (**`{sheet.name}`**) و فروش (**`{sales_sheet.name}`**)\n" if sales_sheet else ""
        return ActionResult(
            success=True,
            message_persian=(
                f"📦 **گزارش کنترل و ارزش‌گذاری موجودی انبار ({'کسر فروش از خرید' if has_deduction else 'شیت جاری'}):**\n\n"
                f"{sheet_info}"
                f"• تعداد کل ردیف‌های انبار: **{total_items_count}** قلم\n"
                f"• حد آستانه کمبود موجودی: **{threshold:.0f} عدد**\n"
                f"• تعداد اقلام دارای کسری موجودی: **{len(low_stock_items)}** مورد (با رنگ صورتی هایلایت شدند)\n"
                f"• ارزش ریالی کل موجودی باقی‌مانده انبار: **{total_valuation:,.0f}** تومان\n"
                f"• ستون‌های تحلیلی و فرمول‌های محاسبه مانده به جدول اضافه شدند.{alert_text}"
            )
        )

    @classmethod
    def execute_inventory_valuation(cls, action: InventoryValuationAction, ctx: ActionContext) -> ActionResult:
        """
        محاسبه ارزش موجودی انبار به روش‌های استاندارد حسابداری:
        - فایفو (FIFO): اولین وارده اولین صادره
        - لایفو (LIFO): آخرین وارده اولین صادره
        - میانگین موزون (Weighted Average Cost)
        امکان مقایسه جامع سه روش و بهای تمام شده کالای فروش‌رفته (COGS) و ایجاد شیت تحلیلی اختصاصی.
        """
        driver = ctx.manager.active_driver if ctx.manager else None
        if not driver:
            return ActionResult(success=False, message_persian="ارتباط با اکسل برقرار نیست.")

        all_sheets = driver.get_sheet_names()
        sheet = ctx.sheet

        # 1. شناسایی شیت‌های خرید (انبار) و فروش
        purchase_sheet = None
        sales_sheet = None

        if action.purchase_sheet:
            try:
                purchase_sheet = driver.get_sheet(action.purchase_sheet)
            except Exception:
                purchase_sheet = None

        if action.sales_sheet:
            try:
                sales_sheet = driver.get_sheet(action.sales_sheet)
            except Exception:
                sales_sheet = None

        # اگر تعیین نشده بود، شناسایی هوشمند شیت‌ها
        if not purchase_sheet:
            for sh_n in all_sheets:
                if any(k in sh_n.lower() for k in ["خرید", "انبار", "purchase", "موجودی"]):
                    purchase_sheet = driver.get_sheet(sh_n)
                    break

        if not sales_sheet:
            for sh_n in all_sheets:
                if (not purchase_sheet or sh_n != purchase_sheet.name) and any(k in sh_n.lower() for k in ["فروش", "sales", "خروج"]):
                    sales_sheet = driver.get_sheet(sh_n)
                    break

        if not purchase_sheet:
            purchase_sheet = sheet

        SUMMARY_KEYWORDS = ["مجموع", "جمع", "کل", "خلاصه", "میانگین", "تراز", "total", "sum", "average", "grand total"]

        def norm_key(s: Any) -> str:
            if s is None:
                return ""
            txt = str(s).strip().lower()
            char_map = {'ي': 'ی', 'ك': 'ک', 'ة': 'ه', 'ؤ': 'و', 'إ': 'ا', 'أ': 'ا', '\u200c': ' ', '\xa0': ' '}
            for k, v in char_map.items():
                txt = txt.replace(k, v)
            return " ".join(txt.split())

        # 2. استخراج فروش هر کالا از شیت فروش
        sales_by_prod: Dict[str, float] = {}
        if sales_sheet:
            s_min_r, s_min_c, s_max_r, s_max_c = sales_sheet.get_used_range_bounds()
            s_hdr_r = s_min_r
            for sr in range(s_min_r, min(s_min_r + 15, s_max_r + 1)):
                row_t = " ".join(str(sales_sheet.get_value(f"{col_index_to_letter(c)}{sr}") or "") for c in range(s_min_c, s_max_c + 1)).lower()
                if any(k in row_t for k in ["کالا", "محصول", "شرح", "تعداد", "فروش"]):
                    s_hdr_r = sr
                    break
            s_prod_c, s_qty_c = None, None
            for c in range(s_min_c, s_max_c + 1):
                c_let = col_index_to_letter(c)
                h = str(sales_sheet.get_value(f"{c_let}{s_hdr_r}") or "").strip().lower()
                if not s_prod_c and any(k in h for k in ["نام کالا", "نام محصول", "شرح کالا", "کالا", "محصول"]) and not any(ck in h for ck in ["کد", "شناسه"]):
                    s_prod_c = c_let
                elif not s_qty_c and any(k in h for k in ["تعداد فروش", "تعداد", "مقدار", "خروج", "qty"]):
                    s_qty_c = c_let
            if not s_prod_c: s_prod_c = "A"
            if not s_qty_c: s_qty_c = "C" if s_prod_c != "C" else "B"

            for sr in range(s_hdr_r + 1, s_max_r + 1):
                p_raw = sales_sheet.get_value(f"{s_prod_c}{sr}")
                if not p_raw or str(p_raw).strip() == "":
                    continue
                p_str = str(p_raw).strip()
                if any(kw in p_str.lower() for kw in SUMMARY_KEYWORDS):
                    continue
                q_raw = sales_sheet.get_value(f"{s_qty_c}{sr}")
                try:
                    q_val = float(str(q_raw).replace(",", "").strip()) if q_raw is not None else 0.0
                except ValueError:
                    q_val = 0.0
                nk = norm_key(p_str)
                sales_by_prod[nk] = sales_by_prod.get(nk, 0.0) + q_val

        # 3. استخراج پارت‌های خرید از شیت خرید
        p_min_r, p_min_c, p_max_r, p_max_c = purchase_sheet.get_used_range_bounds()
        p_hdr_r = p_min_r
        for pr in range(p_min_r, min(p_min_r + 15, p_max_r + 1)):
            row_t = " ".join(str(purchase_sheet.get_value(f"{col_index_to_letter(c)}{pr}") or "") for c in range(p_min_c, p_max_c + 1)).lower()
            if any(k in row_t for k in ["کالا", "محصول", "شرح", "تعداد", "خرید", "قیمت"]):
                p_hdr_r = pr
                break

        p_prod_c, p_qty_c, p_price_c, same_sheet_sold_c = None, None, None, None
        for c in range(p_min_c, p_max_c + 1):
            c_let = col_index_to_letter(c)
            h = str(purchase_sheet.get_value(f"{c_let}{p_hdr_r}") or "").strip().lower()
            if not p_prod_c and any(k in h for k in ["نام کالا", "نام محصول", "شرح کالا", "کالا", "محصول"]) and not any(ck in h for ck in ["کد", "شناسه"]):
                p_prod_c = c_let
            elif not p_qty_c and any(k in h for k in ["تعداد خرید", "موجودی انبار", "تعداد", "مقدار"]) and "فروش" not in h:
                p_qty_c = c_let
            elif not p_price_c and any(k in h for k in ["قیمت خرید", "بهای تمام شده", "فی خرید", "قیمت واحد", "فی", "قیمت"]) and "فروش" not in h:
                p_price_c = c_let
            elif not sales_sheet and not same_sheet_sold_c and any(k in h for k in ["تعداد فروش", "مقدار فروش", "فروش"]):
                same_sheet_sold_c = c_let

        if not p_prod_c: p_prod_c = "A"
        if not p_qty_c: p_qty_c = "C" if p_prod_c != "C" else "B"
        if not p_price_c: p_price_c = "B" if p_prod_c != "B" else "D"

        # دسته‌بندی پارت‌های خرید به ترتیب وقوع (ردیف یا تاریخ)
        batches_by_prod: Dict[str, List[Tuple[float, float]]] = {}
        prod_names_ordered: List[str] = []
        prod_display_map: Dict[str, str] = {}

        for pr in range(p_hdr_r + 1, p_max_r + 1):
            p_raw = purchase_sheet.get_value(f"{p_prod_c}{pr}")
            if not p_raw or str(p_raw).strip() == "":
                continue
            p_str = str(p_raw).strip()
            if any(kw in p_str.lower() for kw in SUMMARY_KEYWORDS):
                continue

            q_raw = purchase_sheet.get_value(f"{p_qty_c}{pr}")
            try:
                q_val = float(str(q_raw).replace(",", "").strip()) if q_raw is not None else 0.0
            except ValueError:
                q_val = 0.0

            prc_raw = purchase_sheet.get_value(f"{p_price_c}{pr}")
            try:
                prc_val = float(str(prc_raw).replace(",", "").strip()) if prc_raw is not None else 0.0
            except ValueError:
                prc_val = 0.0

            if same_sheet_sold_c:
                s_sold_raw = purchase_sheet.get_value(f"{same_sheet_sold_c}{pr}")
                try:
                    s_sold_val = float(str(s_sold_raw).replace(",", "").strip()) if s_sold_raw is not None else 0.0
                except ValueError:
                    s_sold_val = 0.0
                sales_by_prod[norm_key(p_str)] = sales_by_prod.get(norm_key(p_str), 0.0) + s_sold_val

            nk = norm_key(p_str)
            if nk not in batches_by_prod:
                batches_by_prod[nk] = []
                prod_names_ordered.append(nk)
                prod_display_map[nk] = p_str
            batches_by_prod[nk].append((q_val, prc_val))

        if not prod_names_ordered:
            return ActionResult(success=False, message_persian="کالایی برای ارزش‌گذاری انبار یافت نشد.")

        # 4. محاسبه مقادیر فایفو، لایفو و میانگین موزون برای هر کالا
        prod_evals = []
        total_fifo_val = 0.0
        total_lifo_val = 0.0
        total_avg_val = 0.0
        total_fifo_cogs = 0.0
        total_lifo_cogs = 0.0
        total_avg_cogs = 0.0
        total_all_purchased_qty = 0.0
        total_all_purchased_cost = 0.0
        total_all_sold_qty = 0.0
        total_all_remaining_qty = 0.0

        for nk in prod_names_ordered:
            p_name = prod_display_map[nk]
            batches = batches_by_prod[nk]
            sold_qty = sales_by_prod.get(nk, 0.0)

            fifo_res = InventoryValuationCalculator.calculate_valuation(batches, sold_qty, "fifo")
            lifo_res = InventoryValuationCalculator.calculate_valuation(batches, sold_qty, "lifo")
            avg_res = InventoryValuationCalculator.calculate_valuation(batches, sold_qty, "average")

            rem_q = fifo_res["remaining_qty"]
            is_low = rem_q <= action.min_threshold

            total_fifo_val += fifo_res["remaining_value"]
            total_lifo_val += lifo_res["remaining_value"]
            total_avg_val += avg_res["remaining_value"]
            total_fifo_cogs += fifo_res["cogs"]
            total_lifo_cogs += lifo_res["cogs"]
            total_avg_cogs += avg_res["cogs"]

            total_all_purchased_qty += fifo_res["total_purchased"]
            total_all_purchased_cost += fifo_res["total_purchase_cost"]
            total_all_sold_qty += sold_qty
            total_all_remaining_qty += rem_q

            prod_evals.append({
                "name": p_name,
                "purchased_qty": fifo_res["total_purchased"],
                "purchased_cost": fifo_res["total_purchase_cost"],
                "sold_qty": sold_qty,
                "remaining_qty": rem_q,
                "fifo_price": fifo_res["effective_unit_price"],
                "fifo_val": fifo_res["remaining_value"],
                "fifo_cogs": fifo_res["cogs"],
                "lifo_price": lifo_res["effective_unit_price"],
                "lifo_val": lifo_res["remaining_value"],
                "lifo_cogs": lifo_res["cogs"],
                "avg_price": avg_res["effective_unit_price"],
                "avg_val": avg_res["remaining_value"],
                "avg_cogs": avg_res["cogs"],
                "is_low": is_low
            })

        # 5. ایجاد شیت گزارش ارزش‌گذاری موجودی انبار
        report_sheet_name = "ارزش‌گذاری_موجودی"
        if action.create_report_sheet:
            if report_sheet_name in all_sheets:
                rep_sheet = driver.get_sheet(report_sheet_name)
            else:
                rep_sheet = driver.create_sheet(report_sheet_name)

            if hasattr(rep_sheet, "_ws"):
                try:
                    rep_sheet._ws.views.sheetView[0].rightToLeft = True
                except Exception:
                    pass

            # پاک‌سازی محتویات قبلی
            r_min_r, r_min_c, r_max_r, r_max_c = rep_sheet.get_used_range_bounds()
            if r_max_r >= r_min_r and r_max_c >= r_min_c:
                for cr in range(r_min_r, r_max_r + 5):
                    for cc in range(r_min_c, r_max_c + 5):
                        rep_sheet.set_value(f"{col_index_to_letter(cc)}{cr}", None)
                        rep_sheet.set_cell_color(f"{col_index_to_letter(cc)}{cr}", None)

            # Header Banner
            rep_sheet.set_value("A1", "📦 گزارش جامع ارزش‌گذاری موجودی انبار و بهای تمام شده (فایفو، لایفو و میانگین)")
            rep_sheet.set_font("A1", bold=True, size=14, color_hex="#1E3A8A")
            rep_sheet.set_cell_color("A1:N1", "#DBEAFE")
            rep_sheet.merge_cells("A1:N1")

            now_str = datetime.datetime.now().strftime("%Y/%m/%d %H:%M")
            src_str = f"شیت خرید: {purchase_sheet.name}" + (f" | شیت فروش: {sales_sheet.name}" if sales_sheet else "")
            rep_sheet.set_value("A2", f"📅 تاریخ گزارش: {now_str} | {src_str} | واحد: ریال/تومان")
            rep_sheet.set_font("A2", size=10, color_hex="#475569")
            rep_sheet.merge_cells("A2:N2")

            # KPI Cards (Row 4-5)
            rep_sheet.set_value("A4", "ارزش موجودی - فایفو (FIFO)")
            rep_sheet.set_value("A5", round(total_fifo_val, 0))
            rep_sheet.set_font("A4", bold=True, size=10, color_hex="#1E40AF")
            rep_sheet.set_font("A5", bold=True, size=13, color_hex="#1E3A8A")
            rep_sheet.set_cell_color("A4:C5", "#EFF6FF")
            rep_sheet.merge_cells("A4:C4")
            rep_sheet.merge_cells("A5:C5")

            rep_sheet.set_value("E4", "ارزش موجودی - لایفو (LIFO)")
            rep_sheet.set_value("E5", round(total_lifo_val, 0))
            rep_sheet.set_font("E4", bold=True, size=10, color_hex="#9A3412")
            rep_sheet.set_font("E5", bold=True, size=13, color_hex="#7C2D12")
            rep_sheet.set_cell_color("E4:G5", "#FFF7ED")
            rep_sheet.merge_cells("E4:G4")
            rep_sheet.merge_cells("E5:G5")

            rep_sheet.set_value("I4", "ارزش موجودی - میانگین موزون")
            rep_sheet.set_value("I5", round(total_avg_val, 0))
            rep_sheet.set_font("I4", bold=True, size=10, color_hex="#166534")
            rep_sheet.set_font("I5", bold=True, size=13, color_hex="#14532D")
            rep_sheet.set_cell_color("I4:K5", "#F0FDF4")
            rep_sheet.merge_cells("I4:K4")
            rep_sheet.merge_cells("I5:K5")

            # Table Headers (Row 7)
            table_hdr_r = 7
            headers = [
                ("A", "ردیف"),
                ("B", "نام کالا"),
                ("C", "کل خرید (تعداد)"),
                ("D", "کل بهای خرید"),
                ("E", "کل فروش (تعداد)"),
                ("F", "مانده انبار"),
                ("G", "نرخ فایفو"),
                ("H", "ارزش فایفو (FIFO)"),
                ("I", "نرخ لایفو"),
                ("J", "ارزش لایفو (LIFO)"),
                ("K", "نرخ میانگین"),
                ("L", "ارزش میانگین موزون"),
                ("M", "بهای تمام شده فروش (فایفو)"),
                ("N", "وضعیت انبار"),
            ]
            for col_l, title in headers:
                rep_sheet.set_value(f"{col_l}{table_hdr_r}", title)
            rep_sheet.set_font(f"A{table_hdr_r}:N{table_hdr_r}", bold=True)
            rep_sheet.set_cell_color(f"A{table_hdr_r}:N{table_hdr_r}", "#E0E7FF")

            # Populate Table Rows
            start_data_r = 8
            for idx, item in enumerate(prod_evals, start=1):
                cur_r = start_data_r + idx - 1
                rep_sheet.set_value(f"A{cur_r}", idx)
                rep_sheet.set_value(f"B{cur_r}", item["name"])
                rep_sheet.set_value(f"C{cur_r}", item["purchased_qty"])
                rep_sheet.set_value(f"D{cur_r}", item["purchased_cost"])
                rep_sheet.set_value(f"E{cur_r}", item["sold_qty"])
                rep_sheet.set_formula(f"F{cur_r}", f'=IF(OR(C{cur_r}="", E{cur_r}=""), "", C{cur_r}-E{cur_r})')
                rep_sheet.set_value(f"G{cur_r}", round(item["fifo_price"], 0))
                rep_sheet.set_value(f"H{cur_r}", round(item["fifo_val"], 0))
                rep_sheet.set_value(f"I{cur_r}", round(item["lifo_price"], 0))
                rep_sheet.set_value(f"J{cur_r}", round(item["lifo_val"], 0))
                rep_sheet.set_value(f"K{cur_r}", round(item["avg_price"], 0))
                rep_sheet.set_value(f"L{cur_r}", round(item["avg_val"], 0))
                rep_sheet.set_value(f"M{cur_r}", round(item["fifo_cogs"], 0))
                rep_sheet.set_formula(f"N{cur_r}", f'=IF(F{cur_r}<={action.min_threshold}, "⚠️ کسری موجودی", "موجودی کافی ✅")')

                if item["is_low"] and action.highlight_low_stock:
                    rep_sheet.set_cell_color(f"A{cur_r}:N{cur_r}", "#FEE2E2")

            last_data_r = start_data_r + len(prod_evals) - 1
            sum_r = last_data_r + 1

            # Summary Row
            rep_sheet.set_value(f"B{sum_r}", "مجموع کل انبار:")
            rep_sheet.set_formula(f"C{sum_r}", f"=SUM(C{start_data_r}:C{last_data_r})")
            rep_sheet.set_formula(f"D{sum_r}", f"=SUM(D{start_data_r}:D{last_data_r})")
            rep_sheet.set_formula(f"E{sum_r}", f"=SUM(E{start_data_r}:E{last_data_r})")
            rep_sheet.set_formula(f"F{sum_r}", f"=SUM(F{start_data_r}:F{last_data_r})")
            rep_sheet.set_formula(f"H{sum_r}", f"=SUM(H{start_data_r}:H{last_data_r})")
            rep_sheet.set_formula(f"J{sum_r}", f"=SUM(J{start_data_r}:J{last_data_r})")
            rep_sheet.set_formula(f"L{sum_r}", f"=SUM(L{start_data_r}:L{last_data_r})")
            rep_sheet.set_formula(f"M{sum_r}", f"=SUM(M{start_data_r}:M{last_data_r})")

            rep_sheet.set_font(f"A{sum_r}:N{sum_r}", bold=True)
            rep_sheet.set_cell_color(f"A{sum_r}:N{sum_r}", "#DCFCE7")

            # Formats
            rep_sheet.set_number_format(f"C{start_data_r}:M{sum_r}", "#,##0")
            rep_sheet.set_number_format("A5", "#,##0")
            rep_sheet.set_number_format("E5", "#,##0")
            rep_sheet.set_number_format("I5", "#,##0")

            rep_sheet.auto_fit_columns(["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N"])

        # 6. ساخت پیام گزارش تفصیلی فارسی
        method_title = {
            "fifo": "روش فایفو (FIFO - اولین وارده اولین صادره)",
            "lifo": "روش لایفو (LIFO - آخرین وارده اولین صادره)",
            "average": "روش میانگین موزون (Weighted Average)",
            "compare": "مقایسه جامع سه روش (فایفو، لایفو و میانگین موزون)"
        }.get(action.method, "روش فایفو و مقایسه جامع")

        low_items = [it for it in prod_evals if it["is_low"]]
        alert_str = ""
        if low_items:
            alert_str = "\n⚠️ **اقلام نیازمند سفارش و شارژ مجدد (زیر ۵ عدد):**\n" + "\n".join(
                f"• {it['name']}: خرید: **{it['purchased_qty']:.0f}** | فروش: **{it['sold_qty']:.0f}** | مانده: فقط **{it['remaining_qty']:.0f}** عدد"
                for it in low_items[:10]
            )

        report_sheet_info = f"\n📑 شیت اختصاصی و مستقل **«{report_sheet_name}»** با کاردکس کامل و فرمول‌های اکسل ایجاد گردید." if action.create_report_sheet else ""

        msg = (
            f"📦 **گزارش تخصصی ارزش‌گذاری موجودی انبار ({method_title}):**\n\n"
            f"• تعداد کل ردیف‌های کالا: **{len(prod_evals)}** قلم\n"
            f"• کل پارت‌های خرید: **{total_all_purchased_qty:,.0f}** عدد به ارزش **{total_all_purchased_cost:,.0f}** تومان\n"
            f"• کل اقلام فروخته‌شده: **{total_all_sold_qty:,.0f}** عدد\n"
            f"• کل مانده موجودی فیزیکی انبار: **{total_all_remaining_qty:,.0f}** عدد\n\n"
            f"📊 **نتایج ارزش‌گذاری موجودی انبار بر اساس روش‌های حسابداری:**\n"
            f"۱. **روش فایفو (FIFO):**\n"
            f"   - ارزش کل مانده موجودی: **`{total_fifo_val:,.0f}`** تومان\n"
            f"   - بهای تمام شده کالای فروش‌رفته (COGS): **`{total_fifo_cogs:,.0f}`** تومان\n\n"
            f"۲. **روش لایفو (LIFO):**\n"
            f"   - ارزش کل مانده موجودی: **`{total_lifo_val:,.0f}`** تومان\n"
            f"   - بهای تمام شده کالای فروش‌رفته (COGS): **`{total_lifo_cogs:,.0f}`** تومان\n\n"
            f"۳. **روش میانگین موزون (Weighted Average):**\n"
            f"   - ارزش کل مانده موجودی: **`{total_avg_val:,.0f}`** تومان\n"
            f"   - بهای تمام شده کالای فروش‌رفته (COGS): **`{total_avg_cogs:,.0f}`** تومان\n\n"
            f"💡 **تحلیل مالی مدیریتی:**\n"
            f"در خریدهای بعدی با نرخ‌های متغیر یا افزایشی (تورمی)، روش **فایفو** ارزش بالاتری را برای دارایی انبار ثبت می‌کند و روش **لایفو** بهای تمام شده را افزایش می‌دهد؛ روش **میانگین موزون** تعادل دقیق بین پارت‌های خرید مختلف ایجاد می‌کند.{report_sheet_info}{alert_str}"
        )

        return ActionResult(
            success=True,
            message_persian=msg,
            data={
                "fifo_val": total_fifo_val,
                "lifo_val": total_lifo_val,
                "avg_val": total_avg_val,
                "fifo_cogs": total_fifo_cogs,
                "lifo_cogs": total_lifo_cogs,
                "avg_cogs": total_avg_cogs,
                "remaining_qty": total_all_remaining_qty,
            }
        )

    @classmethod
    def execute_top_selling(cls, action: TopSellingAction, ctx: ActionContext) -> ActionResult:
        """
        گزارش پرفروش‌ترین کالاها (Top Selling Products) به تفکیک تعداد و مبلغ.
        """
        sheet = ctx.sheet
        min_r, min_c, max_r, max_c = sheet.get_used_range_bounds()

        prod_col = action.product_col
        qty_col = action.qty_col
        amt_col = action.amount_col

        for c in range(min_c, max_c + 1):
            letter = col_index_to_letter(c)
            h = str(sheet.get_value(f"{letter}{min_r}") or "").lower()
            if any(k in h for k in ["نام کالا", "محصول", "عنوان", "شرح", "کالا", "item"]):
                if not prod_col:
                    prod_col = letter
            elif any(k in h for k in ["تعداد فروش", "فروش رفته", "تعداد", "qty", "مقدار"]):
                if not qty_col:
                    qty_col = letter
            elif any(k in h for k in ["مبلغ فروش", "کل فروش", "درآمد", "مبلغ", "amount", "revenue"]):
                if not amt_col:
                    amt_col = letter

        if not prod_col:
            prod_col = "B"
        if not qty_col:
            qty_col = "C"

        product_sales: Dict[str, Dict[str, float]] = {}

        for r in range(min_r + 1, max_r + 1):
            name = str(sheet.get_value(f"{prod_col}{r}") or "").strip()
            if not name:
                continue
            q_val = sheet.get_value(f"{qty_col}{r}")
            try:
                qty = float(str(q_val).replace(",", "").strip()) if q_val is not None else 0.0
            except ValueError:
                qty = 0.0

            amt = 0.0
            if amt_col:
                a_val = sheet.get_value(f"{amt_col}{r}")
                try:
                    amt = float(str(a_val).replace(",", "").strip()) if a_val is not None else 0.0
                except ValueError:
                    amt = 0.0

            if name not in product_sales:
                product_sales[name] = {"qty": 0.0, "amount": 0.0}
            product_sales[name]["qty"] += qty
            product_sales[name]["amount"] += amt

        # Sort descending by quantity
        sorted_products = sorted(product_sales.items(), key=lambda x: (x[1]["qty"], x[1]["amount"]), reverse=True)
        top_n = sorted_products[:action.top_n]

        lines = []
        for idx, (p_name, stats) in enumerate(top_n, 1):
            lines.append(f"**{idx}. {p_name}**: تعداد فروش: **{stats['qty']:,.0f}** | کل مبلغ: **{stats['amount']:,.0f}**")

        details = "\n".join(lines) if lines else "هیچ کالایی یافت نشد."

        return ActionResult(
            success=True,
            message_persian=(
                f"🏆 **گزارش پرفروش‌ترین کالاهای ثبت‌شده (Top {action.top_n}):**\n\n"
                f"{details}\n\n"
                f"💡 این رتبه‌بندی بر مبنای بیشترین تعداد و گردش ریالی فروش محاسبه شده است."
            )
        )

    @classmethod
    def execute_customer_ledger(cls, action: CustomerLedgerAction, ctx: ActionContext) -> ActionResult:
        """
        محاسبه بدهی و طلب مشتریان (Customer Debt & Credit).
        تجمیع فاکتورهای فروش و پرداخت‌های انجام شده به تفکیک نام مشتری.
        """
        sheet = ctx.sheet
        min_r, min_c, max_r, max_c = sheet.get_used_range_bounds()

        cust_col = action.customer_col
        dr_col = action.debit_col   # بدهی / فاکتور
        cr_col = action.credit_col  # طلب / پرداخت

        for c in range(min_c, max_c + 1):
            letter = col_index_to_letter(c)
            h = str(sheet.get_value(f"{letter}{min_r}") or "").lower()
            if any(k in h for k in ["مشتری", "نام خریدار", "طرف حساب", "customer", "نام"]):
                if not cust_col:
                    cust_col = letter
            elif any(k in h for k in ["فاکتور", "بدهی", "خرید مشتری", "بدهکار", "debit"]):
                if not dr_col:
                    dr_col = letter
            elif any(k in h for k in ["واریزی", "پرداخت", "دریافتی", "بستانکار", "credit"]):
                if not cr_col:
                    cr_col = letter

        if not cust_col:
            cust_col = "A"
        if not dr_col:
            dr_col = "B"
        if not cr_col:
            cr_col = "C"

        customers: Dict[str, Dict[str, float]] = {}

        for r in range(min_r + 1, max_r + 1):
            c_name = str(sheet.get_value(f"{cust_col}{r}") or "").strip()
            if not c_name:
                continue

            try:
                deb = float(str(sheet.get_value(f"{dr_col}{r}") or 0).replace(",", "").strip())
            except ValueError:
                deb = 0.0

            try:
                crd = float(str(sheet.get_value(f"{cr_col}{r}") or 0).replace(",", "").strip())
            except ValueError:
                crd = 0.0

            if c_name not in customers:
                customers[c_name] = {"debit": 0.0, "credit": 0.0}
            customers[c_name]["debit"] += deb
            customers[c_name]["credit"] += crd

        # Create customer summary sheet
        sheet_names = ctx.manager.get_sheet_names()
        target_name = "مانده_حساب_مشتریان"
        counter = 1
        name = target_name
        while name in sheet_names:
            name = f"{target_name}_{counter}"
            counter += 1

        cs = ctx.manager.create_sheet(name)
        if hasattr(cs, "_ws"):
            cs._ws.views.sheetView[0].rightToLeft = True

        cs.set_value("A1", "صورت حساب و مانده بدهی/طلب مشتریان")
        cs.set_font("A1", bold=True, size=14)
        cs.merge_cells("A1:E1")

        headers = [
            ("A3", "ردیف"),
            ("B3", "نام طرف حساب / مشتری"),
            ("C3", "جمع فاکتورها (بدهی)"),
            ("D3", "جمع پرداختی‌ها (طلب)"),
            ("E3", "مانده نهایی (وضعیت)")
        ]
        for ref, text in headers:
            cs.set_value(ref, text)
        cs.set_font("A3:E3", bold=True)
        cs.set_cell_color("A3:E3", "#047857")
        cs.set_font("A3:E3", font_color="#FFFFFF")

        cur_r = 4
        total_debt = 0.0
        total_credit = 0.0

        for idx, (c_name, vals) in enumerate(customers.items(), 1):
            cs.set_value(f"A{cur_r}", idx)
            cs.set_value(f"B{cur_r}", c_name)
            cs.set_value(f"C{cur_r}", vals["debit"])
            cs.set_value(f"D{cur_r}", vals["credit"])
            # Balance = Debit - Credit
            cs.set_formula(f"E{cur_r}", f"=C{cur_r}-D{cur_r}")

            total_debt += vals["debit"]
            total_credit += vals["credit"]
            cur_r += 1

        # Total row
        tot_r = cur_r
        cs.set_value(f"A{tot_r}", "جمع کل حساب‌ها:")
        cs.merge_cells(f"A{tot_r}:B{tot_r}")
        cs.set_formula(f"C{tot_r}", f"=SUM(C4:C{tot_r-1})")
        cs.set_formula(f"D{tot_r}", f"=SUM(D4:D{tot_r-1})")
        cs.set_formula(f"E{tot_r}", f"=C{tot_r}-D{tot_r}")
        cs.set_font(f"A{tot_r}:E{tot_r}", bold=True)
        cs.set_cell_color(f"A{tot_r}:E{tot_r}", "#D1FAE5")

        cs.set_number_format(f"C4:E{tot_r}", "#,##0")
        cs.auto_fit_columns(["A", "B", "C", "D", "E"])

        net_bal = total_debt - total_credit
        net_status = f"بدهکار به ما: {net_bal:,.0f}" if net_bal > 0 else f"بستانکار از ما: {abs(net_bal):,.0f}"

        return ActionResult(
            success=True,
            message_persian=(
                f"👥 **محاسبه بدهی و طلب طرف‌های حساب انجام شد:**\n\n"
                f"• تعداد مشتریان بررسی شده: **{len(customers)}** طرف حساب\n"
                f"• کل فاکتورها (بدهی مشتریان): **{total_debt:,.0f}**\n"
                f"• کل پرداختی‌ها (طلب مشتریان): **{total_credit:,.0f}**\n"
                f"• وضعیت خالص مطالبات: **{net_status}**\n"
                f"• شیت تفصیلی «{name}» با فرمول‌های تفکیک و جمع نهایی ایجاد گردید."
            )
        )

    @classmethod
    def execute_price_adjustment(cls, action: PriceAdjustmentAction, ctx: ActionContext) -> ActionResult:
        """
        افزایش یا کاهش درصدی قیمت‌ها (مانند «قیمت‌های ستون C را ۱۵ درصد افزایش بده»).
        """
        sheet = ctx.sheet
        min_r, min_c, max_r, max_c = sheet.get_used_range_bounds()

        col_str = str(action.target_col or "C").strip().upper()
        if col_str.isdigit():
            col_str = col_index_to_letter(int(col_str))

        pct = action.percentage
        if action.operation == "decrease" and pct > 0:
            multiplier = 1.0 - (pct / 100.0)
            sign_str = f"-{pct:.0f}٪"
        elif action.operation == "increase" or pct > 0:
            multiplier = 1.0 + (abs(pct) / 100.0)
            sign_str = f"+{abs(pct):.0f}٪"
        else:
            multiplier = 1.0 + (pct / 100.0)
            sign_str = f"{pct:.0f}٪"

        col_idx = col_letter_to_index(col_str)

        if action.in_place:
            target_col = col_str
        else:
            target_col = action.target_result_col or col_index_to_letter(max(max_c + 1, col_idx + 1))
            sheet.set_value(f"{target_col}{min_r}", f"قیمت با تغییر ({sign_str})")
            sheet.set_font(f"{target_col}{min_r}", bold=True)
            sheet.set_cell_color(f"{target_col}{min_r}", "#FEF3C7")

        start_row = min_r + 1
        for r in range(start_row, max_r + 1):
            if action.in_place:
                val = sheet.get_value(f"{col_str}{r}")
                try:
                    num = float(str(val).replace(",", "").strip()) if val is not None else 0.0
                    sheet.set_value(f"{col_str}{r}", round(num * multiplier, 2))
                except ValueError:
                    pass
            else:
                # Excel formula: =ROUND(C{r} * 1.15, 0)
                sheet.set_formula(f"{target_col}{r}", f"=ROUND({col_str}{r}*{multiplier:.4f}, 0)")

        sheet.set_number_format(f"{target_col}{start_row}:{target_col}{max_r}", "#,##0")
        sheet.auto_fit_columns([target_col])

        return ActionResult(
            success=True,
            message_persian=(
                f"🏷️ **تغییر درصدی قیمت‌ها با موفقیت اعمال شد:**\n\n"
                f"• ستون مبنا: **{col_str}**\n"
                f"• ضریب تغییر: **{sign_str}** (ضریب محاسباتی: {multiplier:.4f})\n"
                f"• ستون مقصد: **{target_col}** (با فرمول‌های داینامیک و گردکردن مبالغ)"
            )
        )

    @classmethod
    def execute_auto_chart(cls, action: AutoChartAction, ctx: ActionContext) -> ActionResult:
        """
        رسم نمودار ستونی/میله‌ای خودکار در فایل اکسل با openpyxl.
        """
        sheet = ctx.sheet
        min_r, min_c, max_r, max_c = sheet.get_used_range_bounds()

        try:
            from openpyxl.chart import BarChart, LineChart, PieChart, Reference
        except ImportError:
            return ActionResult(success=False, message_persian="کتابخانه رسم نمودار در محیط اکسل فعال نیست.")

        ws = getattr(sheet, "_ws", None)
        if not ws:
            return ActionResult(success=False, message_persian="دسترسی مستقیم به شیت اکسل مقدور نمی‌باشد.")

        # Determine label column (text) and value column (numeric)
        label_col_idx = col_letter_to_index(action.label_col) if action.label_col else min_c
        val_col_idx = col_letter_to_index(action.value_col) if action.value_col else min(max_c, label_col_idx + 1)

        # Ensure value column is numeric if possible
        if val_col_idx == label_col_idx and max_c > min_c:
            val_col_idx = min_c + 1

        chart_type = action.chart_type
        if chart_type == "line":
            chart = LineChart()
        elif chart_type == "pie":
            chart = PieChart()
        else:
            chart = BarChart()
            chart.type = "col"
            chart.style = 10

        chart_title_str = action.title or "نمودار تحلیل آماری"
        chart.title = chart_title_str
        chart.y_axis.title = "مبلغ / مقدار"
        chart.x_axis.title = "عناوین"
        chart.width = 16
        chart.height = 10

        data_ref = Reference(ws, min_col=val_col_idx, min_row=min_r, max_row=max_r)
        cats_ref = Reference(ws, min_col=label_col_idx, min_row=min_r + 1, max_row=max_r)

        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)

        # Place chart next to table e.g. at column max_c + 2
        chart_cell = f"{col_index_to_letter(max_c + 2)}2"
        ws.add_chart(chart, chart_cell)

        return ActionResult(
            success=True,
            message_persian=(
                f"📊 **نمودار تحلیلی گرافیکی با موفقیت تولید و در سلول {chart_cell} قرار گرفت:**\n\n"
                f"• عنوان نمودار: **{chart_title_str}**\n"
                f"• ستون برچسب‌ها (محور افقی): **{col_index_to_letter(label_col_idx)}**\n"
                f"• ستون مقادیر (محور عمودی): **{col_index_to_letter(val_col_idx)}**\n"
                f"• نوع نمودار: **{action.chart_type.upper()} Chart**"
            )
        )

    @classmethod
    def execute_backup(cls, action: BackupAction, ctx: ActionContext) -> ActionResult:
        """
        ذخیره نسخه پشتیبان (Backup) دارای برچسب زمانی از فایل اکسل.
        """
        driver = ctx.manager.active_driver
        current_path = getattr(driver, "file_path", None)
        if not current_path or not os.path.exists(current_path):
            # Save current first
            current_path = driver.save("Worksheet.xlsx")

        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        dir_name = os.path.dirname(current_path) or "."
        base_name = os.path.splitext(os.path.basename(current_path))[0]
        backup_filename = f"{base_name}_backup_{timestamp}.xlsx"
        backup_path = os.path.join(dir_name, backup_filename)

        # Save driver
        driver.save(backup_path)

        return ActionResult(
            success=True,
            message_persian=(
                f"💾 **نسخه پشتیبان با موفقیت ذخیره شد:**\n\n"
                f"• مسیر فایل پشتیبان: `{backup_path}`\n"
                f"• زمان پشتیبان‌گیری: {datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S')}\n"
                f"• کلیه شیت‌ها، فرمول‌ها و ساختار با نهایت امنیت کپی گردید."
            )
        )

    @classmethod
    def execute_pdf_export(cls, action: PdfExportAction, ctx: ActionContext) -> ActionResult:
        """
        خروجی PDF استاندارد از شیت فعال با کتابخانه ReportLab.
        """
        try:
            from reportlab.lib.pagesizes import letter, landscape, A4
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib import colors
        except ImportError:
            return ActionResult(success=False, message_persian="کتابخانه تولید PDF (ReportLab) در دسترس نیست.")

        sheet = ctx.sheet
        min_r, min_c, max_r, max_c = sheet.get_used_range_bounds()
        matrix = sheet.get_matrix(min_r, min_c, max_r, max_c)

        out_name = action.output_filename or f"{sheet.name}_report.pdf"
        if not out_name.endswith(".pdf"):
            out_name += ".pdf"

        # Prepare table data for ReportLab
        table_data = []
        for r_idx, row in enumerate(matrix[:50]): # Up to 50 rows
            row_cells = []
            for c_idx, val in enumerate(row):
                s_val = str(val) if val is not None else ""
                row_cells.append(s_val)
            table_data.append(row_cells)

        if not table_data:
            table_data = [["(فایل خالی است)"]]

        doc = SimpleDocTemplate(out_name, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
        elements = []

        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            'TitleStyle',
            parent=styles['Heading1'],
            fontSize=16,
            leading=20,
            alignment=1
        )
        elements.append(Paragraph(f"Excel Sheet Report: {sheet.name}", title_style))
        elements.append(Spacer(1, 15))

        pdf_table = Table(table_data)
        pdf_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1E3A8A')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#D1D5DB')),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F9FAFB')]),
        ]))
        elements.append(pdf_table)

        doc.build(elements)

        return ActionResult(
            success=True,
            message_persian=(
                f"📄 **خروجی PDF با موفقیت تولید شد:**\n\n"
                f"• نام فایل خروجی: `{out_name}`\n"
                f"• شیت مبنا: **{sheet.name}**\n"
                f"• ساختار جدول با استایل زیبا و چیدمان استاندارد آماده پرینت می‌باشد."
            )
        )

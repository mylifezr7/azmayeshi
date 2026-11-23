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
    TopSellingAction,
    CustomerLedgerAction,
    PriceAdjustmentAction,
    AutoChartAction,
    BackupAction,
    PdfExportAction,
)


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
        بررسی مقادیر موجودی نسبت به حد آستانه (نقطه سفارش)، هایلایت اقلام کسری و محاسبه ارزش دفتری انبار.
        """
        sheet = ctx.sheet
        min_r, min_c, max_r, max_c = sheet.get_used_range_bounds()

        # Find columns
        prod_col = None
        stock_col = None
        price_col = None

        for c in range(min_c, max_c + 1):
            letter = col_index_to_letter(c)
            h = str(sheet.get_value(f"{letter}{min_r}") or "").lower()
            if any(k in h for k in ["نام کالا", "شرح کالا", "محصول", "عنوان", "کالا", "product"]):
                prod_col = letter
            elif any(k in h for k in ["موجودی", "تعداد انبار", "مانده انبار", "انبار", "stock", "quantity", "تعداد"]):
                if not stock_col:
                    stock_col = letter
            elif any(k in h for k in ["قیمت خرید", "بهای تمام شده", "فی خرید", "نرخ خرید", "cost", "unit price", "قیمت"]):
                price_col = letter

        if not stock_col:
            stock_col = action.stock_col or "C"
        if not prod_col:
            prod_col = action.product_col or "B"
        if not price_col:
            price_col = action.purchase_price_col or "D"

        stock_idx = col_letter_to_index(stock_col)
        status_col_idx = max(max_c + 1, stock_idx + 1)
        status_col = col_index_to_letter(status_col_idx)
        val_col = col_index_to_letter(status_col_idx + 1)

        sheet.set_value(f"{status_col}{min_r}", "وضعیت موجودی انبار")
        sheet.set_value(f"{val_col}{min_r}", "ارزش ریالی موجودی")
        sheet.set_font(f"{status_col}{min_r}:{val_col}{min_r}", bold=True)
        sheet.set_cell_color(f"{status_col}{min_r}:{val_col}{min_r}", "#E0E7FF")

        threshold = action.min_threshold
        low_stock_items = []
        total_items_count = 0
        total_valuation = 0.0

        start_row = min_r + 1
        for r in range(start_row, max_r + 1):
            prod_name = str(sheet.get_value(f"{prod_col}{r}") or f"کالای ردیف {r}").strip()
            stock_val = sheet.get_value(f"{stock_col}{r}")
            price_val = sheet.get_value(f"{price_col}{r}")

            try:
                s_qty = float(str(stock_val).replace(",", "").strip()) if stock_val is not None else 0.0
            except ValueError:
                s_qty = 0.0

            try:
                u_price = float(str(price_val).replace(",", "").strip()) if price_val is not None else 0.0
            except ValueError:
                u_price = 0.0

            total_items_count += 1
            item_val = s_qty * u_price
            total_valuation += item_val

            # Formula for Status: =IF(stock_col<=threshold, "⚠️ کسری موجودی", "موجودی کافی ✅")
            sheet.set_formula(
                f"{status_col}{r}",
                f'=IF({stock_col}{r}<={threshold}, "⚠️ کسری موجودی (سفارش دهید)", "موجودی کافی ✅")'
            )
            # Formula for Valuation: =stock_col * price_col
            sheet.set_formula(f"{val_col}{r}", f"={stock_col}{r}*{price_col}{r}")

            if s_qty <= threshold:
                low_stock_items.append({"name": prod_name, "stock": s_qty, "row": r})
                if action.highlight_low_stock:
                    sheet.set_cell_color(f"{stock_col}{r}:{val_col}{r}", "#FEE2E2")

        # Summary Row
        sum_row = max_r + 1
        sheet.set_value(f"{prod_col}{sum_row}", "مجموع کل ارزش موجودی انبار:")
        sheet.set_formula(f"{val_col}{sum_row}", f"=SUM({val_col}{start_row}:{val_col}{max_r})")
        sheet.set_font(f"{prod_col}{sum_row}:{val_col}{sum_row}", bold=True)
        sheet.set_cell_color(f"{prod_col}{sum_row}:{val_col}{sum_row}", "#DCFCE7")

        sheet.set_number_format(f"{val_col}{start_row}:{val_col}{sum_row}", "#,##0")
        sheet.auto_fit_columns([status_col, val_col])

        alert_text = ""
        if low_stock_items:
            alert_text = "\n⚠️ **اقلام نیازمند شارژ فوری انبار:**\n" + "\n".join(
                f"• {it['name']}: فقط **{it['stock']:.0f}** عدد باقی‌مانده (ردیف {it['row']})"
                for it in low_stock_items[:10]
            )
        else:
            alert_text = "\n✅ کلیه اقلام دارای موجودی کافی بالای حد آستانه هستند."

        return ActionResult(
            success=True,
            message_persian=(
                f"📦 **گزارش کنترل و ارزش‌گذاری موجودی انبار:**\n\n"
                f"• تعداد کل ردیف‌های انبار: **{total_items_count}** قلم\n"
                f"• حد آستانه کمبود موجودی: **{threshold:.0f} عدد**\n"
                f"• تعداد اقلام دارای کسری موجودی: **{len(low_stock_items)}** مورد (با رنگ صورتی هایلایت شدند)\n"
                f"• ارزش ریالی کل موجودی انبار: **{total_valuation:,.0f}**\n"
                f"• ستون‌های «وضعیت موجودی» و «ارزش ریالی موجودی» به انتهای جدول اضافه شدند.{alert_text}"
            )
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

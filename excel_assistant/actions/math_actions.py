"""
Mathematical and Statistical Actions for Excel Assistant.
Handles column arithmetic (multiply, add, subtract, divide) and statistical summaries.
"""
from collections import defaultdict
import datetime
import re
import statistics
from typing import Any, List, Optional

from excel_assistant.actions.base import ActionResult, ActionContext
from excel_assistant.nlp.intent_models import (
    ColumnArithmeticAction, ColumnStatsAction, ApplyDiscountAction, SalesSummaryAction
)
from excel_assistant.core.utils import (
    col_letter_to_index, col_index_to_letter, bounds_to_range_str
)


class MathActionHandler:
    """Executes mathematical formulas and statistical calculations on Excel sheets."""

    @classmethod
    def execute_arithmetic(cls, action: ColumnArithmeticAction, ctx: ActionContext) -> ActionResult:
        sheet = ctx.sheet
        min_r, min_c, max_r, max_c = sheet.get_used_range_bounds()

        # Extract and normalize all source columns (support letters or numbers e.g. '2', '4', '5' -> 'B', 'D', 'E')
        if action.columns and len(action.columns) >= 2:
            raw_cols = action.columns
        else:
            raw_cols = [action.col1 or "A", action.col2 or "B"]

        cols: List[str] = []
        for c in raw_cols:
            c_str = str(c).strip().upper()
            if c_str.isdigit():
                c_str = col_index_to_letter(int(c_str))
            if c_str and c_str not in cols:
                cols.append(c_str)

        if len(cols) < 2:
            cols = ["A", "B"]

        # Target column determination (support letter or number e.g. '8' -> 'H')
        if action.target_col:
            t_str = str(action.target_col).strip().upper()
            if t_str.isdigit():
                target_col = col_index_to_letter(int(t_str))
            else:
                target_col = t_str
        else:
            # Safely pick the next empty column after all operands and used range
            max_col_idx = max(col_letter_to_index(c) for c in cols)
            next_idx = max(max_col_idx, max_c) + 1
            target_col = col_index_to_letter(next_idx)

        op_symbols = {"multiply": "*", "add": "+", "subtract": "-", "divide": "/"}
        op_symbol = op_symbols.get(action.operation, "*")
        op_fa_names = {"multiply": "ضرب", "add": "جمع", "subtract": "تفریق", "divide": "تقسیم"}
        op_fa = op_fa_names.get(action.operation, "محاسبه")

        def is_numeric_value(v):
            if v is None:
                return False
            if isinstance(v, (int, float)):
                return True
            try:
                float(str(v).replace(',', '').strip())
                return True
            except ValueError:
                return False

        # Determine whether row 1 contains header text or numeric data
        has_header = False
        for c in cols:
            val_r1 = sheet.get_value(f"{c}1")
            if val_r1 is not None and not is_numeric_value(val_r1):
                has_header = True
                break

        if has_header:
            start_row = 2
            # Add header to target column
            header_title = f"{op_fa} {' و '.join(cols)}"
            sheet.set_value(f"{target_col}1", header_title)
            sheet.set_font(f"{target_col}1", bold=True)
        else:
            start_row = 1

        # Determine last row with actual data in any operand column
        max_data_len = start_row
        for c in cols:
            c_vals = sheet.get_column_values(c, include_header=True)
            for r_i in range(len(c_vals), 0, -1):
                v = c_vals[r_i - 1]
                if v is not None and str(v).strip() != "":
                    max_data_len = max(max_data_len, r_i)
                    break
        end_row = max(start_row, max_data_len)

        # Circular Reference Guard: If target column is one of the source columns
        is_circular = (target_col in cols)

        if is_circular:
            # Compute values in Python and write numeric results so values are never zeroed
            calc_matrix = []
            for r in range(start_row, end_row + 1):
                row_nums = []
                for c in cols:
                    v = sheet.get_value(f"{c}{r}")
                    try:
                        n = float(str(v).replace(',', '').strip()) if v is not None else 0.0
                    except Exception:
                        n = 0.0
                    row_nums.append(n)

                if action.operation == "add":
                    res_val = sum(row_nums)
                elif action.operation == "multiply":
                    res_val = 1.0
                    for n in row_nums:
                        res_val *= n
                elif action.operation == "subtract":
                    res_val = row_nums[0]
                    for n in row_nums[1:]:
                        res_val -= n
                elif action.operation == "divide":
                    res_val = row_nums[0]
                    for n in row_nums[1:]:
                        res_val = (res_val / n) if n != 0 else "#DIV/0!"
                else:
                    res_val = sum(row_nums)

                calc_matrix.append([res_val])
            target_range = f"{target_col}{start_row}:{target_col}{end_row}"
            sheet.set_value(target_range, calc_matrix)
        else:
            # Set exact formula matrix for each row e.g. =B2+D2+E2
            formula_matrix = [
                [f"={op_symbol.join(f'{c}{r}' for c in cols)}"]
                for r in range(start_row, end_row + 1)
            ]
            target_range = f"{target_col}{start_row}:{target_col}{end_row}"
            sheet.set_formula(target_range, formula_matrix)

        sheet.auto_fit_columns([target_col])
        affected = f"{target_col}{start_row}:{target_col}{end_row}"

        # Trigger Excel calculation refresh if running via xlwings COM
        try:
            if hasattr(sheet, "_sheet") and hasattr(sheet._sheet, "book"):
                sheet._sheet.book.app.calculate()
        except Exception:
            pass

        cols_display = "، ".join(cols)
        msg = (
            f"✅ عملیات **{op_fa}** ستون‌های `{cols_display}` با موفقیت انجام شد.\n"
            f"📌 نتیجه در ستون `{target_col}` (محدوده `{affected}`) درج شد."
        )
        return ActionResult(success=True, message_persian=msg, affected_range=affected)

    @classmethod
    def execute_stats(cls, action: ColumnStatsAction, ctx: ActionContext) -> ActionResult:
        sheet = ctx.sheet
        source_col = action.source_col.upper()
        values = sheet.get_column_values(source_col, include_header=False)

        # Extract numeric values
        num_vals: List[float] = []
        for v in values:
            if v is not None and not str(v).startswith('='):
                try:
                    num_vals.append(float(str(v).replace(',', '').strip()))
                except ValueError:
                    pass

        min_r, min_c, max_r, max_c = sheet.get_used_range_bounds()
        start_r = min_r + 1
        end_r = max_r
        range_ref = f"{source_col}{start_r}:{source_col}{end_r}"

        formula_map = {
            "sum": f"=SUM({range_ref})",
            "average": f"=AVERAGE({range_ref})",
            "max": f"=MAX({range_ref})",
            "min": f"=MIN({range_ref})",
            "median": f"=MEDIAN({range_ref})",
            "count": f"=COUNT({range_ref})"
        }

        stat_fa_map = {
            "sum": "مجموع",
            "average": "میانگین",
            "median": "میانه",
            "max": "بیشترین مقدار (حداکثر)",
            "min": "کمترین مقدار (حداقل)",
            "count": "تعداد داده‌ها",
            "percentage": "محاسبه درصد"
        }
        stat_fa = stat_fa_map.get(action.stat_type, "آمار")

        # Compute calculated value in Python
        calc_val = None
        if num_vals:
            if action.stat_type == "sum":
                calc_val = sum(num_vals)
            elif action.stat_type == "average":
                calc_val = statistics.mean(num_vals)
            elif action.stat_type == "median":
                calc_val = statistics.median(num_vals)
            elif action.stat_type == "max":
                calc_val = max(num_vals)
            elif action.stat_type == "min":
                calc_val = min(num_vals)
            elif action.stat_type == "count":
                calc_val = len(num_vals)

        # If user specified a cell or if we place summary at the bottom
        target_cell = action.target_cell
        if not target_cell:
            target_cell = f"{source_col}{max_r + 1}"
            label_cell = f"{col_index_to_letter(max(1, col_letter_to_index(source_col) - 1))}{max_r + 1}"
            sheet.set_value(label_cell, f"{stat_fa}:")
            sheet.set_font(label_cell, bold=True)

        if action.stat_type in formula_map:
            sheet.set_formula(target_cell, formula_map[action.stat_type])
            sheet.set_font(target_cell, bold=True)

        val_display = f"{calc_val:,.2f}".rstrip('0').rstrip('.') if calc_val is not None else "محاسبه با فرمول"
        msg = (
            f"📊 **{stat_fa}** ستون `{source_col}` محاسبه گردید:\n"
            f"🔹 مقدار محاسبه‌شده: **{val_display}**\n"
            f"🔹 فرمول اکسل `{formula_map.get(action.stat_type, '')}` در سلول `{target_cell}` ثبت شد."
        )
        return ActionResult(success=True, message_persian=msg, affected_range=target_cell, data=calc_val)

    @classmethod
    def execute_discount(cls, action: ApplyDiscountAction, ctx: ActionContext) -> ActionResult:
        sheet = ctx.sheet
        min_r, min_c, max_r, max_c = sheet.get_used_range_bounds()

        # 1. Normalize price column (support number e.g. '2' -> 'B' or 'B' -> 'B')
        p_raw = str(action.price_col).strip().upper()
        p_col = col_index_to_letter(int(p_raw)) if p_raw.isdigit() else p_raw

        # 2. Normalize discount column or fixed percent
        d_col = None
        if action.discount_col:
            d_raw = str(action.discount_col).strip().upper()
            d_col = col_index_to_letter(int(d_raw)) if d_raw.isdigit() else d_raw

        pct = action.discount_percent

        # 3. Determine target column (final price after discount)
        if action.target_col:
            t_raw = str(action.target_col).strip().upper()
            target_col = col_index_to_letter(int(t_raw)) if t_raw.isdigit() else t_raw
        else:
            used_indices = [col_letter_to_index(p_col)]
            if d_col:
                used_indices.append(col_letter_to_index(d_col))
            next_idx = max(max(used_indices), max_c) + 1
            target_col = col_index_to_letter(next_idx)

        # 4. Check row 1 for header
        val_r1 = sheet.get_value(f"{p_col}1")
        has_header = False
        if val_r1 is not None:
            try:
                float(str(val_r1).replace(',', '').strip())
            except ValueError:
                has_header = True

        start_row = 2 if has_header else 1

        if has_header:
            sheet.set_value(f"{target_col}1", "قیمت بعد از تخفیف")
            sheet.set_font(f"{target_col}1", bold=True)

        # 5. Determine end row
        p_vals = sheet.get_column_values(p_col, include_header=True)
        end_row = start_row
        for r_i in range(len(p_vals), 0, -1):
            v = p_vals[r_i - 1]
            if v is not None and str(v).strip() != "":
                end_row = max(start_row, r_i)
                break

        # 6. Generate formulas for final discounted price
        formula_matrix = []
        for r in range(start_row, end_row + 1):
            if d_col:
                # Handle both integer percentages (e.g. 20 -> 20%) and decimal percentages (e.g. 0.20)
                row_f = f"=IF({d_col}{r}>1, {p_col}{r}*(1-{d_col}{r}/100), {p_col}{r}*(1-{d_col}{r}))"
            elif pct is not None:
                row_f = f"={p_col}{r}*(1-{pct}/100)"
            else:
                row_f = f"={p_col}{r}"
            formula_matrix.append([row_f])

        target_range = f"{target_col}{start_row}:{target_col}{end_row}"
        sheet.set_formula(target_range, formula_matrix)
        sheet.auto_fit_columns([target_col])

        # 7. If optional discount amount column requested
        if action.discount_amount_col:
            da_raw = str(action.discount_amount_col).strip().upper()
            da_col = col_index_to_letter(int(da_raw)) if da_raw.isdigit() else da_raw
            if has_header:
                sheet.set_value(f"{da_col}1", "مبلغ تخفیف")
                sheet.set_font(f"{da_col}1", bold=True)
            da_matrix = []
            for r in range(start_row, end_row + 1):
                if d_col:
                    da_f = f"=IF({d_col}{r}>1, {p_col}{r}*({d_col}{r}/100), {p_col}{r}*{d_col}{r})"
                elif pct is not None:
                    da_f = f"={p_col}{r}*({pct}/100)"
                else:
                    da_f = "=0"
                da_matrix.append([da_f])
            da_range = f"{da_col}{start_row}:{da_col}{end_row}"
            sheet.set_formula(da_range, da_matrix)
            sheet.auto_fit_columns([da_col])

        # Trigger calculation refresh
        try:
            if hasattr(sheet, "_sheet") and hasattr(sheet._sheet, "book"):
                sheet._sheet.book.app.calculate()
        except Exception:
            pass

        d_desc = f"ستون `{d_col}`" if d_col else f"`{pct}%`"
        msg = (
            f"🏷️ **محاسبه تخفیف کالاها با موفقیت انجام شد.**\n"
            f"📌 ستون قیمت اولیه: `{p_col}`\n"
            f"📌 درصد تخفیف: {d_desc}\n"
            f"📌 ستون قیمت بعد از تخفیف: `{target_col}` (محدوده `{target_range}`)"
        )
        return ActionResult(success=True, message_persian=msg, affected_range=target_range)

    @classmethod
    def execute_sales_summary(cls, action: SalesSummaryAction, ctx: ActionContext) -> ActionResult:
        sheet = ctx.sheet
        min_r, min_c, max_r, max_c = sheet.get_used_range_bounds()

        # 1. Determine Amount & Date Columns
        date_col = None
        amount_col = None

        if action.date_col:
            d_raw = str(action.date_col).strip().upper()
            date_col = col_index_to_letter(int(d_raw)) if d_raw.isdigit() else d_raw
        if action.amount_col:
            a_raw = str(action.amount_col).strip().upper()
            amount_col = col_index_to_letter(int(a_raw)) if a_raw.isdigit() else a_raw

        # Auto-detect if missing
        if not date_col or not amount_col:
            col_scores_date = {}
            col_scores_amount = {}

            date_keywords = ["تاریخ", "date", "روز", "زمان", "time", "day", "month", "سال"]
            amount_keywords = ["فروش", "مبلغ", "درآمد", "قیمت", "sales", "amount", "price", "total", "جمع", "fee"]

            for c_idx in range(min_c, max_c + 1):
                c_let = col_index_to_letter(c_idx)
                hdr = str(sheet.get_value(f"{c_let}{min_r}") or "").lower()

                d_score = 0
                a_score = 0
                for kw in date_keywords:
                    if kw in hdr:
                        d_score += 10
                for kw in amount_keywords:
                    if kw in hdr:
                        a_score += 10

                for r in range(min_r + 1, min(min_r + 6, max_r + 1)):
                    val = sheet.get_value(f"{c_let}{r}")
                    if val is not None:
                        s_val = str(val).strip()
                        if hasattr(val, 'strftime') or re.search(r'\d{4}[/-]\d{1,2}[/-]\d{1,2}', s_val):
                            d_score += 5
                        elif s_val.replace(',', '').replace('.', '').isdigit():
                            a_score += 3

                col_scores_date[c_let] = d_score
                col_scores_amount[c_let] = a_score

            if not date_col:
                best_date = max(col_scores_date, key=col_scores_date.get) if col_scores_date else "A"
                date_col = best_date if col_scores_date.get(best_date, 0) > 0 else "A"

            if not amount_col:
                sorted_amounts = sorted(col_scores_amount.items(), key=lambda x: x[1], reverse=True)
                for c_let, sc in sorted_amounts:
                    if c_let != date_col:
                        amount_col = c_let
                        break
                if not amount_col:
                    amount_col = "B" if date_col != "B" else "C"

        # 2. Extract Data
        date_vals = sheet.get_column_values(date_col, include_header=False)
        amount_vals = sheet.get_column_values(amount_col, include_header=False)

        records = []
        for d_raw, a_raw in zip(date_vals, amount_vals):
            if d_raw is None and a_raw is None:
                continue
            amt = 0.0
            if a_raw is not None:
                s_amt = str(a_raw).replace(',', '').replace('تومان', '').replace('ریال', '').strip()
                try:
                    amt = float(s_amt)
                except ValueError:
                    amt = 0.0

            d_str = None
            if d_raw is not None:
                if hasattr(d_raw, 'strftime'):
                    d_str = d_raw.strftime("%Y/%m/%d")
                else:
                    s_d = str(d_raw).strip()
                    m = re.search(r'(\d{4})[/-](\d{1,2})[/-](\d{1,2})', s_d)
                    if m:
                        d_str = f"{int(m.group(1)):04d}/{int(m.group(2)):02d}/{int(m.group(3)):02d}"
                    elif len(s_d) >= 4:
                        d_str = s_d

            if amt > 0 or d_str:
                records.append({
                    "date": d_str or "نامشخص",
                    "amount": amt
                })

        if not records:
            return ActionResult(
                success=False,
                message_persian=f"⚠️ هیچ دادهٔ معتبری برای فروش در ستون‌های تاریخ `{date_col}` و مبلغ `{amount_col}` یافت نشد."
            )

        # 3. Aggregate
        total_sales = sum(r["amount"] for r in records)
        total_count = len(records)
        avg_sale = total_sales / total_count if total_count > 0 else 0.0
        max_sale = max(r["amount"] for r in records) if records else 0.0

        daily_map = defaultdict(lambda: {"sum": 0.0, "count": 0})
        monthly_map = defaultdict(lambda: {"sum": 0.0, "count": 0})
        yearly_map = defaultdict(lambda: {"sum": 0.0, "count": 0})

        for r in records:
            dt = r["date"]
            amt = r["amount"]
            daily_map[dt]["sum"] += amt
            daily_map[dt]["count"] += 1

            mo_key = dt[:7] if len(dt) >= 7 and dt[4] in ('/', '-') else dt
            monthly_map[mo_key]["sum"] += amt
            monthly_map[mo_key]["count"] += 1

            yr_key = dt[:4] if len(dt) >= 4 and dt[:4].isdigit() else "نامشخص"
            yearly_map[yr_key]["sum"] += amt
            yearly_map[yr_key]["count"] += 1

        sorted_daily = sorted(daily_map.items(), key=lambda x: str(x[0]))
        sorted_monthly = sorted(monthly_map.items(), key=lambda x: str(x[0]))
        sorted_yearly = sorted(yearly_map.items(), key=lambda x: str(x[0]))

        best_day = max(daily_map.items(), key=lambda x: x[1]["sum"]) if daily_map else (None, {"sum": 0, "count": 0})
        best_month = max(monthly_map.items(), key=lambda x: x[1]["sum"]) if monthly_map else (None, {"sum": 0, "count": 0})

        # 4. Generate Report in Excel Sheet
        report_sheet_name = "گزارش_فروش"
        if action.create_report_sheet:
            existing_sheets = ctx.manager.active_driver.get_sheet_names()
            if report_sheet_name in existing_sheets:
                report_sheet = ctx.manager.active_driver.get_sheet(report_sheet_name)
            else:
                report_sheet = ctx.manager.active_driver.create_sheet(report_sheet_name)

            curr_r = 1
            # Banner Title
            report_sheet.set_value(f"A{curr_r}", "📊 داشبورد و گزارش جامع فروش (روزانه، ماهانه، سالانه و کل)")
            report_sheet.set_font(f"A{curr_r}", bold=True, size=14, color_hex="#38BDF8")
            report_sheet.set_cell_color(f"A{curr_r}", "#09132B")
            curr_r += 2

            # Section 1: Summary Table
            report_sheet.set_value(f"A{curr_r}", "📌 خلاصه کل عملکرد فروش")
            report_sheet.set_font(f"A{curr_r}", bold=True, size=12)
            curr_r += 1

            headers_s1 = ["شاخص کلیدی", "مقدار عددی", "توضیحات"]
            for idx, h in enumerate(headers_s1):
                col_c = col_index_to_letter(idx + 1)
                report_sheet.set_value(f"{col_c}{curr_r}", h)
                report_sheet.set_font(f"{col_c}{curr_r}", bold=True)
                report_sheet.set_cell_color(f"{col_c}{curr_r}", "#1E3A8A")
            curr_r += 1

            kpis = [
                ("💰 جمع کل فروش", f"{total_sales:,.0f} تومان", f"بر اساس ستون {amount_col}"),
                ("🔢 تعداد کل فاکتورها / تراکنش‌ها", f"{total_count:,}", "تعداد ردیف‌های ثبت شده"),
                ("📈 میانگین ارزش هر فاکتور", f"{avg_sale:,.0f} تومان", "فروش کل تقسیم بر تعداد"),
                ("🏆 بالاترین مبلغ ثبت‌شده در یک فاکتور", f"{max_sale:,.0f} تومان", "حداکثر مبلغ فروش"),
                ("⭐ پرفروش‌ترین روز", f"{best_day[0]} ({best_day[1]['sum']:,.0f} تومان)", f"{best_day[1]['count']} فاکتور"),
                ("⭐ پرفروش‌ترین ماه", f"{best_month[0]} ({best_month[1]['sum']:,.0f} تومان)", f"{best_month[1]['count']} فاکتور")
            ]
            for title, val, desc in kpis:
                report_sheet.set_value(f"A{curr_r}", title)
                report_sheet.set_value(f"B{curr_r}", val)
                report_sheet.set_value(f"C{curr_r}", desc)
                report_sheet.set_font(f"B{curr_r}", bold=True)
                curr_r += 1
            curr_r += 2

            # Section 2: Yearly Sales Table
            report_sheet.set_value(f"A{curr_r}", "📆 گزارش فروش سالانه")
            report_sheet.set_font(f"A{curr_r}", bold=True, size=12)
            curr_r += 1

            headers_yr = ["سال", "مبلغ فروش سالانه (تومان)", "تعداد فاکتور", "سهم از کل فروش"]
            for idx, h in enumerate(headers_yr):
                col_c = col_index_to_letter(idx + 1)
                report_sheet.set_value(f"{col_c}{curr_r}", h)
                report_sheet.set_font(f"{col_c}{curr_r}", bold=True)
                report_sheet.set_cell_color(f"{col_c}{curr_r}", "#0F766E")
            curr_r += 1

            for yr, data in sorted_yearly:
                share_pct = (data["sum"] / total_sales * 100) if total_sales > 0 else 0
                report_sheet.set_value(f"A{curr_r}", str(yr))
                report_sheet.set_value(f"B{curr_r}", f"{data['sum']:,.0f}")
                report_sheet.set_value(f"C{curr_r}", data["count"])
                report_sheet.set_value(f"D{curr_r}", f"{share_pct:.1f}%")
                curr_r += 1
            curr_r += 2

            # Section 3: Monthly Sales Table
            report_sheet.set_value(f"A{curr_r}", "🗓️ گزارش فروش ماهانه")
            report_sheet.set_font(f"A{curr_r}", bold=True, size=12)
            curr_r += 1

            headers_mo = ["دوره / ماه", "مبلغ فروش ماهانه (تومان)", "تعداد فاکتور", "میانگین هر فاکتور"]
            for idx, h in enumerate(headers_mo):
                col_c = col_index_to_letter(idx + 1)
                report_sheet.set_value(f"{col_c}{curr_r}", h)
                report_sheet.set_font(f"{col_c}{curr_r}", bold=True)
                report_sheet.set_cell_color(f"{col_c}{curr_r}", "#1E40AF")
            curr_r += 1

            for mo, data in sorted_monthly:
                mo_avg = data["sum"] / data["count"] if data["count"] > 0 else 0
                report_sheet.set_value(f"A{curr_r}", str(mo))
                report_sheet.set_value(f"B{curr_r}", f"{data['sum']:,.0f}")
                report_sheet.set_value(f"C{curr_r}", data["count"])
                report_sheet.set_value(f"D{curr_r}", f"{mo_avg:,.0f}")
                curr_r += 1
            curr_r += 2

            # Section 4: Daily Sales Table
            report_sheet.set_value(f"A{curr_r}", "📅 گزارش فروش روزانه")
            report_sheet.set_font(f"A{curr_r}", bold=True, size=12)
            curr_r += 1

            headers_dy = ["تاریخ روز", "مبلغ فروش روزانه (تومان)", "تعداد فاکتور", "میانگین هر فاکتور"]
            for idx, h in enumerate(headers_dy):
                col_c = col_index_to_letter(idx + 1)
                report_sheet.set_value(f"{col_c}{curr_r}", h)
                report_sheet.set_font(f"{col_c}{curr_r}", bold=True)
                report_sheet.set_cell_color(f"{col_c}{curr_r}", "#0369A1")
            curr_r += 1

            for dy, data in sorted_daily:
                dy_avg = data["sum"] / data["count"] if data["count"] > 0 else 0
                report_sheet.set_value(f"A{curr_r}", str(dy))
                report_sheet.set_value(f"B{curr_r}", f"{data['sum']:,.0f}")
                report_sheet.set_value(f"C{curr_r}", data["count"])
                report_sheet.set_value(f"D{curr_r}", f"{dy_avg:,.0f}")
                curr_r += 1

            report_sheet.auto_fit_columns()
            ctx.manager.activate_sheet(report_sheet_name)

        # 5. Format Persian response message
        msg_lines = [
            "📊 **محاسبه و تحلیل جامع فروش با موفقیت انجام شد:**",
            f"💰 **جمع کل فروش:** `{total_sales:,.0f}` تومان (تعداد: `{total_count:,}` تراکنش)",
            f"📈 **میانگین هر فاکتور:** `{avg_sale:,.0f}` تومان | **بالاترین فاکتور:** `{max_sale:,.0f}` تومان",
            "",
            "📅 **خلاصه فروش روزانه:**",
            f"• تعداد روزهای ثبت‌شده: `{len(daily_map)}` روز",
            f"• پرفروش‌ترین روز: `{best_day[0]}` با مبلغ `{best_day[1]['sum']:,.0f}` تومان",
            "",
            "🗓️ **خلاصه فروش ماهانه:**",
            f"• دوره‌های ماهانه: `{len(monthly_map)}` ماه",
            f"• پرفروش‌ترین ماه: `{best_month[0]}` با مجموع `{best_month[1]['sum']:,.0f}` تومان",
            "",
            "📆 **خلاصه فروش سالانه:**"
        ]
        for yr, data in sorted_yearly:
            pct = (data['sum'] / total_sales * 100) if total_sales > 0 else 0
            msg_lines.append(f"• سال `{yr}`: `{data['sum']:,.0f}` تومان (`{data['count']}` فاکتور - سهم: `{pct:.1f}%`)")

        if action.create_report_sheet:
            msg_lines.append("")
            msg_lines.append("📑 شیت اختصاصی **«گزارش_فروش»** با جداول تفکیک‌شده روزانه، ماهانه، سالانه و کل ایجاد گردید.")

        return ActionResult(
            success=True,
            message_persian="\n".join(msg_lines),
            affected_range=report_sheet_name,
            data={
                "total_sales": total_sales,
                "total_count": total_count,
                "avg_sale": avg_sale,
                "daily_count": len(daily_map),
                "monthly_count": len(monthly_map),
                "yearly_count": len(yearly_map)
            }
        )

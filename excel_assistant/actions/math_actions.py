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
    ColumnArithmeticAction, ColumnStatsAction, ApplyDiscountAction, SalesSummaryAction, ProfitCalculationAction
)
from excel_assistant.core.utils import (
    col_letter_to_index, col_index_to_letter, bounds_to_range_str, to_float_safe
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
        op_fa_names = {"multiply": "ضرب", "add": "جمع", "subtract": "تفریق", "divide": "تقسیم", "average": "میانگین"}
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
                elif action.operation == "average":
                    res_val = sum(row_nums) / len(row_nums) if row_nums else 0.0
                else:
                    res_val = sum(row_nums)

                calc_matrix.append([res_val])
            target_range = f"{target_col}{start_row}:{target_col}{end_row}"
            sheet.set_value(target_range, calc_matrix)
        else:
            # Set exact formula matrix for each row e.g. =B2+D2+E2 or =AVERAGE(B2, D2)
            if action.operation == "average":
                formula_matrix = [
                    [f"=AVERAGE({', '.join(f'{c}{r}' for c in cols)})"]
                    for r in range(start_row, end_row + 1)
                ]
            else:
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
        # Determine last data row excluding summary/formula cells at the bottom
        end_r = start_r
        for r_idx in range(start_r, max_r + 1):
            form = sheet.get_formula(f"{source_col}{r_idx}")
            if form and str(form).startswith("="):
                break
            val = sheet.get_value(f"{source_col}{r_idx}")
            if val is not None and str(val).strip() != "":
                end_r = max(end_r, r_idx)

        range_ref = f"{source_col}{start_r}:{source_col}{end_r}"

        formula_map = {
            "sum": f"=SUM({range_ref})",
            "average": f"=AVERAGE({range_ref})",
            "max": f"=MAX({range_ref})",
            "min": f"=MIN({range_ref})",
            "median": f"=MEDIAN({range_ref})",
            "count": f"=COUNT({range_ref})",
            "counta": f"=COUNTA({range_ref})"
        }

        stat_fa_map = {
            "sum": "جمع خودکار (مجموع)",
            "average": "میانگین",
            "median": "میانه",
            "max": "بیشترین مقدار (حداکثر)",
            "min": "کمترین مقدار (حداقل)",
            "count": "تعداد سلول‌های عددی",
            "counta": "تعداد سلول‌های پر (غیرخالی)",
            "percentage": "محاسبه درصد"
        }
        stat_fa = stat_fa_map.get(action.stat_type, "آمار")

        # Compute calculated value in Python
        calc_val = None
        if action.stat_type == "counta":
            calc_val = sum(1 for v in values if v is not None and str(v).strip() != "")
        elif num_vals:
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
        qty_col = None
        price_col = None
        disc_col = None
        computed_desc = None

        if action.date_col:
            d_raw = str(action.date_col).strip().upper()
            date_col = col_index_to_letter(int(d_raw)) if d_raw.isdigit() else d_raw

        # Scan sheet headers for structural detection
        explicit_total_col = None
        total_kws = ["مبلغ کل", "مبلغ فروش", "جمع کل", "مبلغ فاکتور", "خالص فروش", "فروش کل", "مبلغ نهایی", "کل مبلغ", "قیمت بعد از تخفیف", "مبلغ بعد از تخفیف", "مبلغ خالص", "total amount", "total sales"]
        qty_kws = ["تعداد", "مقدار", "حجم", "qty", "quantity"]
        price_kws = ["قیمت واحد", "بهای واحد", "فی", "نرخ", "unit price"]
        disc_kws = ["درصد تخفیف", "تخفیف درصد", "تخفیف", "discount"]
        disc_exclude_kws = ["بعد از تخفیف", "پس از تخفیف", "با تخفیف", "بعد تخفیف", "قیمت بعد", "مبلغ بعد", "خالص پس از تخفیف"]

        for c_idx in range(min_c, max_c + 1):
            c_let = col_index_to_letter(c_idx)
            hdr = str(sheet.get_value(f"{c_let}{min_r}") or "").lower()
            if not explicit_total_col and any(k in hdr for k in total_kws):
                explicit_total_col = c_let
            if not qty_col and any(k in hdr for k in qty_kws):
                qty_col = c_let
            if not price_col and any(k in hdr for k in price_kws):
                price_col = c_let
            if not disc_col and any(k in hdr for k in disc_kws) and not any(k in hdr for k in disc_exclude_kws):
                disc_col = c_let

        if action.amount_col:
            a_raw = str(action.amount_col).strip().upper()
            chosen_col = col_index_to_letter(int(a_raw)) if a_raw.isdigit() else a_raw
            chosen_hdr = str(sheet.get_value(f"{chosen_col}{min_r}") or "").lower()
            # If the user or wizard selected a unit price column (e.g. "قیمت واحد (ریال)") instead of total amount
            if any(k in chosen_hdr for k in price_kws) and not any(k in chosen_hdr for k in total_kws):
                price_col = chosen_col
                amount_col = None
            else:
                amount_col = chosen_col

        if not amount_col:
            if explicit_total_col:
                amount_col = explicit_total_col
            elif qty_col and price_col:
                computed_desc = f"تعداد ({qty_col}) × قیمت واحد ({price_col}){' - تخفیف (' + disc_col + ')' if disc_col else ''}"

        # Auto-detect date_col or amount_col if still missing
        if not date_col or (not amount_col and not computed_desc):
            col_scores_date = {}
            col_scores_amount = {}

            date_keywords = ["تاریخ", "date", "روز", "زمان", "time", "day", "month", "سال"]
            amount_keywords = ["مبلغ", "فروش", "درآمد", "قیمت", "sales", "amount", "price", "total", "جمع", "fee"]

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

            if not amount_col and not computed_desc:
                sorted_amounts = sorted(col_scores_amount.items(), key=lambda x: x[1], reverse=True)
                for c_let, sc in sorted_amounts:
                    if c_let != date_col:
                        amount_col = c_let
                        break
                if not amount_col:
                    amount_col = "B" if date_col != "B" else "C"

        # 2. Extract Data
        date_vals = sheet.get_column_values(date_col, include_header=False)
        if computed_desc and qty_col and price_col:
            qty_vals = sheet.get_column_values(qty_col, include_header=False)
            price_vals = sheet.get_column_values(price_col, include_header=False)
            disc_vals = sheet.get_column_values(disc_col, include_header=False) if disc_col else [0] * len(qty_vals)
            amount_vals = []
            for q_raw, p_raw, d_raw in zip(qty_vals, price_vals, disc_vals):
                q = to_float_safe(q_raw)
                p = to_float_safe(p_raw)
                d = to_float_safe(d_raw)
                if d > 1.0:
                    d = d / 100.0
                d = max(0.0, min(1.0, d))
                amount_vals.append(max(0.0, q * p * (1.0 - d)))
        else:
            amount_vals = sheet.get_column_values(amount_col, include_header=False)

        records = []
        for d_raw, a_raw in zip(date_vals, amount_vals):
            if d_raw is None and a_raw is None:
                continue
            amt = to_float_safe(a_raw) if a_raw is not None else 0.0

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
                try:
                    ub_r1, ub_c1, ub_r2, ub_c2 = report_sheet.get_used_range_bounds()
                    if ub_r2 >= ub_r1 and ub_r2 > 0:
                        report_sheet.delete_rows(list(range(ub_r1, ub_r2 + 1)))
                except Exception:
                    pass
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
                report_sheet.set_font(f"{col_c}{curr_r}", bold=True, color_hex="#FFFFFF")
                report_sheet.set_cell_color(f"{col_c}{curr_r}", "#1E3A8A")
            curr_r += 1

            kpi_basis = f"بر اساس {computed_desc}" if computed_desc else f"بر اساس ستون {amount_col}"
            kpis = [
                ("💰 جمع کل فروش", f"{total_sales:,.0f} تومان", kpi_basis),
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
                report_sheet.set_font(f"{col_c}{curr_r}", bold=True, color_hex="#FFFFFF")
                report_sheet.set_cell_color(f"{col_c}{curr_r}", "#0F766E")
            curr_r += 1

            start_yr_r = curr_r
            for yr, data in sorted_yearly:
                share_pct = (data["sum"] / total_sales * 100) if total_sales > 0 else 0
                report_sheet.set_value(f"A{curr_r}", str(yr))
                report_sheet.set_value(f"B{curr_r}", f"{data['sum']:,.0f}")
                report_sheet.set_value(f"C{curr_r}", data["count"])
                report_sheet.set_value(f"D{curr_r}", f"{share_pct:.1f}%")
                curr_r += 1
            try:
                report_sheet.set_number_format(f"B{start_yr_r}:B{curr_r-1}", "#,##0")
            except Exception:
                pass
            curr_r += 2

            # Section 3: Monthly Sales Table
            report_sheet.set_value(f"A{curr_r}", "🗓️ گزارش فروش ماهانه")
            report_sheet.set_font(f"A{curr_r}", bold=True, size=12)
            curr_r += 1

            headers_mo = ["دوره / ماه", "مبلغ فروش ماهانه (تومان)", "تعداد فاکتور", "میانگین هر فاکتور"]
            for idx, h in enumerate(headers_mo):
                col_c = col_index_to_letter(idx + 1)
                report_sheet.set_value(f"{col_c}{curr_r}", h)
                report_sheet.set_font(f"{col_c}{curr_r}", bold=True, color_hex="#FFFFFF")
                report_sheet.set_cell_color(f"{col_c}{curr_r}", "#1E40AF")
            curr_r += 1

            start_mo_r = curr_r
            for mo, data in sorted_monthly:
                mo_avg = data["sum"] / data["count"] if data["count"] > 0 else 0
                report_sheet.set_value(f"A{curr_r}", str(mo))
                report_sheet.set_value(f"B{curr_r}", f"{data['sum']:,.0f}")
                report_sheet.set_value(f"C{curr_r}", data["count"])
                report_sheet.set_value(f"D{curr_r}", f"{mo_avg:,.0f}")
                curr_r += 1
            try:
                report_sheet.set_number_format(f"B{start_mo_r}:B{curr_r-1}", "#,##0")
                report_sheet.set_number_format(f"D{start_mo_r}:D{curr_r-1}", "#,##0")
            except Exception:
                pass
            curr_r += 2

            # Section 4: Daily Sales Table
            report_sheet.set_value(f"A{curr_r}", "📅 گزارش فروش روزانه")
            report_sheet.set_font(f"A{curr_r}", bold=True, size=12)
            curr_r += 1

            headers_dy = ["تاریخ روز", "مبلغ فروش روزانه (تومان)", "تعداد فاکتور", "میانگین هر فاکتور"]
            for idx, h in enumerate(headers_dy):
                col_c = col_index_to_letter(idx + 1)
                report_sheet.set_value(f"{col_c}{curr_r}", h)
                report_sheet.set_font(f"{col_c}{curr_r}", bold=True, color_hex="#FFFFFF")
                report_sheet.set_cell_color(f"{col_c}{curr_r}", "#0369A1")
            curr_r += 1

            start_dy_r = curr_r
            for dy, data in sorted_daily:
                dy_avg = data["sum"] / data["count"] if data["count"] > 0 else 0
                report_sheet.set_value(f"A{curr_r}", str(dy))
                report_sheet.set_value(f"B{curr_r}", f"{data['sum']:,.0f}")
                report_sheet.set_value(f"C{curr_r}", data["count"])
                report_sheet.set_value(f"D{curr_r}", f"{dy_avg:,.0f}")
                curr_r += 1
            try:
                report_sheet.set_number_format(f"B{start_dy_r}:B{curr_r-1}", "#,##0")
                report_sheet.set_number_format(f"D{start_dy_r}:D{curr_r-1}", "#,##0")
            except Exception:
                pass

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

    @classmethod
    def execute_profit_calculation(cls, action: ProfitCalculationAction, ctx: ActionContext) -> ActionResult:
        driver = ctx.manager.active_driver
        sheet_names = driver.get_sheet_names()

        if not sheet_names:
            return ActionResult(success=False, message_persian="⚠️ هیچ شیتی در ورک‌بوک یافت نشد.")

        # 1. Resolve Sales Sheet (Sheet 1) and Purchase Sheet (Sheet 2) with Smart Auto-Detection
        auto_sales_sheet = None
        auto_purch_sheet = None
        sales_sheet_kws = ["فاکتورهای فروش", "فاکتورها", "فروش", "فاکتور", "sales", "invoice"]
        purch_sheet_kws = ["خرید و موجودی", "موجودی کالا", "خرید", "موجودی", "انبار", "purchase", "inventory", "stock"]

        for sn in sheet_names:
            sn_l = sn.lower()
            if not auto_sales_sheet and any(k in sn_l for k in sales_sheet_kws) and not any(k in sn_l for k in purch_sheet_kws):
                auto_sales_sheet = sn
            elif not auto_purch_sheet and any(k in sn_l for k in purch_sheet_kws) and not any(k in sn_l for k in sales_sheet_kws):
                auto_purch_sheet = sn

        def resolve_sheet_name(target_val, fallback):
            if not target_val:
                return fallback
            val_str = str(target_val).strip()
            if val_str.isdigit():
                idx = int(val_str) - 1
                if 0 <= idx < len(sheet_names):
                    return sheet_names[idx]
            # Exact match first
            for name in sheet_names:
                if name.lower() == val_str.lower():
                    return name
            # Substring match
            for name in sheet_names:
                if val_str.lower() in name.lower():
                    return name
            return fallback

        default_sales = auto_sales_sheet or (ctx.sheet.name if ctx.sheet and ctx.sheet.name in sheet_names else sheet_names[0])
        default_purch = auto_purch_sheet or (sheet_names[1] if len(sheet_names) >= 2 else default_sales)

        sales_sheet_name = resolve_sheet_name(action.sales_sheet, default_sales)
        purchase_sheet_name = resolve_sheet_name(action.purchase_sheet, default_purch)

        val_method = getattr(action, "valuation_method", "fifo") or "fifo"
        if val_method not in ["fifo", "lifo", "average"]:
            val_method = "fifo"

        method_persian = {
            "fifo": "فایفو (FIFO)",
            "lifo": "لایفو (LIFO)",
            "average": "میانگین موزون (Weighted Average)"
        }[val_method]

        s_sheet = ctx.manager.get_sheet(sales_sheet_name)
        p_sheet = ctx.manager.get_sheet(purchase_sheet_name)

        # 2. Inspect Sales Sheet Columns
        s_min_r, s_min_c, s_max_r, s_max_c = s_sheet.get_used_range_bounds()
        p_min_r, p_min_c, p_max_r, p_max_c = p_sheet.get_used_range_bounds()

        def clean_val(v):
            from excel_assistant.core.utils import to_float_safe
            return to_float_safe(v, default=0.0)

        # Auto-detect columns in Sales Sheet
        s_qty_col = None
        s_price_col = None
        s_prod_col = None
        s_date_col = None
        s_disc_col = None
        s_explicit_profit_col = None
        s_explicit_cost_col = None
        s_total_col = None

        cust_keywords = ["مشتری", "خریدار", "طرف حساب", "customer", "client", "فروشنده", "نام مشتری"]
        inv_keywords = ["شماره فاکتور", "شماره سند", "کد فاکتور", "فاکتور", "invoice", "doc"]
        disc_keywords = ["درصد تخفیف", "تخفیف درصد", "تخفیف", "discount", "off"]
        disc_exclude_keywords = [
            "بعد از تخفیف", "پس از تخفیف", "با تخفیف", "بعد تخفیف",
            "قیمت بعد", "مبلغ بعد", "خالص پس از تخفیف", "قیمت با تخفیف", "مبلغ با تخفیف"
        ]
        date_keywords = ["تاریخ", "date", "زمان", "time"]
        qty_keywords = ["تعداد فروش", "تعداد", "مقدار", "حجم", "qty", "quantity", "count"]
        qty_exclude_keywords = ["خرید", "مانده", "موجودی", "قیمت", "مبلغ"]
        price_keywords = ["قیمت واحد", "فی فروش", "قیمت فروش", "مبلغ فروش", "بهای واحد", "قیمت", "فی", "بها", "نرخ", "price", "unit price"]
        price_exclude_keywords = ["خرید", "تمام شده", "بعد از تخفیف", "پس از تخفیف", "با تخفیف", "تخفیف"]
        prod_keywords = ["شرح کالا", "نام کالا", "کالا", "محصول", "شرح", "نام", "item", "product"]
        profit_keywords = ["سود خالص", "سود", "profit", "درآمد خالص", "مارجین"]
        cost_keywords = ["بهای تمام شده", "قیمت تمام شده", "خرید", "هزینه", "cost"]
        total_sales_keywords = [
            "مبلغ کل", "مبلغ فروش", "جمع کل", "مبلغ فاکتور", "خالص فروش",
            "فروش کل", "مبلغ نهایی", "کل مبلغ", "مبلغ", "total amount", "total sales"
        ]

        all_s_hdrs = {}
        for c in range(s_min_c, s_max_c + 1):
            c_let = col_index_to_letter(c)
            hdr = str(s_sheet.get_value(f"{c_let}{s_min_r}") or "").strip().lower()
            if not hdr:
                continue
            all_s_hdrs[c_let] = hdr

            # 1. Discount column: MUST NOT be a "price after discount" or "amount after discount"
            if not s_disc_col and any(k in hdr for k in disc_keywords) and not any(k in hdr for k in disc_exclude_keywords):
                s_disc_col = c_let
            # 2. Date column
            elif not s_date_col and any(k in hdr for k in date_keywords) and not any(k in hdr for k in qty_keywords + inv_keywords + disc_keywords):
                s_date_col = c_let
            # 3. Product column
            elif not s_prod_col and any(k in hdr for k in prod_keywords) and not any(k in hdr for k in cust_keywords + inv_keywords + qty_keywords + price_keywords):
                s_prod_col = c_let
            # 4. Quantity column
            elif not s_qty_col and any(k in hdr for k in qty_keywords) and not any(k in hdr for k in qty_exclude_keywords + price_keywords + date_keywords + disc_keywords):
                s_qty_col = c_let
            # 5. Price column
            elif not s_price_col and any(k in hdr for k in price_keywords) and not any(k in hdr for k in price_exclude_keywords + qty_keywords + disc_keywords + date_keywords):
                s_price_col = c_let
            # 6. Total sales column
            elif not s_total_col and any(k in hdr for k in total_sales_keywords) and not any(k in hdr for k in cust_keywords + inv_keywords + disc_keywords):
                s_total_col = c_let
            # 7. Explicit profit column
            elif not s_explicit_profit_col and any(k in hdr for k in profit_keywords):
                s_explicit_profit_col = c_let
            # 8. Explicit cost column
            elif not s_explicit_cost_col and any(k in hdr for k in cost_keywords) and not any(k in hdr for k in cust_keywords):
                s_explicit_cost_col = c_let

        # Determine whether discount should be subtracted from sales
        apply_invoice_discount = False
        if s_disc_col:
            disc_hdr = all_s_hdrs.get(s_disc_col, "")
            has_explicit_total = any(any(k in h for k in ["مبلغ کل", "جمع کل", "مبلغ فاکتور", "مبلغ فروش"]) or (h == "مبلغ") for h in all_s_hdrs.values())
            has_explicit_net = any(any(k in h for k in ["بعد از تخفیف", "پس از تخفیف", "مبلغ خالص", "قیمت بعد"]) for h in all_s_hdrs.values())

            if has_explicit_total and has_explicit_net:
                # E.g. image-24.png with 'قیمت', 'مبلغ', 'تخفیف', 'قیمت بعد از تخفیف':
                # The user's sales total is 'مبلغ' (91,078,000) and unit price is 'قیمت' (23,000).
                apply_invoice_discount = False
            elif any(k in disc_hdr for k in ["درصد", "%", "percent"]):
                # E.g. test_roadmap with 'تخفیف درصد':
                apply_invoice_discount = True
            elif not has_explicit_total and not has_explicit_net:
                apply_invoice_discount = True

        if action.date_col:
            d_raw = str(action.date_col).strip().upper()
            d_cand = col_index_to_letter(int(d_raw)) if d_raw.isdigit() else d_raw
            # Verify if user-supplied date column is not accidentally quantity or price
            d_hdr = str(s_sheet.get_value(f"{d_cand}{s_min_r}") or "").strip().lower()
            if not any(k in d_hdr for k in qty_keywords + price_keywords):
                s_date_col = d_cand
        if action.profit_col:
            p_raw = str(action.profit_col).strip().upper()
            s_explicit_profit_col = col_index_to_letter(int(p_raw)) if p_raw.isdigit() else p_raw
        if action.purchase_col:
            pc_raw = str(action.purchase_col).strip().upper()
            s_explicit_cost_col = col_index_to_letter(int(pc_raw)) if pc_raw.isdigit() else pc_raw

        if action.sales_qty_col:
            sq_raw = str(action.sales_qty_col).strip().upper()
            s_qty_col = col_index_to_letter(int(sq_raw)) if sq_raw.isdigit() else sq_raw
        if action.sales_price_col:
            sp_raw = str(action.sales_price_col).strip().upper()
            s_price_col = col_index_to_letter(int(sp_raw)) if sp_raw.isdigit() else sp_raw

        if action.sales_col:
            sc_raw = str(action.sales_col).strip().upper()
            action.sales_col = col_index_to_letter(int(sc_raw)) if sc_raw.isdigit() else sc_raw
            if not s_price_col:
                s_price_col = action.sales_col

        if action.purchase_col:
            pc_raw = str(action.purchase_col).strip().upper()
            action.purchase_col = col_index_to_letter(int(pc_raw)) if pc_raw.isdigit() else pc_raw
            s_explicit_cost_col = action.purchase_col

        if s_prod_col in [s_date_col, s_price_col, s_explicit_cost_col, action.sales_col, action.purchase_col]:
            s_prod_col = None

        if not s_prod_col:
            for c in range(s_min_c, s_max_c + 1):
                c_let = col_index_to_letter(c)
                hdr = str(s_sheet.get_value(f"{c_let}{s_min_r}") or "").strip().lower()
                if c_let not in [s_date_col, s_qty_col, s_price_col, s_explicit_cost_col, s_explicit_profit_col, s_disc_col] and not any(k in hdr for k in cust_keywords + inv_keywords):
                    s_prod_col = c_let
                    break

        if not s_qty_col: s_qty_col = "B" if s_prod_col != "B" and s_date_col != "B" else "C"
        if not s_price_col: s_price_col = "C" if s_qty_col != "C" and s_prod_col != "C" and s_date_col != "C" else "D"

        # Auto-detect columns in Purchase Sheet
        p_qty_col = None
        p_price_col = None
        p_prod_col = None

        p_qty_keywords = ["تعداد خرید", "تعداد", "مقدار", "qty", "quantity"]
        p_qty_exclude = ["فروش", "مانده", "موجودی"]
        p_price_keywords = ["بهای واحد خرید", "قیمت واحد خرید", "قیمت خرید", "بهای واحد", "فی خرید", "قیمت", "فی", "بها", "cost", "purchase price"]
        p_price_exclude = ["فروش", "ارزش", "موجودی", "مانده", "ضرب"]

        for c in range(p_min_c, p_max_c + 1):
            c_let = col_index_to_letter(c)
            hdr = str(p_sheet.get_value(f"{c_let}{p_min_r}") or "").strip().lower()
            if not hdr:
                continue
            if not p_prod_col and any(k in hdr for k in prod_keywords) and not any(k in hdr for k in cust_keywords + inv_keywords + p_qty_keywords + p_price_keywords):
                p_prod_col = c_let
            elif not p_qty_col and any(k in hdr for k in p_qty_keywords) and not any(k in hdr for k in p_qty_exclude + p_price_keywords):
                p_qty_col = c_let
            elif not p_price_col and any(k in hdr for k in p_price_keywords) and not any(k in hdr for k in p_price_exclude + p_qty_keywords):
                p_price_col = c_let

        if action.purchase_qty_col:
            pq_raw = str(action.purchase_qty_col).strip().upper()
            p_qty_col = col_index_to_letter(int(pq_raw)) if pq_raw.isdigit() else pq_raw
        if action.purchase_price_col:
            pp_raw = str(action.purchase_price_col).strip().upper()
            p_price_col = col_index_to_letter(int(pp_raw)) if pp_raw.isdigit() else pp_raw

        if not p_prod_col: p_prod_col = "A"
        if not p_qty_col: p_qty_col = "B" if p_prod_col != "B" else "C"
        if not p_price_col: p_price_col = "C" if p_qty_col != "C" and p_prod_col != "C" else "D"

        # 3. Read Data and Match
        sales_records = []
        for r in range(s_min_r + 1, s_max_r + 1):
            prod = s_sheet.get_value(f"{s_prod_col}{r}")
            qty = clean_val(s_sheet.get_value(f"{s_qty_col}{r}")) if s_qty_col else 1.0
            price = clean_val(s_sheet.get_value(f"{s_price_col}{r}")) if s_price_col else 0.0

            disc_val = clean_val(s_sheet.get_value(f"{s_disc_col}{r}")) if (s_disc_col and apply_invoice_discount) else 0.0
            disc_pct = 0.0
            if apply_invoice_discount and s_disc_col and disc_val > 0.0:
                if disc_val <= 1.0:
                    disc_pct = disc_val
                elif disc_val <= 100.0:
                    disc_pct = disc_val / 100.0
                else:
                    gross = (qty * price) if (qty > 0 and price > 0) else 0.0
                    if gross > 0 and disc_val < gross:
                        disc_pct = disc_val / gross
                    else:
                        disc_pct = 0.0
            disc_pct = max(0.0, min(1.0, disc_pct))

            d_str = None
            if s_date_col:
                d_val = s_sheet.get_value(f"{s_date_col}{r}")
                if hasattr(d_val, 'strftime'):
                    d_str = d_val.strftime("%Y/%m/%d")
                elif d_val is not None:
                    sd = str(d_val).strip()
                    m = re.search(r'(\d{4})[/-](\d{1,2})[/-](\d{1,2})', sd)
                    if m:
                        d_str = f"{int(m.group(1)):04d}/{int(m.group(2)):02d}/{int(m.group(3)):02d}"
                    elif len(sd) >= 4 and ("/" in sd or "-" in sd):
                        d_str = sd

            direct_sales = clean_val(s_sheet.get_value(f"{action.sales_col}{r}")) if (action.sales_col and action.sales_col != s_price_col) else 0.0
            if direct_sales > 0:
                tot_sales = max(0.0, direct_sales * (1.0 - disc_pct))
            elif qty > 0 and price > 0:
                tot_sales = max(0.0, (qty * price) * (1.0 - disc_pct))
            elif price > 0:
                tot_sales = max(0.0, price * (1.0 - disc_pct))
            else:
                tot_sales = 0.0

            direct_cost = clean_val(s_sheet.get_value(f"{s_explicit_cost_col}{r}")) if s_explicit_cost_col else 0.0
            direct_profit = clean_val(s_sheet.get_value(f"{s_explicit_profit_col}{r}")) if s_explicit_profit_col else None

            if prod is not None or tot_sales > 0 or direct_cost > 0 or (direct_profit is not None and direct_profit != 0) or d_str:
                p_label = str(prod).strip() if (prod is not None and not str(prod).replace('.', '').isdigit()) else f"ردیف {r}"
                effective_unit_price = (tot_sales / qty) if (qty > 0 and tot_sales > 0 and disc_pct > 0) else price
                sales_records.append({
                    "row": r,
                    "product": p_label,
                    "qty": qty if qty > 0 else 1.0,
                    "price": effective_unit_price if effective_unit_price > 0 else (tot_sales / (qty if qty > 0 else 1.0)),
                    "total_sales": tot_sales,
                    "direct_cost": direct_cost,
                    "direct_profit": direct_profit,
                    "date": d_str
                })

        # Collect purchase batches per product in order of arrival (supports multi-batch FIFO)
        purchase_batches_by_prod = defaultdict(list)
        purchase_records = []
        if sales_sheet_name != purchase_sheet_name:
            for r in range(p_min_r + 1, p_max_r + 1):
                prod = p_sheet.get_value(f"{p_prod_col}{r}")
                qty = clean_val(p_sheet.get_value(f"{p_qty_col}{r}"))
                price = clean_val(p_sheet.get_value(f"{p_price_col}{r}"))
                if prod is not None and (qty > 0 or price > 0):
                    p_name = str(prod).strip()
                    purchase_batches_by_prod[p_name].append((qty, price))
                    purchase_records.append({
                        "row": r,
                        "product": p_name,
                        "qty": qty,
                        "price": price,
                        "total_purchase": qty * price
                    })

        has_named_products = any(sr["product"] and not sr["product"].startswith("ردیف ") for sr in sales_records) and \
                             any(pr["product"] and not pr["product"].startswith("ردیف ") for pr in purchase_records)

        purchase_by_name = {pr["product"]: pr for pr in purchase_records}
        matched_items = []
        consumed_stock_by_prod = defaultdict(float)

        if sales_sheet_name == purchase_sheet_name:
            # Single sheet profit calculation (e.g. Sales col - Cost col or explicit profit)
            for sr in sales_records:
                s_qty = sr["qty"] if sr["qty"] > 0 else 1.0
                s_price = sr["price"] if sr["price"] > 0 else (sr["total_sales"] / s_qty if s_qty > 0 else 0.0)
                s_total = s_qty * s_price if (s_qty > 0 and s_price > 0) else sr["total_sales"]

                if sr["direct_profit"] is not None:
                    profit = sr["direct_profit"]
                    cogs = s_total - profit if s_total > 0 else 0.0
                    p_unit_price = cogs / s_qty if s_qty > 0 else cogs
                elif sr["direct_cost"] > 0:
                    cogs = sr["direct_cost"] * s_qty if (sr["direct_cost"] < s_price and s_qty > 1) else sr["direct_cost"]
                    p_unit_price = cogs / s_qty if s_qty > 0 else cogs
                    profit = s_total - cogs
                else:
                    cogs = 0.0
                    p_unit_price = 0.0
                    profit = s_total

                margin = (profit / s_total * 100) if s_total > 0 else 0.0
                matched_items.append({
                    "product": sr["product"],
                    "s_qty": s_qty,
                    "s_price": s_price,
                    "s_total": s_total,
                    "p_qty": s_qty,
                    "p_price": p_unit_price,
                    "p_total": cogs,
                    "profit": profit,
                    "margin": margin,
                    "remaining_stock": 0.0,
                    "stock_value": 0.0,
                    "date": sr["date"]
                })
        elif has_named_products:
            for sr in sales_records:
                prod = sr["product"]
                s_qty = sr["qty"] if sr["qty"] > 0 else 1.0
                s_price = sr["price"] if sr["price"] > 0 else (sr["total_sales"] / s_qty if s_qty > 0 else 0.0)

                batches = purchase_batches_by_prod.get(prod, [])
                if batches:
                    prior_consumed = consumed_stock_by_prod[prod]
                    consumed_stock_by_prod[prod] += s_qty

                    cogs = 0.0
                    if val_method == "average":
                        tot_b_cost = sum(q * p for q, p in batches)
                        tot_b_qty = sum(q for q, _ in batches)
                        unit_cost = (tot_b_cost / tot_b_qty) if tot_b_qty > 0 else 0.0
                        cogs = s_qty * unit_cost
                    elif val_method == "lifo":
                        # Allocate COGS in LIFO order (latest purchases consumed first)
                        accum_b = 0.0
                        sale_start = prior_consumed
                        sale_end = prior_consumed + s_qty
                        for b_qty, b_price in reversed(batches):
                            b_start = accum_b
                            b_end = accum_b + b_qty
                            accum_b += b_qty
                            overlap = max(0.0, min(sale_end, b_end) - max(sale_start, b_start))
                            if overlap > 0:
                                cogs += overlap * b_price
                    else:
                        # Allocate COGS in FIFO order (earliest purchases consumed first)
                        accum_b = 0.0
                        sale_start = prior_consumed
                        sale_end = prior_consumed + s_qty
                        for b_qty, b_price in batches:
                            b_start = accum_b
                            b_end = accum_b + b_qty
                            accum_b += b_qty
                            overlap = max(0.0, min(sale_end, b_end) - max(sale_start, b_start))
                            if overlap > 0:
                                cogs += overlap * b_price

                    p_total_bought = sum(b[0] for b in batches)
                    p_unit_price = (cogs / s_qty) if s_qty > 0 else 0.0
                else:
                    pr = purchase_by_name.get(prod)
                    p_total_bought = pr["qty"] if pr else 0.0
                    p_unit_price = pr["price"] if pr else 0.0
                    if p_unit_price == 0.0 and pr and pr.get("total_purchase", 0) > 0 and p_total_bought > 0:
                        p_unit_price = pr["total_purchase"] / p_total_bought
                    cogs = s_qty * p_unit_price if p_unit_price > 0 else (pr["total_purchase"] if pr else 0.0)

                s_total = s_qty * s_price if (s_qty > 0 and s_price > 0) else sr["total_sales"]
                profit = s_total - cogs
                margin = (profit / s_total * 100) if s_total > 0 else 0.0
                remaining_stock = max(0.0, p_total_bought - consumed_stock_by_prod[prod]) if p_total_bought > 0 else 0.0

                matched_items.append({
                    "product": prod,
                    "s_qty": s_qty,
                    "s_price": s_price,
                    "s_total": s_total,
                    "p_qty": p_total_bought,
                    "p_price": p_unit_price,
                    "p_total": cogs,
                    "profit": profit,
                    "margin": margin,
                    "remaining_stock": remaining_stock,
                    "stock_value": 0.0,
                    "date": sr["date"]
                })
        else:
            max_len = max(len(sales_records), len(purchase_records))
            for i in range(max_len):
                sr = sales_records[i] if i < len(sales_records) else None
                pr = purchase_records[i] if i < len(purchase_records) else None

                prod = sr["product"] if sr else (pr["product"] if pr else f"آیتم {i+1}")
                s_qty = sr["qty"] if (sr and sr["qty"] > 0) else 1.0
                s_price = sr["price"] if sr else 0.0
                s_total = s_qty * s_price if (s_qty > 0 and s_price > 0) else (sr["total_sales"] if sr else 0.0)

                p_total_bought = pr["qty"] if pr else 0.0
                p_unit_price = pr["price"] if pr else 0.0
                if p_unit_price == 0.0 and pr and pr.get("total_purchase", 0) > 0 and p_total_bought > 0:
                    p_unit_price = pr["total_purchase"] / p_total_bought

                cogs = s_qty * p_unit_price if p_unit_price > 0 else (pr["total_purchase"] if pr else 0.0)
                profit = s_total - cogs
                margin = (profit / s_total * 100) if s_total > 0 else 0.0
                remaining_stock = max(0.0, p_total_bought - s_qty) if p_total_bought > 0 else 0.0
                stock_value = remaining_stock * p_unit_price

                matched_items.append({
                    "product": prod,
                    "s_qty": s_qty,
                    "s_price": s_price,
                    "s_total": s_total,
                    "p_qty": p_total_bought,
                    "p_price": p_unit_price,
                    "p_total": cogs,
                    "profit": profit,
                    "margin": margin,
                    "remaining_stock": remaining_stock,
                    "stock_value": stock_value,
                    "date": sr["date"] if sr else None
                })

        total_sales_all = sum(it["s_total"] for it in matched_items)
        total_purchase_all = sum(it["p_total"] for it in matched_items)
        total_profit_all = total_sales_all - total_purchase_all
        overall_margin = (total_profit_all / total_sales_all * 100) if total_sales_all > 0 else 0.0

        best_item = max(matched_items, key=lambda x: x["profit"]) if matched_items else {"product": "-", "profit": 0}
        loss_items = [it for it in matched_items if it["profit"] < 0]

        # Time-period aggregations (Daily, Monthly, Yearly)
        daily_map = defaultdict(lambda: {"sales": 0.0, "purchase": 0.0, "profit": 0.0, "count": 0})
        monthly_map = defaultdict(lambda: {"sales": 0.0, "purchase": 0.0, "profit": 0.0, "count": 0})
        yearly_map = defaultdict(lambda: {"sales": 0.0, "purchase": 0.0, "profit": 0.0, "count": 0})

        has_dates = any(it.get("date") for it in matched_items)
        if has_dates:
            for it in matched_items:
                dt = it.get("date")
                if not dt:
                    continue
                s_tot = it["s_total"]
                p_tot = it["p_total"]
                prof = it["profit"]

                daily_map[dt]["sales"] += s_tot
                daily_map[dt]["purchase"] += p_tot
                daily_map[dt]["profit"] += prof
                daily_map[dt]["count"] += 1

                mo_key = dt[:7] if len(dt) >= 7 and dt[4] in ('/', '-') else dt
                monthly_map[mo_key]["sales"] += s_tot
                monthly_map[mo_key]["purchase"] += p_tot
                monthly_map[mo_key]["profit"] += prof
                monthly_map[mo_key]["count"] += 1

                yr_key = dt[:4] if len(dt) >= 4 and dt[:4].isdigit() else "نامشخص"
                yearly_map[yr_key]["sales"] += s_tot
                yearly_map[yr_key]["purchase"] += p_tot
                yearly_map[yr_key]["profit"] += prof
                yearly_map[yr_key]["count"] += 1

        best_item = max(matched_items, key=lambda x: x["profit"]) if matched_items else {"product": "-", "profit": 0}
        loss_items = [it for it in matched_items if it["profit"] < 0]

        # 4. Generate 'گزارش_سود' Report Sheet
        report_sheet_name = "گزارش_سود"
        if action.create_report_sheet:
            existing_sheets = driver.get_sheet_names()
            if report_sheet_name in existing_sheets:
                try:
                    if hasattr(driver, 'delete_sheet'):
                        driver.delete_sheet(report_sheet_name)
                    else:
                        report_sheet = driver.get_sheet(report_sheet_name)
                        ub_r1, ub_c1, ub_r2, ub_c2 = report_sheet.get_used_range_bounds()
                        if ub_r2 >= ub_r1 and ub_r2 > 0:
                            report_sheet.delete_rows(list(range(ub_r1, ub_r2 + 1)))
                except Exception:
                    pass
            existing_sheets = driver.get_sheet_names()
            if report_sheet_name not in existing_sheets:
                report_sheet = driver.create_sheet(report_sheet_name)
            else:
                report_sheet = driver.get_sheet(report_sheet_name)

            curr_r = 1
            # Banner
            report_sheet.set_value(f"A{curr_r}", f"📊 داشبورد و گزارش جامع سود و زیان (روش {method_persian})")
            report_sheet.set_font(f"A{curr_r}", bold=True, size=14, color_hex="#38BDF8")
            report_sheet.set_cell_color(f"A{curr_r}", "#09132B")
            curr_r += 2

            # Section 1: KPI Cards
            report_sheet.set_value(f"A{curr_r}", "📌 خلاصه کل وضعیت سود و زیان")
            report_sheet.set_font(f"A{curr_r}", bold=True, size=12)
            curr_r += 1

            headers_kpi = ["شاخص مالی", "مبلغ / مقدار", "شیت مرجع / فرمول"]
            for idx, h in enumerate(headers_kpi):
                col_c = col_index_to_letter(idx + 1)
                report_sheet.set_value(f"{col_c}{curr_r}", h)
                report_sheet.set_font(f"{col_c}{curr_r}", bold=True, color_hex="#FFFFFF")
                report_sheet.set_cell_color(f"{col_c}{curr_r}", "#1E3A8A")
            curr_r += 1

            kpis = [
                ("💰 کل مبلغ فروش", f"{total_sales_all:,.0f} تومان", f"شیت: {sales_sheet_name}"),
                ("🛒 کل مبلغ خرید اقلام فروش‌رفته (COGS)", f"{total_purchase_all:,.0f} تومان", f"محاسبه بر مبنای روش {method_persian}"),
                ("💵 سود خالص کل", f"{total_profit_all:,.0f} تومان", "فروش کل منهای بهای خرید اقلام فروش‌رفته"),
                ("📈 میانگین حاشیه سود", f"{overall_margin:.1f}%", "درصد سود از کل فروش"),
            ]
            # Calculate true unique remaining stock and inventory valuation across products
            product_remaining = {}
            for it in matched_items:
                product_remaining[it["product"]] = it.get("remaining_stock", 0.0)

            tot_rem_stock = sum(product_remaining.values())

            product_stock_val = {}
            if purchase_batches_by_prod:
                for p, rem_q in product_remaining.items():
                    batches = purchase_batches_by_prod.get(p, [])
                    val = 0.0
                    if val_method == "average":
                        tot_b_cost = sum(q * pr for q, pr in batches)
                        tot_b_qty = sum(q for q, _ in batches)
                        unit_cost = (tot_b_cost / tot_b_qty) if tot_b_qty > 0 else 0.0
                        val = rem_q * unit_cost
                    elif val_method == "lifo":
                        # LIFO: remaining inventory comes from earliest batches
                        to_val = rem_q
                        for b_q, b_p in batches:
                            take = min(to_val, b_q)
                            val += take * b_p
                            to_val -= take
                            if to_val <= 0:
                                break
                    else:
                        # FIFO: remaining inventory comes from latest batches
                        to_val = rem_q
                        for b_q, b_p in reversed(batches):
                            take = min(to_val, b_q)
                            val += take * b_p
                            to_val -= take
                            if to_val <= 0:
                                break
                    product_stock_val[p] = val
            tot_stock_val = sum(product_stock_val.values())

            if tot_rem_stock > 0:
                kpis.append(("📦 ارزش موجودی باقیمانده انبار", f"{tot_stock_val:,.0f} تومان", f"{tot_rem_stock:,.0f} عدد کالا در انبار"))

            if has_dates and daily_map:
                best_day = max(daily_map.items(), key=lambda x: x[1]["profit"])
                kpis.append(("📅 پرسودترین روز", f"{best_day[0]} ({best_day[1]['profit']:,.0f} تومان)", "بیشترین سود روزانه"))
            if has_dates and monthly_map:
                best_month = max(monthly_map.items(), key=lambda x: x[1]["profit"])
                kpis.append(("🗓️ پرسودترین ماه", f"{best_month[0]} ({best_month[1]['profit']:,.0f} تومان)", "بیشترین سود ماهانه"))

            kpis.append(("🏆 پر سودترین کالا", f"{best_item['product']} ({best_item['profit']:,.0f} تومان)", "بیشترین سود ریالی"))
            kpis.append(("⚠️ تعداد اقلام زیان‌ده", f"{len(loss_items)} کالا", "اقلام با سود منفی"))

            for title, val, desc in kpis:
                report_sheet.set_value(f"A{curr_r}", title)
                report_sheet.set_value(f"B{curr_r}", val)
                report_sheet.set_value(f"C{curr_r}", desc)
                report_sheet.set_font(f"B{curr_r}", bold=True)
                if "سود خالص" in title:
                    report_sheet.set_cell_color(f"B{curr_r}", "#064E3B" if total_profit_all >= 0 else "#7F1D1D")
                    report_sheet.set_font(f"B{curr_r}", bold=True, color_hex="#FFFFFF")
                curr_r += 1
            curr_r += 2

            # Section 2: Daily Profit Breakdown (if dates exist and requested)
            if has_dates and action.period in ["all", "daily"]:
                report_sheet.set_value(f"A{curr_r}", "📅 جدول تفکیک سود روزانه")
                report_sheet.set_font(f"A{curr_r}", bold=True, size=12)
                curr_r += 1

                d_headers = ["ردیف", "تاریخ", "تعداد تراکنش", "کل فروش روز (تومان)", "کل خرید/هزینه (تومان)", "سود خالص روز (تومان)", "حاشیه سود روز (%)"]
                for idx, h in enumerate(d_headers):
                    col_c = col_index_to_letter(idx + 1)
                    report_sheet.set_value(f"{col_c}{curr_r}", h)
                    report_sheet.set_font(f"{col_c}{curr_r}", bold=True, color_hex="#FFFFFF")
                    report_sheet.set_cell_color(f"{col_c}{curr_r}", "#1E3A8A")
                curr_r += 1

                start_d_r = curr_r
                for i, (dt, d_data) in enumerate(sorted(daily_map.items())):
                    r_num = curr_r
                    report_sheet.set_value(f"A{r_num}", i + 1)
                    report_sheet.set_value(f"B{r_num}", dt)
                    report_sheet.set_value(f"C{r_num}", d_data["count"])
                    report_sheet.set_value(f"D{r_num}", d_data["sales"])
                    report_sheet.set_value(f"E{r_num}", d_data["purchase"])
                    report_sheet.set_formula(f"F{r_num}", f"=D{r_num}-E{r_num}")
                    report_sheet.set_formula(f"G{r_num}", f'=IF(D{r_num}>0, (F{r_num}/D{r_num})*100, 0)')

                    report_sheet.set_font(f"F{r_num}", bold=True, color_hex="#FFFFFF")
                    if d_data["profit"] >= 0:
                        report_sheet.set_cell_color(f"F{r_num}", "#064E3B")
                    else:
                        report_sheet.set_cell_color(f"F{r_num}", "#7F1D1D")
                    curr_r += 1

                last_d_r = curr_r
                report_sheet.set_value(f"A{last_d_r}", "جمع کل روزانه")
                report_sheet.set_font(f"A{last_d_r}", bold=True, color_hex="#FFFFFF")
                report_sheet.set_formula(f"C{last_d_r}", f"=SUM(C{start_d_r}:C{last_d_r-1})")
                report_sheet.set_formula(f"D{last_d_r}", f"=SUM(D{start_d_r}:D{last_d_r-1})")
                report_sheet.set_formula(f"E{last_d_r}", f"=SUM(E{start_d_r}:E{last_d_r-1})")
                report_sheet.set_formula(f"F{last_d_r}", f"=D{last_d_r}-E{last_d_r}")
                report_sheet.set_formula(f"G{last_d_r}", f'=IF(D{last_d_r}>0, (F{last_d_r}/D{last_d_r})*100, 0)')
                for c_idx in range(1, 8):
                    col_c = col_index_to_letter(c_idx)
                    report_sheet.set_font(f"{col_c}{last_d_r}", bold=True, color_hex="#FFFFFF")
                    report_sheet.set_cell_color(f"{col_c}{last_d_r}", "#0B1938")

                try:
                    report_sheet.set_number_format(f"D{start_d_r}:F{last_d_r}", "#,##0")
                    report_sheet.set_number_format(f"G{start_d_r}:G{last_d_r}", "0.0")
                except Exception:
                    pass
                curr_r += 2

            # Section 3: Monthly Profit Breakdown (if dates exist and requested)
            if has_dates and action.period in ["all", "monthly"]:
                report_sheet.set_value(f"A{curr_r}", "🗓️ جدول تفکیک سود ماهانه")
                report_sheet.set_font(f"A{curr_r}", bold=True, size=12)
                curr_r += 1

                m_headers = ["ردیف", "ماه", "تعداد تراکنش", "کل فروش ماه (تومان)", "کل خرید/هزینه (تومان)", "سود خالص ماه (تومان)", "حاشیه سود ماه (%)"]
                for idx, h in enumerate(m_headers):
                    col_c = col_index_to_letter(idx + 1)
                    report_sheet.set_value(f"{col_c}{curr_r}", h)
                    report_sheet.set_font(f"{col_c}{curr_r}", bold=True, color_hex="#FFFFFF")
                    report_sheet.set_cell_color(f"{col_c}{curr_r}", "#7C2D12")
                curr_r += 1

                start_m_r = curr_r
                for i, (mo, m_data) in enumerate(sorted(monthly_map.items())):
                    r_num = curr_r
                    report_sheet.set_value(f"A{r_num}", i + 1)
                    report_sheet.set_value(f"B{r_num}", mo)
                    report_sheet.set_value(f"C{r_num}", m_data["count"])
                    report_sheet.set_value(f"D{r_num}", m_data["sales"])
                    report_sheet.set_value(f"E{r_num}", m_data["purchase"])
                    report_sheet.set_formula(f"F{r_num}", f"=D{r_num}-E{r_num}")
                    report_sheet.set_formula(f"G{r_num}", f'=IF(D{r_num}>0, (F{r_num}/D{r_num})*100, 0)')

                    report_sheet.set_font(f"F{r_num}", bold=True, color_hex="#FFFFFF")
                    if m_data["profit"] >= 0:
                        report_sheet.set_cell_color(f"F{r_num}", "#064E3B")
                    else:
                        report_sheet.set_cell_color(f"F{r_num}", "#7F1D1D")
                    curr_r += 1

                last_m_r = curr_r
                report_sheet.set_value(f"A{last_m_r}", "جمع کل ماهانه")
                report_sheet.set_font(f"A{last_m_r}", bold=True, color_hex="#FFFFFF")
                report_sheet.set_formula(f"C{last_m_r}", f"=SUM(C{start_m_r}:C{last_m_r-1})")
                report_sheet.set_formula(f"D{last_m_r}", f"=SUM(D{start_m_r}:D{last_m_r-1})")
                report_sheet.set_formula(f"E{last_m_r}", f"=SUM(E{start_m_r}:E{last_m_r-1})")
                report_sheet.set_formula(f"F{last_m_r}", f"=D{last_m_r}-E{last_m_r}")
                report_sheet.set_formula(f"G{last_m_r}", f'=IF(D{last_m_r}>0, (F{last_m_r}/D{last_m_r})*100, 0)')
                for c_idx in range(1, 8):
                    col_c = col_index_to_letter(c_idx)
                    report_sheet.set_font(f"{col_c}{last_m_r}", bold=True, color_hex="#FFFFFF")
                    report_sheet.set_cell_color(f"{col_c}{last_m_r}", "#0B1938")

                try:
                    report_sheet.set_number_format(f"D{start_m_r}:F{last_m_r}", "#,##0")
                    report_sheet.set_number_format(f"G{start_m_r}:G{last_m_r}", "0.0")
                except Exception:
                    pass
                curr_r += 2

            # Section 4: Detailed Profit Table by Item
            report_sheet.set_value(f"A{curr_r}", "📋 جدول تفکیک سود و زیان و موجودی به ازای هر کالا")
            report_sheet.set_font(f"A{curr_r}", bold=True, size=12)
            curr_r += 1

            table_headers = [
                "ردیف",
                "نام کالا",
                f"تعداد فروش ({sales_sheet_name})",
                "قیمت فروش هر واحد",
                "مبلغ کل فروش",
                f"قیمت خرید هر واحد ({purchase_sheet_name})",
                "بهای تمام‌شده اقلام فروش‌رفته (COGS)",
                "سود خالص (تومان)",
                "درصد حاشیه سود",
                f"کل خرید اولیه ({purchase_sheet_name})",
                "موجودی باقیمانده انبار",
                "ارزش موجودی انبار"
            ]
            for idx, h in enumerate(table_headers):
                col_c = col_index_to_letter(idx + 1)
                report_sheet.set_value(f"{col_c}{curr_r}", h)
                report_sheet.set_font(f"{col_c}{curr_r}", bold=True, color_hex="#FFFFFF")
                report_sheet.set_cell_color(f"{col_c}{curr_r}", "#047857")
            curr_r += 1

            start_table_r = curr_r
            for i, it in enumerate(matched_items):
                r_num = curr_r
                report_sheet.set_value(f"A{r_num}", i + 1)
                report_sheet.set_value(f"B{r_num}", it["product"])
                report_sheet.set_value(f"C{r_num}", it["s_qty"])
                report_sheet.set_value(f"D{r_num}", it["s_price"])
                report_sheet.set_formula(f"E{r_num}", f"=C{r_num}*D{r_num}")
                report_sheet.set_value(f"F{r_num}", it["p_price"])
                report_sheet.set_formula(f"G{r_num}", f"=C{r_num}*F{r_num}")
                report_sheet.set_formula(f"H{r_num}", f"=E{r_num}-G{r_num}")
                report_sheet.set_formula(f"I{r_num}", f'=IF(E{r_num}>0, (H{r_num}/E{r_num})*100, 0)')
                report_sheet.set_value(f"J{r_num}", it["p_qty"])
                report_sheet.set_formula(f"K{r_num}", f'=IF(J{r_num}>C{r_num}, J{r_num}-C{r_num}, 0)')
                report_sheet.set_formula(f"L{r_num}", f'=K{r_num}*F{r_num}')

                report_sheet.set_font(f"H{r_num}", bold=True, color_hex="#FFFFFF")
                if it["profit"] >= 0:
                    report_sheet.set_cell_color(f"H{r_num}", "#064E3B")
                else:
                    report_sheet.set_cell_color(f"H{r_num}", "#7F1D1D")

                curr_r += 1

            # Total Summary Row at bottom
            last_r = curr_r
            report_sheet.set_value(f"A{last_r}", "جمع کل")
            report_sheet.set_font(f"A{last_r}", bold=True, color_hex="#FFFFFF")
            report_sheet.set_formula(f"C{last_r}", f"=SUM(C{start_table_r}:C{last_r-1})")
            report_sheet.set_formula(f"E{last_r}", f"=SUM(E{start_table_r}:E{last_r-1})")
            report_sheet.set_formula(f"G{last_r}", f"=SUM(G{start_table_r}:G{last_r-1})")
            report_sheet.set_formula(f"H{last_r}", f"=E{last_r}-G{last_r}")
            report_sheet.set_formula(f"I{last_r}", f'=IF(E{last_r}>0, (H{last_r}/E{last_r})*100, 0)')
            report_sheet.set_formula(f"J{last_r}", f"=SUM(J{start_table_r}:J{last_r-1})")
            report_sheet.set_formula(f"K{last_r}", f"=SUM(K{start_table_r}:K{last_r-1})")
            report_sheet.set_formula(f"L{last_r}", f"=SUM(L{start_table_r}:L{last_r-1})")

            for c_idx in range(1, 13):
                col_c = col_index_to_letter(c_idx)
                report_sheet.set_font(f"{col_c}{last_r}", bold=True, color_hex="#FFFFFF")
                report_sheet.set_cell_color(f"{col_c}{last_r}", "#0B1938")

            try:
                report_sheet.set_number_format(f"D{start_table_r}:H{last_r}", "#,##0")
                report_sheet.set_number_format(f"I{start_table_r}:I{last_r}", "0.0")
                report_sheet.set_number_format(f"L{start_table_r}:L{last_r}", "#,##0")
            except Exception:
                pass

            report_sheet.auto_fit_columns()
            ctx.manager.activate_sheet(report_sheet_name)

        # 5. Format Persian message
        profit_status = "🟢 سودآور" if total_profit_all >= 0 else "🔴 زیان‌ده"
        period_title = {
            "all": "روزانه، ماهانه، سالانه و کل",
            "daily": "روزانه",
            "monthly": "ماهانه",
            "yearly": "سالانه",
            "total": "کل",
            "by_item": "تفکیک کالاها"
        }.get(action.period, "روزانه، ماهانه، سالانه و کل")

        tot_sold_qty = sum(it["s_qty"] for it in matched_items)
        msg_lines = [
            f"💵 **محاسبه سود به روش {method_persian} ({period_title}) با موفقیت انجام شد ({profit_status}):**",
            f"💰 **کل مبلغ فروش ({sales_sheet_name}):** `{total_sales_all:,.0f}` تومان (تعداد فروش: `{tot_sold_qty:,.0f}` عدد در `{len(matched_items)}` قلم)",
            f"🛒 **بهای تمام‌شده اقلام فروش‌رفته (COGS):** `{total_purchase_all:,.0f}` تومان (محاسبه‌شده بر مبنای روش {method_persian})",
            "───────────────────────────────",
            f"✨ **سود خالص کل (فروش منهای بهای خرید اقلام فروخته‌شده):** **`{total_profit_all:,.0f}` تومان**",
            f"📈 **میانگین حاشیه سود:** **`{overall_margin:.1f}%`**"
        ]
        if tot_rem_stock > 0:
            msg_lines.append(f"📦 **موجودی باقیمانده در انبار:** `{tot_rem_stock:,.0f}` قلم به ارزش `{tot_stock_val:,.0f}` تومان (ارزیابی شده به روش {method_persian})")
        if has_dates and daily_map:
            best_day = max(daily_map.items(), key=lambda x: x[1]["profit"])
            msg_lines.append(f"📅 **تعداد روزهای ثبت‌شده:** `{len(daily_map)}` روز (پرسودترین روز: `{best_day[0]}` با سود `{best_day[1]['profit']:,.0f}` تومان)")
        if has_dates and monthly_map:
            best_month = max(monthly_map.items(), key=lambda x: x[1]["profit"])
            msg_lines.append(f"🗓️ **تعداد ماه‌های فعال:** `{len(monthly_map)}` ماه (پرسودترین ماه: `{best_month[0]}` با سود `{best_month[1]['profit']:,.0f}` تومان)")
        if best_item and best_item["product"] != "-":
            msg_lines.append(f"🏆 **پرسودترین کالا:** `{best_item['product']}` با سود `{best_item['profit']:,.0f}` تومان")
        if loss_items:
            msg_lines.append(f"⚠️ **اقلام با سود منفی:** `{len(loss_items)}` قلم کالا نیازمند بازنگری در قیمت‌گذاری هستند.")

        if action.create_report_sheet:
            msg_lines.append("")
            msg_lines.append(f"📑 شیت اختصاصی **«{report_sheet_name}»** همراه با داشبورد، موجودی انبار، جداول تفکیک دوره‌ای (روزانه/ماهانه/سالانه) و فرمول‌های اکسل ایجاد گردید.")

        return ActionResult(
            success=True,
            message_persian="\n".join(msg_lines),
            affected_range=report_sheet_name,
            data={
                "total_sales": total_sales_all,
                "total_purchase": total_purchase_all,
                "total_profit": total_profit_all,
                "margin_pct": overall_margin,
                "remaining_stock": tot_rem_stock,
                "stock_value": tot_stock_val,
                "valuation_method": val_method,
                "item_count": len(matched_items),
                "daily_count": len(daily_map) if has_dates else 0,
                "monthly_count": len(monthly_map) if has_dates else 0
            }
        )

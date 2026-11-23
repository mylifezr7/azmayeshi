"""
Workbook and Worksheet Structure Inspector.
Performs comprehensive deep analysis of Excel files, sheets, headers, data distributions, and anomalies.
"""
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

from excel_assistant.core.interfaces import IExcelDriver, ISheet, WorkbookInfo
from excel_assistant.core.utils import col_index_to_letter, bounds_to_range_str
from excel_assistant.analyzer.type_detector import ColumnSummary, analyze_column
from excel_assistant.analyzer.formula_checker import FormulaReport, check_sheet_formulas_and_errors


@dataclass
class SheetAnalysis:
    sheet_name: str
    used_range_str: str
    min_row: int
    min_col: int
    max_row: int
    max_col: int
    row_count: int
    column_count: int
    header_row_index: int
    headers: List[Tuple[str, str, str]]  # (col_letter, header_name, detected_type)
    columns: Dict[str, ColumnSummary]  # col_letter -> ColumnSummary
    header_to_col: Dict[str, str]      # header_name.lower() -> col_letter
    empty_row_indices: List[int]
    duplicate_row_indices: List[int]
    formula_report: FormulaReport
    sample_matrix: List[List[Any]] = field(default_factory=list)


@dataclass
class WorkbookAnalysis:
    workbook_name: str
    file_path: Optional[str]
    sheet_count: int
    sheet_names: List[str]
    active_sheet_name: str
    sheets: Dict[str, SheetAnalysis]
    total_formulas: int = 0
    total_errors: int = 0
    has_empty_rows: bool = False
    has_duplicates: bool = False


class SheetInspector:
    """Analyzes structure and content of worksheets."""

    @classmethod
    def inspect_sheet(cls, sheet: ISheet) -> SheetAnalysis:
        min_r, min_c, max_r, max_c = sheet.get_used_range_bounds()
        used_range_str = bounds_to_range_str(min_r, min_c, max_r, max_c)
        row_count = max_r - min_r + 1
        col_count = max_c - min_c + 1

        # Fetch entire sheet matrix in a single batch call (massive speedup)
        matrix = sheet.get_matrix(min_r, min_c, max_r, max_c)

        # Helper to get cell value from in-memory matrix (1-based coordinates)
        def get_mat_val(row_1b, col_1b):
            r_idx = row_1b - min_r
            c_idx = col_1b - min_c
            if 0 <= r_idx < len(matrix) and 0 <= c_idx < len(matrix[r_idx]):
                return matrix[r_idx][c_idx]
            return None

        # Detect header row: first row with non-empty content
        header_row_idx = min_r
        headers_list: List[Tuple[str, str, str]] = []
        columns_dict: Dict[str, ColumnSummary] = {}
        header_to_col_dict: Dict[str, str] = {}

        # Sample matrix for preview (up to 15 rows)
        sample_rows = matrix[:15] if matrix else []

        # Detect headers and analyze columns in memory
        for c_idx in range(min_c, max_c + 1):
            letter = col_index_to_letter(c_idx)
            header_cell_val = get_mat_val(header_row_idx, c_idx)
            header_title = str(header_cell_val).strip() if header_cell_val is not None else f"Column {letter}"

            # Fetch column values below header from in-memory matrix
            col_vals = []
            for r_idx in range(header_row_idx + 1, max_r + 1):
                col_vals.append(get_mat_val(r_idx, c_idx))

            col_summary = analyze_column(
                column_letter=letter,
                column_index=c_idx,
                header_name=header_title,
                values=col_vals
            )
            columns_dict[letter] = col_summary
            headers_list.append((letter, header_title, col_summary.detected_type))
            header_to_col_dict[header_title.lower()] = letter

        # Check empty rows and duplicate rows in memory
        empty_rows = []
        row_hashes: Dict[str, int] = {}
        duplicate_rows = []

        for r_idx in range(header_row_idx + 1, max_r + 1):
            row_vals = [get_mat_val(r_idx, c) for c in range(min_c, max_c + 1)]
            if all(v is None or str(v).strip() == "" for v in row_vals):
                empty_rows.append(r_idx)
            else:
                row_str = "|".join(str(v).strip() for v in row_vals if v is not None)
                if row_str in row_hashes:
                    duplicate_rows.append(r_idx)
                else:
                    row_hashes[row_str] = r_idx

        # Check formulas and errors
        formula_rep = check_sheet_formulas_and_errors(sheet)

        return SheetAnalysis(
            sheet_name=sheet.name,
            used_range_str=used_range_str,
            min_row=min_r,
            min_col=min_c,
            max_row=max_r,
            max_col=max_c,
            row_count=row_count,
            column_count=col_count,
            header_row_index=header_row_idx,
            headers=headers_list,
            columns=columns_dict,
            header_to_col=header_to_col_dict,
            empty_row_indices=empty_rows,
            duplicate_row_indices=duplicate_rows,
            formula_report=formula_rep,
            sample_matrix=sample_rows
        )


class WorkbookInspector:
    """High-level inspector for the entire active workbook."""

    @classmethod
    def inspect(cls, driver: IExcelDriver) -> WorkbookAnalysis:
        info = driver.get_workbook_info()
        sheet_analyses: Dict[str, SheetAnalysis] = {}

        tot_formulas = 0
        tot_errors = 0
        has_empty = False
        has_dup = False

        for name in info.sheets:
            try:
                sheet = driver.get_sheet(name)
                sa = SheetInspector.inspect_sheet(sheet)
                sheet_analyses[name] = sa
                tot_formulas += sa.formula_report.total_formulas
                tot_errors += len(sa.formula_report.error_cells)
                if sa.empty_row_indices:
                    has_empty = True
                if sa.duplicate_row_indices:
                    has_dup = True
            except Exception:
                pass

        return WorkbookAnalysis(
            workbook_name=info.name,
            file_path=info.path,
            sheet_count=len(info.sheets),
            sheet_names=info.sheets,
            active_sheet_name=info.active_sheet,
            sheets=sheet_analyses,
            total_formulas=tot_formulas,
            total_errors=tot_errors,
            has_empty_rows=has_empty,
            has_duplicates=has_dup
        )

    @classmethod
    def format_persian_summary(cls, analysis: WorkbookAnalysis) -> str:
        """Produces an elegant, rich Persian report of the workbook structure."""
        active = analysis.sheets.get(analysis.active_sheet_name)
        lines = [
            f"📊 **گزارش تحلیل هوشمند فایل اکسل:** `{analysis.workbook_name}`",
            f"📑 تعداد کل شیت‌ها: **{analysis.sheet_count}** ({', '.join(analysis.sheet_names)})",
            f"📌 شیت فعال فعلی: **{analysis.active_sheet_name}**"
        ]

        if active:
            lines.append(f"\n🔍 **مشخصات شیت فعال ({active.sheet_name}):**")
            lines.append(f"- محدوده داده‌ها: `{active.used_range_str}` ({active.row_count} ردیف × {active.column_count} ستون)")

            type_fa = {
                "numeric": "عددی", "integer": "عدد صحیح", "float": "اعشاری",
                "text": "متنی", "date": "تاریخ", "currency": "مالی/ارز",
                "percentage": "درصدی", "boolean": "بولی", "empty": "خالی", "mixed": "ترکیبی"
            }

            lines.append("\n📋 **ستون‌های شناسایی‌شده:**")
            for letter, title, col_type in active.headers:
                summary = active.columns.get(letter)
                dup_info = f" | {summary.duplicate_count} تکراری" if summary and summary.duplicate_count > 0 else ""
                lines.append(f"  • ستون `{letter}`: **{title}** (نوع: {type_fa.get(col_type, col_type)}{dup_info})")

            # Quality metrics
            lines.append("\n⚙️ **کیفیت و وضعیت داده‌ها:**")
            if active.formula_report.total_formulas > 0:
                lines.append(f"  • فرمول‌های موجود: {active.formula_report.total_formulas} سلول")
            else:
                lines.append("  • فرمول: در حال حاضر هیچ فرمولی در این شیت وجود ندارد.")

            if active.empty_row_indices:
                lines.append(f"  • ⚠️ ردیف‌های خالی بین داده‌ها: {len(active.empty_row_indices)} ردیف (ردیف‌های: {', '.join(map(str, active.empty_row_indices[:5]))})")

            if active.duplicate_row_indices:
                lines.append(f"  • ⚠️ ردیف‌های تکراری: {len(active.duplicate_row_indices)} ردیف شناسایی شد.")

            if active.formula_report.error_cells:
                lines.append(f"  • ❌ خطاهای اکسل شناسایی‌شده: {len(active.formula_report.error_cells)} سلول ({', '.join(active.formula_report.error_cells)})")
            else:
                lines.append("  • ✅ هیچ خطای فرمولی یا سیستمی در شیت یافت نشد.")

        return "\n".join(lines)

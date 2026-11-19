"""
LLM Client for Persian Natural Language Understanding in Excel.
Communicates with OpenAI-compatible APIs to interpret complex natural language instructions.
"""
import json
import os
from typing import Any, Dict, List, Optional
import requests

from excel_assistant.nlp.intent_models import ExcelAction


SYSTEM_PROMPT_TEMPLATE = """شما یک دستیار هوشمند و ارشد متخصص در نرم‌افزار Microsoft Excel هستید.
وظیفه شما تبدیل دستورات کاربر به زبان فارسی به عملیات ساختاریافته (JSON Action) بر روی اکسل است.

اطلاعات شیت فعلی اکسل:
- نام شیت: {sheet_name}
- محدوده مورد استفاده: {used_range}
- ستون‌ها: {columns_info}

پاسخ شما باید دقیقاً یک شیء JSON با فیلد `action_type` باشد.
انواع مجاز:
1. "column_arithmetic":
   {{"action_type": "column_arithmetic", "operation": "multiply|add|subtract|divide", "col1": "A", "col2": "B", "target_col": "C", "as_formula": true}}
2. "column_stats":
   {{"action_type": "column_stats", "stat_type": "sum|average|median|max|min|count|percentage", "source_col": "D"}}
3. "highlight_condition":
   {{"action_type": "highlight_condition", "target_col": "D", "condition": "gt|lt|eq", "value": 100, "color_hex": "#90EE90"}}
4. "highlight_duplicates":
   {{"action_type": "highlight_duplicates", "target_col": null, "color_hex": "#FFFF00"}}
5. "delete_empty_rows":
   {{"action_type": "delete_empty_rows"}}
6. "delete_duplicates":
   {{"action_type": "delete_duplicates", "target_col": null}}
7. "sort_data":
   {{"action_type": "sort_data", "key_col": "A", "ascending": true|false}}
8. "autofit_columns":
   {{"action_type": "autofit_columns"}}
9. "create_table":
   {{"action_type": "create_table"}}
10. "apply_formula":
   {{"action_type": "apply_formula", "target_cell_or_range": "E2:E20", "formula_template": "=C2*D2"}}
11. "analyze_sheet":
   {{"action_type": "analyze_sheet"}}
12. "undo_last_action":
   {{"action_type": "undo_last_action"}}

فقط شیء JSON را برگردانید و هیچ متن اضافی دیگری ننویسید.
"""


class LLMClient:
    """Client for calling LLM completions with fallback and error handling."""

    def __init__(self, api_key: Optional[str] = None, base_url: Optional[str] = None, model: str = "gpt-4o-mini"):
        self.api_key = api_key or os.getenv("OPENAI_API_KEY", "")
        self.base_url = (base_url or os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1")).rstrip("/")
        self.model = model

    @property
    def is_configured(self) -> bool:
        return bool(self.api_key)

    def parse_instruction(self, user_command: str, sheet_context: Optional[Dict[str, Any]] = None) -> Optional[Dict[str, Any]]:
        """Sends user command to LLM and parses JSON action response."""
        if not self.is_configured:
            return None

        ctx = sheet_context or {}
        sheet_name = ctx.get("sheet_name", "Sheet1")
        used_range = ctx.get("used_range", "A1:Z100")
        headers = ctx.get("headers", [])
        cols_desc = ", ".join([f"{col} ({title} - {ctype})" for col, title, ctype in headers]) if headers else "نامشخص"

        system_msg = SYSTEM_PROMPT_TEMPLATE.format(
            sheet_name=sheet_name,
            used_range=used_range,
            columns_info=cols_desc
        )

        headers_http = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }

        payload = {
            "model": self.model,
            "messages": [
                {"role": "system", "content": system_msg},
                {"role": "user", "content": user_command}
            ],
            "temperature": 0.1,
            "response_format": {"type": "json_object"}
        }

        try:
            resp = requests.post(f"{self.base_url}/chat/completions", headers=headers_http, json=payload, timeout=15)
            if resp.status_code == 200:
                data = resp.json()
                content = data["choices"][0]["message"]["content"]
                return json.loads(content)
        except Exception:
            pass

        return None

"""
Lightweight, Zero-Dependency Persian Web UI Server for Excel AI Assistant.
Provides a modern live web preview with Excel grid view and Persian AI chat.
"""
import http.server
import json
import os
import socketserver
import urllib.parse
from typing import Optional

from excel_assistant.core.connection import ExcelConnectionManager
from excel_assistant.analyzer.inspector import SheetInspector
from excel_assistant.ui.controller import AssistantUIController
from excel_assistant.core.utils import col_index_to_letter


HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="fa" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>دستیار هوشمند اکسل - Excel AI Assistant</title>
    <link href="https://fonts.googleapis.com/css2?family=Vazirmatn:wght@300;400;600;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --excel-green: #107C41;
            --excel-green-hover: #148A49;
            --bg-dark: #121418;
            --card-dark: #1E222B;
            --border-color: #2F3644;
            --text-light: #F1F5F9;
            --text-muted: #94A3B8;
            --user-bubble: #1E293B;
            --ai-bubble: #132E20;
        }

        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
            font-family: 'Vazirmatn', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
        }

        body {
            background-color: var(--bg-dark);
            color: var(--text-light);
            height: 100vh;
            display: flex;
            flex-direction: column;
            overflow: hidden;
        }

        /* Top Header */
        header {
            background-color: var(--card-dark);
            border-bottom: 2px solid var(--excel-green);
            padding: 10px 24px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            box-shadow: 0 4px 12px rgba(0,0,0,0.3);
        }

        .header-title {
            display: flex;
            align-items: center;
            gap: 12px;
            font-size: 17px;
            font-weight: 700;
        }

        .excel-badge {
            background-color: var(--excel-green);
            color: white;
            padding: 4px 10px;
            border-radius: 6px;
            font-size: 13px;
        }

        .status-bar {
            display: flex;
            align-items: center;
            gap: 16px;
            font-size: 13px;
        }

        .status-pill {
            display: flex;
            align-items: center;
            gap: 6px;
            background-color: rgba(16, 124, 65, 0.15);
            border: 1px solid var(--excel-green);
            padding: 4px 12px;
            border-radius: 20px;
            color: #4ADE80;
        }

        select.sheet-select {
            background-color: var(--bg-dark);
            color: var(--text-light);
            border: 1px solid var(--border-color);
            padding: 6px 12px;
            border-radius: 6px;
            cursor: pointer;
        }

        /* Main Container */
        .main-container {
            flex: 1;
            display: flex;
            overflow: hidden;
        }

        /* Left / Center: Chat Area */
        .chat-section {
            flex: 3;
            display: flex;
            flex-direction: column;
            border-left: 1px solid var(--border-color);
            background-color: #16181F;
        }

        .chat-messages {
            flex: 1;
            padding: 20px;
            overflow-y: auto;
            display: flex;
            flex-direction: column;
            gap: 14px;
        }

        .message-bubble {
            max-width: 85%;
            padding: 12px 16px;
            border-radius: 12px;
            line-height: 1.6;
            animation: fadeIn 0.2s ease-in;
            white-space: pre-wrap;
        }

        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(6px); }
            to { opacity: 1; transform: translateY(0); }
        }

        .message-user {
            align-self: flex-start;
            background-color: var(--user-bubble);
            border: 1px solid #334155;
            border-top-right-radius: 2px;
        }

        .message-ai {
            align-self: flex-end;
            background-color: var(--ai-bubble);
            border: 1px solid var(--excel-green);
            border-top-left-radius: 2px;
        }

        .message-confirm {
            align-self: center;
            background-color: #382718;
            border: 1px solid #D97706;
            width: 90%;
        }

        .confirm-btn-group {
            display: flex;
            gap: 12px;
            margin-top: 10px;
        }

        .btn-confirm {
            background-color: var(--excel-green);
            color: white;
            border: none;
            padding: 6px 16px;
            border-radius: 6px;
            cursor: pointer;
            font-weight: bold;
        }

        .btn-cancel {
            background-color: #991B1B;
            color: white;
            border: none;
            padding: 6px 16px;
            border-radius: 6px;
            cursor: pointer;
            font-weight: bold;
        }

        .sender-tag {
            font-size: 11px;
            color: var(--text-muted);
            margin-bottom: 4px;
            font-weight: 600;
        }

        /* Quick Chips */
        .chips-bar {
            padding: 8px 16px;
            background-color: var(--card-dark);
            display: flex;
            gap: 8px;
            overflow-x: auto;
            border-top: 1px solid var(--border-color);
        }

        .chip {
            background-color: var(--bg-dark);
            border: 1px solid var(--border-color);
            color: var(--text-muted);
            padding: 5px 12px;
            border-radius: 16px;
            font-size: 12px;
            cursor: pointer;
            white-space: nowrap;
            transition: all 0.2s;
        }

        .chip:hover {
            background-color: var(--excel-green);
            color: white;
            border-color: var(--excel-green);
        }

        /* Input Area */
        .input-bar {
            padding: 14px 18px;
            background-color: var(--card-dark);
            display: flex;
            gap: 10px;
            border-top: 1px solid var(--border-color);
        }

        .input-bar input {
            flex: 1;
            background-color: var(--bg-dark);
            border: 1px solid var(--border-color);
            color: white;
            padding: 12px 16px;
            border-radius: 8px;
            font-size: 14px;
            outline: none;
        }

        .input-bar input:focus {
            border-color: var(--excel-green);
        }

        .btn-send {
            background-color: var(--excel-green);
            color: white;
            border: none;
            padding: 0 24px;
            border-radius: 8px;
            cursor: pointer;
            font-weight: 700;
            font-size: 14px;
            transition: background 0.2s;
        }

        .btn-send:hover {
            background-color: var(--excel-green-hover);
        }

        /* Right: Excel Grid & Inspection View */
        .grid-section {
            flex: 2;
            display: flex;
            flex-direction: column;
            background-color: var(--card-dark);
            overflow: hidden;
        }

        .grid-header {
            padding: 12px 16px;
            font-weight: 700;
            font-size: 14px;
            border-bottom: 1px solid var(--border-color);
            display: flex;
            justify-content: space-between;
            align-items: center;
        }

        .grid-container {
            flex: 1;
            overflow: auto;
            padding: 12px;
        }

        table.excel-table {
            border-collapse: collapse;
            width: 100%;
            font-size: 12px;
            background-color: #1A1D24;
        }

        table.excel-table th, table.excel-table td {
            border: 1px solid #333846;
            padding: 6px 10px;
            text-align: center;
        }

        table.excel-table th {
            background-color: #262B36;
            color: #94A3B8;
            position: sticky;
            top: 0;
            font-weight: 600;
        }

        .cell-yellow { background-color: #6B5B00 !important; color: #FFF; }
        .cell-green { background-color: #0E4F29 !important; color: #FFF; }

        /* Undo and Action buttons */
        .btn-undo-panel {
            background-color: #B45309;
            color: white;
            border: none;
            padding: 6px 12px;
            border-radius: 6px;
            font-size: 12px;
            cursor: pointer;
            font-weight: bold;
        }

        .btn-undo-panel:hover {
            background-color: #D97706;
        }
    </style>
</head>
<body>

    <header>
        <div class="header-title">
            <span class="excel-badge">Excel AI</span>
            <span>دستیار هوشمند Microsoft Excel 2024</span>
        </div>
        <div class="status-bar">
            <div class="status-pill" id="statusPill">🟢 متصل به اکسل</div>
            <div id="fileName">📁 فایل: -</div>
            <div>📑 شیت:
                <select id="sheetSelect" class="sheet-select" onchange="switchSheet(this.value)"></select>
            </div>
            <button class="btn-undo-panel" onclick="sendCommand('لغو آخرین عملیات')">↩️ بازگردانی (Undo)</button>
        </div>
    </header>

    <div class="main-container">
        <!-- Chat Panel -->
        <div class="chat-section">
            <div class="chat-messages" id="chatFeed"></div>

            <!-- Quick Action Chips -->
            <div class="chips-bar">
                <button class="chip" onclick="promptArithmetic('add')">➕ جمع ستون‌ها</button>
                <button class="chip" onclick="promptArithmetic('subtract')">➖ تفریق ستون‌ها</button>
                <button class="chip" onclick="promptArithmetic('multiply')">✖️ ضرب ستون‌ها</button>
                <button class="chip" onclick="promptArithmetic('divide')">➗ تقسیم ستون‌ها</button>
                <button class="chip" onclick="sendCommand('تحلیل ساختار فایل و شیت')">📊 تحلیل شیت</button>
                <button class="chip" onclick="sendCommand('داده‌های تکراری را زرد کن')">🟨 داده‌های تکراری را زرد کن</button>
                <button class="chip" onclick="sendCommand('سلول‌های بیشتر از 100 را سبز کن')">🟩 سلول‌های بیشتر از 100 را سبز کن</button>
                <button class="chip" onclick="sendCommand('ردیف‌های خالی را حذف کن')">🧹 ردیف‌های خالی را حذف کن</button>
                <button class="chip" onclick="sendCommand('یک جدول بساز')">📋 تبدیل به جدول</button>
                <button class="chip" onclick="sendCommand('عرض ستون‌ها را خودکار تنظیم کن')">📐 تنظیم عرض ستون‌ها</button>
            </div>

            <!-- Input Bar -->
            <div class="input-bar">
                <input type="text" id="commandInput" placeholder="دستور خود را به زبان فارسی وارد کنید... (مثلاً: ستون A را در ستون B ضرب کن)" onkeydown="if(event.key==='Enter') sendInput();">
                <button class="btn-send" onclick="sendInput()">ارسال دستور</button>
            </div>
        </div>

        <!-- Excel Live Grid Preview -->
        <div class="grid-section">
            <div class="grid-header">
                <span>پیش‌نمایش زنده سلول‌های اکسل</span>
                <button class="chip" onclick="refreshGrid()">🔄 بازخوانی جدول</button>
            </div>
            <div class="grid-container" id="gridContainer">
                <div style="text-align:center; padding: 40px; color: var(--text-muted);">در حال بارگذاری داده‌های شیت...</div>
            </div>
        </div>
    </div>

    <script>
        async function fetchStatus() {
            const res = await fetch('/api/status');
            const data = await res.json();
            document.getElementById('fileName').innerText = '📁 فایل: ' + data.file_name;
            const select = document.getElementById('sheetSelect');
            select.innerHTML = '';
            (data.sheets || []).forEach(s => {
                const opt = document.createElement('option');
                opt.value = s;
                opt.innerText = s;
                if (s === data.active_sheet) opt.selected = true;
                select.appendChild(opt);
            });
            refreshGrid();
        }

        async function refreshGrid() {
            const res = await fetch('/api/sheet_data');
            const data = await res.json();
            const container = document.getElementById('gridContainer');
            if (!data.matrix || data.matrix.length === 0) {
                container.innerHTML = '<div style="text-align:center; padding:20px; color:#94A3B8;">شیت در حال حاضر خالی است.</div>';
                return;
            }

            let html = '<table class="excel-table"><thead><tr><th>#</th>';
            for (let c = 0; c < (data.matrix[0] || []).length; c++) {
                html += `<th>${String.fromCharCode(65 + c)}</th>`;
            }
            html += '</tr></thead><tbody>';

            data.matrix.forEach((row, rIdx) => {
                html += `<tr><th>${rIdx + 1}</th>`;
                row.forEach(val => {
                    const displayVal = val !== null ? val : '';
                    html += `<td>${displayVal}</td>`;
                });
                html += '</tr>';
            });
            html += '</tbody></table>';
            container.innerHTML = html;
        }

        function appendMessage(sender, text, isConfirm = false) {
            const feed = document.getElementById('chatFeed');
            const bubble = document.createElement('div');
            bubble.className = 'message-bubble ' + (sender === 'user' ? 'message-user' : (isConfirm ? 'message-confirm' : 'message-ai'));

            const tag = document.createElement('div');
            tag.className = 'sender-tag';
            tag.innerText = sender === 'user' ? '👤 شما' : '🤖 دستیار هوشمند اکسل';
            bubble.appendChild(tag);

            const content = document.createElement('div');
            content.innerText = text;
            bubble.appendChild(content);

            if (isConfirm) {
                const btnBox = document.createElement('div');
                btnBox.className = 'confirm-btn-group';
                btnBox.innerHTML = `
                    <button class="btn-confirm" onclick="confirmAction()">✅ تأیید و اجرا</button>
                    <button class="btn-cancel" onclick="cancelAction()">❌ انصراف</button>
                `;
                bubble.appendChild(btnBox);
            }

            feed.appendChild(bubble);
            feed.scrollTop = feed.scrollHeight;
        }

        function promptArithmetic(op) {
            const opNames = {add: 'جمع', subtract: 'تفریق', multiply: 'ضرب', divide: 'تقسیم'};
            const opName = opNames[op] || 'محاسبه';
            const cols = prompt(`لطفاً ستون‌های مبدأ برای عملیات ${opName} را وارد کنید (مثلاً: 2, 4, 5 یا B, D, E):`, "2, 4, 5");
            if (!cols || !cols.trim()) return;
            const target = prompt(`نتیجه در کدام ستون درج شود؟ (مثلاً: 8 یا H):`, "8");
            if (!target || !target.trim()) return;
            const cmd = `عملیات ${opName} ستون‌های ${cols.trim()} در ستون ${target.trim()}`;
            sendCommand(cmd);
        }

        async function sendCommand(cmd) {
            appendMessage('user', cmd);
            const res = await fetch('/api/command', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({command: cmd})
            });
            const data = await res.json();
            appendMessage('ai', data.message, data.requires_confirmation);
            refreshGrid();
        }

        function sendInput() {
            const input = document.getElementById('commandInput');
            const val = input.value.trim();
            if (!val) return;
            input.value = '';
            sendCommand(val);
        }

        async function confirmAction() {
            const res = await fetch('/api/confirm', {method: 'POST'});
            const data = await res.json();
            appendMessage('ai', data.message);
            refreshGrid();
        }

        async function cancelAction() {
            const res = await fetch('/api/cancel', {method: 'POST'});
            const data = await res.json();
            appendMessage('ai', data.message);
        }

        async function switchSheet(name) {
            await fetch('/api/switch_sheet', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({sheet: name})
            });
            fetchStatus();
        }

        window.onload = async () => {
            await fetchStatus();
            appendMessage('ai', 'سلام! من دستیار هوشمند Microsoft Excel هستم.\\nهر دستوری مانند «ستون A را در B ضرب کن»، «میانگین ستون D را حساب کن» یا «داده‌های تکراری را زرد کن» را به زبان فارسی بنویسید تا اجرا کنم.');
        };
    </script>
</body>
</html>
"""


class WebAssistantHandler(http.server.BaseHTTPRequestHandler):
    """HTTP request handler serving the Persian UI and Excel REST API."""

    controller: AssistantUIController = None

    def log_message(self, format, *args):
        # Silence default server logging for clean output
        pass

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        if path in ("/", "/index.html"):
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(HTML_TEMPLATE.encode("utf-8"))
            return

        if path == "/api/status":
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.end_headers()
            data = self.controller.get_status()
            self.wfile.write(json.dumps(data, ensure_ascii=False).encode("utf-8"))
            return

        if path == "/api/sheet_data":
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.end_headers()
            try:
                sheet = self.controller.manager.get_active_sheet()
                min_r, min_c, max_r, max_c = sheet.get_used_range_bounds()
                matrix = []
                for r in range(min_r, min(min_r + 50, max_r + 1)):
                    row_vals = []
                    for c in range(min_c, min(min_c + 20, max_c + 1)):
                        val = sheet.get_value(f"{col_index_to_letter(c)}{r}")
                        row_vals.append(val)
                    matrix.append(row_vals)
                resp = {"matrix": matrix, "used_range": f"{min_r}:{min_c} to {max_r}:{max_c}"}
            except Exception as e:
                resp = {"matrix": [], "error": str(e)}

            self.wfile.write(json.dumps(resp, ensure_ascii=False).encode("utf-8"))
            return

        self.send_response(404)
        self.end_headers()

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(length).decode("utf-8") if length > 0 else "{}"
        try:
            req_data = json.loads(body)
        except Exception:
            req_data = {}

        if path == "/api/command":
            cmd = req_data.get("command", "")
            res = self.controller.process_command(cmd)
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.end_headers()
            self.wfile.write(json.dumps({
                "success": res.success,
                "message": res.message_persian,
                "requires_confirmation": res.requires_confirmation
            }, ensure_ascii=False).encode("utf-8"))
            return

        if path == "/api/confirm":
            res = self.controller.confirm_pending_action()
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.end_headers()
            self.wfile.write(json.dumps({"success": res.success, "message": res.message_persian}, ensure_ascii=False).encode("utf-8"))
            return

        if path == "/api/cancel":
            self.controller.cancel_pending_action()
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.end_headers()
            self.wfile.write(json.dumps({"success": True, "message": "عملیات لغو شد."}, ensure_ascii=False).encode("utf-8"))
            return

        if path == "/api/switch_sheet":
            s_name = req_data.get("sheet", "")
            msg = self.controller.switch_sheet(s_name)
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.end_headers()
            self.wfile.write(json.dumps({"success": True, "message": msg}, ensure_ascii=False).encode("utf-8"))
            return

        self.send_response(404)
        self.end_headers()


def start_web_server(host: str = "0.0.0.0", port: int = 8080, workbook_path: Optional[str] = None):
    controller = AssistantUIController()
    controller.connect(workbook_path)

    # Initialize sample data if fresh workbook
    sheet = controller.manager.get_active_sheet()
    if sheet.get_value("A1") is None:
        sheet.set_value("A1", "کالا")
        sheet.set_value("B1", "تعداد")
        sheet.set_value("C1", "قیمت")
        sheet.set_value("D1", "فروش کل")

        sheet.set_value("A2", "لپ‌تاپ")
        sheet.set_value("B2", 5)
        sheet.set_value("C2", 1200)

        sheet.set_value("A3", "موبایل")
        sheet.set_value("B3", 10)
        sheet.set_value("C3", 800)

        sheet.set_value("A4", "تبلت")
        sheet.set_value("B4", 7)
        sheet.set_value("C4", 450)

        sheet.set_value("A5", "موبایل")
        sheet.set_value("B5", 10)
        sheet.set_value("C5", 800)

        # Empty row
        sheet.set_value("A7", "مانیتور")
        sheet.set_value("B7", 3)
        sheet.set_value("C7", 300)

    class CustomServer(socketserver.TCPServer):
        allow_reuse_address = True

    WebAssistantHandler.controller = controller
    with CustomServer((host, port), WebAssistantHandler) as httpd:
        print(f"🚀 Excel AI Assistant Web Server running at http://{host}:{port}")
        httpd.serve_forever()

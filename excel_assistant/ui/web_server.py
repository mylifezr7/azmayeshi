"""
Lightweight, Zero-Dependency Persian Web UI Server for Excel AI Assistant.
Provides a modern live web preview with Excel grid view, deep cosmic glassmorphic theme, and Persian AI chat.
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
    <title>Excel AI Assistant - MIYUM</title>
    <link href="https://fonts.googleapis.com/css2?family=Vazirmatn:wght@300;400;600;700;800&display=swap" rel="stylesheet">
    <style>
        :root {
            --bg-deep: #020617;
            --bg-card: #081026;
            --border-blue: #1D4ED8;
            --border-subtle: #132247;
            --text-light: #F8FAFC;
            --text-muted: #94A3B8;
            --user-bubble: rgba(14, 30, 68, 0.92);
            --ai-bubble: rgba(8, 16, 38, 0.88);
            --accent-cyan: #38BDF8;
            --accent-blue: #2563EB;
        }

        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
            font-family: 'Vazirmatn', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
        }

        body {
            background-color: var(--bg-deep);
            color: var(--text-light);
            height: 100vh;
            display: flex;
            flex-direction: column;
            overflow: hidden;
        }

        /* Top Header */
        header {
            background: linear-gradient(90deg, #09132B, #04091A);
            border-bottom: 2px solid var(--border-blue);
            padding: 10px 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            box-shadow: 0 4px 18px rgba(0, 0, 0, 0.6);
            z-index: 10;
        }

        .header-right {
            display: flex;
            align-items: center;
            gap: 14px;
        }

        .user-badge {
            background: linear-gradient(135deg, #08132B, #0E224F);
            color: var(--accent-cyan);
            border: 1.5px solid var(--accent-blue);
            font-weight: 800;
            padding: 5px 16px;
            border-radius: 14px;
            font-size: 13px;
            letter-spacing: 0.5px;
            display: flex;
            align-items: center;
            gap: 6px;
            box-shadow: 0 0 10px rgba(56, 189, 248, 0.2);
        }

        .status-pill {
            display: flex;
            align-items: center;
            gap: 8px;
            background-color: #081024;
            border: 1px solid #1E293B;
            padding: 5px 14px;
            border-radius: 14px;
            color: #E2E8F0;
            font-size: 12px;
            font-weight: 600;
        }

        .status-pill span.dot {
            color: #10B981;
            font-size: 13px;
        }

        .btn-open-file {
            background: linear-gradient(90deg, #1D4ED8, #3B82F6);
            color: white;
            border: 1px solid #60A5FA;
            padding: 6px 14px;
            border-radius: 10px;
            font-size: 13px;
            font-weight: bold;
            cursor: pointer;
            transition: all 0.2s;
        }

        .btn-open-file:hover {
            background: linear-gradient(90deg, #2563EB, #60A5FA);
            box-shadow: 0 0 12px rgba(59, 130, 246, 0.4);
        }

        .header-left {
            display: flex;
            align-items: center;
            gap: 14px;
            font-size: 12px;
        }

        select.sheet-select {
            background-color: #081024;
            color: var(--text-light);
            border: 1px solid var(--border-blue);
            padding: 5px 12px;
            border-radius: 8px;
            cursor: pointer;
            outline: none;
            font-size: 12px;
        }

        .btn-reconnect {
            background-color: #081024;
            color: #93C5FD;
            border: 1px solid var(--border-blue);
            padding: 5px 12px;
            border-radius: 8px;
            cursor: pointer;
            font-weight: bold;
            font-size: 12px;
            transition: all 0.2s;
        }

        .btn-reconnect:hover {
            background-color: var(--border-blue);
            color: #FFFFFF;
        }

        /* Main Container */
        .main-container {
            flex: 1;
            display: flex;
            overflow: hidden;
            position: relative;
        }

        /* Sidebar Action Panel */
        .sidebar-panel {
            width: 250px;
            background-color: #04091A;
            border-left: 1px solid var(--border-subtle);
            padding: 12px;
            display: flex;
            flex-direction: column;
            gap: 10px;
            overflow-y: auto;
        }

        .sidebar-title {
            font-size: 13px;
            font-weight: 700;
            color: var(--text-muted);
            margin-bottom: 2px;
            display: flex;
            align-items: center;
            gap: 6px;
        }

        .history-box {
            flex: 1;
            background-color: #020617;
            border: 1px solid var(--border-subtle);
            border-radius: 10px;
            padding: 6px;
            overflow-y: auto;
            min-height: 120px;
            font-size: 12px;
            display: flex;
            flex-direction: column;
            gap: 4px;
        }

        .history-item {
            padding: 6px 8px;
            border-radius: 6px;
            color: #CBD5E1;
            cursor: pointer;
            transition: background 0.15s;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .history-item:hover {
            background-color: #0B193D;
            color: var(--accent-cyan);
        }

        /* High-Vibrancy Primary Colored Sidebar Buttons */
        .btn-shortcut {
            width: 100%;
            padding: 10px 12px;
            border-radius: 10px;
            color: #FFFFFF;
            font-weight: 700;
            font-size: 12.5px;
            cursor: pointer;
            text-align: center;
            border: none;
            transition: all 0.2s;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.4);
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 6px;
        }

        .btn-undo {
            background: linear-gradient(90deg, #EA580C, #FB923C);
            border: 1px solid #F97316;
        }
        .btn-undo:hover {
            background: linear-gradient(90deg, #F97316, #FDBA74);
            box-shadow: 0 0 12px rgba(249, 115, 22, 0.6);
        }

        .btn-unique {
            background: linear-gradient(90deg, #6D28D9, #8B5CF6);
            border: 1px solid #8B5CF6;
        }
        .btn-unique:hover {
            background: linear-gradient(90deg, #7C3AED, #A78BFA);
            box-shadow: 0 0 12px rgba(139, 92, 246, 0.6);
        }

        .btn-arithmetic {
            background: linear-gradient(90deg, #047857, #10B981);
            border: 1px solid #10B981;
        }
        .btn-arithmetic:hover {
            background: linear-gradient(90deg, #059669, #34D399);
            box-shadow: 0 0 12px rgba(16, 185, 129, 0.6);
        }

        .btn-discount {
            background: linear-gradient(90deg, #B45309, #F59E0B);
            border: 1px solid #F59E0B;
        }
        .btn-discount:hover {
            background: linear-gradient(90deg, #D97706, #FBBF24);
            box-shadow: 0 0 12px rgba(245, 158, 11, 0.6);
        }

        .btn-date {
            background: linear-gradient(90deg, #0284C7, #38BDF8);
            border: 1px solid #38BDF8;
        }
        .btn-date:hover {
            background: linear-gradient(90deg, #0369A1, #7DD3FC);
            box-shadow: 0 0 12px rgba(56, 189, 248, 0.6);
        }

        .btn-analyze {
            background: linear-gradient(90deg, #1E3A8A, #2563EB);
            border: 1px solid #3B82F6;
        }
        .btn-analyze:hover {
            background: linear-gradient(90deg, #1D4ED8, #60A5FA);
            box-shadow: 0 0 12px rgba(37, 99, 235, 0.6);
        }

        .btn-sales {
            background: linear-gradient(90deg, #BE123C, #F43F5E);
            border: 1px solid #F43F5E;
        }
        .btn-sales:hover {
            background: linear-gradient(90deg, #E11D48, #FB7185);
            box-shadow: 0 0 12px rgba(244, 63, 94, 0.6);
        }

        .btn-profit {
            background: linear-gradient(90deg, #059669, #10B981);
            border: 1px solid #10B981;
        }
        .btn-profit:hover {
            background: linear-gradient(90deg, #10B981, #34D399);
            box-shadow: 0 0 12px rgba(16, 185, 129, 0.6);
        }

        /* Center: Chat Area with Background Wallpaper */
        .chat-section {
            flex: 3;
            display: flex;
            flex-direction: column;
            background-color: var(--bg-deep);
            background-image: url('/resources/background.jpg');
            background-repeat: no-repeat;
            background-position: center bottom;
            background-size: cover;
            position: relative;
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
            padding: 14px 18px;
            border-radius: 14px;
            line-height: 1.6;
            animation: fadeIn 0.2s ease-in;
            white-space: pre-wrap;
            box-shadow: 0 4px 14px rgba(0, 0, 0, 0.4);
        }

        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(6px); }
            to { opacity: 1; transform: translateY(0); }
        }

        .message-user {
            align-self: flex-start;
            background-color: var(--user-bubble);
            border: 1.5px solid var(--accent-blue);
            border-top-right-radius: 2px;
        }

        .message-ai {
            align-self: flex-end;
            background-color: var(--ai-bubble);
            border: 1.5px solid var(--border-blue);
            border-top-left-radius: 2px;
        }

        .message-confirm {
            align-self: center;
            background-color: rgba(45, 22, 10, 0.94);
            border: 1.5px solid #F59E0B;
            width: 90%;
        }

        .confirm-btn-group {
            display: flex;
            gap: 12px;
            margin-top: 10px;
        }

        .btn-confirm {
            background: linear-gradient(90deg, #059669, #10B981);
            color: white;
            border: 1px solid #10B981;
            padding: 7px 18px;
            border-radius: 8px;
            cursor: pointer;
            font-weight: bold;
        }

        .btn-cancel {
            background: linear-gradient(90deg, #991B1B, #DC2626);
            color: white;
            border: 1px solid #EF4444;
            padding: 7px 18px;
            border-radius: 8px;
            cursor: pointer;
            font-weight: bold;
        }

        .sender-tag {
            font-size: 11px;
            color: var(--accent-cyan);
            margin-bottom: 6px;
            font-weight: 700;
        }

        /* Quick Chips */
        .chips-bar {
            padding: 8px 16px;
            background: rgba(8, 16, 38, 0.7);
            display: flex;
            gap: 8px;
            overflow-x: auto;
            border-top: 1px solid var(--border-subtle);
            backdrop-filter: blur(8px);
        }

        .chip {
            background-color: rgba(10, 20, 42, 0.85);
            border: 1px solid var(--border-blue);
            color: #93C5FD;
            padding: 5px 13px;
            border-radius: 14px;
            font-size: 12px;
            cursor: pointer;
            white-space: nowrap;
            transition: all 0.2s;
            font-weight: 500;
        }

        .chip:hover {
            background-color: var(--border-blue);
            color: #FFFFFF;
            border-color: var(--accent-cyan);
            box-shadow: 0 0 10px rgba(56, 189, 248, 0.3);
        }

        /* Input Area */
        .input-bar {
            padding: 12px 18px;
            background: linear-gradient(180deg, #09132B, #04091A);
            display: flex;
            gap: 10px;
            border-top: 1px solid var(--border-subtle);
        }

        .input-bar input {
            flex: 1;
            background-color: rgba(6, 12, 28, 0.9);
            border: 1px solid var(--border-blue);
            color: white;
            padding: 11px 16px;
            border-radius: 10px;
            font-size: 13px;
            outline: none;
            transition: border-color 0.2s;
        }

        .input-bar input:focus {
            border-color: var(--accent-cyan);
            box-shadow: 0 0 8px rgba(56, 189, 248, 0.3);
        }

        .btn-send {
            background: linear-gradient(90deg, #1D4ED8, #3B82F6);
            color: white;
            border: 1px solid #60A5FA;
            padding: 0 22px;
            border-radius: 10px;
            cursor: pointer;
            font-weight: 700;
            font-size: 13px;
            transition: all 0.2s;
        }

        .btn-send:hover {
            background: linear-gradient(90deg, #2563EB, #60A5FA);
            box-shadow: 0 0 12px rgba(59, 130, 246, 0.5);
        }

        .btn-stop {
            background: linear-gradient(90deg, #991B1B, #DC2626);
            color: white;
            border: 1px solid #EF4444;
            padding: 0 18px;
            border-radius: 10px;
            cursor: pointer;
            font-weight: 700;
            font-size: 13px;
            transition: all 0.2s;
        }

        .btn-stop:hover {
            background: linear-gradient(90deg, #B91C1C, #EF4444);
        }

        /* Right: Excel Grid & Inspection View */
        .grid-section {
            flex: 2;
            display: flex;
            flex-direction: column;
            background-color: #04091A;
            border-right: 1px solid var(--border-subtle);
            overflow: hidden;
        }

        .grid-header {
            padding: 12px 16px;
            font-weight: 700;
            font-size: 13px;
            border-bottom: 1px solid var(--border-subtle);
            display: flex;
            justify-content: space-between;
            align-items: center;
            background-color: #081024;
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
            background-color: #060B18;
        }

        table.excel-table th, table.excel-table td {
            border: 1px solid #132247;
            padding: 6px 10px;
            text-align: center;
        }

        table.excel-table th {
            background-color: #0B1938;
            color: #94A3B8;
            position: sticky;
            top: 0;
            font-weight: 600;
        }

        .cell-yellow { background-color: #713F12 !important; color: #FFF; }
        .cell-green { background-color: #064E3B !important; color: #FFF; }
    </style>
</head>
<body>

    <!-- Top Header -->
    <header>
        <div class="header-right">
            <span class="user-badge">M MIYUM</span>
            <div class="status-pill" id="statusPill"><span class="dot">🟢</span> Microsoft Excel 2024 (COM)</div>
            <button class="btn-open-file" onclick="promptOpenFile()">📂 باز کردن فایل اکسل...</button>
        </div>
        <div class="header-left">
            <div id="fileName" style="color: #CBD5E1;">📁 فایل: Book1</div>
            <div style="color: #CBD5E1;">📑 شیت:
                <select id="sheetSelect" class="sheet-select" onchange="switchSheet(this.value)"></select>
            </div>
            <button class="btn-reconnect" onclick="fetchStatus()">📶 اتصال مجدد</button>
        </div>
    </header>

    <div class="main-container">
        <!-- Sidebar Panel: History + High-Vibrancy Primary Colored Shortcuts -->
        <div class="sidebar-panel">
            <div class="sidebar-title">📜 تاریخچه دستورات</div>
            <div class="history-box" id="historyBox"></div>

            <button class="btn-shortcut btn-undo" onclick="sendCommand('لغو آخرین عملیات')">↩️ لغو آخرین تغییر (Undo)</button>
            <button class="btn-shortcut btn-unique" onclick="promptPreventDuplicates()">🔒 قفل یکتایی ستون اکسل...</button>
            <button class="btn-shortcut btn-arithmetic" onclick="promptArithmetic('add')">🧮 ۴ عمل اصلی ستون‌ها...</button>
            <button class="btn-shortcut btn-discount" onclick="promptDiscount()">🏷️ محاسبه تخفیف کالاها...</button>
            <button class="btn-shortcut btn-date" onclick="promptInsertDate()">📅 درج تاریخ در شیت...</button>
            <button class="btn-shortcut btn-analyze" onclick="sendCommand('تحلیل ساختار فایل و شیت')">📊 تحلیل هوشمند شیت</button>
            <button class="btn-shortcut btn-sales" onclick="promptSalesSummary()">💰 محاسبه فروش (روزانه/ماهانه/سالانه/کل)...</button>
            <button class="btn-shortcut btn-profit" onclick="promptProfitCalculation()">💵 محاسبه سود (فروش - خرید)...</button>
        </div>

        <!-- Chat Panel with Wallpaper Backdrop -->
        <div class="chat-section">
            <div class="chat-messages" id="chatFeed"></div>

            <!-- Quick Action Chips -->
            <div class="chips-bar">
                <button class="chip" onclick="promptProfitCalculation()">💵 محاسبه سود (شیت ۱ و ۲)</button>
                <button class="chip" onclick="promptSalesSummary()">💰 محاسبه فروش روزانه و ماهانه</button>
                <button class="chip" onclick="promptArithmetic('add')">➕ جمع ستون‌ها</button>
                <button class="chip" onclick="promptArithmetic('subtract')">➖ تفریق ستون‌ها</button>
                <button class="chip" onclick="promptArithmetic('multiply')">✖️ ضرب ستون‌ها</button>
                <button class="chip" onclick="promptArithmetic('divide')">➗ تقسیم ستون‌ها</button>
                <button class="chip" onclick="promptDiscount()">🏷️ اعمال تخفیف کالاها</button>
                <button class="chip" onclick="promptInsertDate()">📅 درج تاریخ در شیت</button>
                <button class="chip" onclick="promptPreventDuplicates()">🔒 قفل داده تکراری در ستون</button>
                <button class="chip" onclick="sendCommand('داده‌های متنی تکراری را رنگ کن')">🎨 رنگ‌آمیزی متون تکراری</button>
                <button class="chip" onclick="sendCommand('تحلیل ساختار فایل و شیت')">📊 تحلیل هوشمند شیت</button>
                <button class="chip" onclick="sendCommand('سلول‌های بیشتر از 100 را سبز کن')">🟩 سلول‌های بیشتر از 100</button>
                <button class="chip" onclick="sendCommand('ردیف‌های خالی را حذف کن')">🧹 حذف ردیف‌های خالی</button>
                <button class="chip" onclick="sendCommand('یک جدول بساز')">📋 ساخت جدول یکدست</button>
                <button class="chip" onclick="sendCommand('عرض ستون‌ها را خودکار تنظیم کن')">📐 تنظیم عرض ستون‌ها</button>
            </div>

            <!-- Input Bar -->
            <div class="input-bar">
                <input type="text" id="commandInput" placeholder="دستور خود را به زبان فارسی وارد کنید... (مثلاً: ستون A را در ستون B ضرب کن و در C بنویس)" onkeydown="if(event.key==='Enter') sendInput();">
                <button class="btn-send" onclick="sendInput()">ارسال دستور ✈️</button>
                <button class="btn-stop" onclick="cancelAction()">توقف ⏹</button>
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
            try {
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
            } catch(e) {
                console.error("Status error:", e);
            }
        }

        async function refreshGrid() {
            try {
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
            } catch(e) {
                console.error("Grid error:", e);
            }
        }

        function appendMessage(sender, text, isConfirm = false) {
            const feed = document.getElementById('chatFeed');
            const bubble = document.createElement('div');
            bubble.className = 'message-bubble ' + (sender === 'user' ? 'message-user' : (isConfirm ? 'message-confirm' : 'message-ai'));

            const tag = document.createElement('div');
            tag.className = 'sender-tag';
            tag.innerText = sender === 'user' ? '👤 MIYUM' : '🤖 دستیار هوشمند Microsoft Excel';
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

        function addHistoryItem(cmd) {
            const box = document.getElementById('historyBox');
            const item = document.createElement('div');
            item.className = 'history-item';
            item.innerText = cmd;
            item.onclick = () => {
                document.getElementById('commandInput').value = cmd;
            };
            box.insertBefore(item, box.firstChild);
        }

        function promptOpenFile() {
            const path = prompt("مسیر یا نام فایل اکسل مورد نظر را وارد کنید:", "sample.xlsx");
            if (!path || !path.trim()) return;
            sendCommand(`فایل اکسل ${path.trim()} را باز کن`);
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

        function promptDiscount() {
            const col = prompt("ستون قیمت اصلی کالاها را وارد کنید (مثلاً: B یا 2):", "B");
            if (!col || !col.trim()) return;
            const pct = prompt("درصد تخفیف یا ستون درصد (مثلاً: 20 یا 20% یا C):", "20%");
            if (!pct || !pct.trim()) return;
            const target = prompt("نتیجه قیمت نهایی در کدام ستون درج شود؟ (مثلاً: D):", "D");
            if (!target || !target.trim()) return;
            const cmd = `تخفیف ${pct.trim()} ستون ${col.trim()} در ستون ${target.trim()}`;
            sendCommand(cmd);
        }

        function promptInsertDate() {
            const col = prompt("تاریخ در کدام ستون درج شود؟ (مثلاً: A یا 1):", "A");
            if (!col || !col.trim()) return;
            const range = prompt("محدوده سطرها (از سطر تا سطر):", "از سطر 1 تا 30");
            if (!range || !range.trim()) return;
            const date = prompt("تاریخ مورد نظر (مثلاً: 1403/08/29):", "1403/08/29");
            if (!date || !date.trim()) return;
            const cmd = `درج تاریخ ${date.trim()} ${range.trim()} در ستون ${col.trim()}`;
            sendCommand(cmd);
        }

        function promptPreventDuplicates() {
            const col = prompt("ستون مورد نظر برای قفل یکتایی و جلوگیری از تکرار (مثلاً: 12 یا L):", "12");
            if (!col || !col.trim()) return;
            const cmd = `ستون ${col.trim()} نباید داده تکراری داشته باشه`;
            sendCommand(cmd);
        }

        function promptSalesSummary() {
            const dateCol = prompt("ستون تاریخ را وارد کنید (مثلاً: A یا 1):", "A");
            if (!dateCol || !dateCol.trim()) return;
            const amtCol = prompt("ستون مبلغ فروش را وارد کنید (مثلاً: B یا 2 یا D):", "B");
            if (!amtCol || !amtCol.trim()) return;
            const period = prompt("دوره زمانی محاسبه (all: همه | daily: روزانه | monthly: ماهانه | yearly: سالانه | total: کل):", "all");
            const periodMap = {all: "روزانه و ماهانه و سالانه و کل", daily: "روزانه", monthly: "ماهانه", yearly: "سالانه", total: "کل"};
            const pText = periodMap[period] || "روزانه و ماهانه و سالانه و کل";
            const cmd = `محاسبه فروش ${pText} ستون فروش ${amtCol.trim()} بر اساس تاریخ ${dateCol.trim()}`;
            sendCommand(cmd);
        }

        function promptProfitCalculation() {
            const sSheet = prompt("نام یا شماره شیت فروش (پیش‌فرض: شیت اول):", "Sheet1") || "1";
            const pSheet = prompt("نام یا شماره شیت خرید (پیش‌فرض: شیت دوم):", "Sheet2") || "2";
            const cmd = `تعداد فروش و قیمت فروش در ${sSheet.trim()} و تعداد خرید و قیمت خرید در ${pSheet.trim()}، فروش رو منهای خرید کن تا سود به دست بیاد`;
            sendCommand(cmd);
        }

        async function sendCommand(cmd) {
            appendMessage('user', cmd);
            addHistoryItem(cmd);
            try {
                const res = await fetch('/api/command', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({command: cmd})
                });
                const data = await res.json();
                appendMessage('ai', data.message, data.requires_confirmation);
                refreshGrid();
            } catch(e) {
                appendMessage('ai', 'خطا در ارتباط با سرور: ' + e);
            }
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
            appendMessage('ai', 'سلام! من دستیار هوشمند Microsoft Excel هستم.\\nهر سوالی در زمینه فرمول‌ها از Excel تا بپرسید.\\nبرای کار روی فایل اکسل خود و وارد کردن فایل اکسل... در نوار بالا کلیک کنید.');
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

        if path == "/resources/background.jpg":
            bg_path = os.path.join(os.path.dirname(__file__), "resources", "background.jpg")
            if os.path.exists(bg_path):
                self.send_response(200)
                self.send_header("Content-Type", "image/jpeg")
                self.send_header("Cache-Control", "public, max-age=86400")
                self.end_headers()
                with open(bg_path, "rb") as f:
                    self.wfile.write(f.read())
                return
            else:
                self.send_response(404)
                self.end_headers()
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
        body = self.rfile.read(length) if length > 0 else b""

        try:
            req_data = json.loads(body.decode("utf-8")) if body else {}
        except Exception:
            req_data = {}

        if path == "/api/command":
            cmd = req_data.get("command", "")
            res = self.controller.process_command(cmd)
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.end_headers()
            resp = {
                "success": res.success,
                "message": res.message_persian,
                "requires_confirmation": res.requires_confirmation
            }
            self.wfile.write(json.dumps(resp, ensure_ascii=False).encode("utf-8"))
            return

        if path == "/api/confirm":
            res = self.controller.confirm_pending_action()
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.end_headers()
            resp = {
                "success": res.success,
                "message": res.message_persian
            }
            self.wfile.write(json.dumps(resp, ensure_ascii=False).encode("utf-8"))
            return

        if path == "/api/cancel":
            res = self.controller.cancel_pending_action()
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.end_headers()
            resp = {
                "success": res.success,
                "message": res.message_persian
            }
            self.wfile.write(json.dumps(resp, ensure_ascii=False).encode("utf-8"))
            return

        if path == "/api/switch_sheet":
            sheet_name = req_data.get("sheet", "")
            msg = self.controller.switch_sheet(sheet_name)
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.end_headers()
            self.wfile.write(json.dumps({"message": msg}, ensure_ascii=False).encode("utf-8"))
            return

        self.send_response(404)
        self.end_headers()


def start_web_server(host: str = "0.0.0.0", port: int = 8080, workbook_path: Optional[str] = None, controller: Optional[AssistantUIController] = None):
    """Starts the standalone Persian Web UI Server."""
    ctrl = controller or AssistantUIController()
    ctrl.connect(workbook_path)

    # Initialize sample data if fresh workbook
    sheet = ctrl.manager.get_active_sheet()
    if sheet.get_value("A1") is None:
        # Sheet 1: Sales
        sheet.set_value("A1", "کالا")
        sheet.set_value("B1", "تعداد فروش")
        sheet.set_value("C1", "قیمت فروش")
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

        sheet.set_value("A6", "مانیتور")
        sheet.set_value("B6", 3)
        sheet.set_value("C6", 300)

        # Create Sheet 2 for Purchases if not present
        if len(ctrl.manager.active_driver.get_sheet_names()) < 2:
            sheet2 = ctrl.manager.active_driver.create_sheet("Sheet2")
            sheet2.set_value("A1", "کالا")
            sheet2.set_value("B1", "تعداد خرید")
            sheet2.set_value("C1", "قیمت خرید")
            sheet2.set_value("D1", "خرید کل")

            sheet2.set_value("A2", "لپ‌تاپ")
            sheet2.set_value("B2", 5)
            sheet2.set_value("C2", 950)

            sheet2.set_value("A3", "موبایل")
            sheet2.set_value("B3", 20)
            sheet2.set_value("C3", 620)

            sheet2.set_value("A4", "تبلت")
            sheet2.set_value("B4", 7)
            sheet2.set_value("C4", 350)

            sheet2.set_value("A5", "مانیتور")
            sheet2.set_value("B5", 3)
            sheet2.set_value("C5", 220)

    WebAssistantHandler.controller = ctrl
    server_address = (host, port)

    class ReusableTCPServer(socketserver.TCPServer):
        allow_reuse_address = True

    with ReusableTCPServer(server_address, WebAssistantHandler) as httpd:
        print(f"🚀 سرور وب دستیار هوشمند اکسل با تم تیره و مدرن فعال شد:")
        print(f"👉 http://127.0.0.1:{port}")
        httpd.serve_forever()


if __name__ == "__main__":
    start_web_server(8080)

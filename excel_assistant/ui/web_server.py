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
    <title>Excel AI Assistant - MUYUM</title>
    <link href="https://fonts.googleapis.com/css2?family=Vazirmatn:wght@300;400;600;700;800&family=Sahel:wght@400;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --bg-deep: #020617;
            --bg-card: #081026;
            --border-blue: #1D4ED8;
            --border-subtle: #132247;
            --text-light: #F8FAFC;
            --text-muted: #94A3B8;
            --user-bubble: rgba(6, 32, 44, 0.92);
            --ai-bubble: rgba(4, 21, 31, 0.90);
            --accent-cyan: #38BDF8;
            --accent-blue: #2563EB;
            --accent-gold: #F59E0B;
            --accent-emerald: #10B981;
        }

        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
            font-family: 'B Nazanin', 'Sahel', 'Vazirmatn', 'IRANSans', -apple-system, BlinkMacSystemFont, 'Segoe UI', Tahoma, serif;
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
            background: linear-gradient(180deg, #062226 0%, #0B192C 50%, #030A16 100%);
            border-bottom: 2px solid #D97706;
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
            background: linear-gradient(135deg, #78350F, #B45309, #D97706);
            color: #FEF3C7;
            border: 1.5px solid #F59E0B;
            font-weight: 800;
            padding: 5px 16px;
            border-radius: 14px;
            font-size: 13px;
            letter-spacing: 0.5px;
            display: flex;
            align-items: center;
            gap: 6px;
            box-shadow: 0 0 12px rgba(245, 158, 11, 0.3);
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

        .btn-save {
            background: linear-gradient(90deg, #0F766E, #14B8A6);
            border: 1px solid #2DD4BF;
        }
        .btn-save:hover {
            background: linear-gradient(90deg, #115E59, #2DD4BF);
            box-shadow: 0 0 12px rgba(45, 212, 191, 0.6);
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

        .btn-unlock {
            background: linear-gradient(90deg, #991B1B, #DC2626);
            border: 1px solid #DC2626;
        }
        .btn-unlock:hover {
            background: linear-gradient(90deg, #B91C1C, #EF4444);
            box-shadow: 0 0 12px rgba(220, 38, 38, 0.6);
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

        .hotkey-badge {
            background: rgba(0, 0, 0, 0.35);
            border: 1px solid rgba(255, 255, 255, 0.25);
            border-radius: 4px;
            padding: 1px 6px;
            font-size: 10px;
            margin-right: 6px;
            color: #E2E8F0;
            display: inline-block;
        }

        /* In-App Interactive Modal Dialogs */
        .modal-overlay {
            position: fixed;
            top: 0;
            left: 0;
            width: 100vw;
            height: 100vh;
            background: rgba(2, 6, 23, 0.85);
            backdrop-filter: blur(8px);
            z-index: 9999;
            display: flex;
            align-items: center;
            justify-content: center;
            animation: fadeIn 0.2s ease-out;
        }

        .modal-card {
            background: #081026;
            border: 1.5px solid #10B981;
            box-shadow: 0 12px 45px rgba(16, 185, 129, 0.35);
            border-radius: 16px;
            width: 92%;
            max-width: 530px;
            color: var(--text-light);
            display: flex;
            flex-direction: column;
            overflow: hidden;
            animation: slideUp 0.25s cubic-bezier(0.16, 1, 0.3, 1);
        }

        .modal-card-sales {
            border: 1.5px solid #F43F5E;
            box-shadow: 0 12px 45px rgba(244, 63, 94, 0.35);
        }

        @keyframes slideUp {
            from { opacity: 0; transform: translateY(20px) scale(0.97); }
            to { opacity: 1; transform: translateY(0) scale(1); }
        }

        .modal-header {
            padding: 16px 20px;
            background: rgba(14, 30, 68, 0.65);
            border-bottom: 1px solid var(--border-subtle);
            display: flex;
            justify-content: space-between;
            align-items: center;
        }

        .modal-title {
            font-size: 15px;
            font-weight: 700;
            color: #34D399;
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .modal-title-sales {
            color: #FB7185;
        }

        .modal-close {
            background: transparent;
            border: none;
            color: var(--text-muted);
            font-size: 24px;
            cursor: pointer;
            padding: 0 4px;
            line-height: 1;
            transition: color 0.15s;
        }
        .modal-close:hover {
            color: #EF4444;
        }

        .modal-body {
            padding: 20px;
            display: flex;
            flex-direction: column;
            gap: 13px;
            max-height: 75vh;
            overflow-y: auto;
        }

        .modal-group {
            display: flex;
            flex-direction: column;
            gap: 6px;
        }

        .modal-label {
            font-size: 12.5px;
            font-weight: 600;
            color: #E2E8F0;
        }

        .modal-input {
            background: #020617;
            border: 1px solid var(--border-blue);
            border-radius: 8px;
            color: #F8FAFC;
            padding: 9px 12px;
            font-family: inherit;
            font-size: 13px;
            outline: none;
            transition: border-color 0.15s, box-shadow 0.15s;
        }
        .modal-input:focus {
            border-color: var(--accent-cyan);
            box-shadow: 0 0 10px rgba(56, 189, 248, 0.4);
        }

        .modal-radios {
            display: flex;
            flex-direction: column;
            gap: 6px;
            background: rgba(2, 6, 23, 0.5);
            padding: 10px 14px;
            border-radius: 10px;
            border: 1px solid var(--border-subtle);
        }

        .radio-option {
            display: flex;
            align-items: center;
            gap: 8px;
            cursor: pointer;
            font-size: 12px;
            color: #CBD5E1;
        }

        .radio-option input[type="radio"] {
            accent-color: #10B981;
            cursor: pointer;
        }

        .radio-option-sales input[type="radio"] {
            accent-color: #F43F5E;
        }

        .modal-preview {
            background: rgba(4, 9, 26, 0.85);
            border: 1px solid var(--border-subtle);
            border-radius: 8px;
            padding: 10px 12px;
            font-size: 11.5px;
            color: #94A3B8;
            line-height: 1.5;
        }

        .modal-footer {
            padding: 14px 20px;
            background: rgba(14, 30, 68, 0.4);
            border-top: 1px solid var(--border-subtle);
            display: flex;
            justify-content: flex-end;
            gap: 12px;
        }

        .btn-modal-cancel {
            background: #1E293B;
            color: #E2E8F0;
            border: 1px solid #334155;
            padding: 8px 18px;
            border-radius: 8px;
            cursor: pointer;
            font-family: inherit;
            font-weight: 600;
        }
        .btn-modal-cancel:hover {
            background: #334155;
        }

        .btn-modal-submit-profit {
            background: linear-gradient(90deg, #059669, #10B981);
            color: #FFFFFF;
            border: 1px solid #10B981;
            padding: 8px 24px;
            border-radius: 8px;
            cursor: pointer;
            font-family: inherit;
            font-weight: 700;
            box-shadow: 0 0 10px rgba(16, 185, 129, 0.5);
        }
        .btn-modal-submit-profit:hover {
            background: linear-gradient(90deg, #10B981, #34D399);
            box-shadow: 0 0 14px rgba(52, 211, 153, 0.7);
        }

        .btn-modal-submit-sales {
            background: linear-gradient(90deg, #BE123C, #F43F5E);
            color: #FFFFFF;
            border: 1px solid #F43F5E;
            padding: 8px 24px;
            border-radius: 8px;
            cursor: pointer;
            font-family: inherit;
            font-weight: 700;
            box-shadow: 0 0 10px rgba(244, 63, 94, 0.5);
        }
        .btn-modal-submit-sales:hover {
            background: linear-gradient(90deg, #E11D48, #FB7185);
            box-shadow: 0 0 14px rgba(251, 113, 133, 0.7);
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

        /* Quick Action Bar */
        .chips-bar {
            padding: 10px 18px;
            background: rgba(8, 16, 38, 0.85);
            display: flex;
            align-items: center;
            gap: 12px;
            border-top: 1px solid var(--border-subtle);
            backdrop-filter: blur(8px);
        }

        .chip-select {
            background-color: #0A142A;
            border: 2px solid var(--accent-blue);
            color: #E2E8F0;
            padding: 8px 16px;
            border-radius: 10px;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            outline: none;
            white-space: nowrap;
            font-family: inherit;
            transition: all 0.2s;
            flex: 1;
            min-height: 44px;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.25);
        }

        .chip-select:hover, .chip-select:focus {
            border-color: var(--accent-cyan);
            background-color: #112042;
            color: #FFFFFF;
            box-shadow: 0 0 12px rgba(6, 182, 212, 0.35);
        }

        .action-run-btn {
            background: linear-gradient(135deg, #059669, #10B981);
            color: #FFFFFF;
            font-weight: bold;
            font-size: 14px;
            border: 2px solid #34D399;
            padding: 8px 24px;
            border-radius: 10px;
            cursor: pointer;
            outline: none;
            font-family: inherit;
            transition: all 0.2s;
            min-height: 44px;
            box-shadow: 0 3px 10px rgba(16, 185, 129, 0.35);
            flex-shrink: 0;
        }

        .action-run-btn:hover {
            background: linear-gradient(135deg, #047857, #059669);
            border-color: #6EE7B7;
            transform: translateY(-1px);
            box-shadow: 0 4px 14px rgba(16, 185, 129, 0.5);
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
            <span class="user-badge">M MUYUM</span>
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

            <button class="btn-shortcut btn-save" onclick="saveWorkbook()">💾 ذخیره فایل (Save)... <span class="hotkey-badge">Ctrl+S</span></button>
            <button class="btn-shortcut btn-undo" onclick="sendCommand('لغو آخرین عملیات')">↩️ لغو آخرین تغییر (Undo)</button>
            <button class="btn-shortcut btn-arithmetic" onclick="promptArithmetic('add')">🧮 ۴ عمل اصلی ستون‌ها...</button>
        </div>

        <!-- Chat Panel with Wallpaper Backdrop -->
        <div class="chat-section">
            <div class="chat-messages" id="chatFeed"></div>

            <!-- Quick Action Cascading Selector (5 Categories) -->
            <div class="chips-bar">
                <!-- Dropdown 1: انتخاب حوزه ۵ گانه -->
                <select id="catSelect" class="chip-select" onchange="onCategoryChanged(this.value)">
                    <option value="">📁 ۱. انتخاب حوزه (۵ دسته‌بندی اصلی)...</option>
                    <option value="essential">⚡ ۱. قابلیت‌های ضروری</option>
                    <option value="accounting">💼 ۲. قابلیت‌های حسابداری</option>
                    <option value="sales_inventory">🛒 ۳. قابلیت‌های فروش و انبارداری</option>
                    <option value="smart">🤖 ۴. قابلیت‌های هوشمند و کاربردی</option>
                    <option value="management">⚙️ ۵. امکانات مدیریتی خود برنامه</option>
                </select>

                <!-- Dropdown 2: انتخاب عملیات زیرمجموعه -->
                <select id="actSelect" class="chip-select" onchange="onActionSelected(this.value)">
                    <option value="">🎯 ۲. ابتدا حوزه را انتخاب کنید...</option>
                </select>

                <!-- دکمه اجرای عملیات انتخابی -->
                <button class="action-run-btn" onclick="runSelectedAction()">اجرا 🚀</button>
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

    <!-- In-App Interactive Modal: Profit Wizard -->
    <div class="modal-overlay" id="profitModalOverlay" style="display: none;" onclick="if(event.target===this) closeProfitModal()">
        <div class="modal-card">
            <div class="modal-header">
                <span class="modal-title">💵 پنجره تعاملی محاسبه سود (روزانه، ماهانه، سالانه و کل)</span>
                <button class="modal-close" onclick="closeProfitModal()">&times;</button>
            </div>
            <div class="modal-body">
                <div style="display: flex; gap: 10px;">
                    <div class="modal-group" style="flex: 1;">
                        <label class="modal-label">🛒 ۱. شیت فروش (شیلد ۱):</label>
                        <select id="profitSalesSheet" class="modal-input" onchange="updateProfitPreview()"></select>
                    </div>
                    <div class="modal-group" style="flex: 1;">
                        <label class="modal-label">📦 ۲. شیت خرید (شیلد ۲ یا دیگر):</label>
                        <select id="profitPurchaseSheet" class="modal-input" onchange="updateProfitPreview()"></select>
                    </div>
                </div>

                <div style="display: flex; gap: 10px;">
                    <div class="modal-group" style="flex: 1;">
                        <label class="modal-label">📅 ستون تاریخ فاکتور فروش:</label>
                        <input type="text" id="profitDateCol" class="modal-input" value="A" placeholder="مثلاً: A یا 1" oninput="updateProfitPreview()">
                    </div>
                    <div class="modal-group" style="flex: 1;">
                        <label class="modal-label">💰 ستون مبلغ / قیمت فروش:</label>
                        <input type="text" id="profitSalesCol" class="modal-input" value="B" placeholder="مثلاً: B یا 2" oninput="updateProfitPreview()">
                    </div>
                </div>

                <div class="modal-group">
                    <label class="modal-label">🛒 ستون مبلغ خرید / هزینه (در شیت خرید):</label>
                    <input type="text" id="profitPurchaseCol" class="modal-input" value="C" placeholder="مثلاً: C یا B یا 2 یا 3" oninput="updateProfitPreview()">
                </div>

                <div class="modal-group">
                    <label class="modal-label">⏱️ دوره زمانی تفکیک سود:</label>
                    <div class="modal-radios">
                        <label class="radio-option">
                            <input type="radio" name="profitPeriodRadio" value="all" checked onchange="updateProfitPreview()">
                            <span>📊 گزارش جامع (سود کل + تفکیک روزانه + ماهانه + سالانه)</span>
                        </label>
                        <label class="radio-option">
                            <input type="radio" name="profitPeriodRadio" value="total" onchange="updateProfitPreview()">
                            <span>💵 فقط سود کل و شاخص‌های مالی</span>
                        </label>
                        <label class="radio-option">
                            <input type="radio" name="profitPeriodRadio" value="daily" onchange="updateProfitPreview()">
                            <span>📅 تفکیک سود روزانه با فرمول تفریق</span>
                        </label>
                        <label class="radio-option">
                            <input type="radio" name="profitPeriodRadio" value="monthly" onchange="updateProfitPreview()">
                            <span>🗓️ تفکیک سود ماهانه روند سودآوری</span>
                        </label>
                        <label class="radio-option">
                            <input type="radio" name="profitPeriodRadio" value="yearly" onchange="updateProfitPreview()">
                            <span>📆 تفکیک سود سالانه</span>
                        </label>
                    </div>
                </div>
                <div class="modal-preview" id="profitPreviewText">
                    👁️ پیش‌نمایش: محاسبه سود روزانه، ماهانه و کل از ستون فروش B و خرید C بر اساس تاریخ A
                </div>
            </div>
            <div class="modal-footer">
                <button class="btn-modal-cancel" onclick="closeProfitModal()">انصراف</button>
                <button class="btn-modal-submit-profit" onclick="submitProfitModal()">محاسبه سود 💵</button>
            </div>
        </div>
    </div>

    <!-- In-App Interactive Modal: Sales Wizard -->
    <div class="modal-overlay" id="salesModalOverlay" style="display: none;" onclick="if(event.target===this) closeSalesModal()">
        <div class="modal-card modal-card-sales">
            <div class="modal-header">
                <span class="modal-title modal-title-sales">💰 پنجره تعاملی محاسبه فروش (روزانه، ماهانه، سالانه و کل)</span>
                <button class="modal-close" onclick="closeSalesModal()">&times;</button>
            </div>
            <div class="modal-body">
                <div class="modal-group">
                    <label class="modal-label">📅 ستون تاریخ فاکتورها / تراکنش‌ها:</label>
                    <input type="text" id="salesDateCol" class="modal-input" value="A" placeholder="شماره یا حرف (مثلاً: A یا 1)" oninput="updateSalesPreview()">
                </div>
                <div class="modal-group">
                    <label class="modal-label">💳 ستون مبلغ / قیمت فروش:</label>
                    <input type="text" id="salesAmountCol" class="modal-input" value="B" placeholder="شماره یا حرف (مثلاً: B یا 2)" oninput="updateSalesPreview()">
                </div>
                <div class="modal-group">
                    <label class="modal-label">⏱️ دوره زمانی تفکیک فروش:</label>
                    <div class="modal-radios">
                        <label class="radio-option radio-option-sales">
                            <input type="radio" name="salesPeriodRadio" value="all" checked onchange="updateSalesPreview()">
                            <span>📊 گزارش جامع (فروش کل + تفکیک روزانه + ماهانه + سالانه)</span>
                        </label>
                        <label class="radio-option radio-option-sales">
                            <input type="radio" name="salesPeriodRadio" value="total" onchange="updateSalesPreview()">
                            <span>💰 فقط فروش کل و شاخص‌های مالی</span>
                        </label>
                        <label class="radio-option radio-option-sales">
                            <input type="radio" name="salesPeriodRadio" value="daily" onchange="updateSalesPreview()">
                            <span>📅 تفکیک فروش روزانه</span>
                        </label>
                        <label class="radio-option radio-option-sales">
                            <input type="radio" name="salesPeriodRadio" value="monthly" onchange="updateSalesPreview()">
                            <span>🗓️ تفکیک فروش ماهانه</span>
                        </label>
                        <label class="radio-option radio-option-sales">
                            <input type="radio" name="salesPeriodRadio" value="yearly" onchange="updateSalesPreview()">
                            <span>📆 تفکیک فروش سالانه</span>
                        </label>
                    </div>
                </div>
                <div class="modal-preview" id="salesPreviewText">
                    👁️ پیش‌نمایش: محاسبه فروش جامع بر اساس تاریخ ستون A و مبلغ ستون B
                </div>
            </div>
            <div class="modal-footer">
                <button class="btn-modal-cancel" onclick="closeSalesModal()">انصراف</button>
                <button class="btn-modal-submit-sales" onclick="submitSalesModal()">محاسبه فروش 📈</button>
            </div>
        </div>
    </div>

    <script>
        async function fetchStatus() {
            try {
                const res = await fetch('/api/status');
                const data = await res.json();
                window.currentSheets = data.sheets || [];
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

        const ACCOUNTING_KEYWORDS_MAP = {
            'صورت سود و زیان جامع': 'background: rgba(6, 78, 59, 0.45); color: #6EE7B7; border: 1px solid #059669;',
            'صورت سود و زیان': 'background: rgba(6, 78, 59, 0.45); color: #6EE7B7; border: 1px solid #059669;',
            'سود و زیان': 'background: rgba(6, 78, 59, 0.45); color: #6EE7B7; border: 1px solid #059669;',
            'سود ناخالص': 'background: rgba(6, 78, 59, 0.45); color: #6EE7B7; border: 1px solid #059669;',
            'سود خالص': 'background: rgba(6, 78, 59, 0.45); color: #6EE7B7; border: 1px solid #059669;',
            'سود عملیاتی': 'background: rgba(6, 78, 59, 0.45); color: #6EE7B7; border: 1px solid #059669;',
            'پرسودترین': 'background: rgba(6, 78, 59, 0.35); color: #34D399; border: 1px solid #10B981;',
            'سود': 'background: rgba(6, 78, 59, 0.4); color: #34D399; border: 1px solid #10B981;',
            'زیان': 'background: rgba(159, 18, 57, 0.4); color: #FDA4AF; border: 1px solid #E11D48;',
            'ترازنامه': 'background: rgba(120, 53, 15, 0.45); color: #FDE68A; border: 1px solid #D97706;',
            'صورت وضعیت مالی': 'background: rgba(120, 53, 15, 0.45); color: #FDE68A; border: 1px solid #D97706;',
            'تراز آزمایشی': 'background: rgba(120, 53, 15, 0.4); color: #FDE68A; border: 1px solid #D97706;',
            'تراز است': 'background: rgba(6, 78, 59, 0.45); color: #6EE7B7; border: 1px solid #059669;',
            'تراز': 'background: rgba(120, 53, 15, 0.35); color: #FDE68A; border: 1px solid #B45309;',
            'فروش خالص': 'background: rgba(4, 120, 87, 0.4); color: #A7F3D0; border: 1px solid #059669;',
            'برگشت از فروش': 'background: rgba(159, 18, 57, 0.4); color: #FECDD3; border: 1px solid #BE123C;',
            'پرفروش‌ترین': 'background: rgba(4, 120, 87, 0.35); color: #6EE7B7; border: 1px solid #047857;',
            'فروش': 'background: rgba(4, 120, 87, 0.4); color: #6EE7B7; border: 1px solid #059669;',
            'بهای تمام شده': 'background: rgba(146, 64, 14, 0.45); color: #FDE047; border: 1px solid #D97706;',
            'بهای تمام‌شده': 'background: rgba(146, 64, 14, 0.45); color: #FDE047; border: 1px solid #D97706;',
            'COGS': 'background: rgba(146, 64, 14, 0.45); color: #FDE047; border: 1px solid #D97706;',
            'برگشت از خرید': 'background: rgba(194, 65, 12, 0.4); color: #FED7AA; border: 1px solid #C2410C;',
            'خرید': 'background: rgba(194, 65, 12, 0.4); color: #FED7AA; border: 1px solid #EA580C;',
            'کاردکس انبار': 'background: rgba(91, 33, 182, 0.4); color: #DDD6FE; border: 1px solid #7C3AED;',
            'موجودی انبار': 'background: rgba(91, 33, 182, 0.4); color: #DDD6FE; border: 1px solid #7C3AED;',
            'انبارگردانی': 'background: rgba(91, 33, 182, 0.4); color: #DDD6FE; border: 1px solid #7C3AED;',
            'کاردکس': 'background: rgba(91, 33, 182, 0.4); color: #DDD6FE; border: 1px solid #7C3AED;',
            'موجودی کالا': 'background: rgba(91, 33, 182, 0.35); color: #DDD6FE; border: 1px solid #7C3AED;',
            'انبار': 'background: rgba(91, 33, 182, 0.35); color: #C4B5FD; border: 1px solid #6D28D9;',
            'حقوق و دستمزد': 'background: rgba(3, 105, 161, 0.4); color: #BAE6FD; border: 1px solid #0284C7;',
            'فیش حقوقی': 'background: rgba(3, 105, 161, 0.4); color: #BAE6FD; border: 1px solid #0284C7;',
            'حقوق': 'background: rgba(3, 105, 161, 0.35); color: #BAE6FD; border: 1px solid #0284C7;',
            'دستمزد': 'background: rgba(3, 105, 161, 0.35); color: #BAE6FD; border: 1px solid #0284C7;',
            'استهلاک انباشته': 'background: rgba(161, 98, 7, 0.4); color: #FEF08A; border: 1px solid #CA8A04;',
            'استهلاک': 'background: rgba(161, 98, 7, 0.4); color: #FEF08A; border: 1px solid #CA8A04;',
            'دارایی ثابت': 'background: rgba(15, 118, 110, 0.4); color: #99F6E4; border: 1px solid #0D9488;',
            'دارایی': 'background: rgba(15, 118, 110, 0.35); color: #99F6E4; border: 1px solid #0D9488;',
            'بدهی': 'background: rgba(185, 28, 28, 0.35); color: #FECACA; border: 1px solid #DC2626;',
            'سرمایه': 'background: rgba(180, 83, 9, 0.4); color: #FDE68A; border: 1px solid #D97706;',
            'صورت مغایرت': 'background: rgba(2, 132, 199, 0.4); color: #7DD3FC; border: 1px solid #0284C7;',
            'مغایرت بانکی': 'background: rgba(2, 132, 199, 0.4); color: #7DD3FC; border: 1px solid #0284C7;',
            'بانک': 'background: rgba(2, 132, 199, 0.35); color: #7DD3FC; border: 1px solid #0284C7;',
            'ارزش افزوده': 'background: rgba(30, 64, 175, 0.4); color: #BFDBFE; border: 1px solid #2563EB;',
            'مالیات': 'background: rgba(30, 64, 175, 0.4); color: #BFDBFE; border: 1px solid #2563EB;',
            'بیمه': 'background: rgba(6, 95, 70, 0.4); color: #A7F3D0; border: 1px solid #059669;'
        };

        function formatAccountingText(text) {
            if (!text) return "";
            let escaped = String(text)
                .replace(/&/g, "&amp;")
                .replace(/</g, "&lt;")
                .replace(/>/g, "&gt;");

            const sortedKeys = Object.keys(ACCOUNTING_KEYWORDS_MAP).sort((a, b) => b.length - a.length);
            const pat = new RegExp('(?<![<\\\\w/])(' + sortedKeys.map(k => k.replace(/[-\\/\\\\^$*+?.()|[\\]{}]/g, '\\\\$&')).join('|') + ')(?![>\\\\w])', 'g');

            escaped = escaped.replace(pat, (match) => {
                const style = ACCOUNTING_KEYWORDS_MAP[match];
                return `<span class="acc-pill" style="${style} padding: 1px 6px; border-radius: 4px; font-weight: bold; margin: 0 2px; display: inline-block;">${match}</span>`;
            });

            escaped = escaped.replace(/\\*\\*(.+?)\\*\\*/g, '<b>$1</b>');
            escaped = escaped.replace(/`(.+?)`/g, '<code style="background: rgba(6, 78, 59, 0.5); color: #FCD34D; border: 1px solid #D97706; padding: 2px 6px; border-radius: 4px; font-family: Consolas, monospace;">$1</code>');
            escaped = escaped.replace(/\\n/g, '<br>');
            return escaped;
        }

        function appendMessage(sender, text, isConfirm = false) {
            const feed = document.getElementById('chatFeed');
            const bubble = document.createElement('div');
            bubble.className = 'message-bubble ' + (sender === 'user' ? 'message-user' : (isConfirm ? 'message-confirm' : 'message-ai'));

            const tag = document.createElement('div');
            tag.className = 'sender-tag';
            tag.innerText = sender === 'user' ? '👤 MUYUM' : '🤖 دستیار هوشمند Microsoft Excel';
            bubble.appendChild(tag);

            const content = document.createElement('div');
            content.innerHTML = formatAccountingText(text);
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
            const mode = confirm("مدیریت یکتاسازی ستون اکسل:\n\nبرای [اعمال قفل یکتایی] روی OK کلیک کنید.\nبرای [لغو قفل یکتایی] روی Cancel کلیک کنید.");
            if (mode) {
                const col = prompt("📌 ستون مورد نظر برای اعمال قفل یکتایی (مثلاً: L یا 12):", "L");
                if (!col || !col.trim()) return;
                sendCommand(`ستون ${col.trim()} نباید داده تکراری داشته باشه`);
            } else {
                const col = prompt("📌 ستون مورد نظر برای لغو و برداشتن قفل یکتایی (مثلاً: L یا 12):", "L");
                if (!col || !col.trim()) return;
                sendCommand(`قفل یکتایی ستون ${col.trim()} را بردار`);
            }
        }

        function promptRemoveUniqueValidation() {
            const col = prompt("📌 ستون مورد نظر برای لغو و برداشتن قفل یکتایی (مثلاً: L یا 12):", "L");
            if (!col || !col.trim()) return;
            const cmd = `قفل یکتایی ستون ${col.trim()} را بردار`;
            sendCommand(cmd);
        }

        function promptSalesSummary() {
            openSalesModal();
        }

        function promptProfitCalculation() {
            openProfitModal();
        }

        function openProfitModal() {
            const sSelect = document.getElementById('profitSalesSheet');
            const pSelect = document.getElementById('profitPurchaseSheet');
            if (sSelect && pSelect && window.currentSheets) {
                sSelect.innerHTML = '';
                pSelect.innerHTML = '';
                window.currentSheets.forEach((s, idx) => {
                    const opt1 = document.createElement('option');
                    opt1.value = s;
                    opt1.innerText = s;
                    sSelect.appendChild(opt1);

                    const opt2 = document.createElement('option');
                    opt2.value = s;
                    opt2.innerText = s;
                    if (idx === 1 || (idx === 0 && window.currentSheets.length === 1)) opt2.selected = true;
                    pSelect.appendChild(opt2);
                });
            }
            document.getElementById('profitModalOverlay').style.display = 'flex';
            updateProfitPreview();
            setTimeout(() => document.getElementById('profitDateCol').focus(), 50);
        }

        function closeProfitModal() {
            document.getElementById('profitModalOverlay').style.display = 'none';
        }

        function updateProfitPreview() {
            const sSheet = (document.getElementById('profitSalesSheet') ? document.getElementById('profitSalesSheet').value : 'Sheet1') || 'Sheet1';
            const pSheet = (document.getElementById('profitPurchaseSheet') ? document.getElementById('profitPurchaseSheet').value : 'Sheet2') || 'Sheet2';
            const d = (document.getElementById('profitDateCol').value || 'A').trim().toUpperCase();
            const s = (document.getElementById('profitSalesCol').value || 'B').trim().toUpperCase();
            const p = (document.getElementById('profitPurchaseCol').value || 'C').trim().toUpperCase();
            const periodRadios = document.getElementsByName('profitPeriodRadio');
            let period = 'all';
            for (let r of periodRadios) { if (r.checked) { period = r.value; break; } }
            const periodMap = {all: "روزانه، ماهانه، سالانه و کل", daily: "روزانه", monthly: "ماهانه", yearly: "سالانه", total: "کل"};
            const pText = periodMap[period] || "روزانه، ماهانه و کل";
            const formulaText = `سود = (${sSheet}!${s}) - (${pSheet}!${p})`;
            document.getElementById('profitPreviewText').innerText = `👁️ پیش‌نمایش: محاسبه سود (${pText}) با فرمول «${formulaText}» بر اساس تاریخ ستون ${d} و ایجاد شیت «گزارش_سود»`;
        }

        function submitProfitModal() {
            const sSheet = (document.getElementById('profitSalesSheet') ? document.getElementById('profitSalesSheet').value : 'Sheet1') || 'Sheet1';
            const pSheet = (document.getElementById('profitPurchaseSheet') ? document.getElementById('profitPurchaseSheet').value : 'Sheet2') || 'Sheet2';
            const d = (document.getElementById('profitDateCol').value || 'A').trim().toUpperCase();
            const s = (document.getElementById('profitSalesCol').value || 'B').trim().toUpperCase();
            const p = (document.getElementById('profitPurchaseCol').value || 'C').trim().toUpperCase();
            const periodRadios = document.getElementsByName('profitPeriodRadio');
            let period = 'all';
            for (let r of periodRadios) { if (r.checked) { period = r.value; break; } }
            const periodMap = {all: "روزانه و ماهانه و سالانه و کل", daily: "روزانه", monthly: "ماهانه", yearly: "سالانه", total: "کل"};
            const pText = periodMap[period] || "روزانه و ماهانه و سالانه و کل";

            closeProfitModal();
            const cmd = `محاسبه سود ${pText} ستون فروش ${s} در شیت ${sSheet} و ستون خرید ${p} در شیت ${pSheet} بر اساس تاریخ ستون ${d}`;
            sendCommand(cmd);
        }

        function openSalesModal() {
            document.getElementById('salesModalOverlay').style.display = 'flex';
            updateSalesPreview();
            setTimeout(() => document.getElementById('salesDateCol').focus(), 50);
        }

        function closeSalesModal() {
            document.getElementById('salesModalOverlay').style.display = 'none';
        }

        function updateSalesPreview() {
            const d = (document.getElementById('salesDateCol').value || 'A').trim().toUpperCase();
            const a = (document.getElementById('salesAmountCol').value || 'B').trim().toUpperCase();
            const periodRadios = document.getElementsByName('salesPeriodRadio');
            let period = 'all';
            for (let r of periodRadios) { if (r.checked) { period = r.value; break; } }
            const periodMap = {all: "روزانه، ماهانه، سالانه و کل", daily: "روزانه", monthly: "ماهانه", yearly: "سالانه", total: "کل"};
            const pText = periodMap[period] || "روزانه، ماهانه و کل";
            document.getElementById('salesPreviewText').innerText = `👁️ پیش‌نمایش: محاسبه فروش (${pText}) ستون مبلغ ${a} بر اساس تاریخ ستون ${d} و ایجاد شیت «گزارش_فروش»`;
        }

        function submitSalesModal() {
            const d = (document.getElementById('salesDateCol').value || 'A').trim().toUpperCase();
            const a = (document.getElementById('salesAmountCol').value || 'B').trim().toUpperCase();
            const periodRadios = document.getElementsByName('salesPeriodRadio');
            let period = 'all';
            for (let r of periodRadios) { if (r.checked) { period = r.value; break; } }
            const periodMap = {all: "روزانه و ماهانه و سالانه و کل", daily: "روزانه", monthly: "ماهانه", yearly: "سالانه", total: "کل"};
            const pText = periodMap[period] || "روزانه و ماهانه و سالانه و کل";

            closeSalesModal();
            const cmd = `محاسبه فروش ${pText} ستون فروش ${a} بر اساس تاریخ ${d}`;
            sendCommand(cmd);
        }

        // Global Keyboard Shortcuts (میانبرهای کیبورد)
        document.addEventListener('keydown', function(e) {
            if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'p') {
                e.preventDefault();
                openProfitModal();
            }
            if ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key.toLowerCase() === 's') {
                e.preventDefault();
                openSalesModal();
            }
            if (e.key === 'Escape') {
                closeProfitModal();
                closeSalesModal();
            }
        });

        const FIVE_CATEGORIES = {
            "essential": {
                "title": "⚡ ۱. قابلیت‌های ضروری",
                "commands": [
                    {label: "🔍 فیلتر هوشمند اطلاعات", cmd: "فیلتر هوشمند جدول"},
                    {label: "🔄 مرتب‌سازی صعودی داده‌ها", cmd: "مرتب‌سازی صعودی ستون A"},
                    {label: "🔽 مرتب‌سازی نزولی داده‌ها", cmd: "مرتب‌سازی نزولی ستون A"},
                    {label: "🗑️ حذف موارد تکراری", cmd: "حذف داده‌های تکراری"},
                    {label: "🧹 حذف سطرهای خالی", cmd: "ردیف‌های خالی را حذف کن"},
                    {label: "➕ جمع خودکار (AutoSum)", cmd: "جمع خودکار ستون A"},
                    {label: "📊 محاسبه میانگین داده‌ها", cmd: "میانگین ستون A"},
                    {label: "🔺 بیشترین و کمترین (MAX / MIN)", cmd: "بیشترین مقدار ستون A"},
                    {label: "🔢 شمارش اطلاعات و سلول‌های پر", cmd: "شمارش سلول‌های پر ستون A"},
                    {label: "🔢 تبدیل متن به عدد (اصلاح داده‌ها)", cmd: "تبدیل متن به عدد در ستون A"},
                    {label: "🧮 ۴ عمل اصلی ستون‌ها...", cmd: "۴ عمل اصلی"},
                    {label: "📅 درج تاریخ در شیت...", cmd: "درج تاریخ در شیت"}
                ]
            },
            "accounting": {
                "title": "💼 ۲. قابلیت‌های حسابداری",
                "commands": [
                    {label: "💵 محاسبه سود (روزانه/ماهانه/سالانه/کل)...", cmd: "محاسبه سود"},
                    {label: "💰 محاسبه فروش (روزانه/ماهانه/سالانه/کل)...", cmd: "محاسبه فروش"},
                    {label: "🏷️ محاسبه تخفیف کالاها...", cmd: "محاسبه تخفیف"},
                    {label: "⚖️ مغایرت‌گیری دو ستون (Reconciliation)", cmd: "مغایرت‌گیری ستون A و B"},
                    {label: "📊 تهیه تراز آزمایشی چهار ستونه", cmd: "تهیه تراز آزمایشی"},
                    {label: "💳 محاسبه مانده حساب جاری (بدهکار/بستانکار)", cmd: "محاسبه مانده حساب"},
                    {label: "💼 لیست حقوق، اضافه‌کاری، بیمه و مالیات", cmd: "محاسبه حقوق و دستمزد"},
                    {label: "📉 محاسبه جدول استهلاک دارایی‌ها", cmd: "محاسبه استهلاک"},
                    {label: "📈 تهیه صورت سود و زیان جامع دوره", cmd: "صورت سود و زیان جامع"},
                    {label: "🏛️ تهیه ترازنامه مالی (بیلان)", cmd: "ترازنامه"}
                ]
            },
            "sales_inventory": {
                "title": "🛒 ۳. قابلیت‌های فروش و انبارداری",
                "commands": [
                    {label: "🏷️ محاسبه تخفیف کالاها...", cmd: "محاسبه تخفیف"},
                    {label: "🧾 صدور و ثبت فاکتور رسمی فروش", cmd: "ثبت فاکتور فروش"},
                    {label: "🛒 صدور و ثبت فاکتور رسمی خرید", cmd: "ثبت فاکتور خرید"},
                    {label: "📦 کنترل موجودی و هشدار کمبود انبار", cmd: "موجودی انبار و هشدار کمبود موجودی"},
                    {label: "📊 ارزش‌گذاری انبار (فایفو، لایفو و میانگین)", cmd: "ارزش‌گذاری انبار به روش‌های فایفو، لایفو و میانگین موزون"},
                    {label: "💰 محاسبه ارزش ریالی کل موجودی انبار", cmd: "ارزش کل موجودی انبار"},
                    {label: "👥 محاسبه بدهی و طلب مشتریان", cmd: "محاسبه بدهی و طلب مشتریان"},
                    {label: "🏆 گزارش پرفروش‌ترین کالاها", cmd: "گزارش پرفروش‌ترین کالاها"},
                    {label: "🏷️ افزایش ۱۵ درصدی قیمت‌های ستون C", cmd: "قیمت‌های ستون C را ۱۵ درصد افزایش بده"}
                ]
            },
            "smart": {
                "title": "🤖 ۴. قابلیت‌های هوشمند و کاربردی",
                "commands": [
                    {label: "🔒 مدیریت و قفل یکتاسازی ستون...", cmd: "قفل یکتایی ستون"},
                    {label: "🔓 لغو و برداشتن قفل یکتایی ستون", cmd: "قفل یکتایی ستون A را بردار"},
                    {label: "📊 تحلیل هوشمند شیت و داده‌ها", cmd: "تحلیل ساختار فایل و شیت"},
                    {label: "🩺 چکاپ عمیق سلامت و خطاهای فایل", cmd: "سلامت فایل و بررسی خطاهای اکسل"},
                    {label: "🧮 میانگین سطری بین ستون‌ها (B و C در D)", cmd: "از ستون B و C میانگین بگیر و نتیجه را در D بنویس"},
                    {label: "📊 گزارش‌ساز خودکار با نمودار گرافیکی", cmd: "رسم نمودار ستونی"},
                    {label: "📋 تشخیص و ساخت جدول اکسل", cmd: "یک جدول بساز"},
                    {label: "✂️ تفکیک نام و نام خانوادگی", cmd: "تفکیک نام و نام خانوادگی در ستون A"},
                    {label: "🔗 ادغام سلول‌ها (Merge)", cmd: "سلول‌های A1:B1 را ادغام کن"}
                ]
            },
            "management": {
                "title": "⚙️ ۵. امکانات مدیریتی خود برنامه",
                "commands": [
                    {label: "🔒 قفل و مدیریت یکتایی ستون...", cmd: "قفل یکتایی ستون"},
                    {label: "🔓 لغو و برداشتن قفل یکتایی ستون", cmd: "قفل یکتایی ستون A را بردار"},
                    {label: "📅 درج تاریخ در شیت (شمسی/میلادی)...", cmd: "درج تاریخ در شیت"},
                    {label: "📊 تحلیل هوشمند شیت و داده‌ها", cmd: "تحلیل ساختار فایل و شیت"},
                    {label: "💾 ذخیره تغییرات فایل اکسل (Ctrl+S)", cmd: "ذخیره فایل"},
                    {label: "💾 ذخیره نسخه پشتیبان (Backup)", cmd: "ذخیره نسخه پشتیبان"},
                    {label: "📄 تولید خروجی PDF استاندارد", cmd: "خروجی PDF از شیت"},
                    {label: "↩️ لغو آخرین عملیات (Undo چندمرحله‌ای)", cmd: "لغو آخرین عملیات"},
                    {label: "📐 تنظیم خودکار عرض ستون‌ها (AutoFit)", cmd: "عرض ستون‌ها را خودکار تنظیم کن"},
                    {label: "🎨 رنگ‌آمیزی متون تکراری", cmd: "داده‌های متنی تکراری را رنگ کن"},
                    {label: "💲 تغییر فرمت به واحد پول (تومان)", cmd: "فرمت ستون B را به تومان تغییر بده"},
                    {label: "📈 تغییر فرمت به درصد (%)", cmd: "فرمت ستون C را درصدی کن"},
                    {label: "📅 تغییر فرمت به تاریخ", cmd: "فرمت ستون A را به تاریخ تغییر بده"}
                ]
            }
        };

        function onCategoryChanged(catKey) {
            const actSelect = document.getElementById('actSelect');
            if (catKey && FIVE_CATEGORIES[catKey]) {
                actSelect.innerHTML = '<option value="">🎯 ۲. انتخاب عملیات...</option>';
                FIVE_CATEGORIES[catKey].commands.forEach(item => {
                    const opt = document.createElement('option');
                    opt.value = item.cmd;
                    opt.innerText = item.label;
                    actSelect.appendChild(opt);
                });
                setTimeout(() => {
                    try {
                        if (actSelect.showPicker) actSelect.showPicker();
                        else actSelect.focus();
                    } catch (e) {
                        actSelect.focus();
                    }
                }, 100);
            } else {
                actSelect.innerHTML = '<option value="">🎯 ۲. ابتدا حوزه را انتخاب کنید...</option>';
            }
        }

        function onActionSelected(cmd) {
            if (cmd) {
                document.getElementById('commandInput').value = cmd;
                document.getElementById('commandInput').focus();
            }
        }

        function runSelectedAction() {
            const actSelect = document.getElementById('actSelect');
            const cmd = actSelect.value || document.getElementById('commandInput').value.trim();
            if (!cmd) {
                appendMessage('ai', '⚠️ لطفاً ابتدا یک عملیات را از منوی کشویی انتخاب کنید یا دستور خود را بنویسید.');
                return;
            }
            if (cmd === 'قفل یکتایی ستون') {
                promptPreventDuplicates();
            } else if (cmd.includes('قفل یکتایی') && (cmd.includes('بردار') || cmd.includes('لغو') || cmd.includes('حذف'))) {
                promptRemoveUniqueValidation();
            } else if (cmd === 'محاسبه سود') {
                openProfitModal();
            } else if (cmd === 'محاسبه فروش') {
                openSalesModal();
            } else if (cmd.includes('تخفیف')) {
                promptDiscount();
            } else if (cmd.includes('۴ عمل') || cmd.includes('4 عمل')) {
                promptArithmetic('add');
            } else if (cmd === 'درج تاریخ در شیت') {
                promptInsertDate();
            } else {
                sendCommand(cmd);
            }
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

        async function saveWorkbook() {
            appendMessage('user', '💾 ذخیره فایل اکسل');
            addHistoryItem('ذخیره فایل اکسل');
            try {
                const res = await fetch('/api/save', { method: 'POST' });
                const data = await res.json();
                if (data.success) {
                    appendMessage('ai', '💾 ' + data.message);
                } else {
                    appendMessage('ai', '❌ خطا در ذخیره‌سازی: ' + (data.error || 'خطای نامشخص'));
                }
            } catch(e) {
                appendMessage('ai', '❌ خطا در ارتباط با سرور جهت ذخیره: ' + e);
            }
        }

        window.addEventListener('keydown', (e) => {
            if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's' && !e.shiftKey) {
                e.preventDefault();
                saveWorkbook();
            }
        });

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

        if path == "/api/save":
            try:
                saved_path = self.controller.manager.save()
                resp = {
                    "success": True,
                    "message": f"فایل اکسل با موفقیت ذخیره شد: {saved_path}"
                }
            except Exception as e:
                resp = {
                    "success": False,
                    "error": str(e)
                }
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.end_headers()
            self.wfile.write(json.dumps(resp, ensure_ascii=False).encode("utf-8"))
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
        try:
            if hasattr(sheet, '_ws') and hasattr(sheet._ws, 'title'):
                sheet._ws.title = "فروش"
        except Exception:
            pass

        # Sheet 1: Sales with dates
        sheet.set_value("A1", "تاریخ")
        sheet.set_value("B1", "کالا")
        sheet.set_value("C1", "تعداد فروش")
        sheet.set_value("D1", "قیمت فروش")
        sheet.set_value("E1", "فروش کل")

        sheet.set_value("A2", "1403/08/20")
        sheet.set_value("B2", "لپ‌تاپ")
        sheet.set_value("C2", 5)
        sheet.set_value("D2", 1200)

        sheet.set_value("A3", "1403/08/21")
        sheet.set_value("B3", "موبایل")
        sheet.set_value("C3", 10)
        sheet.set_value("D3", 800)

        sheet.set_value("A4", "1403/08/22")
        sheet.set_value("B4", "تبلت")
        sheet.set_value("C4", 7)
        sheet.set_value("D4", 450)

        sheet.set_value("A5", "1403/09/01")
        sheet.set_value("B5", "موبایل")
        sheet.set_value("C5", 10)
        sheet.set_value("D5", 800)

        sheet.set_value("A6", "1403/09/02")
        sheet.set_value("B6", "مانیتور")
        sheet.set_value("C6", 3)
        sheet.set_value("D6", 300)

        # Create Sheet 2 for Purchases if not present
        if len(ctrl.manager.active_driver.get_sheet_names()) < 2:
            sheet2 = ctrl.manager.active_driver.create_sheet("خرید")
            sheet2.set_value("A1", "کالا")
            sheet2.set_value("B1", "تعداد خرید")
            sheet2.set_value("C1", "قیمت خرید")
            sheet2.set_value("D1", "خرید کل")

            sheet2.set_value("A2", "لپ‌تاپ")
            sheet2.set_value("B2", 10)
            sheet2.set_value("C2", 950)

            sheet2.set_value("A3", "موبایل")
            sheet2.set_value("B3", 25)
            sheet2.set_value("C3", 620)

            sheet2.set_value("A4", "تبلت")
            sheet2.set_value("B4", 10)
            sheet2.set_value("C4", 350)

            sheet2.set_value("A5", "مانیتور")
            sheet2.set_value("B5", 5)
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
    start_web_server("0.0.0.0", 8080)

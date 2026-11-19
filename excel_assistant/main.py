"""
Main entry point for Excel AI Assistant.
Supports launching PySide6 modern GUI, responsive Web interface, or interactive CLI.
"""
import argparse
import os
import sys


def main():
    parser = argparse.ArgumentParser(description="دستیار هوشمند اکسل - Excel AI Assistant")
    parser.add_argument("--cli", action="store_true", help="اجرا در حالت خط فرمان (CLI)")
    parser.add_argument("--gui", action="store_true", help="اجرا در حالت رابط گرافیکی دسکتاپ (PySide6 GUI)")
    parser.add_argument("--web", action="store_true", help="اجرا در حالت وب‌سرویس و رابط کاربری تحت وب")
    parser.add_argument("--port", type=int, default=8080, help="پورت وب‌سرور (پیش‌فرض 8080)")
    parser.add_argument("--file", type=str, default=None, help="مسیر فایل اکسل")

    args = parser.parse_args()

    has_display = bool(os.getenv("DISPLAY")) or sys.platform == "win32"

    if args.web:
        from excel_assistant.ui.web_server import start_web_server
        start_web_server(host="0.0.0.0", port=args.port, workbook_path=args.file)
    elif args.cli:
        from excel_assistant.cli import run_cli
        run_cli(args.file)
    elif args.gui or has_display:
        try:
            from excel_assistant.ui.app import run_gui
            run_gui(args.file)
        except Exception as e:
            print(f"⚠️ نمایشگر گرافیکی دسکتاپ در دسترس نیست ({e}). در حال راه‌اندازی رابط کاربری تحت وب...")
            from excel_assistant.ui.web_server import start_web_server
            start_web_server(host="0.0.0.0", port=args.port, workbook_path=args.file)
    else:
        from excel_assistant.ui.web_server import start_web_server
        start_web_server(host="0.0.0.0", port=args.port, workbook_path=args.file)


if __name__ == "__main__":
    main()

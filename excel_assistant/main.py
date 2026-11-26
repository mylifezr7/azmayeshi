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
    parser.add_argument("--no-excel", action="store_true", help="اجرا به صورت مستقل بدون باز کردن نرم‌افزار Microsoft Excel")
    parser.add_argument("--file", type=str, default=None, help="مسیر فایل اکسل")
    parser.add_argument("file_arg", nargs="?", default=None, help="مسیر فایل اکسل")

    args = parser.parse_args()
    target_file = args.file or args.file_arg

    if not target_file:
        for candidate in ["Challenge_5_Accounting.xlsx", "Challenge_4_Accounting.xlsx", "Challenge_3_Accounting.xlsx"]:
            if os.path.exists(candidate):
                target_file = candidate
                break

    pref_driver = "openpyxl" if args.no_excel else None

    has_display = bool(os.getenv("DISPLAY")) or sys.platform == "win32"

    if args.web:
        from excel_assistant.ui.web_server import start_web_server
        start_web_server(host="0.0.0.0", port=args.port, workbook_path=target_file, preferred_driver=pref_driver)
    elif args.cli:
        from excel_assistant.cli import run_cli
        run_cli(target_file, preferred_driver=pref_driver)
    elif args.gui or has_display:
        gui_ok = False
        try:
            from excel_assistant.ui.app import run_gui
            res = run_gui(target_file)
            if res == 0:
                gui_ok = True
        except Exception as e:
            print(f"\n⚠️ خطا در اجرای رابط گرافیکی دسکتاپ: {e}")

        if not gui_ok:
            if args.gui:
                print("\n❌ باز شدن پنجره دسکتاپ با خطا مواجه شد.")
                print("💡 برای نصب یا تعمیر وابستگی گرافیکی، دستور زیر را در CMD اجرا کنید:")
                print("   python -m pip install PySide6\n")
                sys.exit(1)
            print("🚀 در حال راه‌اندازی و باز کردن خودکار برنامه در مرورگر (Web UI)...")
            from excel_assistant.ui.web_server import start_web_server
            start_web_server(host="0.0.0.0", port=args.port, workbook_path=target_file, preferred_driver=pref_driver)
    else:
        from excel_assistant.ui.web_server import start_web_server
        start_web_server(host="0.0.0.0", port=args.port, workbook_path=target_file, preferred_driver=pref_driver)


if __name__ == "__main__":
    main()

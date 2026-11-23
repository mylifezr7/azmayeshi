"""
Quick entry point for Excel AI Assistant - MIYUM.
Usage:
    python run.py             # Runs Desktop GUI (or Web UI if in headless environment)
    python run.py --gui       # Runs Desktop PySide6 GUI
    python run.py --web       # Runs Web Live Preview on port 8080
    python run.py --cli       # Runs Persian Terminal CLI
"""
import sys
import os

def ensure_windows_excel_dependencies():
    """On Windows, ensures xlwings is installed for live Microsoft Excel automation."""
    if sys.platform == "win32":
        try:
            import xlwings
        except ImportError:
            print("\n" + "=" * 60)
            print("⏳ در حال نصب خودکار پکیج xlwings برای باز شدن هم‌زمان نرم‌افزار اکسل...")
            print("=" * 60)
            try:
                import subprocess
                subprocess.check_call([sys.executable, "-m", "pip", "install", "xlwings"])
                print("✅ پکیج xlwings با موفقیت نصب شد! در حال اجرای برنامه و باز کردن اکسل...\n")
            except Exception as e:
                print(f"⚠️ نصب خودکار xlwings با خطا مواجه شد: {e}")
                print("💡 برای باز شدن خودکار اکسل، دستور زیر را در CMD اجرا کنید:")
                print("pip install xlwings\n")


if __name__ == "__main__":
    ensure_windows_excel_dependencies()

    from excel_assistant.main import main

    if len(sys.argv) == 1:
        if sys.platform != "win32" and not os.environ.get("DISPLAY") and not os.environ.get("WAYLAND_DISPLAY"):
            sys.argv.append("--web")
        else:
            sys.argv.append("--gui")

    main()

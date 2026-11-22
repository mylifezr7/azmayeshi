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

if __name__ == "__main__":
    from excel_assistant.main import main

    if len(sys.argv) == 1:
        if sys.platform != "win32" and not os.environ.get("DISPLAY") and not os.environ.get("WAYLAND_DISPLAY"):
            sys.argv.append("--web")
        else:
            sys.argv.append("--gui")

    main()

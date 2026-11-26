#!/bin/bash
# MUYUM Excel Assistant - Linux/macOS GUI Launcher
set -e
cd "$(dirname "$0")"
export PYTHONPATH=".:$PYTHONPATH"

# Pull latest updates if in git repo
if command -v git >/dev/null 2>&1; then
    echo "🔄 دریافت آخرین آپدیت MUYUM..."
    git pull origin arena/814f030e-azmayeshi 2>/dev/null || true
fi

# Detect Python
PYTHON_CMD="python3"
if [ -f ".venv/bin/python" ]; then
    PYTHON_CMD=".venv/bin/python"
elif [ -f "venv/bin/python" ]; then
    PYTHON_CMD="venv/bin/python"
fi

$PYTHON_CMD -m excel_assistant.main --gui "$@"

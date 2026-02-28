#!/bin/bash
# ===========================================
# NIPT Report Generator - Linux/Mac Launcher
# ===========================================
set -e

echo "==========================================="
echo "  NIPT Professional Reporting Suite"
echo "==========================================="
echo ""

DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$DIR"

# Create venv if missing
if [ ! -d ".venv" ]; then
    echo "📦 Creating virtual environment..."
    python3 -m venv .venv
fi

# Activate venv (same as PGT-A pattern)
source .venv/bin/activate

# Install deps if needed
if ! python3 -c "import PyQt6" 2>/dev/null; then
    echo "📥 Installing dependencies..."
    pip install -r requirements.txt
fi

show_error() {
    EXIT_CODE=$?
    if [ $EXIT_CODE -ne 0 ] && [ $EXIT_CODE -ne 130 ]; then
        echo ""
        echo "❌ Application crashed (exit code: $EXIT_CODE)"
        if [ -s "run_error.log" ]; then
            echo "Error details:"
            echo "-------------------------------------------"
            cat run_error.log
            echo "-------------------------------------------"
        fi
        echo ""
        echo "Common fixes:"
        echo "  1. pip install -r requirements.txt --upgrade"
        echo "  2. rm -rf .venv && bash launch_nipt.sh"
        read -p "Press enter to exit"
    fi
}
trap show_error EXIT

echo "🚀 Starting NIPT Report Generator..."
echo "==========================================="
echo ""

# Use system python3 (activated via venv source) - same as PGT-A
python3 nipt_report_generator.py 2>run_error.log

exit $?

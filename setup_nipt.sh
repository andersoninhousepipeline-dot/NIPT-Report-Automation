#!/bin/bash
# NIPT Report Generator - Linux Setup

# Get script directory
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$DIR"

echo "==========================================="
echo "  NIPT Report Generator - SETUP (Linux)"
echo "==========================================="
echo ""

# Check for Python
if ! command -v python3 &> /dev/null
then
    echo "[ERROR] Python3 not found. Please install it: sudo apt install python3 python3-venv"
    exit 1
fi

echo "[OK] Python3 found"
echo ""

# Create venv
echo "[*] Creating virtual environment..."
python3 -m venv .venv
if [ $? -ne 0 ]; then
    echo "[ERROR] Failed to create virtual environment"
    exit 1
fi
echo "[OK] Created"
echo ""

# Activate and install
echo "[*] Installing dependencies..."
source .venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
if [ $? -ne 0 ]; then
    echo "[ERROR] Failed to install dependencies"
    exit 1
fi

echo "[OK] Setup complete!"
echo ""

#!/bin/bash
# DWF Ireland — Document Generator — Start Script

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

echo ""
echo "══════════════════════════════════════════"
echo "  DWF Ireland — Document Generator"
echo "══════════════════════════════════════════"

# Check Python 3
if ! command -v python3 &>/dev/null; then
  echo "  ERROR: Python 3 is not installed."
  exit 1
fi

# Create virtual environment if it doesn't exist
if [ ! -d "venv" ]; then
  echo "  Setting up environment (first run only)..."
  python3 -m venv venv
  source venv/bin/activate
  pip install -q --upgrade pip
  pip install -q -r requirements.txt
  echo "  Setup complete."
else
  source venv/bin/activate
fi

echo ""
echo "  Open your browser at:  http://localhost:5050"
echo "  Press Ctrl+C to stop."
echo ""

python3 app.py

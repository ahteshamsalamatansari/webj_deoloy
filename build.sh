#!/usr/bin/env bash
# ──────────────────────────────────────────────────────────────────
#  Render.com build script
#  Runs once when a new deploy is triggered (GitHub push).
#  Installs Python packages AND Playwright's Chromium browser
#  with all required OS-level dependencies.
# ──────────────────────────────────────────────────────────────────
set -e   # exit immediately on any error

echo "──────────────────────────────────────────"
echo "  📦  Installing Python dependencies..."
echo "──────────────────────────────────────────"
pip install --upgrade pip
pip install -r requirements.txt

echo ""
echo "──────────────────────────────────────────"
echo "  🎭  Installing Playwright Chromium..."
echo "  (includes all OS-level dependencies)"
echo "──────────────────────────────────────────"
playwright install --with-deps chromium

echo ""
echo "✅  Build complete."

#!/usr/bin/env python3
"""
gui_main.py
===========
Entry point for the NVISH Sales Forecast Automation desktop app.

Run from source:
    python gui_main.py

Build a distributable Windows .exe:
    pyinstaller SalesForecastGUI.spec

This file intentionally contains no business logic -- it only makes
sure the project root is importable and then hands off to gui.app.
"""
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


def main() -> None:
    from gui.app import main as gui_main

    gui_main()


if __name__ == "__main__":
    main()

"""
gui/branding.py
================
Central place for NVISH visual identity: colours, fonts, and asset path
resolution. Keeping this separate from app.py makes it trivial to
re-theme the app or swap in a real corporate logo later without
touching any layout code.
"""
from __future__ import annotations

import sys
from pathlib import Path

# --------------------------------------------------------------------------
# Palette (matches gui/assets/generate_assets.py)
# --------------------------------------------------------------------------
NAVY = "#0B2A47"
NAVY_DARK = "#071C30"
TEAL = "#00A896"
TEAL_DARK = "#00857A"
LIGHT_BG = "#F4F6F8"
CARD_BG = "#FFFFFF"
BORDER = "#DDE3E8"
TEXT_PRIMARY = "#1B2733"
TEXT_MUTED = "#5B6B79"
SUCCESS = "#1E8E5A"
WARNING = "#B8860B"
DANGER = "#C0392B"

APP_TITLE = "NVISH Sales Forecast Automation"
COMPANY_NAME = "NVISH Solutions Inc."
APP_SUBTITLE = "Sales & Forecast Summary Generator"
FOOTER_TEXT = "© NVISH Solutions Inc. — Internal Tool"

FONT_FAMILY = "Segoe UI"  # graceful fallback to a default sans-serif on non-Windows


def resource_path(relative: str) -> Path:
    """Resolve a bundled asset's path, whether running from source or
    from a PyInstaller-frozen executable (onefile mode extracts data
    files to a temporary `sys._MEIPASS` directory at runtime)."""
    base = getattr(sys, "_MEIPASS", None)
    if base:
        return Path(base) / relative
    return Path(__file__).resolve().parent / relative


def logo_path() -> Path:
    return resource_path("assets/nvish_logo.png")


def icon_path() -> Path:
    return resource_path("assets/nvish_icon.ico")

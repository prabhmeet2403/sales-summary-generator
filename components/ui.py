"""
components/ui.py
=================
Reusable, presentation-only building blocks for the Streamlit front
end: cards, badges, KPI/stat tiles, the top bar, page headers, empty
states, and the live processing checklist. Adapted from the SFAE v2
design system (same CSS classes, same visual language) with all
forecasting/staffing/placement-specific pieces removed.

Nothing here computes a single number that ends up in the Sales
Summary -- every function only renders values it is handed.
"""
from __future__ import annotations

import base64
import os
from typing import List, Optional, Tuple

import streamlit as st

_HERE = os.path.dirname(__file__)
_CSS_PATH = os.path.join(_HERE, "..", "styles", "enterprise.css")
_LOGO_PATH = os.path.join(_HERE, "..", "assets", "nvish_logo.png")


def load_css() -> None:
    """Load enterprise.css with explicit UTF-8 encoding (avoids a
    Windows cp1252 decode error some environments hit otherwise)."""
    with open(_CSS_PATH, encoding="utf-8") as f:
        st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)


@st.cache_data(show_spinner=False)
def _logo_base64() -> Optional[str]:
    try:
        with open(_LOGO_PATH, "rb") as f:
            return base64.b64encode(f.read()).decode("ascii")
    except OSError:
        return None


def topbar(status_label: Optional[str] = None) -> None:
    """Sticky top application bar with the real NVISH logo (topbar has
    a light background, unlike the dark sidebar, so the logo's dark
    wordmark reads correctly)."""
    logo_b64 = _logo_base64()
    logo_html = (
        f'<img src="data:image/png;base64,{logo_b64}" class="topbar-logo-img" alt="NVISH"/>'
        if logo_b64
        else '<span class="topbar-logo">NVISH</span>'
    )

    right = '<span class="topbar-org">NVISH Solutions Inc.</span>'
    if status_label:
        right = (
            '<span class="topbar-badge">'
            '<span class="topbar-badge-dot"></span>'
            f"{status_label}"
            "</span>" + right
        )

    st.markdown(
        '<div class="topbar">'
        '<div class="topbar-left">'
        f"{logo_html}"
        '<span class="topbar-sep"></span>'
        '<span class="topbar-prod">Sales Summary Generator</span>'
        "</div>"
        f'<div class="topbar-right">{right}</div>'
        "</div>",
        unsafe_allow_html=True,
    )


def page_header(title: str, description: str = "", status: Optional[str] = None) -> None:
    """Page title + description + optional status badge."""
    status_html = ""
    if status:
        label = "Ready" if status == "ready" else str(status)
        cls = "ready" if status == "ready" else "idle"
        status_html = f'<span class="page-status {cls}">{label}</span>'
    desc_html = f'<p class="page-desc">{description}</p>' if description else ""
    st.markdown(
        '<div class="page-header">'
        f"<div><div class=\"page-title\">{title}</div>{desc_html}</div>"
        + status_html
        + "</div>",
        unsafe_allow_html=True,
    )


def card_open(title: str = "", subtitle: str = "", badge_html: str = "") -> None:
    head = ""
    if title:
        sub = f'<div class="card-subtitle">{subtitle}</div>' if subtitle else ""
        head = (
            '<div class="card-head">'
            f'<div><div class="card-title">{title}</div>{sub}</div>'
            f"{badge_html}</div>"
        )
    st.markdown(f'<div class="card">{head}<div class="card-body">', unsafe_allow_html=True)


def card_close() -> None:
    st.markdown("</div></div>", unsafe_allow_html=True)


def badge(text: str, variant: str = "blue") -> str:
    return f'<span class="badge badge-{variant}">{text}</span>'


def stat_tile(label: str, value: str, sub: str = "") -> None:
    """Small labelled value tile -- used for Detected Year / Detected
    Sheets / Number of Groups."""
    sub_html = f'<div class="stat-tile-sub">{sub}</div>' if sub else ""
    st.markdown(
        f'<div class="stat-tile"><div class="stat-tile-label">{label}</div>'
        f'<div class="stat-tile-value">{value}</div>{sub_html}</div>',
        unsafe_allow_html=True,
    )


def kpi_card(label: str, value, accent: str = "#0057B8", fmt: str = "plain") -> None:
    """KPI metric card for the post-generation Validation Summary."""
    if isinstance(value, (int, float)):
        if fmt == "seconds":
            v = f"{value:.1f}s"
        else:
            v = f"{int(value):,}"
    else:
        v = str(value)
    st.markdown(
        f'<div class="sfae-metric" style="--metric-accent:{accent};">'
        f'<div class="sfae-metric-top"><p class="sfae-metric-label">{label}</p></div>'
        f'<p class="sfae-metric-value">{v}</p></div>',
        unsafe_allow_html=True,
    )


def empty_state(title: str, description: str = "", icon: str = "") -> None:
    icon_html = f'<div class="empty-icon">{icon}</div>' if icon else ""
    desc_html = f'<p class="empty-desc">{description}</p>' if description else ""
    st.markdown(
        f'<div class="card"><div class="card-body p0"><div class="empty">{icon_html}'
        f'<p class="empty-title">{title}</p>{desc_html}</div></div></div>',
        unsafe_allow_html=True,
    )


def info_box(text: str, variant: str = "blue") -> None:
    colors = {
        "blue": ("#EEF4FC", "#0057B8", "#C7DCF5"),
        "green": ("#DCFCE7", "#16803C", "#86EFAC"),
        "amber": ("#FEF3C7", "#B45309", "#FDE68A"),
        "red": ("#FEE2E2", "#B91C1C", "#FECACA"),
    }
    bg, tx, bd = colors.get(variant, colors["blue"])
    st.markdown(
        f'<div style="background:{bg};border:1px solid {bd};border-radius:6px;'
        f'padding:10px 14px;font-size:13px;color:{tx};margin-bottom:12px;">{text}</div>',
        unsafe_allow_html=True,
    )


def upload_source_card(title: str, description: str, meta: List[Tuple[str, str]], status_html: str) -> None:
    """Single-source upload card header (the Master Workbook slot).
    Mirrors SFAE's numbered source row but simplified to one source."""
    meta_html = "".join(f'<span class="src-chip"><strong>{k}:</strong> {v}</span>' for k, v in meta)
    st.markdown(
        f"""
    <div class="src-row">
      <div class="src-head">
        <div class="src-num">1</div>
        <div class="src-info">
          <div class="src-title">{title}</div>
          <div class="src-desc">{description}</div>
          <div class="src-meta">{meta_html}</div>
        </div>
      </div>
      <div class="src-divider"></div>
      <div class="src-status">{status_html}</div>
    </div>
    """,
        unsafe_allow_html=True,
    )


# --------------------------------------------------------------------
# Live processing checklist
# --------------------------------------------------------------------
PROCESS_STEPS = [
    "Loading workbook...",
    "Reading rows...",
    "Grouping projects...",
    "Calculating quarters...",
    "Calculating margins...",
    "Generating summary...",
    "Finished.",
]


def render_process_log(placeholder, current_index: int, failed: bool = False, error_message: str = "") -> None:
    """Render the live step checklist into a pre-created st.empty()
    placeholder. Steps before `current_index` are shown done (check
    mark), the step at `current_index` is shown active (spinner) unless
    `failed` is True (shown with an error mark instead), and steps
    after are shown pending (hollow dot)."""
    rows = []
    for i, label in enumerate(PROCESS_STEPS):
        if failed and i == current_index:
            rows.append(f'<div class="plog"><span class="plog-err">&#10007;</span> {label}</div>')
        elif i < current_index or (i == current_index and current_index == len(PROCESS_STEPS) - 1 and not failed):
            rows.append(f'<div class="plog"><span class="plog-ok">&#10003;</span> {label}</div>')
        elif i == current_index:
            rows.append(
                f'<div class="plog plog-active"><span class="plog-icon"></span> {label}</div>'
            )
        else:
            rows.append(f'<div class="plog plog-pending"><span class="plog-icon"></span> {label}</div>')

    extra = ""
    if failed and error_message:
        extra = (
            f'<div class="plog"><span class="plog-err">&#10007;</span> '
            f"<strong>{error_message}</strong></div>"
        )

    placeholder.markdown(
        '<div class="card" style="margin-top:0;">'
        '<div class="card-head"><div class="card-title">Processing Status</div></div>'
        f'<div class="card-body">{"".join(rows)}{extra}</div></div>',
        unsafe_allow_html=True,
    )

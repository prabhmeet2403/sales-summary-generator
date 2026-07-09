"""
components/sidebar.py
======================
NVISH sidebar: brand mark, navigation (Upload & Generate / Settings /
About), and a compact workbook-status readout. Adapted from SFAE v2's
shared_sidebar.py -- same dark nav shell, same active-item styling via
the hidden ".sb-active-mark" + CSS :has() trick -- with all
forecast/placement/RFQ file-tracking replaced by a single Master
Workbook status line.

This module renders UI only; it does not read, aggregate, or compute
anything from the workbook itself.
"""
from __future__ import annotations

import streamlit as st

_NAV = [
    # (session_key, label, group_label_shown_before_this_item)
    ("upload", "Upload & Generate", None),
    ("ai", "Chatbot Prabh", None),
    ("settings", "Settings", "System"),
    ("about", "About", None),
]


def render_sidebar() -> None:
    current = st.session_state.get("nav", "upload")

    with st.sidebar:
        # -- Brand --------------------------------------------------
        st.markdown(
            '<div class="sb-brand">'
            '<div class="sb-mark">NV</div>'
            '<div><div class="sb-name">NVISH</div>'
            '<div class="sb-sub">Sales Summary Generator</div></div>'
            "</div>",
            unsafe_allow_html=True,
        )

        # -- Navigation -----------------------------------------------
        for key, label, group in _NAV:
            if group:
                st.markdown(f'<div class="sb-section">{group}</div>', unsafe_allow_html=True)
            if key == current:
                st.markdown('<span class="sb-active-mark"></span>', unsafe_allow_html=True)
            if st.button(label, key=f"_nav_{key}", use_container_width=True):
                st.session_state["nav"] = key
                st.rerun()

        # -- Workbook status -------------------------------------------
        st.markdown('<hr class="sb-div">', unsafe_allow_html=True)
        st.markdown('<div class="sb-lbl">Master Workbook</div>', unsafe_allow_html=True)

        preview = st.session_state.get("preview")
        upload_name = st.session_state.get("upload_name")
        preview_error = st.session_state.get("_preview_error")
        if preview is not None:
            dot, text = "ok", (upload_name or "Workbook loaded")
        elif preview_error is not None:
            dot, text = "warn", (upload_name or "Invalid workbook")
        elif upload_name:
            dot, text = "ok", upload_name
        else:
            dot, text = "miss", "Not uploaded"
        st.markdown(
            f'<div class="sb-files"><div class="sb-file">'
            f'<span class="sb-dot {dot}"></span><span>{text}</span></div></div>',
            unsafe_allow_html=True,
        )

        result = st.session_state.get("gen_result")
        if result is not None and getattr(result, "success", False):
            st.markdown('<hr class="sb-div">', unsafe_allow_html=True)
            st.markdown('<div class="sb-lbl">Last Run</div>', unsafe_allow_html=True)
            year = st.session_state.get("target_year", "")
            st.markdown(
                f'<div class="sb-files"><div class="sb-file">'
                f'<span class="sb-dot ok"></span><span>Summary {year} generated</span></div></div>',
                unsafe_allow_html=True,
            )

        # -- Footer -------------------------------------------------
        st.markdown(
            '<div class="sb-footer">NVISH Solutions Inc.<br>Internal Tool &middot; v1.0</div>',
            unsafe_allow_html=True,
        )

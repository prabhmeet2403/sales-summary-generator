"""
app.py
======
NVISH Sales Summary Generator -- Streamlit front end.

This file renders pages and wires up widgets. It performs NO Sales
Summary calculation itself: every number shown anywhere in this app
comes from streamlit_bridge.py, which in turn calls the existing,
unmodified backend (config, excel_reader, aggregator, comment_mapper,
historical_lookup, summary_writer, validator) and the existing,
unmodified gui.runner orchestration wrapper. See streamlit_bridge.py's
module docstring for the exact boundary between "new UI glue" and
"untouched calculation engine".

Run with:
    streamlit run app.py
"""
from __future__ import annotations

import os
import shutil
import tempfile
import time
from datetime import datetime
from pathlib import Path
from typing import Optional

from dotenv import load_dotenv

# Loads a .env file from the project root into the process environment,
# for local development convenience (see .env.example). Does nothing if
# no .env file is present, and never overrides a variable that is
# already set in the real environment -- a production deployment that
# sets these as actual environment variables is unaffected either way.
load_dotenv()

import streamlit as st

from components.ui import (
    load_css,
    topbar,
    page_header,
    card_open,
    card_close,
    badge,
    stat_tile,
    kpi_card,
    empty_state,
    info_box,
    upload_source_card,
    render_process_log,
    PROCESS_STEPS,
)
from components.sidebar import render_sidebar
import streamlit_bridge as bridge

st.set_page_config(
    page_title="NVISH -- Sales Summary Generator",
    page_icon="\U0001F4CA",
    layout="wide",
    initial_sidebar_state="expanded",
)


# =============================================================================
# Helpers (presentation-only)
# =============================================================================
def _fmt_size(num_bytes: int) -> str:
    if num_bytes >= 1_000_000:
        return f"{num_bytes/1_000_000:.1f} MB"
    return f"{num_bytes/1_000:.1f} KB"


def _fmt_seconds(seconds: float) -> str:
    return f"{seconds:.1f}s"


class _StepDriver:
    """Maps the real, unmodified backend's progress_cb messages onto the
    seven user-facing checklist labels, advancing forward only, never
    fabricating a step that hasn't genuinely started. See module
    docstring in components/ui.py (PROCESS_STEPS) for the label list.
    """

    def __init__(self, placeholder):
        self.placeholder = placeholder
        self.current = 0
        self._agg_calls = 0
        render_process_log(self.placeholder, self.current)

    def _advance(self, target: int, pause: float = 0.0) -> None:
        if target > self.current:
            self.current = target
            render_process_log(self.placeholder, self.current)
            if pause:
                time.sleep(pause)

    def handle(self, message: str) -> None:
        m = message.lower()
        if "loading master workbook" in m:
            self._advance(0)
        elif "reading project rows" in m:
            self._advance(1)
        elif "aggregating section" in m:
            self._agg_calls += 1
            if self._agg_calls == 1:
                # Rules 1-4 (group, quarters, margin) all happen inside
                # this single backend call for each section; these three
                # labels are shown in sequence with a brief pause purely
                # so the (genuinely occurring, just not separately
                # instrumented) sub-steps are visible to the user.
                self._advance(2, pause=0.15)
                self._advance(3, pause=0.15)
                self._advance(4, pause=0.15)
        elif "building the summary workbook" in m:
            self._advance(5)
        elif message.strip().lower() == "done.":
            self._advance(6)

    def fail(self, error_message: str) -> None:
        render_process_log(self.placeholder, self.current, failed=True, error_message=error_message)


def _reset_for_new_upload() -> None:
    old_tmp = st.session_state.get("tmp_dir")
    if old_tmp and os.path.isdir(old_tmp):
        shutil.rmtree(old_tmp, ignore_errors=True)
    for key in (
        "master_path", "tmp_dir", "preview", "_base_preview", "_preview_error",
        "gen_result", "elapsed_seconds", "generated_at", "upload_name", "year_input", "target_year",
    ):
        st.session_state.pop(key, None)


# =============================================================================
# Upload & Generate page
# =============================================================================
def _resolve_upload_state() -> None:
    """Resolve the upload/preview state for this run BEFORE anything is
    rendered (sidebar included), so every badge and card -- in the
    sidebar and in the main page -- reflects the true current state on
    the same run instead of lagging one rerun behind. Pure state
    management; no rendering happens here."""
    new_file = st.session_state.get("master_upload")
    # The `master_upload` file_uploader widget only exists while the
    # Upload & Generate page is actually being rendered -- on any other
    # page, reading its session_state key back as None means "this
    # widget wasn't instantiated this run," not "the user removed the
    # file." Without this guard, simply navigating to a different page
    # (Settings, About, or the AI Assistant added in Phase 2) would be
    # misread as a file removal and silently clear gen_result along
    # with everything else _reset_for_new_upload() clears.
    on_upload_page = st.session_state.get("nav", "upload") == "upload"

    if new_file is not None:
        fingerprint = (new_file.name, new_file.size)
        if st.session_state.get("_upload_fingerprint") != fingerprint:
            _reset_for_new_upload()
            tmp = tempfile.mkdtemp(prefix="sfae_streamlit_")
            master_path = os.path.join(tmp, new_file.name)
            with open(master_path, "wb") as fh:
                fh.write(new_file.getvalue())
            st.session_state["tmp_dir"] = tmp
            st.session_state["master_path"] = master_path
            st.session_state["upload_name"] = f"{new_file.name} \u00b7 {_fmt_size(new_file.size)}"
            st.session_state["_upload_fingerprint"] = fingerprint
    elif (
        new_file is None
        and on_upload_page
        and st.session_state.get("_upload_fingerprint")
        and not (st.session_state.get("gen_result") and st.session_state["gen_result"].success)
    ):
        # User removed the file from the uploader widget before
        # generating anything. Once a Summary has already been
        # successfully generated, the widget reporting "no file" is
        # ambiguous -- it's at least as likely to be Streamlit not
        # preserving the widget's value across a period it wasn't
        # rendered (e.g. the user visited another page and came back)
        # as it is a deliberate attempt to discard a completed result.
        # Uploading a genuinely different file is unaffected by this
        # guard and still resets everything correctly via the
        # fingerprint check above.
        _reset_for_new_upload()
        st.session_state.pop("_upload_fingerprint", None)

    master_path = st.session_state.get("master_path")
    base_preview = st.session_state.get("_base_preview")
    preview_error = None

    if master_path and base_preview is None:
        try:
            base_preview = bridge.preview_workbook(master_path)
            st.session_state["_base_preview"] = base_preview
        except bridge.GenerationError as exc:
            preview_error = exc

    year_selected = st.session_state.get("year_input")
    if year_selected is None and base_preview is not None:
        year_selected = base_preview.target_year

    preview: Optional[bridge.PreviewResult] = None
    if master_path and base_preview is not None:
        try:
            preview = bridge.preview_workbook(master_path, year=int(year_selected))
            st.session_state["target_year"] = preview.target_year
        except bridge.GenerationError as exc:
            preview_error = exc

    st.session_state["preview"] = preview
    st.session_state["_preview_error"] = preview_error
    st.session_state["_base_preview"] = base_preview


def _render_upload_page() -> None:
    base_preview = st.session_state.get("_base_preview")
    preview: Optional[bridge.PreviewResult] = st.session_state.get("preview")
    preview_error = st.session_state.get("_preview_error")
    year_selected = st.session_state.get("year_input") or (
        base_preview.target_year if base_preview else None
    )
    is_ok = preview is not None

    # =====================================================================
    # Phase 2: render, using the fully-resolved state above.
    # =====================================================================
    st.markdown(
        '<div class="card" style="margin-bottom:0;border-bottom-left-radius:0;border-bottom-right-radius:0;">'
        '<div class="card-head"><div><div class="card-title">Data Source</div>'
        '<div class="card-subtitle">Upload the Master workbook to generate this month\'s '
        "Sales &amp; Forecast Summary.</div></div>"
        + (badge("Validated", "green") if is_ok else badge("Not uploaded", "gray"))
        + "</div></div>",
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div style="background:#fff;border:1px solid #E2E8F0;border-top:none;'
        'border-radius:0 0 10px 10px;box-shadow:0 1px 2px rgba(15,23,42,0.05);'
        'margin-bottom:20px;overflow:hidden;">',
        unsafe_allow_html=True,
    )

    if is_ok:
        status_html = (
            '<span class="badge badge-green">&#10003; Validated</span>'
            f'<span class="src-fname">{st.session_state.get("upload_name","")}</span>'
        )
    else:
        status_html = '<span class="badge badge-gray">Not uploaded</span>'

    upload_source_card(
        "Master Workbook",
        "The monthly Sales &amp; Forecast master workbook -- the same file used by the "
        "desktop application, unchanged.",
        [("Format", "Excel (.xlsx)"), ("Sheet pattern", "Sales by Customer- &lt;year&gt;")],
        status_html,
    )

    ucol, _ = st.columns([1, 3])
    with ucol:
        st.file_uploader("Master Workbook", type=["xlsx"], key="master_upload", label_visibility="collapsed")
    st.markdown("</div>", unsafe_allow_html=True)

    if preview_error is not None:
        info_box(f"<strong>{preview_error.title}</strong> -- {preview_error.message}", "red")

    # ---- Once uploaded: workbook facts + editable year -------------------
    if base_preview is not None:
        card_open(
            "Workbook Loaded",
            "Detected automatically from the Master workbook -- the year is editable.",
            badge("&#10003; Loaded", "green"),
        )
        top_row = st.columns([1, 1, 1])
        with top_row[0]:
            st.markdown('<div class="stat-tile-label">Detected Year</div>', unsafe_allow_html=True)
            st.number_input(
                "Detected Year",
                min_value=2000,
                max_value=2100,
                step=1,
                value=int(year_selected),
                key="year_input",
                label_visibility="collapsed",
            )

        if preview is not None:
            sheet_names = [preview.main_sheet]
            if preview.comments_sheet:
                sheet_names.append(preview.comments_sheet)
            sheet_names.extend(preview.prior_year_sheets.values())
            sheets_value = f"{len(sheet_names)} of {preview.total_sheets_in_workbook}"
            sheets_sub = ", ".join(sheet_names)

            with top_row[1]:
                stat_tile("Detected Sheets", sheets_value, sheets_sub)
            with top_row[2]:
                stat_tile("Number of Groups", f"{preview.num_groups:,}", f"{preview.num_rows:,} project rows read")

        card_close()

    # ---- Generate button ---------------------------------------------
    can_generate = preview is not None
    gen_clicked = st.button(
        "\u25b6  Generate Summary",
        type="primary",
        use_container_width=True,
        disabled=not can_generate,
    )
    if not can_generate:
        st.caption("Upload a valid Master workbook above to enable Generate Summary.")

    if gen_clicked and can_generate:
        _run_generation(st.session_state["master_path"], preview.target_year)

    # ---- Validation Summary (after generation) ------------------------
    result: bridge.GenerationResult | None = st.session_state.get("gen_result")
    if result is not None:
        _render_result(result)
    elif not can_generate and not st.session_state.get("master_path"):
        empty_state(
            "Ready to generate",
            "Upload the Master workbook above, confirm the detected year, then click Generate Summary.",
            "\U0001F4C4",
        )


def _run_generation(master_path: str, year: int) -> None:
    output_dir = os.path.join(st.session_state["tmp_dir"], "output")
    os.makedirs(output_dir, exist_ok=True)

    placeholder = st.empty()
    driver = _StepDriver(placeholder)

    start = time.time()
    result = bridge.generate_summary(
        master_path, output_dir, year, progress_cb=driver.handle
    )
    elapsed = time.time() - start

    if not result.success:
        driver.fail(result.error_message or "Generation failed.")
    st.session_state["gen_result"] = result
    st.session_state["elapsed_seconds"] = elapsed
    # Captured once, here, at the moment generation actually completes --
    # not read again until the NEXT successful generation overwrites it.
    # The Download button (below) reuses this stored value for the
    # filename rather than calling datetime.now() at download time, so
    # repeated downloads of the same generated workbook keep the same
    # filename no matter how much later they happen.
    st.session_state["generated_at"] = datetime.now()
    st.session_state["target_year"] = year
    st.rerun()


def _render_result(result: "bridge.GenerationResult") -> None:
    report = result.report
    elapsed = st.session_state.get("elapsed_seconds", 0.0)

    if result.success:
        card_open("Validation Summary", "Generation completed successfully.", badge("Generated", "green"))
    else:
        card_open("Validation Summary", "Generation failed -- see details below.", badge("Failed", "red"))

    cols = st.columns(4)
    with cols[0]:
        kpi_card("Groups Processed", report.total_groups_processed if report else 0)
    with cols[1]:
        kpi_card("Comments Matched", report.total_comments_matched if report else 0)
    with cols[2]:
        kpi_card("Warnings", len(report.warnings) if report else 0, accent="#B45309")
    with cols[3]:
        kpi_card("Time Taken", elapsed, fmt="seconds", accent="#16803C" if result.success else "#B91C1C")
    card_close()

    if not result.success:
        info_box(f"<strong>{result.error_title}</strong> -- {result.error_message}", "red")

    # ---- Download -------------------------------------------------
    if result.success and result.output_path and os.path.isfile(result.output_path):
        card_open("Download Results")
        # Filename uses the stored `generated_at` timestamp (captured once,
        # in _run_generation, at the moment this workbook was produced) --
        # NOT datetime.now() here, so the name stays identical across any
        # number of downloads until the user runs Generate again. Falls
        # back to the on-disk name only in the defensive case where
        # `generated_at` is somehow missing (e.g. a pre-existing session).
        generated_at = st.session_state.get("generated_at")
        if generated_at is not None:
            stem = Path(result.output_path).stem
            suffix = Path(result.output_path).suffix
            download_name = f"{stem}_{generated_at:%Y-%m-%d_%H-%M}{suffix}"
        else:
            download_name = Path(result.output_path).name
        with open(result.output_path, "rb") as fh:
            st.download_button(
                "\U0001F4E5 Download Summary Workbook",
                data=fh.read(),
                file_name=download_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        card_close()

    if report is not None:
        with st.expander("Full Validation Report"):
            st.text(report.render())


# =============================================================================
# Settings / About pages
# =============================================================================
def _settings_page() -> None:
    empty_state("Coming Soon", "Settings will be available in a future release.", "\u2699")


def _about_page() -> None:
    card_open("Application")
    st.markdown(
        """
        <div style="font-size:13px;color:#334155;line-height:1.9;">
        <strong>Product:</strong> Sales Forecast Automation Engine (SFAE)<br>
        <strong>Built by:</strong> NVISH Solutions Inc.<br>
        <strong>Calculation engine:</strong> Unmodified backend shared with the desktop
        (Tkinter) application -- config.py, excel_reader.py, aggregator.py,
        comment_mapper.py, historical_lookup.py, summary_writer.py, validator.py.<br>
        <strong>Pipeline:</strong> Master workbook &rarr; grouped by Group/Sub-Group &rarr;
        quarters &amp; margin calculated &rarr; historical years attached &rarr; client
        comments matched &rarr; consolidated Summary workbook.<br>
        <strong>Interfaces:</strong> Desktop app (gui_main.py) and this Streamlit app
        (app.py) both call the exact same calculation engine, so results are identical
        regardless of which one you use.
        </div>
        """,
        unsafe_allow_html=True,
    )
    card_close()


# =============================================================================
# Entry point
# =============================================================================
load_css()

if "nav" not in st.session_state:
    st.session_state["nav"] = "upload"

_resolve_upload_state()
render_sidebar()

from ai.ui.floating_widget import render_floating_assistant

_result = st.session_state.get("gen_result")
render_floating_assistant(_result)
_status = "ready" if (_result and _result.success) else None
topbar(status_label="Summary Ready" if _status == "ready" else None)

_nav = st.session_state.get("nav", "upload")

if _nav == "upload":
    page_header(
        "Sales Summary Generator",
        "Upload the Master Workbook to generate this month's Sales &amp; Forecast Summary.",
        status=_status,
    )
    _render_upload_page()
elif _nav == "ai":
    page_header("Chatbot Prabh", "Ask questions about the generated Sales &amp; Forecast Summary.")
    from ai.ui import chat_page

    chat_page.render(_result)
elif _nav == "settings":
    page_header("Settings", "Application configuration.")
    _settings_page()
elif _nav == "about":
    page_header("About", "Sales Forecast Automation Engine.")
    _about_page()

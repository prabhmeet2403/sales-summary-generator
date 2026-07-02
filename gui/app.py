"""
gui/app.py
==========
The main desktop window for the Sales Forecast Automation Engine.

This module is pure presentation: it collects user input (Master
workbook, output folder, target year), runs gui.runner.generate_summary
on a background thread so the UI never freezes, and renders whatever
that function returns. It performs no aggregation, matching, or
Excel-writing logic itself -- all of that lives, unmodified, in the
existing backend modules (see gui/runner.py's module docstring).
"""
from __future__ import annotations

import os
import platform
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog
from tkinter.scrolledtext import ScrolledText
from typing import Dict, Optional

import ttkbootstrap as tb
from ttkbootstrap.constants import BOTH, LEFT, RIGHT, TOP, BOTTOM, X, Y, YES

from . import branding
from . import runner
from .dialogs import show_error_dialog

try:
    from PIL import Image, ImageTk
except ImportError:  # pragma: no cover - Pillow ships with ttkbootstrap anyway
    Image = None
    ImageTk = None


DEFAULT_OUTPUT_DIR = Path.home() / "Documents" / "NVISH Sales Summary Output"

STAT_KEYS = [
    ("Groups Processed", "total_groups_processed"),
    ("Comments Matched", "total_comments_matched"),
    ("Missing Comments", "total_missing_comments"),
    ("Skipped Blank", "total_skipped_blank_groups"),
]


class SFAEApp(tb.Window):
    def __init__(self) -> None:
        super().__init__(
            title=branding.APP_TITLE,
            themename="flatly",
            size=(900, 840),
            resizable=(True, True),
        )
        self.minsize(860, 680)
        self._center_on_screen()
        self._set_window_icon()

        self._task_queue: "queue.Queue" = queue.Queue()
        self._busy = False
        self._last_result: Optional[runner.GenerationResult] = None
        self._logo_image = None
        self.stat_vars: Dict[str, tk.StringVar] = {}

        self._build_layout()
        self.after(100, self._poll_queue)

    # ------------------------------------------------------------------
    # Layout construction
    # ------------------------------------------------------------------
    def _center_on_screen(self) -> None:
        self.update_idletasks()
        w, h = 900, 840
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 3
        self.geometry(f"{w}x{h}+{max(0, x)}+{max(0, y)}")

    def _set_window_icon(self) -> None:
        try:
            self.iconbitmap(str(branding.icon_path()))
        except Exception:
            pass  # non-Windows platforms / missing icon: fail silently

    def _build_layout(self) -> None:
        self._build_header()

        body = tb.Frame(self, padding=(22, 14, 22, 8))
        body.pack(fill=BOTH, expand=YES)

        self._build_input_section(body)
        self._build_output_section(body)
        self._build_year_section(body)
        self._build_generate_section(body)
        self._build_progress_section(body)
        self._build_summary_section(body)
        self._build_actions_section(body)

        self._build_footer()

    def _build_header(self) -> None:
        header = tk.Frame(self, bg=branding.NAVY, height=76)
        header.pack(fill=X, side=TOP)
        header.pack_propagate(False)

        content = tk.Frame(header, bg=branding.NAVY)
        content.pack(side=LEFT, padx=20, pady=10)

        if Image is not None:
            try:
                img = Image.open(branding.logo_path()).resize((52, 52), Image.LANCZOS)
                self._logo_image = ImageTk.PhotoImage(img)
                tk.Label(content, image=self._logo_image, bg=branding.NAVY).pack(side=LEFT, padx=(0, 14))
            except Exception:
                pass

        text_col = tk.Frame(content, bg=branding.NAVY)
        text_col.pack(side=LEFT)
        tk.Label(
            text_col,
            text=branding.COMPANY_NAME,
            font=(branding.FONT_FAMILY, 16, "bold"),
            fg="white",
            bg=branding.NAVY,
        ).pack(anchor="w")
        tk.Label(
            text_col,
            text=branding.APP_SUBTITLE,
            font=(branding.FONT_FAMILY, 9),
            fg=branding.TEAL,
            bg=branding.NAVY,
        ).pack(anchor="w")

    def _build_input_section(self, parent: tk.Misc) -> None:
        frame = tb.Labelframe(parent, text=" 1.  Master Workbook ", padding=10)
        frame.pack(fill=X, pady=(0, 8))
        frame.columnconfigure(0, weight=1)

        self.input_var = tk.StringVar()
        entry = tb.Entry(frame, textvariable=self.input_var, state="readonly")
        entry.grid(row=0, column=0, sticky="ew", padx=(0, 10))

        self.input_browse_btn = tb.Button(
            frame, text="Browse…", bootstyle="secondary-outline", width=12, command=self._browse_input
        )
        self.input_browse_btn.grid(row=0, column=1)

        self.year_status_var = tk.StringVar(
            value="Select the Master workbook (.xlsx) to generate the Summary from."
        )
        tb.Label(frame, textvariable=self.year_status_var, bootstyle="secondary", font=(branding.FONT_FAMILY, 9)).grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(6, 0)
        )

    def _build_output_section(self, parent: tk.Misc) -> None:
        frame = tb.Labelframe(parent, text=" 2.  Output Folder ", padding=10)
        frame.pack(fill=X, pady=(0, 8))
        frame.columnconfigure(0, weight=1)

        self.output_var = tk.StringVar(value=str(DEFAULT_OUTPUT_DIR))
        entry = tb.Entry(frame, textvariable=self.output_var)
        entry.grid(row=0, column=0, sticky="ew", padx=(0, 10))

        self.output_browse_btn = tb.Button(
            frame, text="Browse…", bootstyle="secondary-outline", width=12, command=self._browse_output
        )
        self.output_browse_btn.grid(row=0, column=1)

        tb.Label(
            frame,
            text="The generated workbook and validation report are both saved here.",
            bootstyle="secondary",
            font=(branding.FONT_FAMILY, 9),
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))

    def _build_year_section(self, parent: tk.Misc) -> None:
        frame = tb.Labelframe(parent, text=" 3.  Target Year ", padding=10)
        frame.pack(fill=X, pady=(0, 10))

        row = tb.Frame(frame)
        row.pack(fill=X)
        tb.Label(row, text="Year:").pack(side=LEFT, padx=(0, 8))

        self.year_var = tk.StringVar()
        self.year_combo = tb.Combobox(row, textvariable=self.year_var, values=[], width=10, state="disabled")
        self.year_combo.pack(side=LEFT, padx=(0, 14))

        self.year_hint_var = tk.StringVar(value="Auto-detected once a workbook is selected — editable.")
        tb.Label(row, textvariable=self.year_hint_var, bootstyle="secondary", font=(branding.FONT_FAMILY, 9)).pack(
            side=LEFT
        )

    def _build_generate_section(self, parent: tk.Misc) -> None:
        self.generate_btn = tb.Button(
            parent,
            text="▶   Generate Summary",
            bootstyle="success",
            command=self._on_generate_clicked,
            padding=(10, 10),
        )
        self.generate_btn.pack(fill=X, pady=(0, 10))

    def _build_progress_section(self, parent: tk.Misc) -> None:
        frame = tb.Frame(parent)
        frame.pack(fill=X, pady=(0, 10))

        self.progress_bar = tb.Progressbar(frame, mode="indeterminate", bootstyle="info-striped")
        self.progress_bar.pack(fill=X)

        self.status_var = tk.StringVar(value="Ready.")
        tb.Label(frame, textvariable=self.status_var, bootstyle="secondary", font=(branding.FONT_FAMILY, 9)).pack(
            anchor="w", pady=(5, 0)
        )

    def _build_summary_section(self, parent: tk.Misc) -> None:
        frame = tb.Labelframe(parent, text=" Validation Summary ", padding=10)
        frame.pack(fill=BOTH, expand=YES, pady=(0, 10))

        self.summary_banner_var = tk.StringVar(value="No summary yet — generate a workbook to see results here.")
        self.summary_banner = tb.Label(
            frame, textvariable=self.summary_banner_var, font=(branding.FONT_FAMILY, 11, "bold"), bootstyle="secondary"
        )
        self.summary_banner.pack(anchor="w", pady=(0, 8))

        stats_row = tb.Frame(frame)
        stats_row.pack(fill=X, pady=(0, 8))
        for i, (label, _attr) in enumerate(STAT_KEYS):
            stats_row.columnconfigure(i, weight=1)
            cell = tb.Frame(stats_row, bootstyle="light", padding=6)
            cell.grid(row=0, column=i, sticky="ew", padx=(0 if i == 0 else 6, 0))
            var = tk.StringVar(value="—")
            tb.Label(cell, textvariable=var, font=(branding.FONT_FAMILY, 16, "bold"), bootstyle="primary").pack()
            tb.Label(cell, text=label, font=(branding.FONT_FAMILY, 8), bootstyle="secondary").pack()
            self.stat_vars[label] = var

        text_container = tb.Frame(frame)
        text_container.pack(fill=BOTH, expand=YES)
        self.report_text = ScrolledText(
            text_container, height=6, font=("Consolas", 9), wrap="word", state="disabled",
            background="white", foreground=branding.TEXT_PRIMARY, relief="flat", borderwidth=1,
        )
        self.report_text.pack(fill=BOTH, expand=YES)

    def _build_actions_section(self, parent: tk.Misc) -> None:
        frame = tb.Frame(parent)
        frame.pack(fill=X)

        self.open_workbook_btn = tb.Button(
            frame, text="📂  Open Workbook", bootstyle="primary-outline", state="disabled", command=self._open_workbook
        )
        self.open_workbook_btn.pack(side=LEFT, expand=YES, fill=X, padx=(0, 6))

        self.open_report_btn = tb.Button(
            frame,
            text="📄  Open Validation Report",
            bootstyle="primary-outline",
            state="disabled",
            command=self._open_report,
        )
        self.open_report_btn.pack(side=LEFT, expand=YES, fill=X, padx=6)

        self.open_folder_btn = tb.Button(
            frame,
            text="📁  Open Output Folder",
            bootstyle="primary-outline",
            state="normal",
            command=self._open_output_folder,
        )
        self.open_folder_btn.pack(side=LEFT, expand=YES, fill=X, padx=(6, 0))

    def _build_footer(self) -> None:
        footer = tk.Frame(self, bg=branding.LIGHT_BG)
        footer.pack(fill=X, side=BOTTOM)
        tk.Label(
            footer,
            text=f"{branding.FOOTER_TEXT}    ·    v1.0",
            bg=branding.LIGHT_BG,
            fg=branding.TEXT_MUTED,
            font=(branding.FONT_FAMILY, 8),
        ).pack(pady=6)

    # ------------------------------------------------------------------
    # Event handlers
    # ------------------------------------------------------------------
    def _browse_input(self) -> None:
        path = filedialog.askopenfilename(
            title="Select Master Workbook",
            filetypes=[("Excel Workbook", "*.xlsx"), ("All files", "*.*")],
        )
        if not path:
            return
        self.input_var.set(path)
        self._start_year_discovery(path)

    def _browse_output(self) -> None:
        initial = self.output_var.get().strip() or str(Path.home())
        path = filedialog.askdirectory(title="Select Output Folder", initialdir=initial)
        if not path:
            return
        self.output_var.set(path)
        self.open_folder_btn.configure(state="normal")

    def _start_year_discovery(self, path: str) -> None:
        self.year_combo.configure(state="disabled")
        self.year_hint_var.set("Detecting available years…")

        def worker() -> None:
            try:
                years = runner.discover_years(path)
                self._task_queue.put(("years_ok", years))
            except runner.GenerationError as exc:
                self._task_queue.put(("years_err", (exc.title, exc.message)))
            except Exception as exc:  # noqa: BLE001
                self._task_queue.put(("years_err", ("Unexpected Error", str(exc))))

        threading.Thread(target=worker, daemon=True).start()

    def _on_generate_clicked(self) -> None:
        if self._busy:
            return

        input_path = self.input_var.get().strip()
        output_dir = self.output_var.get().strip()

        if not input_path:
            show_error_dialog(self, "Master Workbook Required", "Please select a Master workbook (.xlsx) before generating.")
            return
        if not output_dir:
            show_error_dialog(self, "Output Folder Required", "Please select an output folder before generating.")
            return

        year_text = self.year_var.get().strip()
        year = None
        if year_text:
            try:
                year = int(year_text)
            except ValueError:
                show_error_dialog(self, "Invalid Year", f"'{year_text}' is not a valid 4-digit year.")
                return

        self._set_busy(True)
        self._reset_summary_panel()
        self.status_var.set("Starting…")
        self.progress_bar.start(12)

        def worker() -> None:
            result = runner.generate_summary(
                input_path,
                output_dir,
                year,
                progress_cb=lambda msg: self._task_queue.put(("progress", msg)),
            )
            self._task_queue.put(("result", result))

        threading.Thread(target=worker, daemon=True).start()

    def _poll_queue(self) -> None:
        try:
            while True:
                kind, payload = self._task_queue.get_nowait()
                if kind == "progress":
                    self.status_var.set(payload)
                elif kind == "years_ok":
                    self._populate_years(payload)
                elif kind == "years_err":
                    title, message = payload
                    self.year_hint_var.set("Could not auto-detect years — you can still type one manually.")
                    self.year_combo.configure(state="normal")
                    show_error_dialog(self, title, message)
                elif kind == "result":
                    self._handle_result(payload)
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)

    def _populate_years(self, years) -> None:
        values = [str(y) for y in sorted(years)]
        self.year_combo.configure(values=values, state="normal")
        self.year_var.set(str(max(years)))
        self.year_hint_var.set(f"Detected: {', '.join(values)}  (editable — you may type a different year)")

    def _handle_result(self, result: runner.GenerationResult) -> None:
        self.progress_bar.stop()
        self._set_busy(False)
        self._last_result = result
        self._populate_summary(result)

        if result.success:
            self.status_var.set("Generation completed successfully.")
            self.open_workbook_btn.configure(state="normal")
            self.open_folder_btn.configure(state="normal")
        else:
            self.status_var.set("Generation failed. See details below.")
            show_error_dialog(
                self,
                result.error_title or "Generation Failed",
                result.error_message or "An unknown error occurred.",
                details=result.error_details or None,
            )

        self.open_report_btn.configure(state="normal" if result.report_path else "disabled")

    def _reset_summary_panel(self) -> None:
        self.summary_banner_var.set("⏳  Generating…")
        self.summary_banner.configure(bootstyle="secondary")
        for var in self.stat_vars.values():
            var.set("—")
        self.report_text.configure(state="normal")
        self.report_text.delete("1.0", "end")
        self.report_text.configure(state="disabled")
        self.open_workbook_btn.configure(state="disabled")
        self.open_report_btn.configure(state="disabled")

    def _populate_summary(self, result: runner.GenerationResult) -> None:
        if result.success:
            self.summary_banner_var.set("✅  Generation Successful")
            self.summary_banner.configure(bootstyle="success")
        else:
            self.summary_banner_var.set("❌  Generation Failed")
            self.summary_banner.configure(bootstyle="danger")

        report = result.report
        if report is not None:
            for label, attr in STAT_KEYS:
                self.stat_vars[label].set(str(getattr(report, attr)))
            text = report.render()
        else:
            for var in self.stat_vars.values():
                var.set("—")
            text = result.error_message or "No report available."

        self.report_text.configure(state="normal")
        self.report_text.delete("1.0", "end")
        self.report_text.insert("1.0", text)
        self.report_text.configure(state="disabled")

    def _set_busy(self, busy: bool) -> None:
        self._busy = busy
        state = "disabled" if busy else "normal"
        self.generate_btn.configure(state=state)
        self.input_browse_btn.configure(state=state)
        self.output_browse_btn.configure(state=state)
        if not busy:
            has_years = bool(self.year_combo["values"])
            self.year_combo.configure(state="normal" if has_years or True else "disabled")
        else:
            self.year_combo.configure(state="disabled")

    # ------------------------------------------------------------------
    # "Open ..." actions
    # ------------------------------------------------------------------
    def _open_path(self, path) -> None:
        path = Path(path)
        try:
            if platform.system() == "Windows":
                os.startfile(str(path))  # type: ignore[attr-defined]
            elif platform.system() == "Darwin":
                subprocess.run(["open", str(path)], check=False)
            else:
                subprocess.run(["xdg-open", str(path)], check=False)
        except Exception as exc:  # noqa: BLE001
            show_error_dialog(self, "Could Not Open", f"Could not open:\n{path}\n\n{exc}")

    def _open_workbook(self) -> None:
        if self._last_result and self._last_result.output_path:
            self._open_path(self._last_result.output_path)

    def _open_report(self) -> None:
        if self._last_result and self._last_result.report_path:
            self._open_path(self._last_result.report_path)

    def _open_output_folder(self) -> None:
        folder = self.output_var.get().strip()
        if folder:
            Path(folder).mkdir(parents=True, exist_ok=True)
            self._open_path(folder)


def main() -> None:
    app = SFAEApp()
    app.mainloop()


if __name__ == "__main__":
    main()

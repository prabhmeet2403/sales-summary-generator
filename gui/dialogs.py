"""
gui/dialogs.py
===============
Professional modal dialogs used in place of raw console tracebacks:
- show_error_dialog: a concise message plus an optional, collapsible
  "Technical Details" panel (scrollable, copyable) for unexpected
  exceptions.
- show_confirm_dialog: a simple Yes/No confirmation styled to match.
"""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk
from typing import Optional

import ttkbootstrap as tb
from ttkbootstrap.constants import BOTH, LEFT, RIGHT, X, YES, DANGER, SECONDARY, PRIMARY

from . import branding


def _center_on_parent(win: tk.Toplevel, parent: tk.Misc, width: int, height: int) -> None:
    win.update_idletasks()
    px, py = parent.winfo_rootx(), parent.winfo_rooty()
    pw, ph = parent.winfo_width(), parent.winfo_height()
    x = px + max(0, (pw - width) // 2)
    y = py + max(0, (ph - height) // 3)
    win.geometry(f"{width}x{height}+{x}+{y}")


def show_error_dialog(
    parent: tk.Misc,
    title: str,
    message: str,
    details: Optional[str] = None,
) -> None:
    win = tk.Toplevel(parent)
    win.title(title)
    win.transient(parent)
    win.grab_set()
    win.resizable(True, True)
    try:
        win.iconbitmap(str(branding.icon_path()))
    except Exception:
        pass

    outer = tb.Frame(win, padding=20)
    outer.pack(fill=BOTH, expand=YES)

    header = tb.Frame(outer)
    header.pack(fill=X)
    tb.Label(header, text="⚠", font=(branding.FONT_FAMILY, 28), bootstyle=DANGER).pack(side=LEFT, padx=(0, 12))
    text_frame = tb.Frame(header)
    text_frame.pack(side=LEFT, fill=X, expand=YES)
    tb.Label(text_frame, text=title, font=(branding.FONT_FAMILY, 13, "bold")).pack(anchor="w")
    tb.Label(
        text_frame, text=message, wraplength=440, justify=LEFT, bootstyle=SECONDARY
    ).pack(anchor="w", pady=(4, 0))

    if details:
        sep = ttk.Separator(outer)
        sep.pack(fill=X, pady=14)

        toggle_state = {"expanded": False}
        details_frame = tb.Frame(outer)

        def toggle_details() -> None:
            if toggle_state["expanded"]:
                details_frame.pack_forget()
                toggle_btn.configure(text="▸ Show Technical Details")
                _center_on_parent(win, parent, 520, 220)
            else:
                details_frame.pack(fill=BOTH, expand=YES, pady=(10, 0))
                toggle_btn.configure(text="▾ Hide Technical Details")
                _center_on_parent(win, parent, 620, 480)
            toggle_state["expanded"] = not toggle_state["expanded"]

        toggle_btn = tb.Button(
            outer, text="▸ Show Technical Details", bootstyle="link", command=toggle_details
        )
        toggle_btn.pack(anchor="w")

        text_widget = tk.Text(details_frame, wrap="word", height=14, font=("Consolas", 9))
        scrollbar = ttk.Scrollbar(details_frame, orient="vertical", command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        text_widget.insert("1.0", details)
        text_widget.configure(state="disabled")
        text_widget.pack(side=LEFT, fill=BOTH, expand=YES)
        scrollbar.pack(side=RIGHT, fill="y")

        def copy_details() -> None:
            win.clipboard_clear()
            win.clipboard_append(details)

        button_row = tb.Frame(outer)
        button_row.pack(fill=X, pady=(14, 0))
        tb.Button(button_row, text="Copy Details", bootstyle="secondary-outline", command=copy_details).pack(
            side=LEFT
        )
        tb.Button(button_row, text="Close", bootstyle=PRIMARY, command=win.destroy).pack(side=RIGHT)
    else:
        button_row = tb.Frame(outer)
        button_row.pack(fill=X, pady=(18, 0))
        tb.Button(button_row, text="Close", bootstyle=PRIMARY, command=win.destroy).pack(side=RIGHT)

    _center_on_parent(win, parent, 520, 220)
    win.wait_window()

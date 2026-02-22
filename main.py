"""
Substack Archiver
=================
A GUI application to download Substack newsletters and convert them to ePub.

Uses sbstck-dl for downloading and pandoc for ePub conversion.
No command-line knowledge required.

Dependencies (install via pip):
    pip install customtkinter tkcalendar
    pip install sbstck-dl
"""

import os
import sys
import json
import queue
import threading
import subprocess
import re
import tkinter as tk
import tkinter.filedialog as filedialog
import tkinter.messagebox as messagebox

import customtkinter as ctk
from tkcalendar import DateEntry


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

DEFAULTS: dict = {
    "sbstckdl_path": "",
    "pandoc_path": r"C:\Program Files\Pandoc\pandoc.exe",
    "last_url": "",
    "last_output_dir": "",
    "last_format": "Markdown (.md)",
    "last_epub_source_dir": "",
    "last_epub_output_file": "",
    "last_author": "",
    "window_geometry": "1050x800",
}

FORMAT_DISPLAY_TO_FLAG = {
    "Markdown (.md)": "md",
    "HTML (.html)": "html",
    "Plain Text (.txt)": "txt",
}


def load_config() -> dict:
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            # Merge: new keys added in future versions get their defaults
            return {**DEFAULTS, **data}
        except (json.JSONDecodeError, OSError):
            pass
    return dict(DEFAULTS)


def save_config(cfg: dict) -> None:
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=4)
    except OSError as e:
        messagebox.showerror("Config Error", f"Could not save settings:\n{e}")


# ---------------------------------------------------------------------------
# Helper: section header label
# ---------------------------------------------------------------------------

def section_label(parent, text: str, row: int, col: int = 0, colspan: int = 2) -> ctk.CTkLabel:
    lbl = ctk.CTkLabel(
        parent,
        text=text,
        font=ctk.CTkFont(size=13, weight="bold"),
        anchor="w",
    )
    lbl.grid(row=row, column=col, columnspan=colspan, sticky="w", padx=6, pady=(14, 2))
    return lbl


def field_label(parent, text: str, row: int, col: int = 0) -> ctk.CTkLabel:
    lbl = ctk.CTkLabel(parent, text=text, anchor="w")
    lbl.grid(row=row, column=col, sticky="w", padx=(6, 4), pady=3)
    return lbl


# ---------------------------------------------------------------------------
# Main Application
# ---------------------------------------------------------------------------

class SubstackArchiverApp(ctk.CTk):

    def __init__(self) -> None:
        super().__init__()

        self.config = load_config()
        self._is_running = False
        self._log_queue: queue.Queue = queue.Queue()
        self._cookies_visible = False
        self._cookie_val_visible = False

        self._setup_window()
        self._create_variables()
        self._build_ui()
        self._attach_traces()
        self._poll_log_queue()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        # Populate previews after UI is built
        self.after(100, self._update_download_command)
        self.after(100, self._update_epub_command)
        self.after(100, self._update_epub_files_preview)

    # ------------------------------------------------------------------
    # Window setup
    # ------------------------------------------------------------------

    def _setup_window(self) -> None:
        self.title("Substack Archiver")
        geo = self.config.get("window_geometry", DEFAULTS["window_geometry"])
        self.geometry(geo)
        self.minsize(800, 620)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0)
        self.grid_columnconfigure(0, weight=1)

    # ------------------------------------------------------------------
    # Variables
    # ------------------------------------------------------------------

    def _create_variables(self) -> None:
        cfg = self.config

        # Download tab
        self.url_var            = ctk.StringVar(value=cfg["last_url"])
        self.output_dir_var     = ctk.StringVar(value=cfg["last_output_dir"])
        self.format_var         = ctk.StringVar(value=cfg.get("last_format", "Markdown (.md)"))
        self.dates_enabled_var  = ctk.BooleanVar(value=False)
        self.after_date_var     = ctk.StringVar()
        self.before_date_var    = ctk.StringVar()
        self.dl_images_var      = ctk.BooleanVar(value=False)
        self.image_quality_var  = ctk.StringVar(value="low")
        self.images_dir_var     = ctk.StringVar(value="images")
        self.dl_files_var       = ctk.BooleanVar(value=False)
        self.file_exts_var      = ctk.StringVar()
        self.files_dir_var      = ctk.StringVar(value="files")
        self.add_source_var     = ctk.BooleanVar(value=True)
        self.create_archive_var = ctk.BooleanVar(value=False)
        self.rate_var           = ctk.StringVar(value="1")
        self.verbose_var        = ctk.BooleanVar(value=False)
        self.dry_run_var        = ctk.BooleanVar(value=False)
        self.cookie_name_var    = ctk.StringVar(value="substack.sid")
        self.cookie_val_var     = ctk.StringVar()

        # ePub tab
        self.epub_source_var    = ctk.StringVar(value=cfg["last_epub_source_dir"])
        self.epub_output_var    = ctk.StringVar(value=cfg["last_epub_output_file"])
        self.epub_title_var     = ctk.StringVar()
        self.epub_author_var    = ctk.StringVar(value=cfg["last_author"])
        self.epub_toc_var       = ctk.BooleanVar(value=True)
        self.epub_split_var     = ctk.StringVar(value="1")

        # Settings tab
        self.sbstckdl_path_var  = ctk.StringVar(value=cfg["sbstckdl_path"])
        self.pandoc_path_var    = ctk.StringVar(value=cfg["pandoc_path"])

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        # Tab view (top, expands)
        self.tabview = ctk.CTkTabview(self, corner_radius=8)
        self.tabview.grid(row=0, column=0, sticky="nsew", padx=10, pady=(10, 4))

        self.tabview.add("Download")
        self.tabview.add("ePub Conversion")
        self.tabview.add("Settings")

        for tab_name in ("Download", "ePub Conversion", "Settings"):
            tab = self.tabview.tab(tab_name)
            tab.grid_rowconfigure(0, weight=1)
            tab.grid_columnconfigure(0, weight=1)

        self._build_download_tab(self.tabview.tab("Download"))
        self._build_epub_tab(self.tabview.tab("ePub Conversion"))
        self._build_settings_tab(self.tabview.tab("Settings"))

        # Log section (bottom, fixed height)
        self._build_log_section()

    # ------------------------------------------------------------------
    # Download tab
    # ------------------------------------------------------------------

    def _build_download_tab(self, tab: ctk.CTkFrame) -> None:
        scroll = ctk.CTkScrollableFrame(tab, corner_radius=6)
        scroll.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        scroll.grid_columnconfigure(1, weight=1)
        f = scroll  # shorthand

        row = 0

        # ── Source ──────────────────────────────────────────────────
        section_label(f, "Source", row); row += 1

        field_label(f, "Substack URL:", row)
        self.url_entry = ctk.CTkEntry(f, textvariable=self.url_var,
                                      placeholder_text="https://yourname.substack.com/")
        self.url_entry.grid(row=row, column=1, columnspan=2, sticky="ew", padx=6, pady=3)
        row += 1

        # ── Destination ─────────────────────────────────────────────
        section_label(f, "Destination", row); row += 1

        field_label(f, "Output folder:", row)
        self.output_dir_entry = ctk.CTkEntry(f, textvariable=self.output_dir_var,
                                              placeholder_text="Choose a folder…")
        self.output_dir_entry.grid(row=row, column=1, sticky="ew", padx=(6, 2), pady=3)
        ctk.CTkButton(f, text="Browse…", width=80,
                      command=lambda: self._browse_folder(self.output_dir_var)
                      ).grid(row=row, column=2, padx=(2, 6), pady=3)
        row += 1

        field_label(f, "Format:", row)
        self.format_menu = ctk.CTkOptionMenu(
            f, variable=self.format_var,
            values=list(FORMAT_DISPLAY_TO_FLAG.keys()),
            command=lambda _: self._update_download_command(),
        )
        self.format_menu.grid(row=row, column=1, sticky="w", padx=6, pady=3)
        row += 1

        # ── Date Range ──────────────────────────────────────────────
        section_label(f, "Date Range (optional)", row); row += 1

        self.dates_cb = ctk.CTkCheckBox(
            f, text="Filter by date range",
            variable=self.dates_enabled_var,
            command=self._toggle_dates,
        )
        self.dates_cb.grid(row=row, column=0, columnspan=3, sticky="w", padx=6, pady=3)
        row += 1

        self.date_frame = ctk.CTkFrame(f, fg_color="transparent")
        self.date_frame.grid(row=row, column=0, columnspan=3, sticky="ew", padx=6, pady=2)
        self.date_frame.grid_columnconfigure(1, weight=0)
        self.date_frame.grid_columnconfigure(3, weight=0)

        ctk.CTkLabel(self.date_frame, text="After:").grid(row=0, column=0, sticky="w", padx=(0, 4))
        self.after_date_picker = DateEntry(
            self.date_frame, width=12, date_pattern="yyyy-mm-dd", state="disabled",
            background="gray30", foreground="white", borderwidth=2,
        )
        self.after_date_picker.grid(row=0, column=1, sticky="w", padx=(0, 16))
        self.after_date_picker.bind("<<DateEntrySelected>>",
                                    lambda _: self._update_download_command())

        ctk.CTkLabel(self.date_frame, text="Before:").grid(row=0, column=2, sticky="w", padx=(0, 4))
        self.before_date_picker = DateEntry(
            self.date_frame, width=12, date_pattern="yyyy-mm-dd", state="disabled",
            background="gray30", foreground="white", borderwidth=2,
        )
        self.before_date_picker.grid(row=0, column=3, sticky="w")
        self.before_date_picker.bind("<<DateEntrySelected>>",
                                     lambda _: self._update_download_command())

        self.date_frame.grid_remove()
        row += 1

        # ── Image Options ───────────────────────────────────────────
        section_label(f, "Image Options", row); row += 1

        self.dl_images_cb = ctk.CTkCheckBox(
            f, text="Download images locally",
            variable=self.dl_images_var,
            command=self._toggle_image_options,
        )
        self.dl_images_cb.grid(row=row, column=0, columnspan=3, sticky="w", padx=6, pady=3)
        row += 1

        self.image_options_frame = ctk.CTkFrame(f, fg_color="transparent")
        self.image_options_frame.grid(row=row, column=0, columnspan=3, sticky="ew", padx=24, pady=2)
        self.image_options_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(self.image_options_frame, text="Quality:").grid(
            row=0, column=0, sticky="w", padx=(0, 4), pady=2)
        ctk.CTkOptionMenu(
            self.image_options_frame, variable=self.image_quality_var,
            values=["low", "medium", "high"],
            width=120, command=lambda _: self._update_download_command(),
        ).grid(row=0, column=1, sticky="w", pady=2)

        ctk.CTkLabel(self.image_options_frame, text="Images subfolder:").grid(
            row=1, column=0, sticky="w", padx=(0, 4), pady=2)
        ctk.CTkEntry(self.image_options_frame, textvariable=self.images_dir_var,
                     width=160).grid(row=1, column=1, sticky="w", pady=2)

        self.image_options_frame.grid_remove()
        row += 1

        # ── File Attachments ─────────────────────────────────────────
        section_label(f, "File Attachments", row); row += 1

        self.dl_files_cb = ctk.CTkCheckBox(
            f, text="Download file attachments",
            variable=self.dl_files_var,
            command=self._toggle_file_options,
        )
        self.dl_files_cb.grid(row=row, column=0, columnspan=3, sticky="w", padx=6, pady=3)
        row += 1

        self.file_options_frame = ctk.CTkFrame(f, fg_color="transparent")
        self.file_options_frame.grid(row=row, column=0, columnspan=3, sticky="ew", padx=24, pady=2)
        self.file_options_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(self.file_options_frame, text="Extensions (blank = all):").grid(
            row=0, column=0, sticky="w", padx=(0, 4), pady=2)
        ctk.CTkEntry(self.file_options_frame, textvariable=self.file_exts_var,
                     placeholder_text="pdf,docx,mp3", width=200).grid(
            row=0, column=1, sticky="w", pady=2)

        ctk.CTkLabel(self.file_options_frame, text="Files subfolder:").grid(
            row=1, column=0, sticky="w", padx=(0, 4), pady=2)
        ctk.CTkEntry(self.file_options_frame, textvariable=self.files_dir_var,
                     width=160).grid(row=1, column=1, sticky="w", pady=2)

        self.file_options_frame.grid_remove()
        row += 1

        # ── Advanced Options ─────────────────────────────────────────
        section_label(f, "Advanced Options", row); row += 1

        ctk.CTkCheckBox(f, text="Add source URL to each post",
                        variable=self.add_source_var).grid(
            row=row, column=0, columnspan=3, sticky="w", padx=6, pady=2)
        row += 1

        ctk.CTkCheckBox(f, text="Create archive index page (index.md / index.html)",
                        variable=self.create_archive_var).grid(
            row=row, column=0, columnspan=3, sticky="w", padx=6, pady=2)
        row += 1

        ctk.CTkCheckBox(f, text="Verbose output",
                        variable=self.verbose_var).grid(
            row=row, column=0, columnspan=3, sticky="w", padx=6, pady=2)
        row += 1

        ctk.CTkCheckBox(f, text="Dry run (preview command only — no actual download)",
                        variable=self.dry_run_var).grid(
            row=row, column=0, columnspan=3, sticky="w", padx=6, pady=2)
        row += 1

        rate_row_frame = ctk.CTkFrame(f, fg_color="transparent")
        rate_row_frame.grid(row=row, column=0, columnspan=3, sticky="w", padx=6, pady=2)
        ctk.CTkLabel(rate_row_frame, text="Rate limit (requests/sec):").pack(side="left")
        ctk.CTkEntry(rate_row_frame, textvariable=self.rate_var, width=70).pack(
            side="left", padx=(8, 0))
        row += 1

        # ── Cookie Authentication (collapsible) ─────────────────────
        section_label(f, "Paid Content Authentication", row); row += 1

        ctk.CTkLabel(
            f,
            text="Only needed if downloading articles from a paid Substack you subscribe to.",
            text_color="gray60",
            anchor="w",
        ).grid(row=row, column=0, columnspan=3, sticky="w", padx=6, pady=(0, 4))
        row += 1

        self.cookie_toggle_btn = ctk.CTkButton(
            f,
            text="Show Cookie Settings",
            width=200,
            fg_color="transparent",
            border_width=1,
            command=self._toggle_cookie_section,
        )
        self.cookie_toggle_btn.grid(row=row, column=0, columnspan=3, sticky="w", padx=6, pady=4)
        row += 1

        # Cookie inner frame
        self.cookie_frame = ctk.CTkFrame(f)
        self.cookie_frame.grid(row=row, column=0, columnspan=3, sticky="ew", padx=6, pady=4)
        self.cookie_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(self.cookie_frame, text="Cookie name:", anchor="w").grid(
            row=0, column=0, sticky="w", padx=8, pady=4)
        ctk.CTkOptionMenu(
            self.cookie_frame, variable=self.cookie_name_var,
            values=["substack.sid", "connect.sid"],
            width=160,
        ).grid(row=0, column=1, sticky="w", padx=8, pady=4)

        ctk.CTkLabel(self.cookie_frame, text="Cookie value:", anchor="w").grid(
            row=1, column=0, sticky="w", padx=8, pady=4)
        self.cookie_val_entry = ctk.CTkEntry(
            self.cookie_frame,
            textvariable=self.cookie_val_var,
            show="*",
            placeholder_text="Paste your session cookie value here…",
        )
        self.cookie_val_entry.grid(row=1, column=1, sticky="ew", padx=(8, 4), pady=4)
        self.show_cookie_btn = ctk.CTkButton(
            self.cookie_frame, text="Show", width=60,
            command=self._toggle_cookie_visibility,
        )
        self.show_cookie_btn.grid(row=1, column=2, padx=(2, 8), pady=4)

        ctk.CTkLabel(
            self.cookie_frame,
            text=(
                "How to find your cookie:\n"
                "1. Log into Substack in your browser\n"
                "2. Open DevTools (F12) → Application → Cookies → substack.com\n"
                "3. Copy the value of 'substack.sid' (or 'connect.sid')"
            ),
            text_color="gray60",
            justify="left",
            anchor="w",
        ).grid(row=2, column=0, columnspan=3, sticky="w", padx=8, pady=(0, 8))

        self.cookie_frame.grid_remove()
        row += 1

        # ── Command Preview ───────────────────────────────────────────
        section_label(f, "Command Preview", row); row += 1

        self.download_cmd_preview = ctk.CTkTextbox(
            f, height=70, state="disabled", wrap="none",
            font=ctk.CTkFont(family="Courier New", size=11),
        )
        self.download_cmd_preview.grid(row=row, column=0, columnspan=3, sticky="ew",
                                       padx=6, pady=4)
        row += 1

        # ── Action button ────────────────────────────────────────────
        self.download_btn = ctk.CTkButton(
            f, text="Start Download", height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self._on_download_click,
        )
        self.download_btn.grid(row=row, column=0, columnspan=3, sticky="ew",
                               padx=6, pady=(10, 6))
        row += 1

    # ------------------------------------------------------------------
    # ePub tab
    # ------------------------------------------------------------------

    def _build_epub_tab(self, tab: ctk.CTkFrame) -> None:
        scroll = ctk.CTkScrollableFrame(tab, corner_radius=6)
        scroll.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        scroll.grid_columnconfigure(1, weight=1)
        f = scroll

        row = 0

        # ── Source ──────────────────────────────────────────────────
        section_label(f, "Source", row); row += 1

        ctk.CTkLabel(
            f,
            text="Point this to the folder where you saved your downloaded Markdown files.",
            text_color="gray60", anchor="w",
        ).grid(row=row, column=0, columnspan=3, sticky="w", padx=6, pady=(0, 4))
        row += 1

        field_label(f, "Source folder (.md files):", row)
        self.epub_source_entry = ctk.CTkEntry(f, textvariable=self.epub_source_var,
                                               placeholder_text="Folder containing .md files…")
        self.epub_source_entry.grid(row=row, column=1, sticky="ew", padx=(6, 2), pady=3)
        ctk.CTkButton(f, text="Browse…", width=80,
                      command=lambda: self._browse_folder(self.epub_source_var)
                      ).grid(row=row, column=2, padx=(2, 6), pady=3)
        row += 1

        # ── Files found preview ──────────────────────────────────────
        field_label(f, "Files found:", row)
        self.epub_files_preview = ctk.CTkTextbox(
            f, height=110, state="disabled",
            font=ctk.CTkFont(family="Courier New", size=11),
        )
        self.epub_files_preview.grid(row=row, column=1, columnspan=2, sticky="ew",
                                     padx=(6, 6), pady=3)
        row += 1

        # ── Destination ──────────────────────────────────────────────
        section_label(f, "Destination", row); row += 1

        field_label(f, "Output .epub file:", row)
        self.epub_output_entry = ctk.CTkEntry(f, textvariable=self.epub_output_var,
                                               placeholder_text="Save ePub as…")
        self.epub_output_entry.grid(row=row, column=1, sticky="ew", padx=(6, 2), pady=3)
        ctk.CTkButton(f, text="Browse…", width=80,
                      command=self._browse_output_epub,
                      ).grid(row=row, column=2, padx=(2, 6), pady=3)
        row += 1

        # ── Book Metadata ─────────────────────────────────────────────
        section_label(f, "Book Metadata", row); row += 1

        field_label(f, "Book title:", row)
        ctk.CTkEntry(f, textvariable=self.epub_title_var,
                     placeholder_text="e.g. My Substack Archive").grid(
            row=row, column=1, columnspan=2, sticky="ew", padx=6, pady=3)
        row += 1

        field_label(f, "Author:", row)
        ctk.CTkEntry(f, textvariable=self.epub_author_var,
                     placeholder_text="e.g. Jane Smith").grid(
            row=row, column=1, columnspan=2, sticky="ew", padx=6, pady=3)
        row += 1

        # ── ePub Options ─────────────────────────────────────────────
        section_label(f, "ePub Options", row); row += 1

        ctk.CTkCheckBox(f, text="Include Table of Contents",
                        variable=self.epub_toc_var).grid(
            row=row, column=0, columnspan=3, sticky="w", padx=6, pady=3)
        row += 1

        split_frame = ctk.CTkFrame(f, fg_color="transparent")
        split_frame.grid(row=row, column=0, columnspan=3, sticky="w", padx=6, pady=3)
        ctk.CTkLabel(split_frame,
                     text="Chapter split level (1 = each article is a chapter):").pack(side="left")
        ctk.CTkEntry(split_frame, textvariable=self.epub_split_var, width=60).pack(
            side="left", padx=(8, 0))
        row += 1

        # ── Command Preview ───────────────────────────────────────────
        section_label(f, "Command Preview", row); row += 1

        self.epub_cmd_preview = ctk.CTkTextbox(
            f, height=90, state="disabled", wrap="none",
            font=ctk.CTkFont(family="Courier New", size=11),
        )
        self.epub_cmd_preview.grid(row=row, column=0, columnspan=3, sticky="ew",
                                   padx=6, pady=4)
        row += 1

        # ── Action button ────────────────────────────────────────────
        self.convert_btn = ctk.CTkButton(
            f, text="Convert to ePub", height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self._on_convert_click,
        )
        self.convert_btn.grid(row=row, column=0, columnspan=3, sticky="ew",
                              padx=6, pady=(10, 6))
        row += 1

    # ------------------------------------------------------------------
    # Settings tab
    # ------------------------------------------------------------------

    def _build_settings_tab(self, tab: ctk.CTkFrame) -> None:
        scroll = ctk.CTkScrollableFrame(tab, corner_radius=6)
        scroll.grid(row=0, column=0, sticky="nsew")
        scroll.grid_columnconfigure(1, weight=1)
        f = scroll

        row = 0
        section_label(f, "Executable Paths", row); row += 1

        ctk.CTkLabel(
            f,
            text=(
                "If sbstck-dl or pandoc are on your system PATH you can leave these blank.\n"
                "Otherwise browse to the executable file."
            ),
            text_color="gray60", anchor="w", justify="left",
        ).grid(row=row, column=0, columnspan=3, sticky="w", padx=6, pady=(0, 8))
        row += 1

        field_label(f, "sbstck-dl path:", row)
        ctk.CTkEntry(f, textvariable=self.sbstckdl_path_var,
                     placeholder_text="Leave blank to use system PATH").grid(
            row=row, column=1, sticky="ew", padx=(6, 2), pady=3)
        ctk.CTkButton(f, text="Browse…", width=80,
                      command=lambda: self._browse_exe(self.sbstckdl_path_var),
                      ).grid(row=row, column=2, padx=(2, 6), pady=3)
        row += 1

        field_label(f, "pandoc path:", row)
        ctk.CTkEntry(f, textvariable=self.pandoc_path_var,
                     placeholder_text="Leave blank to use system PATH").grid(
            row=row, column=1, sticky="ew", padx=(6, 2), pady=3)
        ctk.CTkButton(f, text="Browse…", width=80,
                      command=lambda: self._browse_exe(self.pandoc_path_var),
                      ).grid(row=row, column=2, padx=(2, 6), pady=3)
        row += 1

        self.save_settings_btn = ctk.CTkButton(
            f, text="Save Settings", width=160,
            command=self._save_settings,
        )
        self.save_settings_btn.grid(row=row, column=0, columnspan=3, sticky="w",
                                    padx=6, pady=12)
        row += 1

        # ── Install guide ─────────────────────────────────────────────
        section_label(f, "First-Time Setup", row); row += 1

        guide = (
            "Install the required Python packages by opening a Command Prompt and running:\n\n"
            "    pip install customtkinter tkcalendar\n\n"
            "Install sbstck-dl:\n\n"
            "    pip install sbstck-dl\n\n"
            "After installing sbstck-dl via pip, it is usually found automatically\n"
            "(no path needed above). If it is not found, use Browse to locate it.\n\n"
            "Pandoc is installed at its default location and pre-filled above.\n"
            "If you installed it elsewhere, update the path and click Save Settings."
        )
        guide_box = ctk.CTkTextbox(f, height=220, wrap="word", state="normal",
                                   font=ctk.CTkFont(family="Courier New", size=11))
        guide_box.grid(row=row, column=0, columnspan=3, sticky="ew", padx=6, pady=4)
        guide_box.insert("1.0", guide)
        guide_box.configure(state="disabled")
        row += 1

    # ------------------------------------------------------------------
    # Log section (always visible)
    # ------------------------------------------------------------------

    def _build_log_section(self) -> None:
        log_frame = ctk.CTkFrame(self, corner_radius=8)
        log_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 10))
        log_frame.grid_columnconfigure(0, weight=1)

        header = ctk.CTkFrame(log_frame, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=8, pady=(6, 0))
        header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(header, text="Output Log",
                     font=ctk.CTkFont(size=13, weight="bold")).grid(
            row=0, column=0, sticky="w")
        ctk.CTkButton(header, text="Clear Log", width=90,
                      command=self._clear_log).grid(row=0, column=1, sticky="e")

        self.log_box = ctk.CTkTextbox(
            log_frame, height=160, state="disabled", wrap="word",
            font=ctk.CTkFont(family="Courier New", size=11),
        )
        self.log_box.grid(row=1, column=0, sticky="ew", padx=8, pady=(4, 8))

    # ------------------------------------------------------------------
    # Traces (live command preview)
    # ------------------------------------------------------------------

    def _attach_traces(self) -> None:
        dl_vars = [
            self.url_var, self.output_dir_var, self.format_var,
            self.after_date_var, self.before_date_var,
            self.images_dir_var, self.files_dir_var,
            self.file_exts_var, self.rate_var,
            self.cookie_name_var, self.cookie_val_var,
            self.sbstckdl_path_var,
        ]
        for v in dl_vars:
            v.trace_add("write", lambda *_: self._update_download_command())

        bool_dl_vars = [
            self.dates_enabled_var, self.dl_images_var, self.image_quality_var,
            self.dl_files_var, self.add_source_var, self.create_archive_var,
            self.verbose_var, self.dry_run_var,
        ]
        for v in bool_dl_vars:
            v.trace_add("write", lambda *_: self._update_download_command())

        epub_vars = [
            self.epub_source_var, self.epub_output_var,
            self.epub_title_var, self.epub_author_var,
            self.epub_split_var, self.pandoc_path_var,
        ]
        for v in epub_vars:
            v.trace_add("write", lambda *_: self._update_epub_command())

        self.epub_toc_var.trace_add("write", lambda *_: self._update_epub_command())
        self.epub_source_var.trace_add("write", lambda *_: self._update_epub_files_preview())

    # ------------------------------------------------------------------
    # Command builders
    # ------------------------------------------------------------------

    def _build_download_cmd(self) -> list:
        exe = self.sbstckdl_path_var.get().strip() or "sbstck-dl"
        cmd = [exe, "download"]

        url = self.url_var.get().strip()
        if url:
            cmd += ["--url", url]

        out = self.output_dir_var.get().strip()
        if out:
            cmd += ["-o", out]

        fmt = FORMAT_DISPLAY_TO_FLAG.get(self.format_var.get(), "md")
        cmd += ["-f", fmt]

        if self.dates_enabled_var.get():
            after = self.after_date_picker.get().strip()
            before = self.before_date_picker.get().strip()
            if after:
                cmd += ["--after", after]
            if before:
                cmd += ["--before", before]

        if self.dl_images_var.get():
            cmd.append("--download-images")
            cmd += ["--image-quality", self.image_quality_var.get()]
            images_dir = self.images_dir_var.get().strip()
            if images_dir and images_dir != "images":
                cmd += ["--images-dir", images_dir]

        if self.dl_files_var.get():
            cmd.append("--download-files")
            exts = self.file_exts_var.get().strip()
            if exts:
                cmd += ["--file-extensions", exts]
            files_dir = self.files_dir_var.get().strip()
            if files_dir and files_dir != "files":
                cmd += ["--files-dir", files_dir]

        if self.add_source_var.get():
            cmd.append("--add-source-url")

        if self.create_archive_var.get():
            cmd.append("--create-archive")

        rate = self.rate_var.get().strip()
        if rate and rate != "1":
            try:
                cmd += ["-r", str(float(rate))]
            except ValueError:
                pass

        if self.verbose_var.get():
            cmd.append("-v")

        if self.dry_run_var.get():
            cmd.append("-d")

        cookie_val = self.cookie_val_var.get().strip()
        if cookie_val:
            cmd += ["--cookie_name", self.cookie_name_var.get()]
            cmd += ["--cookie_val", cookie_val]

        return cmd

    def _build_epub_cmd(self) -> list | None:
        source_dir = self.epub_source_var.get().strip()
        if not source_dir or not os.path.isdir(source_dir):
            return None

        md_files = self._get_md_files(source_dir)
        if not md_files:
            return None

        pandoc_exe = self.pandoc_path_var.get().strip() or "pandoc"
        output = self.epub_output_var.get().strip()
        title = self.epub_title_var.get().strip() or "Substack Archive"
        author = self.epub_author_var.get().strip() or "Unknown"
        split = self.epub_split_var.get().strip() or "1"

        cmd = [pandoc_exe]
        cmd += [os.path.join(source_dir, f) for f in md_files]
        if output:
            cmd += ["-o", output]
        cmd += ["--metadata", f"title={title}"]
        cmd += ["--metadata", f"author={author}"]
        if self.epub_toc_var.get():
            cmd.append("--toc")
        cmd.append(f"--split-level={split}")

        return cmd

    @staticmethod
    def _get_md_files(source_dir: str) -> list:
        try:
            files = sorted([
                f for f in os.listdir(source_dir)
                if f.lower().endswith(".md") and f.lower() != "index.md"
            ])
        except OSError:
            files = []
        return files

    @staticmethod
    def _cmd_to_display_string(cmd: list) -> str:
        parts = []
        for part in cmd:
            if " " in part or not part or any(c in part for c in '&|<>()"'):
                parts.append(f'"{part}"')
            else:
                parts.append(part)
        return " ".join(parts)

    # ------------------------------------------------------------------
    # Command preview update
    # ------------------------------------------------------------------

    def _update_download_command(self, *_) -> None:
        cmd = self._build_download_cmd()
        text = self._cmd_to_display_string(cmd)
        self._set_textbox(self.download_cmd_preview, text)

    def _update_epub_command(self, *_) -> None:
        cmd = self._build_epub_cmd()
        if cmd is None:
            text = "(Select a source folder containing .md files and an output path)"
        else:
            source_dir = self.epub_source_var.get().strip()
            md_files = self._get_md_files(source_dir) if source_dir else []
            count = len(md_files)

            # Build abbreviated display — show exe, first file, "... N files total", then options
            pandoc_exe = self.pandoc_path_var.get().strip() or "pandoc"
            output = self.epub_output_var.get().strip() or "<output.epub>"
            title = self.epub_title_var.get().strip() or "Substack Archive"
            author = self.epub_author_var.get().strip() or "Unknown"
            split = self.epub_split_var.get().strip() or "1"

            lines = [f"{pandoc_exe}"]
            if md_files:
                first = os.path.join(source_dir, md_files[0])
                lines.append(f'  "{first}"')
                if count > 1:
                    lines.append(f"  ... ({count} .md files total, sorted by name)")
            lines.append(f'  -o "{output}"')
            lines.append(f'  --metadata title="{title}"')
            lines.append(f'  --metadata author="{author}"')
            if self.epub_toc_var.get():
                lines.append("  --toc")
            lines.append(f"  --split-level={split}")
            text = " \\\n".join(lines)

        self._set_textbox(self.epub_cmd_preview, text)

    def _update_epub_files_preview(self, *_) -> None:
        source = self.epub_source_var.get().strip()
        if not source:
            self._set_textbox(self.epub_files_preview, "(No folder selected)")
            return
        if not os.path.isdir(source):
            self._set_textbox(self.epub_files_preview, "(Folder does not exist)")
            return

        md_files = self._get_md_files(source)
        if not md_files:
            self._set_textbox(
                self.epub_files_preview,
                "No .md files found (excluding index.md).\n"
                "Make sure you downloaded using Markdown format first.",
            )
        else:
            body = f"Found {len(md_files)} file(s):\n" + "\n".join(md_files)
            self._set_textbox(self.epub_files_preview, body)

        self._update_epub_command()

    @staticmethod
    def _set_textbox(box: ctk.CTkTextbox, text: str) -> None:
        box.configure(state="normal")
        box.delete("1.0", "end")
        box.insert("end", text)
        box.configure(state="disabled")

    # ------------------------------------------------------------------
    # Toggle helpers
    # ------------------------------------------------------------------

    def _toggle_dates(self) -> None:
        if self.dates_enabled_var.get():
            self.date_frame.grid()
        else:
            self.date_frame.grid_remove()
        self._update_download_command()

    def _toggle_image_options(self) -> None:
        if self.dl_images_var.get():
            self.image_options_frame.grid()
        else:
            self.image_options_frame.grid_remove()
        self._update_download_command()

    def _toggle_file_options(self) -> None:
        if self.dl_files_var.get():
            self.file_options_frame.grid()
        else:
            self.file_options_frame.grid_remove()
        self._update_download_command()

    def _toggle_cookie_section(self) -> None:
        self._cookies_visible = not self._cookies_visible
        if self._cookies_visible:
            self.cookie_frame.grid()
            self.cookie_toggle_btn.configure(text="Hide Cookie Settings")
        else:
            self.cookie_frame.grid_remove()
            self.cookie_toggle_btn.configure(text="Show Cookie Settings")
        self._update_download_command()

    def _toggle_cookie_visibility(self) -> None:
        self._cookie_val_visible = not self._cookie_val_visible
        self.cookie_val_entry.configure(show="" if self._cookie_val_visible else "*")
        self.show_cookie_btn.configure(
            text="Hide" if self._cookie_val_visible else "Show"
        )

    # ------------------------------------------------------------------
    # File / folder dialogs
    # ------------------------------------------------------------------

    def _browse_folder(self, var: ctk.StringVar) -> None:
        initial = var.get().strip() or os.path.expanduser("~")
        path = filedialog.askdirectory(initialdir=initial, title="Select Folder")
        if path:
            var.set(path)

    def _browse_output_epub(self) -> None:
        initial_dir = os.path.dirname(self.epub_output_var.get().strip()) or os.path.expanduser("~")
        path = filedialog.asksaveasfilename(
            initialdir=initial_dir,
            title="Save ePub As",
            defaultextension=".epub",
            filetypes=[("ePub files", "*.epub"), ("All files", "*.*")],
        )
        if path:
            self.epub_output_var.set(path)

    def _browse_exe(self, var: ctk.StringVar) -> None:
        initial = os.path.dirname(var.get().strip()) or r"C:\Program Files"
        path = filedialog.askopenfilename(
            initialdir=initial,
            title="Select Executable",
            filetypes=[("Executables", "*.exe"), ("All files", "*.*")],
        )
        if path:
            var.set(path)

    # ------------------------------------------------------------------
    # Validation
    # ------------------------------------------------------------------

    def _validate_download(self) -> str | None:
        url = self.url_var.get().strip()
        if not url:
            return "Substack URL is required."
        if not url.startswith("http"):
            return "URL must start with http:// or https://"

        out = self.output_dir_var.get().strip()
        if not out:
            return "Please select an output folder."

        rate = self.rate_var.get().strip()
        if rate:
            try:
                r = float(rate)
                if r <= 0:
                    return "Rate must be a positive number."
            except ValueError:
                return "Rate must be a number (e.g. 1 or 0.5)."

        if self.dates_enabled_var.get():
            date_pattern = re.compile(r"^\d{4}-\d{2}-\d{2}$")
            after = self.after_date_picker.get().strip()
            before = self.before_date_picker.get().strip()
            if after and not date_pattern.match(after):
                return "After date must be in YYYY-MM-DD format."
            if before and not date_pattern.match(before):
                return "Before date must be in YYYY-MM-DD format."

        return None

    def _validate_epub(self) -> str | None:
        source = self.epub_source_var.get().strip()
        if not source:
            return "Please select a source folder containing .md files."
        if not os.path.isdir(source):
            return f"Source folder does not exist:\n{source}"

        md_files = self._get_md_files(source)
        if not md_files:
            return (
                "No .md files found in the source folder (excluding index.md).\n"
                "Make sure you downloaded in Markdown format first."
            )

        output = self.epub_output_var.get().strip()
        if not output:
            return "Please choose an output .epub file path."
        if not output.lower().endswith(".epub"):
            return "Output file must have a .epub extension."

        return None

    # ------------------------------------------------------------------
    # Action handlers
    # ------------------------------------------------------------------

    def _on_download_click(self) -> None:
        if self._is_running:
            self._log("[WARNING] A process is already running. Please wait.\n")
            return

        err = self._validate_download()
        if err:
            messagebox.showerror("Validation Error", err)
            return

        cmd = self._build_download_cmd()
        self._start_command(cmd, self.download_btn, "Start Download", "Downloading…")

    def _on_convert_click(self) -> None:
        if self._is_running:
            self._log("[WARNING] A process is already running. Please wait.\n")
            return

        err = self._validate_epub()
        if err:
            messagebox.showerror("Validation Error", err)
            return

        cmd = self._build_epub_cmd()
        if cmd is None:
            messagebox.showerror("Error", "Could not build pandoc command.")
            return

        self._start_command(cmd, self.convert_btn, "Convert to ePub", "Converting…")

    def _start_command(self, cmd: list, btn: ctk.CTkButton,
                       normal_text: str, running_text: str) -> None:
        self._is_running = True
        self._active_btn = btn
        self._active_btn_normal_text = normal_text
        btn.configure(state="disabled", text=running_text)
        t = threading.Thread(target=self._run_command, args=(cmd,), daemon=True)
        t.start()

    # ------------------------------------------------------------------
    # Background command execution
    # ------------------------------------------------------------------

    def _run_command(self, cmd: list) -> None:
        display = self._cmd_to_display_string(cmd)
        self._log(f"\n{'─' * 60}\n Running: {display}\n{'─' * 60}\n\n")

        kwargs: dict = {
            "stdout": subprocess.PIPE,
            "stderr": subprocess.STDOUT,
            "text": True,
            "encoding": "utf-8",
            "errors": "replace",
        }
        if sys.platform == "win32":
            kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW

        try:
            process = subprocess.Popen(cmd, **kwargs)
            for line in process.stdout:
                self._log(line)
            process.wait()
            if process.returncode == 0:
                self._log(f"\n✓ Completed successfully (exit code 0)\n")
            else:
                self._log(f"\n✗ Process exited with code {process.returncode}\n")
        except FileNotFoundError:
            self._log(f"\n[ERROR] Could not find executable: {cmd[0]}\n")
            self._log("  → Check the Settings tab and make sure the path is correct.\n")
            self._log(f"  → If using sbstck-dl, install it with:  pip install sbstck-dl\n")
        except Exception as exc:
            self._log(f"\n[ERROR] {exc}\n")
        finally:
            self.after(0, self._on_command_finished)

    def _on_command_finished(self) -> None:
        self._is_running = False
        if hasattr(self, "_active_btn"):
            self._active_btn.configure(
                state="normal", text=self._active_btn_normal_text
            )

    # ------------------------------------------------------------------
    # Logging (thread-safe via queue)
    # ------------------------------------------------------------------

    def _poll_log_queue(self) -> None:
        try:
            while True:
                message = self._log_queue.get_nowait()
                self._append_to_log(message)
        except queue.Empty:
            pass
        self.after(100, self._poll_log_queue)

    def _log(self, message: str) -> None:
        """Thread-safe: may be called from any thread."""
        self._log_queue.put(message)

    def _append_to_log(self, message: str) -> None:
        """Must only be called from the main thread."""
        self.log_box.configure(state="normal")
        self.log_box.insert("end", message)
        # Trim if log exceeds 2000 lines to prevent slowdown on very large downloads
        line_count = int(self.log_box.index("end-1c").split(".")[0])
        if line_count > 2000:
            self.log_box.delete("1.0", f"{line_count - 2000}.0")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _clear_log(self) -> None:
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

    # ------------------------------------------------------------------
    # Settings
    # ------------------------------------------------------------------

    def _save_settings(self) -> None:
        self.config["sbstckdl_path"] = self.sbstckdl_path_var.get().strip()
        self.config["pandoc_path"] = self.pandoc_path_var.get().strip()
        save_config(self.config)
        self._log("[Settings saved]\n")
        self.save_settings_btn.configure(text="Saved!")
        self.after(2000, lambda: self.save_settings_btn.configure(text="Save Settings"))

    # ------------------------------------------------------------------
    # Shutdown
    # ------------------------------------------------------------------

    def _on_close(self) -> None:
        self.config["window_geometry"] = self.geometry()
        self.config["last_url"] = self.url_var.get()
        self.config["last_output_dir"] = self.output_dir_var.get()
        self.config["last_format"] = self.format_var.get()
        self.config["last_epub_source_dir"] = self.epub_source_var.get()
        self.config["last_epub_output_file"] = self.epub_output_var.get()
        self.config["last_author"] = self.epub_author_var.get()
        # Note: cookie value is intentionally NOT saved (security)
        save_config(self.config)
        self.destroy()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    ctk.set_appearance_mode("System")   # Follows Windows dark/light mode
    ctk.set_default_color_theme("blue")

    app = SubstackArchiverApp()
    app.mainloop()

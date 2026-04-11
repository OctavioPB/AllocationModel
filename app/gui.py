"""
gui.py
------
AllocationModel desktop GUI — built with CustomTkinter.

Three-screen wizard flow:
    Screen 1  Load Data    — file picker, data preview, ML auto-classify
    Screen 2  Configure    — solver parameters, run optimisation, live log
    Screen 3  Results      — allocation table, metrics, What-If analysis, export

Navigation is frame-swap (all screens live in one window).
The optimiser and sensitivity analysis run in background threads so the UI
remains responsive during long solves.
"""

from __future__ import annotations

import os
import threading
import tkinter as tk
import tkinter.filedialog as fd
import tkinter.messagebox as mb
from pathlib import Path
from typing import Optional

import customtkinter as ctk
import pandas as pd

from app.data_loader import AllocationInput, AllocationResult, load_file, get_raw_dataframes
from app.ml_classifier import classify_deals, apply_classification, ClassificationSummary
from app.optimizer import optimize
from app.sensitivity import compute_sensitivity, what_if, SensitivityModel
from app.exporter import export_csv, export_excel

# ---------------------------------------------------------------------------
# Theme
# ---------------------------------------------------------------------------
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

_FONT_TITLE   = ("Segoe UI", 20, "bold")
_FONT_HEADING = ("Segoe UI", 14, "bold")
_FONT_BODY    = ("Segoe UI", 13)
_FONT_MONO    = ("Courier New", 11)
_FONT_SMALL   = ("Segoe UI", 11)

_COLOR_SUCCESS  = "#2fa84f"
_COLOR_WARNING  = "#e07b00"
_COLOR_ERROR    = "#d63031"
_COLOR_NEUTRAL  = "#888888"
_COLOR_UNALLOC  = "#e17055"

_STEP_LABELS = ["  1  Load Data  ", "  2  Configure  ", "  3  Results  "]


# ---------------------------------------------------------------------------
# Reusable widgets
# ---------------------------------------------------------------------------

class _SectionLabel(ctk.CTkLabel):
    def __init__(self, parent, text: str, **kwargs):
        super().__init__(parent, text=text, font=_FONT_HEADING, anchor="w", **kwargs)


class _Divider(ctk.CTkFrame):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, height=2, fg_color=("gray80", "gray30"), **kwargs)


class _ScrollableTable(ctk.CTkScrollableFrame):
    """Simple read-only table built from a pandas DataFrame."""

    def __init__(self, parent, df: pd.DataFrame, highlight_col: Optional[str] = None, **kwargs):
        super().__init__(parent, **kwargs)
        self._render(df, highlight_col)

    def _render(self, df: pd.DataFrame, highlight_col: Optional[str]):
        # Header row
        for col_idx, col in enumerate(df.columns):
            ctk.CTkLabel(
                self, text=str(col), font=_FONT_BODY,
                fg_color=("gray85", "gray25"),
                corner_radius=4, padx=8, pady=4, anchor="w",
            ).grid(row=0, column=col_idx, sticky="ew", padx=2, pady=2)

        # Data rows
        for row_idx, row in df.iterrows():
            for col_idx, col in enumerate(df.columns):
                val  = row[col]
                is_unalloc = highlight_col and col == highlight_col and str(val) == "Unallocated"
                text_color = _COLOR_UNALLOC if is_unalloc else ("gray10", "gray90")
                ctk.CTkLabel(
                    self, text=str(val), font=_FONT_SMALL,
                    text_color=text_color, anchor="w", padx=8,
                ).grid(row=row_idx + 1, column=col_idx, sticky="ew", padx=2, pady=1)

        for col_idx in range(len(df.columns)):
            self.grid_columnconfigure(col_idx, weight=1)


class _StepBar(ctk.CTkFrame):
    """Top progress bar showing the 3 wizard steps."""

    def __init__(self, parent, **kwargs):
        super().__init__(parent, fg_color="transparent", **kwargs)
        self._buttons: list[ctk.CTkButton] = []
        for i, label in enumerate(_STEP_LABELS):
            btn = ctk.CTkButton(
                self, text=label, font=_FONT_SMALL,
                state="disabled", fg_color="transparent",
                border_width=1, corner_radius=6,
                text_color=("gray60", "gray50"),
            )
            btn.grid(row=0, column=i, padx=6, pady=6)
            self._buttons.append(btn)
        self.grid_columnconfigure((0, 1, 2), weight=1)

    def set_active(self, step: int):
        """Highlight the active step (0-indexed)."""
        for i, btn in enumerate(self._buttons):
            if i == step:
                btn.configure(
                    fg_color=("gray20", "gray80"),
                    text_color=("white", "gray10"),
                    border_color=("gray20", "gray80"),
                )
            elif i < step:
                btn.configure(
                    fg_color="transparent",
                    text_color=(_COLOR_SUCCESS, _COLOR_SUCCESS),
                    border_color=(_COLOR_SUCCESS, _COLOR_SUCCESS),
                )
            else:
                btn.configure(
                    fg_color="transparent",
                    text_color=("gray60", "gray50"),
                    border_color=("gray70", "gray40"),
                )


# ---------------------------------------------------------------------------
# Screen 1 — Load Data
# ---------------------------------------------------------------------------

class Screen1(ctk.CTkFrame):
    """File loading, data preview, and ML auto-classification."""

    def __init__(self, parent: "App", **kwargs):
        super().__init__(parent, fg_color="transparent", **kwargs)
        self._app = parent
        self._deals_path: Optional[Path] = None
        self._purchasers_path: Optional[Path] = None
        self._deals_df: Optional[pd.DataFrame] = None
        self._purchasers_df: Optional[pd.DataFrame] = None
        self._ml_summary: Optional[ClassificationSummary] = None
        self._build()

    def _build(self):
        self.grid_columnconfigure(0, weight=1)

        # ---- File section ----
        _SectionLabel(self, text="Load Input Files").grid(
            row=0, column=0, sticky="w", padx=20, pady=(20, 4))
        _Divider(self).grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 12))

        file_frame = ctk.CTkFrame(self, fg_color="transparent")
        file_frame.grid(row=2, column=0, sticky="ew", padx=20, pady=4)
        file_frame.grid_columnconfigure(1, weight=1)

        # Mode toggle: CSV or Excel
        ctk.CTkLabel(file_frame, text="Format:", font=_FONT_BODY).grid(
            row=0, column=0, sticky="w", padx=(0, 10))
        self._format_var = ctk.StringVar(value="csv")
        ctk.CTkSegmentedButton(
            file_frame, values=["CSV (2 files)", "Excel (1 file)"],
            variable=self._format_var,
            command=self._on_format_change, font=_FONT_SMALL,
        ).grid(row=0, column=1, sticky="w")

        # Deals row
        self._deals_lbl = ctk.CTkLabel(file_frame, text="Deals file:", font=_FONT_BODY)
        self._deals_lbl.grid(row=1, column=0, sticky="w", padx=(0, 10), pady=(10, 4))
        self._deals_path_var = ctk.StringVar(value="No file selected")
        ctk.CTkLabel(file_frame, textvariable=self._deals_path_var,
                     font=_FONT_SMALL, text_color=("gray50", "gray60")).grid(
            row=1, column=1, sticky="w")
        ctk.CTkButton(file_frame, text="Browse", width=90, font=_FONT_SMALL,
                      command=self._browse_deals).grid(row=1, column=2, padx=(10, 0))

        # Purchasers row (CSV only)
        self._purch_lbl = ctk.CTkLabel(file_frame, text="Purchasers file:", font=_FONT_BODY)
        self._purch_lbl.grid(row=2, column=0, sticky="w", padx=(0, 10), pady=4)
        self._purch_path_var = ctk.StringVar(value="No file selected")
        ctk.CTkLabel(file_frame, textvariable=self._purch_path_var,
                     font=_FONT_SMALL, text_color=("gray50", "gray60")).grid(
            row=2, column=1, sticky="w")
        ctk.CTkButton(file_frame, text="Browse", width=90, font=_FONT_SMALL,
                      command=self._browse_purchasers).grid(row=2, column=2, padx=(10, 0))

        ctk.CTkButton(
            file_frame, text="Load & Preview", font=_FONT_BODY,
            command=self._load_files,
        ).grid(row=3, column=0, columnspan=3, sticky="w", pady=(12, 0))

        # Status label
        self._status_var = ctk.StringVar(value="")
        self._status_lbl = ctk.CTkLabel(self, textvariable=self._status_var,
                                         font=_FONT_SMALL, anchor="w")
        self._status_lbl.grid(row=3, column=0, sticky="w", padx=20, pady=(6, 0))

        # ---- Preview section ----
        _SectionLabel(self, text="Data Preview").grid(
            row=4, column=0, sticky="w", padx=20, pady=(20, 4))
        _Divider(self).grid(row=5, column=0, sticky="ew", padx=20, pady=(0, 8))

        self._preview_frame = ctk.CTkFrame(self, fg_color="transparent")
        self._preview_frame.grid(row=6, column=0, sticky="nsew", padx=20)
        self._preview_frame.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(6, weight=1)

        self._preview_placeholder = ctk.CTkLabel(
            self._preview_frame,
            text="Load a file to see a preview.",
            font=_FONT_SMALL, text_color=("gray55", "gray55"),
        )
        self._preview_placeholder.grid(row=0, column=0, pady=30)

        # ---- ML section ----
        ml_frame = ctk.CTkFrame(self, fg_color="transparent")
        ml_frame.grid(row=7, column=0, sticky="ew", padx=20, pady=(16, 0))
        ml_frame.grid_columnconfigure(1, weight=1)

        self._ml_btn = ctk.CTkButton(
            ml_frame, text="Auto-classify Deals with ML",
            font=_FONT_BODY, state="disabled",
            command=self._run_ml,
        )
        self._ml_btn.grid(row=0, column=0, padx=(0, 16))

        self._ml_status_var = ctk.StringVar(value="")
        ctk.CTkLabel(ml_frame, textvariable=self._ml_status_var,
                     font=_FONT_SMALL, anchor="w").grid(row=0, column=1, sticky="w")

        # ---- Navigation ----
        nav = ctk.CTkFrame(self, fg_color="transparent")
        nav.grid(row=8, column=0, sticky="e", padx=20, pady=20)
        self._next_btn = ctk.CTkButton(
            nav, text="Next: Configure  →", font=_FONT_BODY,
            state="disabled", command=self._go_next,
        )
        self._next_btn.grid(row=0, column=0)

    # ---- Event handlers ----

    def _on_format_change(self, value: str):
        is_csv = value == "CSV (2 files)"
        state = "normal" if is_csv else "disabled"
        self._purch_lbl.configure(text_color=("gray10", "gray90") if is_csv else ("gray60", "gray50"))

    def _browse_deals(self):
        fmt = self._format_var.get()
        if fmt == "CSV (2 files)":
            filetypes = [("CSV files", "*.csv")]
        else:
            filetypes = [("Excel files", "*.xlsx *.xlsm")]
        path = fd.askopenfilename(filetypes=filetypes)
        if path:
            self._deals_path = Path(path)
            self._deals_path_var.set(self._deals_path.name)

    def _browse_purchasers(self):
        path = fd.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if path:
            self._purchasers_path = Path(path)
            self._purch_path_var.set(self._purchasers_path.name)

    def _load_files(self):
        fmt = self._format_var.get()
        try:
            if fmt == "CSV (2 files)":
                if not self._deals_path:
                    self._set_status("Please select the deals CSV file.", error=True)
                    return
                if not self._purchasers_path:
                    self._set_status("Please select the purchasers CSV file.", error=True)
                    return
                self._deals_df, self._purchasers_df = get_raw_dataframes(
                    self._deals_path, self._purchasers_path)
            else:
                if not self._deals_path:
                    self._set_status("Please select the Excel file.", error=True)
                    return
                self._deals_df, self._purchasers_df = get_raw_dataframes(self._deals_path)

            self._set_status(
                f"Loaded {len(self._deals_df)} deals and {len(self._purchasers_df)} purchasers.",
                error=False,
            )
            self._render_preview()
            self._ml_btn.configure(state="normal")
            self._next_btn.configure(state="normal")

        except Exception as exc:
            self._set_status(f"Error loading file: {exc}", error=True)

    def _render_preview(self):
        for widget in self._preview_frame.winfo_children():
            widget.destroy()

        # Deals preview
        ctk.CTkLabel(self._preview_frame, text=f"Deals ({len(self._deals_df)} rows)",
                     font=_FONT_SMALL, text_color=_COLOR_NEUTRAL).grid(
            row=0, column=0, sticky="w", pady=(0, 4))
        _ScrollableTable(
            self._preview_frame,
            df=self._deals_df.head(20),
            height=160,
        ).grid(row=1, column=0, sticky="ew", pady=(0, 12))

        # Purchasers preview
        ctk.CTkLabel(self._preview_frame,
                     text=f"Purchasers ({len(self._purchasers_df)} rows)",
                     font=_FONT_SMALL, text_color=_COLOR_NEUTRAL).grid(
            row=2, column=0, sticky="w", pady=(0, 4))
        _ScrollableTable(
            self._preview_frame,
            df=self._purchasers_df,
            height=160,
        ).grid(row=3, column=0, sticky="ew")

        self._preview_frame.grid_columnconfigure(0, weight=1)

    def _run_ml(self):
        if self._deals_df is None:
            return
        df = self._deals_df.copy()
        df.columns = [c.strip().lower() for c in df.columns]

        deal_ids   = df["deal_id"].astype(str).tolist()
        deal_vals  = df["deal_value"].astype(float).tolist()
        orig_types = df["deal_type"].astype(str).tolist() if "deal_type" in df.columns else None

        self._ml_summary = classify_deals(deal_ids, deal_vals, orig_types)
        s = self._ml_summary

        # Update deal_type column in the preview dataframe
        suggestions = {r.deal_id: r.suggested_type for r in s.results}
        self._deals_df["deal_type"] = self._deals_df.apply(
            lambda row: suggestions.get(str(row.get("deal_id", row.get("Deal ID", ""))).strip(),
                                        row.get("deal_type", row.get("Deal Type", "Prepay"))),
            axis=1,
        )
        self._render_preview()

        msg = (
            f"ML classified {s.prepay_count} Prepay / {s.ppa_count} PPA. "
            f"{s.override_count} override(s) suggested."
        )
        if s.warning:
            msg += f"  Note: {s.warning}"
        self._ml_status_var.set(msg)

    def _set_status(self, msg: str, error: bool = False):
        self._status_var.set(msg)
        color = _COLOR_ERROR if error else _COLOR_SUCCESS
        self._status_lbl.configure(text_color=(color, color))

    def _go_next(self):
        try:
            fmt = self._format_var.get()
            if fmt == "CSV (2 files)":
                data = load_file(self._deals_path, self._purchasers_path)
            else:
                data = load_file(self._deals_path)

            # Apply ML suggestions if classifier was run
            if self._ml_summary:
                data.deals_type = apply_classification(
                    data.deals, data.deals_type, self._ml_summary)

            self._app.ctx.data = data
            self._app.go_to(1)

        except Exception as exc:
            self._set_status(f"Validation error: {exc}", error=True)


# ---------------------------------------------------------------------------
# Screen 2 — Configure & Run
# ---------------------------------------------------------------------------

class Screen2(ctk.CTkFrame):
    """Solver configuration and run screen with live log output."""

    def __init__(self, parent: "App", **kwargs):
        super().__init__(parent, fg_color="transparent", **kwargs)
        self._app = parent
        self._build()

    def _build(self):
        self.grid_columnconfigure(0, weight=1)

        # ---- Parameters ----
        _SectionLabel(self, text="Optimiser Settings").grid(
            row=0, column=0, sticky="w", padx=20, pady=(20, 4))
        _Divider(self).grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 16))

        params = ctk.CTkFrame(self, fg_color="transparent")
        params.grid(row=2, column=0, sticky="ew", padx=20)
        params.grid_columnconfigure(1, weight=1)

        # Time limit
        ctk.CTkLabel(params, text="Time limit (seconds):", font=_FONT_BODY).grid(
            row=0, column=0, sticky="w", padx=(0, 16), pady=8)
        self._time_var = ctk.IntVar(value=60)
        self._time_display = ctk.CTkLabel(params, text="60 s", font=_FONT_BODY, width=50)
        self._time_display.grid(row=0, column=2, padx=(8, 0))
        ctk.CTkSlider(
            params, from_=10, to=300, number_of_steps=29,
            variable=self._time_var,
            command=lambda v: self._time_display.configure(text=f"{int(v)} s"),
        ).grid(row=0, column=1, sticky="ew")

        # Min deal toggle
        ctk.CTkLabel(params, text="Each purchaser gets at least 1 deal:", font=_FONT_BODY).grid(
            row=1, column=0, sticky="w", padx=(0, 16), pady=8)
        self._min_deal_var = ctk.BooleanVar(value=True)
        ctk.CTkSwitch(params, text="", variable=self._min_deal_var).grid(
            row=1, column=1, sticky="w")

        # Pref penalty toggle
        ctk.CTkLabel(params, text="Penalise preference mismatches:", font=_FONT_BODY).grid(
            row=2, column=0, sticky="w", padx=(0, 16), pady=8)
        self._pref_var = ctk.BooleanVar(value=True)
        ctk.CTkSwitch(params, text="", variable=self._pref_var).grid(
            row=2, column=1, sticky="w")

        # ---- Run button ----
        run_row = ctk.CTkFrame(self, fg_color="transparent")
        run_row.grid(row=3, column=0, sticky="w", padx=20, pady=(20, 8))
        self._run_btn = ctk.CTkButton(
            run_row, text="Run Optimisation", font=_FONT_HEADING,
            width=200, height=44, command=self._run,
        )
        self._run_btn.grid(row=0, column=0)
        self._spinner_lbl = ctk.CTkLabel(run_row, text="", font=_FONT_SMALL, width=200)
        self._spinner_lbl.grid(row=0, column=1, padx=16)

        # ---- Log ----
        _SectionLabel(self, text="Solver Log").grid(
            row=4, column=0, sticky="w", padx=20, pady=(16, 4))
        _Divider(self).grid(row=5, column=0, sticky="ew", padx=20, pady=(0, 8))

        self._log = ctk.CTkTextbox(self, font=_FONT_MONO, state="disabled", height=240)
        self._log.grid(row=6, column=0, sticky="nsew", padx=20)
        self.grid_rowconfigure(6, weight=1)

        # ---- Navigation ----
        nav = ctk.CTkFrame(self, fg_color="transparent")
        nav.grid(row=7, column=0, sticky="ew", padx=20, pady=20)
        nav.grid_columnconfigure(1, weight=1)
        ctk.CTkButton(nav, text="← Back", font=_FONT_BODY,
                      fg_color="transparent", border_width=1,
                      command=lambda: self._app.go_to(0)).grid(row=0, column=0)
        self._next_btn = ctk.CTkButton(
            nav, text="View Results  →", font=_FONT_BODY,
            state="disabled", command=lambda: self._app.go_to(2),
        )
        self._next_btn.grid(row=0, column=2)

    def _log_append(self, msg: str):
        self._log.configure(state="normal")
        self._log.insert("end", msg + "\n")
        self._log.see("end")
        self._log.configure(state="disabled")

    def _run(self):
        data = self._app.ctx.data
        if data is None:
            mb.showerror("No data", "Please load data first.")
            return

        data.min_deal    = self._min_deal_var.get()
        data.pref_penalty = self._pref_var.get()
        time_limit       = int(self._time_var.get())

        self._run_btn.configure(state="disabled")
        self._next_btn.configure(state="disabled")
        self._spinner_lbl.configure(text="Solving...", text_color=_COLOR_WARNING)
        self._log_append(f"Starting optimisation  (time limit: {time_limit}s) ...")

        def worker():
            try:
                result = optimize(
                    data,
                    time_limit=time_limit,
                    verbose=False,
                    progress_callback=lambda msg: self.after(0, self._log_append, msg),
                )
                self._app.ctx.result = result
                self.after(0, self._on_done, result)
            except Exception as exc:
                self.after(0, self._on_error, str(exc))

        threading.Thread(target=worker, daemon=True).start()

    def _on_done(self, result: AllocationResult):
        self._run_btn.configure(state="normal")
        self._spinner_lbl.configure(text="Done!", text_color=_COLOR_SUCCESS)
        self._log_append(
            f"\nResult: {result.status} | "
            f"Value: ${result.total_value:,.0f} | "
            f"Unallocated: {result.unallocated_count}/{len(result.allocations)}"
        )
        self._next_btn.configure(state="normal")

    def _on_error(self, msg: str):
        self._run_btn.configure(state="normal")
        self._spinner_lbl.configure(text="Error", text_color=_COLOR_ERROR)
        self._log_append(f"ERROR: {msg}")
        mb.showerror("Optimiser error", msg)


# ---------------------------------------------------------------------------
# Screen 3 — Results & What-If
# ---------------------------------------------------------------------------

class Screen3(ctk.CTkFrame):
    """Results table, metrics, What-If analysis, and export."""

    def __init__(self, parent: "App", **kwargs):
        super().__init__(parent, fg_color="transparent", **kwargs)
        self._app = parent
        self._sensitivity: Optional[SensitivityModel] = None
        self._build()

    def _build(self):
        self.grid_columnconfigure(0, weight=1)

        # ---- Summary metrics ----
        _SectionLabel(self, text="Summary").grid(
            row=0, column=0, sticky="w", padx=20, pady=(20, 4))
        _Divider(self).grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 8))

        self._metrics_frame = ctk.CTkFrame(self, fg_color="transparent")
        self._metrics_frame.grid(row=2, column=0, sticky="ew", padx=20, pady=(0, 8))

        # ---- Results table ----
        _SectionLabel(self, text="Allocation Results").grid(
            row=3, column=0, sticky="w", padx=20, pady=(12, 4))
        _Divider(self).grid(row=4, column=0, sticky="ew", padx=20, pady=(0, 8))

        self._table_frame = ctk.CTkFrame(self, fg_color="transparent")
        self._table_frame.grid(row=5, column=0, sticky="nsew", padx=20)
        self._table_frame.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(5, weight=1)

        # ---- What-If ----
        _SectionLabel(self, text="What-If Analysis").grid(
            row=6, column=0, sticky="w", padx=20, pady=(16, 4))
        _Divider(self).grid(row=7, column=0, sticky="ew", padx=20, pady=(0, 8))

        self._whatif_frame = ctk.CTkFrame(self, fg_color="transparent")
        self._whatif_frame.grid(row=8, column=0, sticky="ew", padx=20)
        self._whatif_frame.grid_columnconfigure(3, weight=1)
        self._build_whatif()

        # ---- Export + navigation ----
        action_bar = ctk.CTkFrame(self, fg_color="transparent")
        action_bar.grid(row=9, column=0, sticky="ew", padx=20, pady=20)
        action_bar.grid_columnconfigure(2, weight=1)

        ctk.CTkButton(action_bar, text="← Back", font=_FONT_BODY,
                      fg_color="transparent", border_width=1,
                      command=lambda: self._app.go_to(1)).grid(row=0, column=0, padx=(0, 8))
        ctk.CTkButton(action_bar, text="Export CSV", font=_FONT_BODY,
                      command=self._export_csv).grid(row=0, column=1, padx=(0, 8))
        ctk.CTkButton(action_bar, text="Export Excel", font=_FONT_BODY,
                      command=self._export_excel).grid(row=0, column=2, sticky="w")
        ctk.CTkButton(
            action_bar, text="Run Again with New Data", font=_FONT_BODY,
            fg_color="transparent", border_width=1,
            command=self._run_again,
        ).grid(row=0, column=3, sticky="e")

    def _build_whatif(self):
        f = self._whatif_frame

        ctk.CTkLabel(f, text="Purchaser:", font=_FONT_BODY).grid(
            row=0, column=0, sticky="w", padx=(0, 8), pady=6)
        self._wi_purchaser_var = ctk.StringVar()
        self._wi_dropdown = ctk.CTkOptionMenu(
            f, variable=self._wi_purchaser_var,
            values=["(load results first)"], width=220, font=_FONT_SMALL,
        )
        self._wi_dropdown.grid(row=0, column=1, padx=(0, 20), pady=6)

        ctk.CTkLabel(f, text="Budget change:", font=_FONT_BODY).grid(
            row=0, column=2, sticky="w", padx=(0, 8))
        self._wi_pct_var = ctk.DoubleVar(value=0.10)
        self._wi_pct_display = ctk.CTkLabel(f, text="+10%", font=_FONT_BODY, width=50)
        self._wi_pct_display.grid(row=0, column=4, padx=(8, 0))
        ctk.CTkSlider(
            f, from_=-0.30, to=0.30, number_of_steps=12,
            variable=self._wi_pct_var, width=200,
            command=self._on_wi_slider,
        ).grid(row=0, column=3)

        self._wi_analyse_btn = ctk.CTkButton(
            f, text="Analyse", font=_FONT_BODY, width=100,
            command=self._run_whatif, state="disabled",
        )
        self._wi_analyse_btn.grid(row=1, column=0, columnspan=2, sticky="w", pady=(8, 0))

        self._wi_exact_btn = ctk.CTkButton(
            f, text="Run Exact", font=_FONT_BODY, width=100,
            fg_color="transparent", border_width=1,
            command=self._run_whatif_exact, state="disabled",
        )
        self._wi_exact_btn.grid(row=1, column=2, sticky="w", pady=(8, 0))

        self._wi_result_var = ctk.StringVar(value="")
        ctk.CTkLabel(
            f, textvariable=self._wi_result_var,
            font=_FONT_SMALL, wraplength=700, justify="left",
        ).grid(row=2, column=0, columnspan=5, sticky="w", pady=(10, 0))

    def _on_wi_slider(self, val):
        pct = float(val)
        sign = "+" if pct >= 0 else ""
        self._wi_pct_display.configure(text=f"{sign}{pct:.0%}")

    # ---- Populate on entry ----

    def populate(self):
        """Called by App.go_to() each time Screen3 becomes visible."""
        result = self._app.ctx.result
        data   = self._app.ctx.data
        if result is None or data is None:
            return
        self._render_metrics(data, result)
        self._render_table(data, result)
        self._update_whatif_controls(data)

    def _render_metrics(self, data: AllocationInput, result: AllocationResult):
        for w in self._metrics_frame.winfo_children():
            w.destroy()

        total_deals   = len(result.allocations)
        allocated     = total_deals - result.unallocated_count
        pct_allocated = allocated / total_deals * 100 if total_deals else 0
        total_budget  = sum(data.purchasers)
        pct_budget    = result.total_value / total_budget * 100 if total_budget else 0

        metrics = [
            ("Status",           result.status,                  _COLOR_SUCCESS if result.status == "Optimal" else _COLOR_WARNING),
            ("Total Value",      f"${result.total_value:,.0f}",  ("gray10", "gray90")),
            ("Deals Allocated",  f"{allocated}/{total_deals}  ({pct_allocated:.0f}%)", ("gray10", "gray90")),
            ("Budget Used",      f"{pct_budget:.1f}%",           ("gray10", "gray90")),
        ]

        for col, (label, value, color) in enumerate(metrics):
            card = ctk.CTkFrame(self._metrics_frame, corner_radius=8)
            card.grid(row=0, column=col, padx=(0, 12), pady=4, sticky="nsew")
            self._metrics_frame.grid_columnconfigure(col, weight=1)
            ctk.CTkLabel(card, text=label, font=_FONT_SMALL,
                         text_color=("gray55", "gray55")).pack(padx=16, pady=(10, 2))
            ctk.CTkLabel(card, text=value, font=_FONT_HEADING,
                         text_color=color).pack(padx=16, pady=(0, 10))

    def _render_table(self, data: AllocationInput, result: AllocationResult):
        for w in self._table_frame.winfo_children():
            w.destroy()

        deal_type_map  = dict(data.deals_type)
        type_labels    = {0: "Prepay", 1: "PPA"}
        rows = []
        for i, (deal_id, deal_value) in enumerate(data.deals):
            p_idx  = result.allocations[i]
            p_name = data.purchaser_ids[p_idx - 1] if p_idx > 0 else "Unallocated"
            rows.append({
                "Deal ID":        deal_id,
                "Value ($)":      f"{deal_value:,}",
                "Type":           type_labels.get(deal_type_map.get(deal_id, 0), "?"),
                "Assigned To":    p_name,
            })

        df = pd.DataFrame(rows)
        _ScrollableTable(
            self._table_frame, df=df,
            highlight_col="Assigned To", height=220,
        ).grid(row=0, column=0, sticky="ew")
        self._table_frame.grid_columnconfigure(0, weight=1)

    def _update_whatif_controls(self, data: AllocationInput):
        names = data.purchaser_ids
        self._wi_dropdown.configure(values=names)
        self._wi_purchaser_var.set(names[0] if names else "")
        self._wi_analyse_btn.configure(state="normal")
        self._wi_exact_btn.configure(state="normal")
        self._sensitivity = None  # Reset on new results

    # ---- What-If handlers ----

    def _run_whatif(self):
        data   = self._app.ctx.data
        result = self._app.ctx.result
        if data is None or result is None:
            return

        p_name = self._wi_purchaser_var.get()
        try:
            p_idx = data.purchaser_ids.index(p_name)
        except ValueError:
            return

        pct = float(self._wi_pct_var.get())
        self._wi_result_var.set("Computing sensitivity model... (this may take ~30s)")
        self._wi_analyse_btn.configure(state="disabled")
        self._wi_exact_btn.configure(state="disabled")

        def worker():
            try:
                sens = compute_sensitivity(
                    data, p_idx, result.total_value,
                    solver_time_limit=10,
                    progress_callback=lambda m: self.after(
                        0, self._wi_result_var.set, f"Analysing... {m}"),
                )
                wi = what_if(sens, pct)
                self._sensitivity = sens
                self.after(0, self._wi_show, wi.explanation)
            except Exception as exc:
                self.after(0, self._wi_show, f"Error: {exc}")
            finally:
                self.after(0, self._wi_analyse_btn.configure, {"state": "normal"})
                self.after(0, self._wi_exact_btn.configure, {"state": "normal"})

        threading.Thread(target=worker, daemon=True).start()

    def _run_whatif_exact(self):
        from app.sensitivity import what_if_exact
        data   = self._app.ctx.data
        result = self._app.ctx.result
        if data is None or result is None or self._sensitivity is None:
            self._wi_result_var.set(
                "Please run 'Analyse' first to build the sensitivity model.")
            return

        pct = float(self._wi_pct_var.get())
        self._wi_result_var.set("Running exact solve...")
        self._wi_analyse_btn.configure(state="disabled")
        self._wi_exact_btn.configure(state="disabled")

        def worker():
            try:
                wi = what_if_exact(
                    data, self._sensitivity, pct,
                    solver_time_limit=30,
                    progress_callback=lambda m: self.after(
                        0, self._wi_result_var.set, f"Solving... {m}"),
                )
                self.after(0, self._wi_show, wi.explanation)
            except Exception as exc:
                self.after(0, self._wi_show, f"Error: {exc}")
            finally:
                self.after(0, self._wi_analyse_btn.configure, {"state": "normal"})
                self.after(0, self._wi_exact_btn.configure, {"state": "normal"})

        threading.Thread(target=worker, daemon=True).start()

    def _wi_show(self, text: str):
        self._wi_result_var.set(text)

    # ---- Export ----

    def _export_csv(self):
        data, result = self._app.ctx.data, self._app.ctx.result
        if data is None or result is None:
            return
        path = fd.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialfile="allocation_results.csv",
        )
        if path:
            try:
                export_csv(data, result, path)
                mb.showinfo("Export", f"CSV saved to:\n{path}")
            except Exception as exc:
                mb.showerror("Export error", str(exc))

    def _export_excel(self):
        data, result = self._app.ctx.data, self._app.ctx.result
        if data is None or result is None:
            return
        path = fd.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="allocation_results.xlsx",
        )
        if path:
            try:
                export_excel(data, result, path)
                mb.showinfo("Export", f"Excel saved to:\n{path}")
            except Exception as exc:
                mb.showerror("Export error", str(exc))

    def _run_again(self):
        self._app.ctx.data   = None
        self._app.ctx.result = None
        self._sensitivity      = None
        self._app.go_to(0)


# ---------------------------------------------------------------------------
# Application state
# ---------------------------------------------------------------------------

class AppState:
    data:   Optional[AllocationInput]  = None
    result: Optional[AllocationResult] = None


# ---------------------------------------------------------------------------
# Main App window
# ---------------------------------------------------------------------------

class App(ctk.CTk):
    """Root window that hosts the step bar and all three screens."""

    def __init__(self):
        super().__init__()
        self.title("AllocationModel")
        self.geometry("960x780")
        self.minsize(860, 660)

        self.ctx = AppState()

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # Step bar
        self._step_bar = _StepBar(self)
        self._step_bar.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 0))

        # Screens
        self._screens: list[ctk.CTkFrame] = [
            Screen1(self),
            Screen2(self),
            Screen3(self),
        ]
        for screen in self._screens:
            screen.grid(row=1, column=0, sticky="nsew")

        self.go_to(0)

    def go_to(self, step: int):
        """Show the given step screen and hide all others."""
        for i, screen in enumerate(self._screens):
            if i == step:
                screen.tkraise()
            else:
                screen.lower()
        self._step_bar.set_active(step)

        # Trigger populate() on Screen3 every time it becomes visible
        if step == 2:
            self._screens[2].populate()


# ---------------------------------------------------------------------------
# Entry point (also called from main.py)
# ---------------------------------------------------------------------------

def run():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    run()

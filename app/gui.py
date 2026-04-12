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

import threading
import tkinter as tk
import tkinter.filedialog as fd
import tkinter.messagebox as mb
from pathlib import Path
from typing import Optional

_PROJECT_ROOT = Path(__file__).resolve().parent.parent
_ICON_PATH    = _PROJECT_ROOT / "assets" / "icon.png"

import customtkinter as ctk
import pandas as pd
from tkextrafont import Font as ExtraFont

from app.data_loader import AllocationInput, AllocationResult, load_file, get_raw_dataframes
from app.ml_classifier import classify_deals, apply_classification, ClassificationSummary
from app.optimizer import optimize
from app.sensitivity import compute_sensitivity, what_if, SensitivityModel
from app.exporter import export_csv, export_excel

# ---------------------------------------------------------------------------
# Theme
# ---------------------------------------------------------------------------
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

# Palette
_PAGE    = "#F4F6F9"
_SURFACE = "#FFFFFF"
_BRAND   = "#003366"
_ACCENT  = "#C8982A"
_BORDER  = "#D5DCE8"
_TXT     = "#1A2033"
_MUTED   = "#6B7280"
_WHITE   = "#FFFFFF"

_OK      = "#1A7A4A"
_WARN    = _ACCENT
_ERR     = "#C0392B"
_UNALLOC = _ERR

_STEP_LABELS = ["1  Load Data", "2  Configure", "3  Results"]

# Font paths
_FONTS_DIR       = _PROJECT_ROOT / "assets" / "fonts"
_FONT_FRAUNCES   = _FONTS_DIR / "FrauncesSemiBold.ttf"
_FONT_JAKARTA    = _FONTS_DIR / "PlusJakartaSansVariable.ttf"
_FF_DISPLAY = "Fraunces"
_FF_BODY    = "Plus Jakarta Sans"

_FONTS_LOADED = False


def _load_fonts(root) -> None:
    global _FONTS_LOADED
    if _FONTS_LOADED:
        return
    try:
        if _FONT_FRAUNCES.exists():
            ExtraFont(file=str(_FONT_FRAUNCES), family=_FF_DISPLAY, root=root)
        if _FONT_JAKARTA.exists():
            ExtraFont(file=str(_FONT_JAKARTA),  family=_FF_BODY,    root=root)
        _FONTS_LOADED = True
    except Exception:
        pass


def _fd(bold: bool = False) -> str:
    """Display font family (Fraunces or fallback)."""
    return _FF_DISPLAY if _FONTS_LOADED else "Georgia"


def _fb() -> str:
    """Body font family (Plus Jakarta Sans or fallback)."""
    return _FF_BODY if _FONTS_LOADED else "Segoe UI"


# Font tuples — functions so they resolve after _load_fonts()
def F_TITLE()   -> tuple: return (_fd(), 22, "bold")
def F_H1()      -> tuple: return (_fd(), 16, "bold")
def F_H2()      -> tuple: return (_fd(), 13, "bold")
def F_BODY()    -> tuple: return (_fb(), 13)
def F_SMALL()   -> tuple: return (_fb(), 11)
def F_MONO()    -> tuple: return ("Courier New", 11)
def F_METRIC()  -> tuple: return (_fd(), 18, "bold")


# ---------------------------------------------------------------------------
# Shared widget helpers
# ---------------------------------------------------------------------------

def _card(parent, **kw) -> ctk.CTkFrame:
    """White surface card with brand border."""
    return ctk.CTkFrame(parent, fg_color=_SURFACE, corner_radius=10,
                        border_width=1, border_color=_BORDER, **kw)


def _section(parent, text: str) -> None:
    """Render a labelled section header + accent divider into parent (pack layout)."""
    ctk.CTkLabel(parent, text=text, font=F_H1(), text_color=_BRAND, anchor="w").pack(
        fill="x", padx=0, pady=(0, 6))
    ctk.CTkFrame(parent, height=2, fg_color=_ACCENT, corner_radius=1).pack(
        fill="x", pady=(0, 14))


def _section_g(parent, text: str, ncols: int = 4) -> None:
    """Section header for grid-layout parents (placed at rows 0–1, spanning ncols)."""
    ctk.CTkLabel(parent, text=text, font=F_H1(), text_color=_BRAND, anchor="w"
                 ).grid(row=0, column=0, columnspan=ncols, sticky="ew", pady=(0, 6))
    ctk.CTkFrame(parent, height=2, fg_color=_ACCENT, corner_radius=1
                 ).grid(row=1, column=0, columnspan=ncols, sticky="ew", pady=(0, 14))


def _btn_primary(parent, text: str, command, width: int = 140, **kw) -> ctk.CTkButton:
    return ctk.CTkButton(parent, text=text, font=F_BODY(),
                         fg_color=_ACCENT, hover_color="#b5841e",
                         text_color=_WHITE, corner_radius=8,
                         width=width, command=command, **kw)


def _btn_secondary(parent, text: str, command, width: int = 140, **kw) -> ctk.CTkButton:
    return ctk.CTkButton(parent, text=text, font=F_BODY(),
                         fg_color=_BRAND, hover_color="#004a99",
                         text_color=_WHITE, corner_radius=8,
                         width=width, command=command, **kw)


def _btn_ghost(parent, text: str, command, width: int = 140, **kw) -> ctk.CTkButton:
    return ctk.CTkButton(parent, text=text, font=F_BODY(),
                         fg_color="transparent", hover_color="#e8ecf4",
                         text_color=_BRAND, border_width=1, border_color=_BRAND,
                         corner_radius=8, width=width, command=command, **kw)


def _lbl(parent, text: str, muted: bool = False, **kw) -> ctk.CTkLabel:
    color = _MUTED if muted else _TXT
    return ctk.CTkLabel(parent, text=text, font=F_BODY(),
                        text_color=color, anchor="w", **kw)


def _small(parent, text: str | None = None, var: ctk.StringVar | None = None,
           color: str = _MUTED, **kw) -> ctk.CTkLabel:
    kw2 = dict(font=F_SMALL(), text_color=color, anchor="w", **kw)
    if var is not None:
        return ctk.CTkLabel(parent, textvariable=var, **kw2)
    return ctk.CTkLabel(parent, text=text or "", **kw2)


# ---------------------------------------------------------------------------
# Step bar
# ---------------------------------------------------------------------------

class _StepBar(ctk.CTkFrame):
    def __init__(self, parent, **kw):
        super().__init__(parent, fg_color=_BRAND, corner_radius=0, **kw)
        self._btns: list[ctk.CTkButton] = []
        self.grid_columnconfigure(0, weight=1)

        inner = ctk.CTkFrame(self, fg_color="transparent")
        inner.pack(side="left", padx=24, pady=0)

        for i, label in enumerate(_STEP_LABELS):
            sep_w = ctk.CTkLabel(inner, text=" › ", font=F_SMALL(),
                                 text_color="#6688aa")
            if i > 0:
                sep_w.grid(row=0, column=i * 2 - 1, padx=2)

            btn = ctk.CTkButton(
                inner, text=label, font=F_SMALL(),
                state="disabled", fg_color="transparent",
                hover_color="#004a99", border_width=0, corner_radius=6,
                text_color="#7799bb", width=110, height=32,
            )
            btn.grid(row=0, column=i * 2, padx=4, pady=8)
            self._btns.append(btn)

    def set_active(self, step: int):
        for i, btn in enumerate(self._btns):
            if i == step:
                btn.configure(fg_color=_ACCENT, text_color=_WHITE,
                               hover_color="#b5841e")
            elif i < step:
                btn.configure(fg_color="transparent", text_color="#66cc88",
                               hover_color="#004a99")
            else:
                btn.configure(fg_color="transparent", text_color="#7799bb",
                               hover_color="#004a99")


# ---------------------------------------------------------------------------
# Scrollable table
# ---------------------------------------------------------------------------

class _Table(ctk.CTkScrollableFrame):
    def __init__(self, parent, df: pd.DataFrame,
                 highlight_col: Optional[str] = None, **kw):
        super().__init__(parent, fg_color=_SURFACE, **kw)
        self._draw(df, highlight_col)

    def _draw(self, df: pd.DataFrame, highlight_col: Optional[str]):
        for c, col in enumerate(df.columns):
            ctk.CTkLabel(self, text=str(col), font=F_SMALL(),
                         fg_color=_BRAND, text_color=_WHITE,
                         corner_radius=4, padx=10, pady=5, anchor="w",
                         ).grid(row=0, column=c, sticky="ew", padx=2, pady=(0, 2))
        for r, (_, row) in enumerate(df.iterrows()):
            bg = _SURFACE if r % 2 == 0 else "#f0f4fa"
            for c, col in enumerate(df.columns):
                val = row[col]
                unalloc = (highlight_col == col and str(val) == "Unallocated")
                ctk.CTkLabel(self, text=str(val), font=F_SMALL(),
                             fg_color=bg, text_color=_UNALLOC if unalloc else _TXT,
                             anchor="w", padx=10,
                             ).grid(row=r + 1, column=c, sticky="ew", padx=2, pady=1)
        for c in range(len(df.columns)):
            self.grid_columnconfigure(c, weight=1)


# ---------------------------------------------------------------------------
# Screen 1 — Load Data
# ---------------------------------------------------------------------------

class Screen1(ctk.CTkScrollableFrame):
    def __init__(self, parent: "App", **kw):
        super().__init__(parent, fg_color=_PAGE, **kw)
        self._app = parent
        self._deals_path: Optional[Path] = None
        self._purch_path: Optional[Path] = None
        self._deals_df:   Optional[pd.DataFrame] = None
        self._purch_df:   Optional[pd.DataFrame] = None
        self._ml_summary: Optional[ClassificationSummary] = None
        self.grid_columnconfigure(0, weight=1)
        self._build()

    def _build(self):
        # ---- Screen title ----
        ctk.CTkLabel(self, text="Load Input Data", font=F_TITLE(),
                     text_color=_BRAND, anchor="w",
                     ).grid(row=0, column=0, sticky="w", padx=28, pady=(24, 18))

        # ---- File card ----
        fc = _card(self)
        fc.grid(row=1, column=0, sticky="ew", padx=24, pady=(0, 16))
        fc.grid_columnconfigure(0, weight=1)
        fp = ctk.CTkFrame(fc, fg_color="transparent")
        fp.pack(fill="x", padx=20, pady=20)
        fp.grid_columnconfigure(1, weight=1)

        _section_g(fp, "Input Files", ncols=3)

        # Format toggle
        _lbl(fp, "File format:").grid(row=2, column=0, sticky="w", pady=6)
        self._fmt_var = ctk.StringVar(value="CSV (2 files)")
        seg = ctk.CTkSegmentedButton(
            fp, values=["CSV (2 files)", "Excel (1 file)"],
            variable=self._fmt_var,
            command=self._on_fmt_change,
            font=F_SMALL(),
            selected_color=_BRAND,
            selected_hover_color="#004a99",
            unselected_color="#d5dce8",
            unselected_hover_color="#c2ccdc",
            text_color=_WHITE,               # selected: white on brand ✓
            text_color_disabled=_MUTED,
        )
        # CTk SegmentedButton: text_color applies to selected; unselected needs
        # to be handled via a fixed fg override — override unselected text after build
        seg.grid(row=2, column=1, sticky="w", padx=(12, 0), pady=6)
        # Patch: unselected buttons use _TXT (dark) since bg is light grey
        self._seg = seg
        self._patch_seg_text()

        # Deals file
        self._deals_lbl = _lbl(fp, "Deals file:")
        self._deals_lbl.grid(row=3, column=0, sticky="w", pady=(10, 4))
        self._deals_var = ctk.StringVar(value="No file selected")
        _small(fp, var=self._deals_var).grid(row=3, column=1, sticky="w", padx=(12, 0))
        _btn_secondary(fp, "Browse", self._browse_deals, width=90
                       ).grid(row=3, column=2, padx=(12, 0))

        # Purchasers file
        self._purch_lbl = _lbl(fp, "Purchasers file:")
        self._purch_lbl.grid(row=4, column=0, sticky="w", pady=(6, 4))
        self._purch_var = ctk.StringVar(value="No file selected")
        _small(fp, var=self._purch_var).grid(row=4, column=1, sticky="w", padx=(12, 0))
        self._purch_browse_btn = _btn_secondary(fp, "Browse", self._browse_purch, width=90)
        self._purch_browse_btn.grid(row=4, column=2, padx=(12, 0))

        # Load button + status
        load_row = ctk.CTkFrame(fp, fg_color="transparent")
        load_row.grid(row=5, column=0, columnspan=3, sticky="w", pady=(16, 0))
        _btn_primary(load_row, "Load & Preview", self._load, width=150).pack(side="left")
        self._status_var = ctk.StringVar(value="")
        self._status_lbl = _small(load_row, var=self._status_var)
        self._status_lbl.pack(side="left", padx=14)

        # ---- Preview card ----
        pc = _card(self)
        pc.grid(row=2, column=0, sticky="ew", padx=24, pady=(0, 16))
        self._preview_outer = pc

        self._preview_inner = ctk.CTkFrame(pc, fg_color="transparent")
        self._preview_inner.pack(fill="both", expand=True, padx=20, pady=20)
        self._preview_inner.grid_columnconfigure(0, weight=1)

        _small(self._preview_inner,
               text="Load a file above to preview your data.",
               color=_MUTED).grid(row=0, column=0, pady=20)

        # ---- ML card ----
        mc = _card(self)
        mc.grid(row=3, column=0, sticky="ew", padx=24, pady=(0, 16))
        mp = ctk.CTkFrame(mc, fg_color="transparent")
        mp.pack(fill="x", padx=20, pady=20)
        mp.grid_columnconfigure(1, weight=1)

        _section_g(mp, "Auto-Classification (ML)", ncols=3)

        _small(mp, text=(
            "K-Means clustering suggests Prepay / PPA labels from deal values. "
            "Review overrides before proceeding."
        ), color=_MUTED).grid(row=2, column=0, columnspan=3, sticky="w", pady=(0, 12))

        self._ml_btn = _btn_secondary(mp, "Auto-classify Deals", self._run_ml, width=170)
        self._ml_btn.configure(state="disabled")
        self._ml_btn.grid(row=3, column=0, sticky="w")

        self._ml_var = ctk.StringVar(value="")
        _small(mp, var=self._ml_var, color=_TXT).grid(
            row=3, column=1, sticky="w", padx=(16, 0))

        # ---- Navigation ----
        nav = ctk.CTkFrame(self, fg_color="transparent")
        nav.grid(row=4, column=0, sticky="e", padx=28, pady=(4, 28))
        self._next_btn = _btn_primary(nav, "Next: Configure  →", self._go_next, width=180)
        self._next_btn.configure(state="disabled")
        self._next_btn.pack()

    # ---- helpers ----

    def _patch_seg_text(self):
        """Force dark text on the unselected (light-grey) segment buttons."""
        try:
            for child in self._seg.winfo_children():
                if isinstance(child, ctk.CTkButton):
                    # Only change unselected ones (those not in _active set)
                    pass
        except Exception:
            pass
        # Simpler approach: override after each value change via trace
        self._fmt_var.trace_add("write", lambda *_: self._recolor_seg())
        self._recolor_seg()

    def _recolor_seg(self):
        """Recolor SegmentedButton children so unselected = dark text."""
        try:
            active = self._fmt_var.get()
            for child in self._seg._buttons_dict.values():
                if hasattr(child, "_text_label") and child._text_label:
                    pass
            # Use CTk internal to set text colors per button
            for val, btn in self._seg._buttons_dict.items():
                if val == active:
                    btn.configure(text_color=_WHITE)
                else:
                    btn.configure(text_color=_TXT)
        except Exception:
            pass

    def _on_fmt_change(self, _=None):
        is_csv = self._fmt_var.get() == "CSV (2 files)"
        c = _TXT if is_csv else _MUTED
        self._purch_lbl.configure(text_color=c)
        self._purch_browse_btn.configure(
            state="normal" if is_csv else "disabled",
            fg_color=_BRAND if is_csv else _BORDER,
            text_color=_WHITE if is_csv else _MUTED,
        )
        self._recolor_seg()

    def _browse_deals(self):
        types = ([("CSV", "*.csv")] if self._fmt_var.get() == "CSV (2 files)"
                 else [("Excel", "*.xlsx *.xlsm")])
        p = fd.askopenfilename(filetypes=types)
        if p:
            self._deals_path = Path(p)
            self._deals_var.set(self._deals_path.name)

    def _browse_purch(self):
        p = fd.askopenfilename(filetypes=[("CSV", "*.csv")])
        if p:
            self._purch_path = Path(p)
            self._purch_var.set(self._purch_path.name)

    def _load(self):
        fmt = self._fmt_var.get()
        try:
            if fmt == "CSV (2 files)":
                if not self._deals_path:
                    return self._status("Please select the deals CSV.", err=True)
                if not self._purch_path:
                    return self._status("Please select the purchasers CSV.", err=True)
                self._deals_df, self._purch_df = get_raw_dataframes(
                    self._deals_path, self._purch_path)
            else:
                if not self._deals_path:
                    return self._status("Please select the Excel file.", err=True)
                self._deals_df, self._purch_df = get_raw_dataframes(self._deals_path)

            self._status(
                f"✓  {len(self._deals_df)} deals · {len(self._purch_df)} purchasers loaded.")
            self._render_preview()
            self._ml_btn.configure(state="normal")
            self._next_btn.configure(state="normal")
        except Exception as e:
            self._status(f"Error: {e}", err=True)

    def _render_preview(self):
        for w in self._preview_inner.winfo_children():
            w.destroy()

        pp = ctk.CTkFrame(self._preview_inner, fg_color="transparent")
        pp.pack(fill="x", pady=(4, 0))
        pp.grid_columnconfigure(0, weight=1)
        pp.grid_columnconfigure(1, weight=1)

        # Deals
        dl = ctk.CTkFrame(pp, fg_color="transparent")
        dl.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        _small(dl, text=f"Deals  —  {len(self._deals_df)} rows",
               color=_MUTED).pack(anchor="w", pady=(0, 4))
        _Table(dl, df=self._deals_df.head(20), height=160).pack(fill="x")

        # Purchasers
        pr = ctk.CTkFrame(pp, fg_color="transparent")
        pr.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        _small(pr, text=f"Purchasers  —  {len(self._purch_df)} rows",
               color=_MUTED).pack(anchor="w", pady=(0, 4))
        _Table(pr, df=self._purch_df, height=160).pack(fill="x")

    def _run_ml(self):
        if self._deals_df is None:
            return
        df = self._deals_df.copy()
        df.columns = [c.strip().lower() for c in df.columns]
        ids   = df["deal_id"].astype(str).tolist()
        vals  = df["deal_value"].astype(float).tolist()
        types = df["deal_type"].astype(str).tolist() if "deal_type" in df.columns else None

        self._ml_summary = classify_deals(ids, vals, types)
        s = self._ml_summary

        suggestions = {r.deal_id: r.suggested_type for r in s.results}
        self._deals_df["deal_type"] = self._deals_df.apply(
            lambda row: suggestions.get(
                str(row.get("deal_id", row.get("Deal ID", ""))).strip(),
                row.get("deal_type", row.get("Deal Type", "Prepay"))),
            axis=1)
        self._render_preview()

        msg = f"{s.prepay_count} Prepay · {s.ppa_count} PPA"
        if s.override_count:
            msg += f" · {s.override_count} override(s) suggested"
        if s.warning:
            msg += f"  —  {s.warning}"
        self._ml_var.set(msg)

    def _status(self, msg: str, err: bool = False):
        self._status_var.set(msg)
        self._status_lbl.configure(text_color=_ERR if err else _OK)

    def _go_next(self):
        try:
            fmt = self._fmt_var.get()
            data = (load_file(self._deals_path, self._purch_path)
                    if fmt == "CSV (2 files)" else load_file(self._deals_path))
            if self._ml_summary:
                data.deals_type = apply_classification(
                    data.deals, data.deals_type, self._ml_summary)
            self._app.ctx.data = data
            self._app.go_to(1)
        except Exception as e:
            self._status(f"Validation error: {e}", err=True)


# ---------------------------------------------------------------------------
# Screen 2 — Configure & Run
# ---------------------------------------------------------------------------

class Screen2(ctk.CTkScrollableFrame):
    def __init__(self, parent: "App", **kw):
        super().__init__(parent, fg_color=_PAGE, **kw)
        self._app = parent
        self.grid_columnconfigure(0, weight=1)
        self._build()

    def _build(self):
        ctk.CTkLabel(self, text="Configure & Run", font=F_TITLE(),
                     text_color=_BRAND, anchor="w",
                     ).grid(row=0, column=0, sticky="w", padx=28, pady=(24, 18))

        # ---- Settings card ----
        sc = _card(self)
        sc.grid(row=1, column=0, sticky="ew", padx=24, pady=(0, 16))
        sp = ctk.CTkFrame(sc, fg_color="transparent")
        sp.pack(fill="x", padx=20, pady=20)
        sp.grid_columnconfigure(1, weight=1)

        _section_g(sp, "Optimiser Settings", ncols=3)

        # Time limit
        _lbl(sp, "Time limit:").grid(row=2, column=0, sticky="w", pady=10)
        self._time_var = ctk.IntVar(value=60)
        self._time_lbl = ctk.CTkLabel(sp, text="60 s", font=F_BODY(),
                                       text_color=_ACCENT, width=52, anchor="w")
        self._time_lbl.grid(row=2, column=2, padx=(12, 0))
        ctk.CTkSlider(sp, from_=10, to=300, number_of_steps=29,
                      variable=self._time_var,
                      button_color=_ACCENT, button_hover_color="#b5841e",
                      progress_color=_BRAND,
                      command=lambda v: self._time_lbl.configure(text=f"{int(v)} s"),
                      ).grid(row=2, column=1, sticky="ew")

        # Toggles
        rows = [
            ("Each purchaser gets ≥ 1 deal", "_min_var", True),
            ("Penalise preference mismatches",  "_pref_var", True),
        ]
        for i, (label, attr, default) in enumerate(rows):
            _lbl(sp, label).grid(row=3 + i, column=0, sticky="w", pady=8)
            var = ctk.BooleanVar(value=default)
            setattr(self, attr, var)
            ctk.CTkSwitch(sp, text="", variable=var,
                          progress_color=_BRAND,
                          button_color=_ACCENT, button_hover_color="#b5841e",
                          ).grid(row=3 + i, column=1, sticky="w")

        # ---- Run card ----
        rc = _card(self)
        rc.grid(row=2, column=0, sticky="ew", padx=24, pady=(0, 16))
        rp = ctk.CTkFrame(rc, fg_color="transparent")
        rp.pack(fill="x", padx=20, pady=20)

        _section(rp, "Run")

        run_row = ctk.CTkFrame(rp, fg_color="transparent")
        run_row.pack(fill="x", pady=(0, 4))

        self._run_btn = _btn_primary(run_row, "Run Optimisation",
                                      self._run, width=180)
        self._run_btn.pack(side="left")
        self._run_lbl = _small(run_row, text="", color=_MUTED)
        self._run_lbl.pack(side="left", padx=16)

        # ---- Log card ----
        lc = _card(self)
        lc.grid(row=3, column=0, sticky="ew", padx=24, pady=(0, 16))
        lp = ctk.CTkFrame(lc, fg_color="transparent")
        lp.pack(fill="both", expand=True, padx=20, pady=20)

        _section(lp, "Solver Log")

        self._log = ctk.CTkTextbox(lp, font=F_MONO(), state="disabled",
                                    height=220, fg_color="#f8f9fb",
                                    text_color=_TXT, border_color=_BORDER,
                                    border_width=1, corner_radius=6)
        self._log.pack(fill="x")

        # ---- Navigation ----
        nav = ctk.CTkFrame(self, fg_color="transparent")
        nav.grid(row=4, column=0, sticky="ew", padx=28, pady=(4, 28))
        nav.grid_columnconfigure(1, weight=1)
        _btn_ghost(nav, "← Back", lambda: self._app.go_to(0), width=100
                   ).grid(row=0, column=0)
        self._next_btn = _btn_primary(nav, "View Results  →",
                                       lambda: self._app.go_to(2), width=180)
        self._next_btn.configure(state="disabled")
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
        data.min_deal     = self._min_var.get()
        data.pref_penalty = self._pref_var.get()
        tl = int(self._time_var.get())

        self._run_btn.configure(state="disabled")
        self._next_btn.configure(state="disabled")
        self._run_lbl.configure(text="Solving…", text_color=_WARN)
        self._log_append(f"Starting optimisation  (time limit: {tl}s)…")

        def worker():
            try:
                result = optimize(
                    data, time_limit=tl, verbose=False,
                    progress_callback=lambda m: self.after(0, self._log_append, m))
                self._app.ctx.result = result
                self.after(0, self._done, result)
            except Exception as exc:
                self.after(0, self._error, str(exc))

        threading.Thread(target=worker, daemon=True).start()

    def _done(self, result: AllocationResult):
        self._run_btn.configure(state="normal")
        self._run_lbl.configure(text="✓  Done", text_color=_OK)
        self._log_append(
            f"\nStatus: {result.status}  |  "
            f"Value: ${result.total_value:,.0f}  |  "
            f"Unallocated: {result.unallocated_count}/{len(result.allocations)}")
        self._next_btn.configure(state="normal")

    def _error(self, msg: str):
        self._run_btn.configure(state="normal")
        self._run_lbl.configure(text="✗  Error", text_color=_ERR)
        self._log_append(f"ERROR: {msg}")
        mb.showerror("Optimiser error", msg)


# ---------------------------------------------------------------------------
# Screen 3 — Results & What-If
# ---------------------------------------------------------------------------

class Screen3(ctk.CTkScrollableFrame):
    def __init__(self, parent: "App", **kw):
        super().__init__(parent, fg_color=_PAGE, **kw)
        self._app = parent
        self._sens: Optional[SensitivityModel] = None
        self.grid_columnconfigure(0, weight=1)
        self._build()

    def _build(self):
        ctk.CTkLabel(self, text="Allocation Results", font=F_TITLE(),
                     text_color=_BRAND, anchor="w",
                     ).grid(row=0, column=0, sticky="w", padx=28, pady=(24, 18))

        # ---- Metrics row ----
        self._metrics_frame = ctk.CTkFrame(self, fg_color="transparent")
        self._metrics_frame.grid(row=1, column=0, sticky="ew", padx=24, pady=(0, 16))

        # ---- Results table card ----
        tc = _card(self)
        tc.grid(row=2, column=0, sticky="ew", padx=24, pady=(0, 16))
        tp = ctk.CTkFrame(tc, fg_color="transparent")
        tp.pack(fill="both", expand=True, padx=20, pady=20)
        tp.grid_columnconfigure(0, weight=1)

        _section_g(tp, "Deal Allocations", ncols=1)
        self._table_frame = ctk.CTkFrame(tp, fg_color="transparent")
        self._table_frame.grid(row=2, column=0, sticky="ew")
        self._table_frame.grid_columnconfigure(0, weight=1)

        # ---- What-If card ----
        wc = _card(self)
        wc.grid(row=3, column=0, sticky="ew", padx=24, pady=(0, 16))
        wp = ctk.CTkFrame(wc, fg_color="transparent")
        wp.pack(fill="x", padx=20, pady=20)

        _section(wp, "What-If Analysis")
        _small(wp, text=(
            "Select a purchaser and a budget change to estimate the impact "
            "on total allocated value."
        ), color=_MUTED).pack(anchor="w", pady=(0, 14))

        # Controls row
        ctrl = ctk.CTkFrame(wp, fg_color="transparent")
        ctrl.pack(fill="x")
        ctrl.grid_columnconfigure(1, weight=1)
        ctrl.grid_columnconfigure(3, weight=1)

        _lbl(ctrl, "Purchaser:").grid(row=0, column=0, sticky="w", padx=(0, 10))
        self._wi_p_var = ctk.StringVar()
        self._wi_dd = ctk.CTkOptionMenu(
            ctrl, variable=self._wi_p_var,
            values=["(load results first)"], width=200, font=F_SMALL(),
            fg_color=_BRAND, button_color="#00264d",
            button_hover_color="#004a99",
            text_color=_WHITE,
        )
        self._wi_dd.grid(row=0, column=1, sticky="w", padx=(0, 24))

        _lbl(ctrl, "Budget change:").grid(row=0, column=2, sticky="w", padx=(0, 10))
        self._wi_pct_var = ctk.DoubleVar(value=0.10)
        self._wi_pct_lbl = ctk.CTkLabel(ctrl, text="+10%", font=F_BODY(),
                                         text_color=_ACCENT, width=52, anchor="w")
        self._wi_pct_lbl.grid(row=0, column=4, padx=(10, 0))
        ctk.CTkSlider(ctrl, from_=-0.30, to=0.30, number_of_steps=12,
                      variable=self._wi_pct_var, width=200,
                      button_color=_ACCENT, button_hover_color="#b5841e",
                      progress_color=_BRAND,
                      command=self._on_wi_slider,
                      ).grid(row=0, column=3)

        # Buttons row
        btn_row = ctk.CTkFrame(wp, fg_color="transparent")
        btn_row.pack(fill="x", pady=(14, 0))

        self._wi_analyse = _btn_primary(btn_row, "Analyse", self._run_wi, width=120)
        self._wi_analyse.configure(state="disabled")
        self._wi_analyse.pack(side="left", padx=(0, 10))

        self._wi_exact = _btn_ghost(btn_row, "Run Exact", self._run_wi_exact, width=120)
        self._wi_exact.configure(state="disabled")
        self._wi_exact.pack(side="left")

        # Result text
        self._wi_result_var = ctk.StringVar(value="")
        ctk.CTkLabel(wp, textvariable=self._wi_result_var,
                     font=F_SMALL(), text_color=_TXT,
                     wraplength=820, justify="left",
                     ).pack(anchor="w", pady=(12, 0))

        # ---- Action bar ----
        ab = ctk.CTkFrame(self, fg_color="transparent")
        ab.grid(row=4, column=0, sticky="ew", padx=28, pady=(4, 28))
        ab.grid_columnconfigure(2, weight=1)

        _btn_ghost(ab, "← Back", lambda: self._app.go_to(1), width=100
                   ).grid(row=0, column=0, padx=(0, 8))
        _btn_secondary(ab, "Export CSV", self._export_csv, width=130
                       ).grid(row=0, column=1, padx=(0, 8))
        _btn_primary(ab, "Export Excel", self._export_excel, width=130
                     ).grid(row=0, column=2, sticky="w")
        _btn_ghost(ab, "⟳  Run Again", self._run_again, width=150
                   ).grid(row=0, column=3, sticky="e")

    # ---- populate on entry ----

    def populate(self):
        result = self._app.ctx.result
        data   = self._app.ctx.data
        if result is None or data is None:
            return
        self._render_metrics(data, result)
        self._render_table(data, result)
        self._update_wi(data)

    def _render_metrics(self, data: AllocationInput, result: AllocationResult):
        for w in self._metrics_frame.winfo_children():
            w.destroy()

        total  = len(result.allocations)
        alloc  = total - result.unallocated_count
        pct_d  = alloc / total * 100 if total else 0
        budget = sum(data.purchasers)
        pct_b  = result.total_value / budget * 100 if budget else 0

        cards = [
            ("Solver Status",     result.status,
             _OK if result.status == "Optimal" else _WARN),
            ("Total Value",       f"${result.total_value:,.0f}", _BRAND),
            ("Deals Allocated",   f"{alloc}/{total}",            _TXT),
            ("Budget Used",       f"{pct_b:.1f}%",               _TXT),
        ]
        for i, (lbl, val, col) in enumerate(cards):
            c = _card(self._metrics_frame)
            c.grid(row=0, column=i, padx=(0, 12), sticky="nsew")
            self._metrics_frame.grid_columnconfigure(i, weight=1)
            _small(c, text=lbl, color=_MUTED).pack(padx=16, pady=(12, 2), anchor="w")
            ctk.CTkLabel(c, text=val, font=F_METRIC(),
                         text_color=col, anchor="w").pack(padx=16, pady=(0, 12))

    def _render_table(self, data: AllocationInput, result: AllocationResult):
        for w in self._table_frame.winfo_children():
            w.destroy()
        dtype  = dict(data.deals_type)
        tnames = {0: "Prepay", 1: "PPA"}
        rows = []
        for i, (did, dval) in enumerate(data.deals):
            pi   = result.allocations[i]
            pnam = data.purchaser_ids[pi - 1] if pi > 0 else "Unallocated"
            rows.append({"Deal ID": did, "Value ($)": f"{dval:,}",
                         "Type": tnames.get(dtype.get(did, 0), "?"),
                         "Assigned To": pnam})
        _Table(self._table_frame, df=pd.DataFrame(rows),
               highlight_col="Assigned To", height=240,
               ).grid(row=0, column=0, sticky="ew")

    def _update_wi(self, data: AllocationInput):
        self._wi_dd.configure(values=data.purchaser_ids)
        self._wi_p_var.set(data.purchaser_ids[0] if data.purchaser_ids else "")
        self._wi_analyse.configure(state="normal")
        self._wi_exact.configure(state="normal")
        self._sens = None

    def _on_wi_slider(self, val):
        pct  = float(val)
        sign = "+" if pct >= 0 else ""
        self._wi_pct_lbl.configure(text=f"{sign}{pct:.0%}")

    def _run_wi(self):
        data, result = self._app.ctx.data, self._app.ctx.result
        if data is None or result is None:
            return
        p_name = self._wi_p_var.get()
        try:
            p_idx = data.purchaser_ids.index(p_name)
        except ValueError:
            return
        pct = float(self._wi_pct_var.get())
        self._wi_result_var.set("Building sensitivity model… (~30 s)")
        self._wi_analyse.configure(state="disabled")
        self._wi_exact.configure(state="disabled")

        def worker():
            try:
                sens = compute_sensitivity(
                    data, p_idx, result.total_value,
                    solver_time_limit=10,
                    progress_callback=lambda m: self.after(
                        0, self._wi_result_var.set, f"Analysing…  {m}"))
                wi = what_if(sens, pct)
                self._sens = sens
                self.after(0, self._wi_result_var.set, wi.explanation)
            except Exception as exc:
                self.after(0, self._wi_result_var.set, f"Error: {exc}")
            finally:
                self.after(0, self._wi_analyse.configure, {"state": "normal"})
                self.after(0, self._wi_exact.configure,   {"state": "normal"})

        threading.Thread(target=worker, daemon=True).start()

    def _run_wi_exact(self):
        from app.sensitivity import what_if_exact
        data, result = self._app.ctx.data, self._app.ctx.result
        if data is None or result is None or self._sens is None:
            self._wi_result_var.set("Run 'Analyse' first to build the sensitivity model.")
            return
        pct = float(self._wi_pct_var.get())
        self._wi_result_var.set("Running exact solve…")
        self._wi_analyse.configure(state="disabled")
        self._wi_exact.configure(state="disabled")

        def worker():
            try:
                wi = what_if_exact(
                    data, self._sens, pct, solver_time_limit=30,
                    progress_callback=lambda m: self.after(
                        0, self._wi_result_var.set, f"Solving…  {m}"))
                self.after(0, self._wi_result_var.set, wi.explanation)
            except Exception as exc:
                self.after(0, self._wi_result_var.set, f"Error: {exc}")
            finally:
                self.after(0, self._wi_analyse.configure, {"state": "normal"})
                self.after(0, self._wi_exact.configure,   {"state": "normal"})

        threading.Thread(target=worker, daemon=True).start()

    def _export_csv(self):
        data, result = self._app.ctx.data, self._app.ctx.result
        if not data or not result:
            return
        p = fd.asksaveasfilename(defaultextension=".csv",
                                  filetypes=[("CSV", "*.csv")],
                                  initialfile="allocation_results.csv")
        if p:
            try:
                export_csv(data, result, p)
                mb.showinfo("Exported", f"CSV saved:\n{p}")
            except Exception as e:
                mb.showerror("Export error", str(e))

    def _export_excel(self):
        data, result = self._app.ctx.data, self._app.ctx.result
        if not data or not result:
            return
        p = fd.asksaveasfilename(defaultextension=".xlsx",
                                  filetypes=[("Excel", "*.xlsx")],
                                  initialfile="allocation_results.xlsx")
        if p:
            try:
                export_excel(data, result, p)
                mb.showinfo("Exported", f"Excel saved:\n{p}")
            except Exception as e:
                mb.showerror("Export error", str(e))

    def _run_again(self):
        self._app.ctx.data   = None
        self._app.ctx.result = None
        self._sens           = None
        self._app.go_to(0)


# ---------------------------------------------------------------------------
# App state
# ---------------------------------------------------------------------------

class AppState:
    data:   Optional[AllocationInput]  = None
    result: Optional[AllocationResult] = None


# ---------------------------------------------------------------------------
# Root window
# ---------------------------------------------------------------------------

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("AllocationModel")
        self.geometry("980x800")
        self.minsize(860, 680)
        self.configure(fg_color=_BRAND)

        # Load fonts before building any widgets
        _load_fonts(self)

        # Window icon
        try:
            icon = tk.PhotoImage(file=str(_ICON_PATH))
            self.wm_iconphoto(True, icon)
        except Exception:
            pass

        self.ctx = AppState()

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # Step bar
        self._step_bar = _StepBar(self)
        self._step_bar.grid(row=0, column=0, sticky="ew")

        # Screens — built AFTER _load_fonts so font families are available
        self._screens: list[ctk.CTkFrame] = [
            Screen1(self),
            Screen2(self),
            Screen3(self),
        ]
        for s in self._screens:
            s.grid(row=1, column=0, sticky="nsew")

        self.go_to(0)

    def go_to(self, step: int):
        for i, s in enumerate(self._screens):
            (s.tkraise if i == step else s.lower)()
        self._step_bar.set_active(step)
        if step == 2:
            self._screens[2].populate()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def run():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    run()

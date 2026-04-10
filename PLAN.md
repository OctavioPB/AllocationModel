# AllocationModel — Production Ready Plan

## Overview

Transform the existing ILP-based deal allocation script into a production-ready,
cross-platform desktop application usable by non-technical users.

**Current state:** Single Python script, Windows-only, no GUI, hardcoded parameters, requires manual solver installation.  
**Target state:** Cross-platform desktop app (Mac + Windows) with a friendly GUI, CSV/Excel support, ML-assisted classification, What-If analysis, and one-command installation.

---

## Constraints & Decisions

| Decision | Choice | Reason |
|---|---|---|
| GUI framework | CustomTkinter | Modern look, pure Python, no extra system deps |
| Data I/O | pandas + openpyxl | Cross-platform, handles CSV and Excel |
| Solver | PuLP + CBC (bundled) | CBC ships with `pip install pulp`, no manual install |
| ML framework | scikit-learn | Lightweight, well-known, no GPU needed |
| Entry point | `main.py` | Double-click or `python main.py` to launch |
| Language | English | GUI, README, comments, docstrings |

---

## Project Structure (Target)

```
AllocationModel/
├── app/
│   ├── __init__.py
│   ├── gui.py              # CustomTkinter interface (3 screens)
│   ├── optimizer.py        # ILP model — refactored, cross-platform
│   ├── data_loader.py      # CSV + Excel reader (pandas/openpyxl)
│   ├── ml_classifier.py    # K-Means deal auto-classification
│   ├── sensitivity.py      # What-If regression analysis
│   └── exporter.py         # Export results to CSV / Excel
├── assets/
│   └── icon.png
├── sample_data/
│   └── example_deals.csv   # Derived from Allocation_Model_test.xlsm
├── tests/
│   ├── test_optimizer.py
│   ├── test_data_loader.py
│   └── test_ml_classifier.py
├── main.py                 # Entry point
├── requirements.txt
├── README.md
├── PLAN.md
└── .gitignore
```

---

## Sprints

---

### Sprint 1 — Foundation & Project Setup
**Goal:** Clean repo structure, cross-platform data layer, solver swap. No GUI yet.

| # | Task | File(s) | Notes |
|---|---|---|---|
| 1.1 | Create project folder structure | `app/`, `tests/`, `assets/`, `sample_data/` | Scaffold all `__init__.py` files |
| 1.2 | Replace `win32com` with `pandas` + `openpyxl` | `app/data_loader.py` | Support `.xlsx`, `.xlsm`, `.csv` inputs |
| 1.3 | Define standard internal data schema | `app/data_loader.py` | Normalize all inputs to a single `AllocationInput` dataclass |
| 1.4 | Convert `Allocation_Model_test.xlsm` to `example_deals.csv` | `sample_data/` | Strip any sensitive data; document expected column names |
| 1.5 | Swap GLPK solver for CBC (PuLP bundled) | `app/optimizer.py` | Verify CBC is available on Mac + Windows after `pip install pulp` |
| 1.6 | Refactor `Optimization.py` into `app/optimizer.py` | `app/optimizer.py` | Fix typos (`get_vaiables` → `get_variables`), add docstrings, remove hardcoded filenames |
| 1.7 | Create `requirements.txt` | `requirements.txt` | Pin versions; include `pulp`, `pandas`, `openpyxl`, `customtkinter`, `scikit-learn` |
| 1.8 | Create `.gitignore` | `.gitignore` | Ignore `.xlsm`, `__pycache__`, `.env`, etc. |

**Exit criteria:** `app/optimizer.py` runs headlessly on Mac and Windows using a CSV input, produces correct allocation output, no `win32com` dependency anywhere.

---

### Sprint 2 — ML Module
**Goal:** Two ML-powered features: deal auto-classification and What-If sensitivity analysis.

| # | Task | File(s) | Notes |
|---|---|---|---|
| 2.1 | Implement K-Means deal auto-classifier | `app/ml_classifier.py` | Cluster deals by value into 2 groups (Prepay / PPA). Return labels + confidence score per deal |
| 2.2 | Add classifier fallback logic | `app/ml_classifier.py` | If only 1 deal type in data, skip clustering and label all as that type |
| 2.3 | Implement What-If sensitivity analysis | `app/sensitivity.py` | Perturb purchaser capacities ±5/10/20%, re-run optimizer, fit a linear regression on (capacity delta → value delta). Expose as `what_if(purchaser_id, capacity_change_pct)` |
| 2.4 | Add result explanation helper | `app/sensitivity.py` | Plain-English output: "Increasing Purchaser A's budget by 10% would allocate ~$X more in deals" |
| 2.5 | Write unit tests for ML module | `tests/test_ml_classifier.py` | Test with synthetic data; assert cluster labels are stable for clear cases |

**Exit criteria:** `ml_classifier.py` correctly suggests deal types on `example_deals.csv`. `sensitivity.py` produces human-readable What-If results for at least 3 capacity perturbation levels.

---

### Sprint 3 — GUI
**Goal:** Full 3-screen CustomTkinter desktop application wired to optimizer + ML modules.

#### Screen 1 — Load Data
| # | Task | Notes |
|---|---|---|
| 3.1 | File picker button (`.csv`, `.xlsx`, `.xlsm`) | Uses `tkinter.filedialog`, no hardcoded paths |
| 3.2 | Data preview table (scrollable) | Show first 20 rows of deals + purchasers after load |
| 3.3 | "Auto-classify deals with ML" button | Calls `ml_classifier.py`; updates Deal Type column in preview |
| 3.4 | Manual override of deal type in table | User can correct ML suggestions before running |
| 3.5 | Input validation feedback | Inline error messages (missing columns, empty values, etc.) |

#### Screen 2 — Configure & Run
| # | Task | Notes |
|---|---|---|
| 3.6 | Time limit slider (10–300 seconds) | Default: 60s |
| 3.7 | Toggle: Minimum 1 deal per purchaser | Default: ON |
| 3.8 | Toggle: Penalize preference violations | Default: ON |
| 3.9 | "Run Optimization" button | Disables UI during solve; shows animated progress indicator |
| 3.10 | Solver status log panel | Live text output (status messages from optimizer) |

#### Screen 3 — Results & What-If
| # | Task | Notes |
|---|---|---|
| 3.11 | Results table: Deal → Assigned Purchaser | Highlight unallocated deals in a distinct color |
| 3.12 | Summary metrics panel | Total allocated value, % of deals allocated, % of capacity used per purchaser |
| 3.13 | What-If Analysis panel | Dropdown: select purchaser. Slider: capacity change %. Display plain-English output from `sensitivity.py` |
| 3.14 | "Export to CSV" button | Calls `exporter.py`; saves results + metrics |
| 3.15 | "Export to Excel" button | Calls `exporter.py`; saves styled `.xlsx` with separate sheets for results and summary |
| 3.16 | "Run Again" button | Returns to Screen 1 with data pre-loaded |

**Exit criteria:** Full end-to-end flow works without touching the terminal. All three screens navigate correctly. Export produces readable files.

---

### Sprint 4 — Packaging & Documentation
**Goal:** Any user with Python 3.9+ can install and run in under 5 minutes.

| # | Task | File(s) | Notes |
|---|---|---|---|
| 4.1 | Write `README.md` | `README.md` | See structure below |
| 4.2 | Write `main.py` entry point | `main.py` | Single file: imports `app.gui`, launches window |
| 4.3 | Add app icon | `assets/icon.png` | Used in CustomTkinter window title bar |
| 4.4 | Document CSV input format | `README.md` + `sample_data/` | Table of required columns, types, example values |
| 4.5 | Write installation troubleshooting section | `README.md` | Common issues: CBC not found, Python version, openpyxl missing |
| 4.6 | Final cross-platform smoke test | — | Run on Mac (or simulate) and Windows; verify no import errors |
| 4.7 | Remove `Optimization.py` (original) | — | Logic fully migrated; keep only in git history |
| 4.8 | Tag release `v1.0.0` | git | Clean commit history with meaningful messages |

#### README.md Structure
```
# AllocationModel
Short description of the problem and what the tool does.

## Requirements
Python 3.9+

## Installation
git clone ...
pip install -r requirements.txt
python main.py

## Input Format
Table: column name | type | description

## How to Use
Step-by-step with screenshots of each screen

## ML Features
Explain auto-classification and What-If analysis in plain English

## Troubleshooting
Common errors and fixes
```

**Exit criteria:** A colleague with no prior context can clone the repo, follow the README, and produce an allocation result within 5 minutes.

---

## Dependency Map (Sprint order)

```
Sprint 1 (Foundation)
    └── Sprint 2 (ML)          ← needs data_loader schema
          └── Sprint 3 (GUI)   ← needs optimizer + ML modules
                └── Sprint 4 (Docs & Packaging)
```

---

## Input CSV Format (Spec)

The application will expect the following columns. Column names are case-insensitive.

| Column | Type | Description |
|---|---|---|
| `deal_id` | string / int | Unique identifier for each deal |
| `deal_value` | float | Monetary value of the deal |
| `deal_type` | string | `"Prepay"` or `"PPA"` (optional if using ML auto-classify) |
| `purchaser_id` | string / int | Unique identifier for each purchaser |
| `purchaser_max` | float | Maximum allocation budget for the purchaser |
| `purchaser_preference` | string | `"Prepay"` or `"PPA"` — preferred deal type |

Two-table format: deals rows + purchaser rows can be in separate sheets (Excel) or separate CSV files. The GUI will handle both.

---

## Out of Scope (v1.0)

- Web deployment (Streamlit / cloud hosting)
- User authentication
- Database persistence of runs
- Real-time collaboration
- Executable (`.exe` / `.app`) packaging — can be Sprint 5 if needed

"""
data_loader.py
--------------
Cross-platform data loading for the AllocationModel.

Replaces the original win32com.client Excel automation with pandas + openpyxl,
enabling support on both Windows and macOS.

Supported input formats
-----------------------
CSV (two files):
    - deals file     : columns deal_id, deal_value, deal_type
    - purchasers file: columns purchaser_id, purchaser_max, purchaser_preference

Excel (.xlsx / .xlsm, one file, two sheets):
    - "Deals"      sheet: columns deal_id, deal_value, deal_type
    - "Purchasers" sheet: columns purchaser_id, purchaser_max, purchaser_preference

Column names are case-insensitive and leading/trailing whitespace is stripped.

Deal types and purchaser preferences accept: "Prepay" or "PPA" (case-insensitive).
"""

from __future__ import annotations

import pandas as pd
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional


# ---------------------------------------------------------------------------
# Data schema
# ---------------------------------------------------------------------------

@dataclass
class AllocationInput:
    """Normalised input for the optimiser.

    Attributes
    ----------
    deals : list of (deal_id: str, deal_value: int)
    deals_type : list of (deal_id: str, type_code: int)
        type_code 0 = Prepay, 1 = PPA
    purchasers : list of int
        Maximum allocation budget for each purchaser (same order as allocation_pref).
    allocation_pref : list of int
        Preferred deal type per purchaser: 0 = Prepay, 1 = PPA.
    purchaser_ids : list of str
        Human-readable purchaser identifiers (for display / export).
    target_gap : float
        MIP gap tolerance (reserved for future use).
    min_deal : bool
        If True, each purchaser with sufficient budget must receive at least one deal.
    pref_penalty : bool
        If True, soft penalty is applied when a purchaser receives a non-preferred deal type.
    """
    deals: list[tuple[str, int]] = field(default_factory=list)
    deals_type: list[tuple[str, int]] = field(default_factory=list)
    purchasers: list[int] = field(default_factory=list)
    allocation_pref: list[int] = field(default_factory=list)
    purchaser_ids: list[str] = field(default_factory=list)
    target_gap: float = 0.01
    min_deal: bool = True
    pref_penalty: bool = True


@dataclass
class AllocationResult:
    """Output produced by the optimiser.

    Attributes
    ----------
    allocations : list of int
        Purchaser index (1-based) assigned to each deal.  0 means unallocated.
    status : str
        Solver status string (e.g. "Optimal", "Feasible", "Infeasible").
    total_value : float
        Sum of values of all allocated deals.
    unallocated_count : int
        Number of deals that could not be allocated.
    deal_ids : list of str
        Deal identifiers in the same order as `allocations`.
    purchaser_ids : list of str
        Purchaser identifiers for the 1-based index lookup.
    """
    allocations: list[int] = field(default_factory=list)
    status: str = ""
    total_value: float = 0.0
    unallocated_count: int = 0
    deal_ids: list[str] = field(default_factory=list)
    purchaser_ids: list[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

_DEAL_TYPE_MAP = {"prepay": 0, "ppa": 1}
_DEAL_TYPE_LABELS = {0: "Prepay", 1: "PPA"}

DEALS_REQUIRED_COLS = {"deal_id", "deal_value", "deal_type"}
PURCHASERS_REQUIRED_COLS = {"purchaser_id", "purchaser_max", "purchaser_preference"}


def _normalise_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Lower-case column names and strip whitespace."""
    df.columns = [c.strip().lower() for c in df.columns]
    return df


def _parse_deal_type(value: str, row_label: str) -> int:
    """Convert a deal-type string to its integer code."""
    key = str(value).strip().lower()
    if key not in _DEAL_TYPE_MAP:
        raise ValueError(
            f"Unknown deal type '{value}' in row '{row_label}'. "
            "Expected 'Prepay' or 'PPA'."
        )
    return _DEAL_TYPE_MAP[key]


def _validate_columns(df: pd.DataFrame, required: set[str], source: str) -> None:
    missing = required - set(df.columns)
    if missing:
        raise ValueError(
            f"Missing columns in {source}: {sorted(missing)}. "
            f"Found: {sorted(df.columns)}"
        )


def _build_allocation_input(deals_df: pd.DataFrame, purchasers_df: pd.DataFrame) -> AllocationInput:
    """Convert validated DataFrames into an AllocationInput dataclass."""
    deals_df = _normalise_columns(deals_df.copy())
    purchasers_df = _normalise_columns(purchasers_df.copy())

    _validate_columns(deals_df, DEALS_REQUIRED_COLS, "deals data")
    _validate_columns(purchasers_df, PURCHASERS_REQUIRED_COLS, "purchasers data")

    # --- Deals ---
    deals: list[tuple[str, int]] = []
    deals_type: list[tuple[str, int]] = []

    for _, row in deals_df.iterrows():
        deal_id = str(row["deal_id"]).strip()
        try:
            deal_value = int(float(row["deal_value"]))
        except (ValueError, TypeError):
            raise ValueError(f"deal_value for '{deal_id}' is not numeric: '{row['deal_value']}'")

        type_code = _parse_deal_type(row["deal_type"], deal_id)
        deals.append((deal_id, deal_value))
        deals_type.append((deal_id, type_code))

    # --- Purchasers ---
    purchasers: list[int] = []
    allocation_pref: list[int] = []
    purchaser_ids: list[str] = []

    for _, row in purchasers_df.iterrows():
        p_id = str(row["purchaser_id"]).strip()
        try:
            p_max = int(float(row["purchaser_max"]))
        except (ValueError, TypeError):
            raise ValueError(f"purchaser_max for '{p_id}' is not numeric: '{row['purchaser_max']}'")

        pref_code = _parse_deal_type(row["purchaser_preference"], p_id)
        purchasers.append(p_max)
        allocation_pref.append(pref_code)
        purchaser_ids.append(p_id)

    return AllocationInput(
        deals=deals,
        deals_type=deals_type,
        purchasers=purchasers,
        allocation_pref=allocation_pref,
        purchaser_ids=purchaser_ids,
    )


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def load_csv(deals_path: str | Path, purchasers_path: str | Path) -> AllocationInput:
    """Load allocation input from two CSV files.

    Parameters
    ----------
    deals_path : path to the deals CSV file.
    purchasers_path : path to the purchasers CSV file.

    Returns
    -------
    AllocationInput
    """
    deals_df = pd.read_csv(deals_path)
    purchasers_df = pd.read_csv(purchasers_path)
    return _build_allocation_input(deals_df, purchasers_df)


def load_excel(
    file_path: str | Path,
    deals_sheet: str = "Deals",
    purchasers_sheet: str = "Purchasers",
) -> AllocationInput:
    """Load allocation input from an Excel file (.xlsx or .xlsm).

    The workbook must contain two sheets named (by default) "Deals" and
    "Purchasers".  Sheet names are case-sensitive.

    Parameters
    ----------
    file_path : path to the Excel file.
    deals_sheet : name of the sheet containing deal data.
    purchasers_sheet : name of the sheet containing purchaser data.

    Returns
    -------
    AllocationInput
    """
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"Excel file not found: {path}")

    xl = pd.ExcelFile(path, engine="openpyxl")
    available_sheets = xl.sheet_names

    if deals_sheet not in available_sheets:
        raise ValueError(
            f"Sheet '{deals_sheet}' not found in '{path.name}'. "
            f"Available sheets: {available_sheets}"
        )
    if purchasers_sheet not in available_sheets:
        raise ValueError(
            f"Sheet '{purchasers_sheet}' not found in '{path.name}'. "
            f"Available sheets: {available_sheets}"
        )

    deals_df = xl.parse(deals_sheet)
    purchasers_df = xl.parse(purchasers_sheet)
    return _build_allocation_input(deals_df, purchasers_df)


def load_file(
    file_path: str | Path,
    purchasers_path: Optional[str | Path] = None,
) -> AllocationInput:
    """Auto-detect file format and load allocation input.

    If `file_path` is a CSV, `purchasers_path` must also be provided.
    If `file_path` is an Excel file, `purchasers_path` is ignored.

    Parameters
    ----------
    file_path : path to the primary data file.
    purchasers_path : path to purchasers CSV (required when file_path is CSV).

    Returns
    -------
    AllocationInput
    """
    path = Path(file_path)
    suffix = path.suffix.lower()

    if suffix in (".xlsx", ".xlsm"):
        return load_excel(path)

    if suffix == ".csv":
        if purchasers_path is None:
            raise ValueError(
                "A purchasers CSV path must be provided when loading from CSV."
            )
        return load_csv(path, purchasers_path)

    raise ValueError(
        f"Unsupported file format: '{suffix}'. "
        "Supported formats: .csv, .xlsx, .xlsm"
    )


def get_raw_dataframes(
    file_path: str | Path,
    purchasers_path: Optional[str | Path] = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Return raw DataFrames (deals, purchasers) without parsing into AllocationInput.

    Useful for previewing data in the GUI before running the optimiser.

    Returns
    -------
    (deals_df, purchasers_df)
    """
    path = Path(file_path)
    suffix = path.suffix.lower()

    if suffix in (".xlsx", ".xlsm"):
        xl = pd.ExcelFile(path, engine="openpyxl")
        return xl.parse("Deals"), xl.parse("Purchasers")

    if suffix == ".csv":
        if purchasers_path is None:
            raise ValueError("purchasers_path is required when loading CSV.")
        return pd.read_csv(path), pd.read_csv(purchasers_path)

    raise ValueError(f"Unsupported file format: '{suffix}'.")

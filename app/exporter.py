"""
exporter.py
-----------
Export AllocationResult to CSV and styled Excel (.xlsx).

CSV output
----------
Single flat file with columns:
    deal_id, deal_value, deal_type, assigned_purchaser, assigned_purchaser_id

Excel output
------------
Two sheets:
    "Allocation Results" — per-deal table (same columns as CSV)
    "Summary"           — metrics: total value, % allocated, capacity usage per purchaser
"""

from __future__ import annotations

from pathlib import Path
from typing import Optional

import pandas as pd

from app.data_loader import AllocationInput, AllocationResult


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _build_results_df(
    data: AllocationInput,
    result: AllocationResult,
) -> pd.DataFrame:
    """Build a flat DataFrame with one row per deal."""
    rows = []
    deal_type_map = dict(data.deals_type)
    type_labels   = {0: "Prepay", 1: "PPA"}

    for i, (deal_id, deal_value) in enumerate(data.deals):
        p_idx = result.allocations[i]
        p_name = data.purchaser_ids[p_idx - 1] if p_idx > 0 else "Unallocated"
        rows.append({
            "Deal ID":         deal_id,
            "Deal Value ($)":  deal_value,
            "Deal Type":       type_labels.get(deal_type_map.get(deal_id, 0), "Unknown"),
            "Purchaser":       p_name,
            "Purchaser Index": p_idx if p_idx > 0 else "",
        })

    return pd.DataFrame(rows)


def _build_summary_df(
    data: AllocationInput,
    result: AllocationResult,
) -> pd.DataFrame:
    """Build a summary DataFrame with per-purchaser metrics."""
    # Per-purchaser totals
    purchaser_totals: dict[str, float] = {p: 0.0 for p in data.purchaser_ids}
    for i, p_idx in enumerate(result.allocations):
        if p_idx > 0:
            p_name = data.purchaser_ids[p_idx - 1]
            purchaser_totals[p_name] += data.deals[i][1]

    rows = []
    for idx, p_name in enumerate(data.purchaser_ids):
        budget    = data.purchasers[idx]
        allocated = purchaser_totals[p_name]
        pct_used  = (allocated / budget * 100) if budget > 0 else 0.0
        rows.append({
            "Purchaser":           p_name,
            "Budget ($)":          budget,
            "Allocated ($)":       round(allocated, 2),
            "Remaining ($)":       round(budget - allocated, 2),
            "Capacity Used (%)":   round(pct_used, 1),
        })

    # Totals row
    total_budget    = sum(data.purchasers)
    total_allocated = result.total_value
    rows.append({
        "Purchaser":         "TOTAL",
        "Budget ($)":        total_budget,
        "Allocated ($)":     round(total_allocated, 2),
        "Remaining ($)":     round(total_budget - total_allocated, 2),
        "Capacity Used (%)": round(total_allocated / total_budget * 100, 1) if total_budget > 0 else 0.0,
    })

    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def export_csv(
    data: AllocationInput,
    result: AllocationResult,
    output_path: str | Path,
) -> Path:
    """Export allocation results to a CSV file.

    Parameters
    ----------
    data        : Original AllocationInput.
    result      : AllocationResult from the optimiser.
    output_path : Destination file path (.csv).

    Returns
    -------
    Resolved output path.
    """
    path = Path(output_path)
    df = _build_results_df(data, result)
    df.to_csv(path, index=False)
    return path.resolve()


def export_excel(
    data: AllocationInput,
    result: AllocationResult,
    output_path: str | Path,
) -> Path:
    """Export allocation results to a styled Excel file.

    Creates two sheets:
      - "Allocation Results": per-deal table
      - "Summary": per-purchaser capacity usage metrics

    Parameters
    ----------
    data        : Original AllocationInput.
    result      : AllocationResult from the optimiser.
    output_path : Destination file path (.xlsx).

    Returns
    -------
    Resolved output path.
    """
    path = Path(output_path)
    results_df = _build_results_df(data, result)
    summary_df = _build_summary_df(data, result)

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        results_df.to_excel(writer, sheet_name="Allocation Results", index=False)
        summary_df.to_excel(writer, sheet_name="Summary", index=False)

        # Basic column width formatting
        for sheet_name, df in [("Allocation Results", results_df), ("Summary", summary_df)]:
            ws = writer.sheets[sheet_name]
            for col_idx, col in enumerate(df.columns, start=1):
                max_len = max(len(str(col)), df[col].astype(str).str.len().max())
                ws.column_dimensions[ws.cell(1, col_idx).column_letter].width = min(max_len + 4, 40)

    return path.resolve()

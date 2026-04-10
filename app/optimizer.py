"""
optimizer.py
------------
Core ILP optimisation logic for the AllocationModel.

Refactored from the original Optimization.py with the following changes:
  - Removed win32com dependency (cross-platform).
  - Replaced GLPK solver with CBC (bundled with PuLP via pip install pulp).
  - Accepts AllocationInput / returns AllocationResult (no Excel I/O).
  - Fixed typo: get_vaiables -> logic now lives in data_loader.py.
  - All parameters are explicit arguments; nothing is hardcoded.
  - Added structured logging via the standard library instead of bare prints.

The mathematical model is unchanged from the original.
"""

from __future__ import annotations

import logging
from typing import Callable, Optional

import pulp

from app.data_loader import AllocationInput, AllocationResult

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _compute_min_deal_value(deals: list[tuple[str, int]]) -> int:
    """Return the smallest non-zero deal value."""
    values = [v for _, v in deals if v > 0]
    return min(values) if values else 0


def _format_output(
    assigned: list[tuple[str, int]],
    deals: list[tuple[str, int]],
) -> list[int]:
    """Convert a sparse list of (deal_id, purchaser_idx) pairs to a dense list.

    Returns a list of length len(deals) where each element is the
    1-based purchaser index assigned to that deal, or 0 if unallocated.
    """
    assigned_dict = dict(assigned)
    return [
        assigned_dict[deal_id] + 1 if deal_id in assigned_dict else 0
        for deal_id, _ in deals
    ]


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def optimize(
    data: AllocationInput,
    time_limit: int = 60,
    verbose: bool = False,
    progress_callback: Optional[Callable[[str], None]] = None,
) -> AllocationResult:
    """Find the optimal allocation of deals to purchasers.

    Solves a Binary Integer Linear Programme that maximises total allocated
    deal value subject to:
      - Hard constraint: total value assigned to a purchaser <= purchaser budget.
      - Hard constraint: each deal is assigned to at most one purchaser.
      - Hard constraint (optional): each eligible purchaser receives >= 1 deal.
      - Soft constraint (optional): penalise mismatches between deal type and
        purchaser preference (elastic constraint, penalty = -1000).

    Parameters
    ----------
    data : AllocationInput
        Normalised problem data produced by data_loader.
    time_limit : int
        Maximum solver wall-clock time in seconds.
    verbose : bool
        If True, emit detailed solver output to the logger.
    progress_callback : callable, optional
        Called with a status string at key solver milestones.
        Useful for updating a GUI progress indicator.

    Returns
    -------
    AllocationResult
    """
    deals = data.deals
    purchasers = data.purchasers
    deals_type = data.deals_type
    allocation_pref = data.allocation_pref
    min_deal = data.min_deal
    pref_penalty = data.pref_penalty

    deal_count = len(deals)
    num_purchasers = len(purchasers)
    min_deal_value = _compute_min_deal_value(deals)

    def _log(msg: str) -> None:
        logger.info(msg)
        if progress_callback:
            progress_callback(msg)

    _log("Building optimisation model...")

    # --- Decision variables ---
    # y[i] = 1 if purchaser i receives at least one deal
    y = pulp.LpVariable.dicts(
        "BinUsed",
        range(num_purchasers),
        lowBound=0, upBound=1,
        cat=pulp.LpInteger,
    )

    # x[(deal_id, purchaser_idx)] = 1 if deal is assigned to purchaser
    possible_assignments = [
        (deal[0], p_idx)
        for deal in deals
        for p_idx in range(num_purchasers)
    ]
    x = pulp.LpVariable.dicts(
        "itemInBin",
        possible_assignments,
        lowBound=0, upBound=1,
        cat=pulp.LpInteger,
    )

    # --- Model ---
    prob = pulp.LpProblem("Deal_Allocation_Problem", pulp.LpMaximize)

    # Objective: maximise total allocated deal value
    prob += pulp.lpSum(
        deals[j][1] * x[(deals[j][0], i)]
        for i in range(num_purchasers)
        for j in range(deal_count)
    )

    # --- Hard constraints ---

    # 1. Capacity: total value allocated to purchaser i <= budget_i
    for i in range(num_purchasers):
        if purchasers[i] > 0:
            prob += (
                pulp.lpSum(deals[j][1] * x[(deals[j][0], i)] for j in range(deal_count))
                <= purchasers[i] * y[i]
            )
        else:
            prob += (
                pulp.lpSum(deals[j][1] * x[(deals[j][0], i)] for j in range(deal_count))
                == 0
            )

    # 2. Uniqueness: each deal is assigned to at most one purchaser
    for idx, item in enumerate(deals):
        if deals[idx][1] > 0:
            prob += pulp.lpSum(x[(item[0], j)] for j in range(num_purchasers)) <= 1
        else:
            prob += pulp.lpSum(x[(item[0], j)] for j in range(num_purchasers)) == 0

    # 3. Minimum deal: each purchaser with sufficient budget gets >= 1 deal
    if min_deal:
        for i in range(num_purchasers):
            if purchasers[i] >= min_deal_value:
                eligible = [
                    x[(deals[j][0], i)]
                    for j in range(deal_count)
                    if deals[j][1] > 0
                ]
                prob += pulp.lpSum(eligible) >= 1

    # --- Soft constraints (preference penalties) ---
    if pref_penalty:
        penalty = -10 * 100   # -1000 per violated preference

        type_1_deals = [(deals_type[idx], deals[idx]) for idx, item in enumerate(deals_type) if item[1] == 1]
        type_0_deals = [(deals_type[idx], deals[idx]) for idx, item in enumerate(deals_type) if item[1] == 0]

        for i, pref in enumerate(allocation_pref):
            if pref == 2 or purchasers[i] == 0:
                continue

            # Soft constraint for PPA deals (type 1)
            s = 1 * pref
            k_1_vars = [x[(type_1_deals[j][0][0], i)] for j in range(len(type_1_deals))]
            z_1 = [
                pulp.LpVariable(str(k_1_vars[j]), lowBound=0, upBound=1, cat=pulp.LpInteger)
                for j in range(len(type_1_deals))
            ]
            lhs_1 = pulp.LpAffineExpression([(z_1[j], 1) for j in range(len(type_1_deals))])
            c_1 = pulp.LpConstraint(e=lhs_1, sense=s, name=f"elastic_{i}_1", rhs=s)
            prob.extend(c_1.makeElasticSubProblem(penalty=penalty))

            # Soft constraint for Prepay deals (type 0)
            s_0 = 0 if s == 1 else 1
            k_0_vars = [x[(type_0_deals[j][0][0], i)] for j in range(len(type_0_deals))]
            z_0 = [
                pulp.LpVariable(str(k_0_vars[j]), lowBound=0, upBound=1, cat=pulp.LpInteger)
                for j in range(len(type_0_deals))
            ]
            lhs_0 = pulp.LpAffineExpression([(z_0[j], 1) for j in range(len(type_0_deals))])
            c_0 = pulp.LpConstraint(e=lhs_0, sense=s_0, name=f"elastic_{i}_0", rhs=s_0)
            prob.extend(c_0.makeElasticSubProblem(penalty=penalty))

    # --- Solve ---
    _log(f"Solving with CBC (time limit: {time_limit}s)...")

    solver = pulp.PULP_CBC_CMD(timeLimit=time_limit, msg=1 if verbose else 0)

    try:
        prob.solve(solver)
    except Exception as exc:
        logger.error("Solver error: %s", exc)
        raise RuntimeError(f"Solver failed: {exc}") from exc

    status = pulp.LpStatus[prob.status]
    _log(f"Solver status: {status}")

    if status == "Infeasible":
        return AllocationResult(
            allocations=[0] * deal_count,
            status=status,
            total_value=0.0,
            unallocated_count=deal_count,
            deal_ids=[d[0] for d in deals],
            purchaser_ids=data.purchaser_ids,
        )

    # --- Extract solution ---
    assigned = [
        key
        for key, var in x.items()
        if var.value() is not None and var.value() >= 0.5
    ]

    allocations = _format_output(assigned, deals)

    total_value = sum(
        deals[i][1]
        for i, p_idx in enumerate(allocations)
        if p_idx > 0
    )
    unallocated_count = allocations.count(0)

    _log(
        f"Done. Total allocated value: {total_value:,.0f} | "
        f"Unallocated deals: {unallocated_count}/{deal_count}"
    )

    return AllocationResult(
        allocations=allocations,
        status=status,
        total_value=float(total_value),
        unallocated_count=unallocated_count,
        deal_ids=[d[0] for d in deals],
        purchaser_ids=data.purchaser_ids,
    )

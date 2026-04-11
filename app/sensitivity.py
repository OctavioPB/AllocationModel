"""
sensitivity.py
--------------
What-If sensitivity analysis for the AllocationModel.

Given a base optimal result, this module answers the question:
    "What happens to the total allocated value if Purchaser X changes
     their budget by Y%?"

Approach
--------
1. Run the optimiser with the original data to obtain a baseline.
2. For a chosen purchaser, perturb their capacity by a fixed set of deltas
   (default: -20%, -10%, -5%, +5%, +10%, +20%).
3. Re-run the optimiser (with a short time limit) for each perturbation.
4. Fit a LinearRegression on (capacity_delta_pct → total_value_delta).
5. Use the regression to give instant estimates for any requested delta.

The analysis runs on demand (user picks a purchaser and clicks "Analyse").
A shorter time limit (default: 10s) is used for perturbation runs to keep
the interaction snappy while still finding good-quality solutions.

Public API
----------
    compute_sensitivity(data, purchaser_idx, ...)  -> SensitivityModel
    what_if(model, capacity_change_pct)            -> WhatIfResult
    explain(result)                                -> str (plain English)
"""

from __future__ import annotations

import copy
import logging
from dataclasses import dataclass, field
from typing import Callable, Optional

import numpy as np
from sklearn.linear_model import LinearRegression

from app.data_loader import AllocationInput, AllocationResult
from app.optimizer import optimize

logger = logging.getLogger(__name__)

# Default perturbation levels (fractional, e.g. 0.10 = 10%)
_DEFAULT_PERTURBATIONS = [-0.20, -0.10, -0.05, 0.05, 0.10, 0.20]


# ---------------------------------------------------------------------------
# Output schema
# ---------------------------------------------------------------------------

@dataclass
class PerturbationPoint:
    """A single data point used to build the regression model.

    Attributes
    ----------
    capacity_change_pct : float  The perturbation applied, e.g. 0.10 = +10%.
    new_capacity        : int    Purchaser budget after perturbation.
    total_value         : float  Optimal total allocated value at this capacity.
    value_delta         : float  Change vs. baseline (positive = improvement).
    solver_status       : str    CBC solver status for this run.
    """
    capacity_change_pct: float
    new_capacity: int
    total_value: float
    value_delta: float
    solver_status: str


@dataclass
class SensitivityModel:
    """Fitted sensitivity model for one purchaser.

    Attributes
    ----------
    purchaser_idx       : int   0-based index in AllocationInput.purchasers.
    purchaser_id        : str   Human-readable purchaser name.
    base_capacity       : int   Original budget.
    base_total_value    : float Baseline total allocated value.
    points              : list  Perturbation data points used to fit the model.
    r_squared           : float R² of the linear fit (1.0 = perfect linear).
    slope               : float Marginal value per unit of capacity added ($/$).
    intercept           : float Regression intercept.
    warning             : str   Non-empty if the fit is poor or data is sparse.
    """
    purchaser_idx: int
    purchaser_id: str
    base_capacity: int
    base_total_value: float
    points: list[PerturbationPoint] = field(default_factory=list)
    r_squared: float = 0.0
    slope: float = 0.0
    intercept: float = 0.0
    warning: str = ""


@dataclass
class WhatIfResult:
    """Result of a single What-If query.

    Attributes
    ----------
    purchaser_id          : str   Purchaser name.
    capacity_change_pct   : float Requested capacity change (e.g. 0.10 = +10%).
    original_capacity     : int   Budget before the change.
    new_capacity          : int   Budget after the change.
    estimated_value_delta : float Predicted change in total allocated value.
    estimated_total_value : float Predicted total allocated value.
    explanation           : str   Plain-English summary.
    is_estimate           : bool  True if derived from regression (not exact solve).
    """
    purchaser_id: str
    capacity_change_pct: float
    original_capacity: int
    new_capacity: int
    estimated_value_delta: float
    estimated_total_value: float
    explanation: str
    is_estimate: bool = True


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _perturb_data(data: AllocationInput, purchaser_idx: int, delta_pct: float) -> AllocationInput:
    """Return a deep copy of AllocationInput with one purchaser's capacity scaled."""
    perturbed = copy.deepcopy(data)
    original = perturbed.purchasers[purchaser_idx]
    perturbed.purchasers[purchaser_idx] = max(0, int(original * (1 + delta_pct)))
    return perturbed


def _fit_regression(
    deltas_pct: list[float],
    value_deltas: list[float],
) -> tuple[LinearRegression, float]:
    """Fit LinearRegression on (capacity_change_pct → value_delta).

    Returns (fitted_model, r_squared).
    """
    X = np.array(deltas_pct).reshape(-1, 1)
    y = np.array(value_deltas)
    model = LinearRegression()
    model.fit(X, y)
    r_sq = float(model.score(X, y))
    return model, r_sq


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def compute_sensitivity(
    data: AllocationInput,
    purchaser_idx: int,
    base_total_value: float,
    perturbations: list[float] = _DEFAULT_PERTURBATIONS,
    solver_time_limit: int = 10,
    progress_callback: Optional[Callable[[str], None]] = None,
) -> SensitivityModel:
    """Compute sensitivity of total allocated value to one purchaser's capacity.

    Runs the optimiser once per perturbation level. Uses a short time limit
    to keep the analysis interactive.

    Parameters
    ----------
    data             : AllocationInput — original problem data (not modified).
    purchaser_idx    : 0-based index of the purchaser to analyse.
    base_total_value : Total value from the baseline optimisation run.
    perturbations    : List of fractional capacity changes to test.
    solver_time_limit: Max seconds per perturbation run (default: 10).
    progress_callback: Optional callable for GUI progress updates.

    Returns
    -------
    SensitivityModel
    """
    purchaser_id = (
        data.purchaser_ids[purchaser_idx]
        if purchaser_idx < len(data.purchaser_ids)
        else f"Purchaser {purchaser_idx + 1}"
    )
    base_capacity = data.purchasers[purchaser_idx]

    def _log(msg: str) -> None:
        logger.info(msg)
        if progress_callback:
            progress_callback(msg)

    _log(f"Computing sensitivity for '{purchaser_id}' (base capacity: {base_capacity:,})...")

    points: list[PerturbationPoint] = []
    deltas_pct: list[float] = []
    value_deltas: list[float] = []

    for delta in perturbations:
        perturbed_data = _perturb_data(data, purchaser_idx, delta)
        new_cap = perturbed_data.purchasers[purchaser_idx]

        _log(f"  Testing {delta:+.0%} capacity ({new_cap:,})...")

        try:
            result = optimize(perturbed_data, time_limit=solver_time_limit, verbose=False)
        except Exception as exc:
            logger.warning("Perturbation run failed for delta=%.2f: %s", delta, exc)
            continue

        v_delta = result.total_value - base_total_value

        points.append(PerturbationPoint(
            capacity_change_pct=delta,
            new_capacity=new_cap,
            total_value=result.total_value,
            value_delta=v_delta,
            solver_status=result.status,
        ))
        deltas_pct.append(delta)
        value_deltas.append(v_delta)

    # Fit regression
    warning = ""
    slope, intercept, r_sq = 0.0, 0.0, 0.0

    if len(points) < 2:
        warning = "Not enough perturbation points to fit a regression model."
        _log(f"  Warning: {warning}")
    else:
        reg_model, r_sq = _fit_regression(deltas_pct, value_deltas)
        slope = float(reg_model.coef_[0])
        intercept = float(reg_model.intercept_)

        if r_sq < 0.7:
            warning = (
                f"The relationship between capacity and value is non-linear "
                f"(R²={r_sq:.2f}). Estimates may be less accurate at extreme values."
            )

    _log(
        f"  Sensitivity analysis complete: slope={slope:,.0f}, R²={r_sq:.2f}"
    )

    return SensitivityModel(
        purchaser_idx=purchaser_idx,
        purchaser_id=purchaser_id,
        base_capacity=base_capacity,
        base_total_value=base_total_value,
        points=points,
        r_squared=r_sq,
        slope=slope,
        intercept=intercept,
        warning=warning,
    )


def what_if(
    model: SensitivityModel,
    capacity_change_pct: float,
) -> WhatIfResult:
    """Estimate the effect of a capacity change using the fitted regression.

    Parameters
    ----------
    model                : SensitivityModel from compute_sensitivity().
    capacity_change_pct  : Fractional change to apply, e.g. 0.10 for +10%.

    Returns
    -------
    WhatIfResult with plain-English explanation.
    """
    new_capacity = max(0, int(model.base_capacity * (1 + capacity_change_pct)))

    # Use regression to estimate value delta
    estimated_delta = model.slope * capacity_change_pct + model.intercept
    estimated_total = model.base_total_value + estimated_delta

    result = WhatIfResult(
        purchaser_id=model.purchaser_id,
        capacity_change_pct=capacity_change_pct,
        original_capacity=model.base_capacity,
        new_capacity=new_capacity,
        estimated_value_delta=round(estimated_delta, 2),
        estimated_total_value=round(max(0, estimated_total), 2),
        explanation=_build_explanation(
            purchaser_id=model.purchaser_id,
            capacity_change_pct=capacity_change_pct,
            original_capacity=model.base_capacity,
            new_capacity=new_capacity,
            estimated_delta=estimated_delta,
            estimated_total=estimated_total,
            is_estimate=True,
            r_squared=model.r_squared,
        ),
        is_estimate=True,
    )
    return result


def what_if_exact(
    data: AllocationInput,
    model: SensitivityModel,
    capacity_change_pct: float,
    solver_time_limit: int = 30,
    progress_callback: Optional[Callable[[str], None]] = None,
) -> WhatIfResult:
    """Run an exact optimisation for a specific capacity change.

    Slower than what_if() but gives the precise answer from the solver.

    Parameters
    ----------
    data                 : Original AllocationInput (not modified).
    model                : SensitivityModel (used for metadata).
    capacity_change_pct  : Fractional change to apply.
    solver_time_limit    : Max seconds for this solve.
    progress_callback    : Optional GUI progress callback.

    Returns
    -------
    WhatIfResult with is_estimate=False and exact solver output.
    """
    perturbed = _perturb_data(data, model.purchaser_idx, capacity_change_pct)
    new_capacity = perturbed.purchasers[model.purchaser_idx]

    result = optimize(
        perturbed,
        time_limit=solver_time_limit,
        verbose=False,
        progress_callback=progress_callback,
    )

    exact_delta = result.total_value - model.base_total_value

    return WhatIfResult(
        purchaser_id=model.purchaser_id,
        capacity_change_pct=capacity_change_pct,
        original_capacity=model.base_capacity,
        new_capacity=new_capacity,
        estimated_value_delta=round(exact_delta, 2),
        estimated_total_value=round(result.total_value, 2),
        explanation=_build_explanation(
            purchaser_id=model.purchaser_id,
            capacity_change_pct=capacity_change_pct,
            original_capacity=model.base_capacity,
            new_capacity=new_capacity,
            estimated_delta=exact_delta,
            estimated_total=result.total_value,
            is_estimate=False,
            r_squared=None,
        ),
        is_estimate=False,
    )


def _build_explanation(
    purchaser_id: str,
    capacity_change_pct: float,
    original_capacity: int,
    new_capacity: int,
    estimated_delta: float,
    estimated_total: float,
    is_estimate: bool,
    r_squared: Optional[float],
) -> str:
    """Build a plain-English explanation of a What-If result."""
    direction = "increase" if capacity_change_pct >= 0 else "decrease"
    abs_pct = abs(capacity_change_pct) * 100
    capacity_diff = abs(new_capacity - original_capacity)

    if abs(estimated_delta) < 1:
        impact_line = (
            f"This change would have no meaningful impact on total allocated value."
        )
    elif estimated_delta > 0:
        impact_line = (
            f"Total allocated value would {'increase' if estimated_delta > 0 else 'decrease'} "
            f"by approximately ${abs(estimated_delta):,.0f}, "
            f"reaching ${max(0, estimated_total):,.0f}."
        )
    else:
        impact_line = (
            f"Total allocated value would decrease "
            f"by approximately ${abs(estimated_delta):,.0f}, "
            f"reaching ${max(0, estimated_total):,.0f}."
        )

    source_note = (
        "This is an estimate based on a linear regression of nearby scenarios."
        if is_estimate
        else "This is an exact result from the optimiser."
    )

    confidence_note = ""
    if is_estimate and r_squared is not None and r_squared < 0.7:
        confidence_note = (
            f" Note: the linear fit has R²={r_squared:.2f}, "
            "so this estimate may be less reliable. "
            "Use 'Run exact analysis' for a precise answer."
        )

    return (
        f"If {purchaser_id}'s budget {direction}s by {abs_pct:.0f}% "
        f"(${original_capacity:,} → ${new_capacity:,}, a ${capacity_diff:,} {direction}): "
        f"{impact_line} "
        f"{source_note}{confidence_note}"
    )

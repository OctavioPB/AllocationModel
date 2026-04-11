"""
ml_classifier.py
----------------
K-Means based deal type auto-classifier.

Given a list of deals with known or unknown types, this module clusters them
into two groups (Prepay / PPA) using K-Means on deal value.

Typical use case: the user provides a CSV without a `deal_type` column, or
wants to verify that their manual labels are consistent with the value
distribution before running the optimiser.

Output per deal
---------------
- suggested_type : str   "Prepay" or "PPA"
- confidence     : float 0.0–1.0  (1.0 = centroid distance is 0)
- is_override    : bool  True if the suggestion differs from the original label

Design decisions
----------------
- Features used: log-scaled deal_value.
  Log scaling makes K-Means robust when one cluster contains very large values
  and the other small values (e.g. PPA deals 70k-310k vs Prepay deals 4k-46k).
- Cluster → label mapping: the cluster with the higher centroid mean
  is assigned "PPA"; the lower is assigned "Prepay".
- Confidence: based on normalised softmax of negative distances to both
  centroids, so it reflects how far into its cluster a deal sits.
- Seed: fixed random_state=42 for reproducibility.
"""

from __future__ import annotations

import logging
import math
from dataclasses import dataclass, field

import numpy as np
from sklearn.cluster import KMeans
from sklearn.metrics import silhouette_score
from sklearn.preprocessing import StandardScaler

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Output schema
# ---------------------------------------------------------------------------

@dataclass
class ClassificationResult:
    """Classification output for a single deal.

    Attributes
    ----------
    deal_id         : str   Original deal identifier.
    original_type   : str   Type provided in input data ("" if unknown).
    suggested_type  : str   Type suggested by the classifier.
    confidence      : float How confidently the deal belongs to suggested_type (0–1).
    is_override     : bool  True if suggestion differs from original_type.
    """
    deal_id: str
    original_type: str
    suggested_type: str
    confidence: float
    is_override: bool


@dataclass
class ClassificationSummary:
    """Full output of the auto-classification step.

    Attributes
    ----------
    results         : list of ClassificationResult, one per deal.
    method_used     : str  "kmeans" | "passthrough" | "single_type_fallback".
    override_count  : int  Number of deals whose suggested type differs from input.
    prepay_count    : int  Deals classified as Prepay.
    ppa_count       : int  Deals classified as PPA.
    warning         : str  Non-empty when a fallback was triggered.
    """
    results: list[ClassificationResult] = field(default_factory=list)
    method_used: str = "kmeans"
    override_count: int = 0
    prepay_count: int = 0
    ppa_count: int = 0
    warning: str = ""


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

_LABEL_PREPAY = "Prepay"
_LABEL_PPA = "PPA"
_LABELS = [_LABEL_PREPAY, _LABEL_PPA]

_VALID_TYPES = {t.lower() for t in _LABELS}


def _log_scale(values: list[float]) -> np.ndarray:
    """Apply log1p scaling to deal values for K-Means feature engineering."""
    return np.array([math.log1p(max(v, 0)) for v in values]).reshape(-1, 1)


def _softmax_confidence(d_assigned: float, d_other: float) -> float:
    """Convert two centroid distances into a confidence score [0, 1].

    Uses a softmax on negative distances: closer to own centroid = higher confidence.
    """
    # Avoid division by zero
    if d_assigned + d_other == 0:
        return 1.0
    exp_own   = math.exp(-d_assigned)
    exp_other = math.exp(-d_other)
    return exp_own / (exp_own + exp_other)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def classify_deals(
    deal_ids: list[str],
    deal_values: list[float],
    original_types: list[str] | None = None,
) -> ClassificationSummary:
    """Classify deals into Prepay / PPA using K-Means clustering.

    Parameters
    ----------
    deal_ids      : Unique identifier for each deal.
    deal_values   : Monetary value of each deal.
    original_types: Existing type labels (optional). Pass None or empty strings
                    when types are unknown. Used to detect overrides.

    Returns
    -------
    ClassificationSummary
    """
    n = len(deal_ids)
    if n == 0:
        return ClassificationSummary(warning="No deals provided.")

    if original_types is None:
        original_types = [""] * n

    if len(deal_values) != n or len(original_types) != n:
        raise ValueError("deal_ids, deal_values, and original_types must have the same length.")

    # --- Fallback: only one deal ---
    if n == 1:
        label = _LABEL_PPA if deal_values[0] > 0 else _LABEL_PREPAY
        result = ClassificationResult(
            deal_id=deal_ids[0],
            original_type=original_types[0],
            suggested_type=label,
            confidence=1.0,
            is_override=(original_types[0].lower() not in ("", label.lower())),
        )
        return ClassificationSummary(
            results=[result],
            method_used="single_deal_fallback",
            override_count=int(result.is_override),
            prepay_count=int(label == _LABEL_PREPAY),
            ppa_count=int(label == _LABEL_PPA),
            warning="Only one deal provided; classification is based on value sign only.",
        )

    # --- Fallback: all values are zero ---
    if all(v == 0 for v in deal_values):
        results = [
            ClassificationResult(
                deal_id=deal_ids[i],
                original_type=original_types[i],
                suggested_type=_LABEL_PREPAY,
                confidence=1.0,
                is_override=(original_types[i].lower() not in ("", _LABEL_PREPAY.lower())),
            )
            for i in range(n)
        ]
        return ClassificationSummary(
            results=results,
            method_used="zero_value_fallback",
            override_count=sum(r.is_override for r in results),
            prepay_count=n,
            ppa_count=0,
            warning="All deal values are zero; all deals labelled as Prepay.",
        )

    # --- Fallback: only one unique non-zero value level ---
    unique_values = set(v for v in deal_values if v > 0)
    if len(unique_values) == 1:
        label = _LABEL_PPA if list(unique_values)[0] > 50_000 else _LABEL_PREPAY
        results = [
            ClassificationResult(
                deal_id=deal_ids[i],
                original_type=original_types[i],
                suggested_type=label,
                confidence=1.0,
                is_override=(original_types[i].lower() not in ("", label.lower())),
            )
            for i in range(n)
        ]
        return ClassificationSummary(
            results=results,
            method_used="single_type_fallback",
            override_count=sum(r.is_override for r in results),
            prepay_count=sum(r.suggested_type == _LABEL_PREPAY for r in results),
            ppa_count=sum(r.suggested_type == _LABEL_PPA for r in results),
            warning=f"All deals have the same value; all labelled as {label}.",
        )

    # --- K-Means clustering ---
    X = _log_scale(deal_values)

    # Determine optimal k: use k=2 only when the data has meaningful cluster
    # separation (silhouette score >= 0.3). If the data is unimodal (e.g. all
    # Prepay-range values), k=1 prevents the classifier from fabricating a PPA
    # label for the higher-valued Prepay sub-group.
    value_range = max(deal_values) - min(v for v in deal_values if v > 0)
    if value_range == 0 or n < 2:
        k = 1
    elif n >= 4:
        # Quick silhouette check to confirm two real clusters exist
        km_test = KMeans(n_clusters=2, random_state=42, n_init=10)
        labels_test = km_test.fit_predict(X)
        try:
            sil = silhouette_score(X, labels_test)
        except ValueError:
            sil = 0.0
        k = 2 if sil >= 0.3 else 1
    else:
        k = 2

    kmeans = KMeans(n_clusters=k, random_state=42, n_init=10)
    kmeans.fit(X)
    labels_raw = kmeans.labels_  # 0 or 1

    centroids_log = kmeans.cluster_centers_.flatten()

    if k == 1:
        # Only one cluster — determine label from centroid value
        centroid_val = math.expm1(centroids_log[0])
        single_label = _LABEL_PPA if centroid_val > 50_000 else _LABEL_PREPAY
        results = [
            ClassificationResult(
                deal_id=deal_ids[i],
                original_type=original_types[i],
                suggested_type=single_label,
                confidence=1.0,
                is_override=(original_types[i].lower() not in ("", single_label.lower())),
            )
            for i in range(n)
        ]
        return ClassificationSummary(
            results=results,
            method_used="kmeans_single_cluster",
            override_count=sum(r.is_override for r in results),
            prepay_count=sum(r.suggested_type == _LABEL_PREPAY for r in results),
            ppa_count=sum(r.suggested_type == _LABEL_PPA for r in results),
            warning="K-Means found only one cluster; all deals labelled the same.",
        )

    # Map cluster index → label: higher centroid = PPA
    cluster_to_label = {
        int(np.argmax(centroids_log)): _LABEL_PPA,
        int(np.argmin(centroids_log)): _LABEL_PREPAY,
    }

    # Compute distances to both centroids for confidence scoring
    distances = kmeans.transform(X)  # shape (n, 2)

    results = []
    for i in range(n):
        cluster_idx = int(labels_raw[i])
        suggested = cluster_to_label[cluster_idx]
        other_idx = 1 - cluster_idx

        confidence = _softmax_confidence(
            d_assigned=distances[i, cluster_idx],
            d_other=distances[i, other_idx],
        )

        orig = original_types[i].strip()
        is_override = orig.lower() not in ("", suggested.lower()) and orig.lower() in _VALID_TYPES

        results.append(ClassificationResult(
            deal_id=deal_ids[i],
            original_type=orig,
            suggested_type=suggested,
            confidence=round(confidence, 4),
            is_override=is_override,
        ))

        if is_override:
            logger.debug(
                "Override: deal '%s' labelled '%s' but K-Means suggests '%s' "
                "(confidence=%.2f)",
                deal_ids[i], orig, suggested, confidence,
            )

    override_count = sum(r.is_override for r in results)
    prepay_count   = sum(r.suggested_type == _LABEL_PREPAY for r in results)
    ppa_count      = sum(r.suggested_type == _LABEL_PPA for r in results)

    warning = ""
    if override_count > 0:
        warning = (
            f"{override_count} deal(s) have a suggested type different from "
            "their original label. Review them before running the optimiser."
        )

    logger.info(
        "Classification complete: %d Prepay, %d PPA, %d overrides.",
        prepay_count, ppa_count, override_count,
    )

    return ClassificationSummary(
        results=results,
        method_used="kmeans",
        override_count=override_count,
        prepay_count=prepay_count,
        ppa_count=ppa_count,
        warning=warning,
    )


def apply_classification(
    data_deals: list[tuple[str, int]],
    data_deals_type: list[tuple[str, int]],
    summary: ClassificationSummary,
) -> list[tuple[str, int]]:
    """Return a new deals_type list with the classifier's suggested labels applied.

    Parameters
    ----------
    data_deals      : Original deals list from AllocationInput.
    data_deals_type : Original deals_type list from AllocationInput.
    summary         : Output from classify_deals().

    Returns
    -------
    Updated deals_type list (same format as AllocationInput.deals_type).
    """
    _type_map = {"prepay": 0, "ppa": 1}
    suggestion_by_id = {r.deal_id: r.suggested_type for r in summary.results}

    updated = []
    for deal_id, _ in data_deals:
        suggested = suggestion_by_id.get(deal_id, "Prepay")
        updated.append((deal_id, _type_map[suggested.lower()]))

    return updated

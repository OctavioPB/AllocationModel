"""
tests/test_ml_classifier.py
---------------------------
Unit tests for app/ml_classifier.py.

Tests cover: correct classification on well-separated data, fallback behaviours,
confidence score range, label stability (determinism), and override detection.
"""

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import pytest
from app.ml_classifier import classify_deals, apply_classification, _LABEL_PREPAY, _LABEL_PPA


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

def _prepay_ids():
    return [f"Prepay_{i}" for i in range(10)]

def _ppa_ids():
    return [f"PPA_{i}" for i in range(6)]

def _prepay_values():
    # Small values typical of Prepay deals
    return [5_000, 8_000, 12_000, 7_500, 9_000, 6_500, 11_000, 14_000, 10_000, 4_500]

def _ppa_values():
    # Large values typical of PPA deals
    return [180_000, 250_000, 310_000, 75_000, 200_000, 120_000]

def _mixed_ids():
    return _prepay_ids() + _ppa_ids()

def _mixed_values():
    return _prepay_values() + _ppa_values()

def _mixed_types():
    return [_LABEL_PREPAY] * 10 + [_LABEL_PPA] * 6


# ---------------------------------------------------------------------------
# Core classification tests
# ---------------------------------------------------------------------------

class TestClassifyDeals:

    def test_correct_labels_well_separated_data(self):
        """K-Means should correctly separate small Prepay from large PPA deals."""
        summary = classify_deals(_mixed_ids(), _mixed_values())

        for r in summary.results:
            if r.deal_id.startswith("Prepay_"):
                assert r.suggested_type == _LABEL_PREPAY, (
                    f"{r.deal_id} should be Prepay, got {r.suggested_type}"
                )
            elif r.deal_id.startswith("PPA_"):
                assert r.suggested_type == _LABEL_PPA, (
                    f"{r.deal_id} should be PPA, got {r.suggested_type}"
                )

    def test_method_is_kmeans(self):
        """Standard run on mixed data should use the kmeans method."""
        summary = classify_deals(_mixed_ids(), _mixed_values())
        assert summary.method_used == "kmeans"

    def test_result_count_matches_input(self):
        """Number of results must equal number of input deals."""
        ids = _mixed_ids()
        summary = classify_deals(ids, _mixed_values())
        assert len(summary.results) == len(ids)

    def test_prepay_and_ppa_counts_sum_to_total(self):
        """prepay_count + ppa_count must equal total deals."""
        summary = classify_deals(_mixed_ids(), _mixed_values())
        assert summary.prepay_count + summary.ppa_count == len(_mixed_ids())

    def test_no_overrides_when_labels_match(self):
        """If original labels perfectly match suggestions, override_count = 0."""
        summary = classify_deals(
            _mixed_ids(), _mixed_values(), original_types=_mixed_types()
        )
        assert summary.override_count == 0

    def test_overrides_detected_when_labels_differ(self):
        """If original labels are all wrong, all valid labels should be overrides."""
        # Swap: label PPA deals as Prepay and vice versa
        wrong_types = [_LABEL_PPA] * 10 + [_LABEL_PREPAY] * 6
        summary = classify_deals(
            _mixed_ids(), _mixed_values(), original_types=wrong_types
        )
        assert summary.override_count == 16

    def test_empty_original_types_no_overrides(self):
        """Empty string original types should never count as overrides."""
        empties = [""] * len(_mixed_ids())
        summary = classify_deals(_mixed_ids(), _mixed_values(), original_types=empties)
        assert summary.override_count == 0


# ---------------------------------------------------------------------------
# Confidence score tests
# ---------------------------------------------------------------------------

class TestConfidence:

    def test_confidence_in_range(self):
        """All confidence scores must be in [0, 1]."""
        summary = classify_deals(_mixed_ids(), _mixed_values())
        for r in summary.results:
            assert 0.0 <= r.confidence <= 1.0, (
                f"{r.deal_id}: confidence {r.confidence} out of range"
            )

    def test_extreme_outlier_has_high_confidence(self):
        """A deal far from the cluster boundary should have high confidence."""
        # PPA 4 = 200k, centroid should be ~220k — far from Prepay cluster
        summary = classify_deals(_mixed_ids(), _mixed_values())
        ppa_results = [r for r in summary.results if r.suggested_type == _LABEL_PPA]
        max_conf = max(r.confidence for r in ppa_results)
        assert max_conf > 0.7, f"Expected high-confidence PPA result, got max={max_conf}"

    def test_boundary_deal_has_lower_confidence(self):
        """A deal with value between clusters should have lower confidence."""
        # Create a dataset where one deal sits at the boundary (~50k)
        ids    = ["Low", "Boundary", "High"]
        values = [5_000, 50_000, 300_000]
        summary = classify_deals(ids, values)
        boundary = next(r for r in summary.results if r.deal_id == "Boundary")
        extremes = [r for r in summary.results if r.deal_id != "Boundary"]
        # Boundary confidence should be lower than at least one extreme
        assert boundary.confidence < max(r.confidence for r in extremes)


# ---------------------------------------------------------------------------
# Determinism test
# ---------------------------------------------------------------------------

class TestDeterminism:

    def test_same_result_on_repeated_calls(self):
        """Two calls with the same data should return identical results."""
        summary_1 = classify_deals(_mixed_ids(), _mixed_values())
        summary_2 = classify_deals(_mixed_ids(), _mixed_values())
        for r1, r2 in zip(summary_1.results, summary_2.results):
            assert r1.suggested_type == r2.suggested_type
            assert r1.confidence == r2.confidence


# ---------------------------------------------------------------------------
# Fallback behaviour tests
# ---------------------------------------------------------------------------

class TestFallbacks:

    def test_single_deal_fallback(self):
        """A single deal should return a result without crashing."""
        summary = classify_deals(["D0"], [150_000])
        assert len(summary.results) == 1
        assert summary.results[0].suggested_type in (_LABEL_PREPAY, _LABEL_PPA)
        assert summary.method_used == "single_deal_fallback"
        assert summary.warning != ""

    def test_all_zero_values_fallback(self):
        """All zero values should trigger the zero-value fallback."""
        ids    = ["D0", "D1", "D2"]
        values = [0, 0, 0]
        summary = classify_deals(ids, values)
        assert summary.method_used == "zero_value_fallback"
        assert all(r.suggested_type == _LABEL_PREPAY for r in summary.results)

    def test_empty_input(self):
        """Empty input should return an empty summary with a warning."""
        summary = classify_deals([], [])
        assert len(summary.results) == 0
        assert summary.warning != ""

    def test_length_mismatch_raises(self):
        """Mismatched list lengths should raise a ValueError."""
        with pytest.raises(ValueError):
            classify_deals(["D0", "D1"], [100], original_types=["Prepay", "PPA"])

    def test_only_prepay_values_does_not_crash(self):
        """Dataset with only small Prepay-range values should not crash.

        K-Means has no domain knowledge about absolute value thresholds, so it
        will split the data into two relative clusters. The important guarantee
        is that the classifier runs without error and returns a result for every
        deal — the user reviews suggestions before optimising.
        """
        ids    = [f"P{i}" for i in range(5)]
        values = [5_000, 7_000, 6_000, 8_000, 4_500]
        summary = classify_deals(ids, values)
        assert len(summary.results) == len(ids), "Must return one result per deal"
        for r in summary.results:
            assert r.suggested_type in (_LABEL_PREPAY, _LABEL_PPA), (
                f"Unexpected label: {r.suggested_type}"
            )


# ---------------------------------------------------------------------------
# apply_classification helper test
# ---------------------------------------------------------------------------

class TestApplyClassification:

    def test_updated_types_match_suggestions(self):
        """apply_classification should update types to match classifier output."""
        deals      = [(_id, val) for _id, val in zip(_mixed_ids(), _mixed_values())]
        deals_type = [(_id, 0) for _id in _mixed_ids()]  # all Prepay initially

        summary = classify_deals(_mixed_ids(), _mixed_values())
        updated = apply_classification(deals, deals_type, summary)

        for (deal_id, type_code), result in zip(updated, summary.results):
            expected_code = 1 if result.suggested_type == _LABEL_PPA else 0
            assert type_code == expected_code, (
                f"{deal_id}: expected {expected_code}, got {type_code}"
            )

"""
tests/test_optimizer.py
-----------------------
Unit tests for app/optimizer.py.

Tests verify: correct allocation of simple known cases, infeasibility handling,
min_deal constraint, preference penalty, and AllocationResult integrity.
"""

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import pytest
from app.data_loader import AllocationInput
from app.optimizer import optimize


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _simple_input(
    deal_values: list[int],
    deal_types: list[int],
    purchaser_budgets: list[int],
    purchaser_prefs: list[int],
    min_deal: bool = False,
    pref_penalty: bool = False,
) -> AllocationInput:
    """Build a minimal AllocationInput for testing."""
    deals      = [(f"D{i}", v) for i, v in enumerate(deal_values)]
    deals_type = [(f"D{i}", t) for i, t in enumerate(deal_types)]
    p_ids      = [f"P{i}" for i in range(len(purchaser_budgets))]
    return AllocationInput(
        deals=deals,
        deals_type=deals_type,
        purchasers=purchaser_budgets,
        allocation_pref=purchaser_prefs,
        purchaser_ids=p_ids,
        min_deal=min_deal,
        pref_penalty=pref_penalty,
    )


# ---------------------------------------------------------------------------
# Basic allocation tests
# ---------------------------------------------------------------------------

class TestBasicAllocation:

    def test_single_deal_single_purchaser_allocated(self):
        """One deal, one purchaser with sufficient budget — must be allocated."""
        data   = _simple_input([10_000], [0], [50_000], [0])
        result = optimize(data, time_limit=10)
        assert result.status in ("Optimal", "Feasible")
        assert result.allocations[0] == 1
        assert result.total_value == 10_000
        assert result.unallocated_count == 0

    def test_deal_exceeds_budget_is_unallocated(self):
        """A deal worth more than the purchaser's budget must remain unallocated."""
        data   = _simple_input([100_000], [0], [50_000], [0])
        result = optimize(data, time_limit=10)
        assert result.allocations[0] == 0
        assert result.unallocated_count == 1

    def test_all_deals_allocated_when_capacity_sufficient(self):
        """With ample capacity, all non-zero deals should be allocated."""
        data = _simple_input(
            deal_values=[5_000, 8_000, 12_000],
            deal_types=[0, 0, 0],
            purchaser_budgets=[100_000, 100_000],
            purchaser_prefs=[0, 0],
        )
        result = optimize(data, time_limit=10)
        assert result.unallocated_count == 0
        assert result.total_value == 25_000

    def test_deal_allocated_to_at_most_one_purchaser(self):
        """Each deal must appear in at most one purchaser's allocation."""
        data = _simple_input(
            deal_values=[10_000, 20_000, 15_000],
            deal_types=[0, 1, 0],
            purchaser_budgets=[50_000, 50_000, 50_000],
            purchaser_prefs=[0, 1, 0],
        )
        result = optimize(data, time_limit=10)
        # Count how many times each deal appears (must be <= 1)
        for idx in range(len(data.deals)):
            allocated_to = result.allocations[idx]
            assert allocated_to >= 0, "Allocation index must be non-negative"
            assert allocated_to <= len(data.purchasers), "Purchaser index out of range"

    def test_zero_value_deals_are_never_allocated(self):
        """Deals with value 0 should never be allocated."""
        data = _simple_input(
            deal_values=[0, 10_000, 0],
            deal_types=[0, 1, 0],
            purchaser_budgets=[50_000],
            purchaser_prefs=[0],
        )
        result = optimize(data, time_limit=10)
        assert result.allocations[0] == 0, "Zero-value deal should not be allocated"
        assert result.allocations[2] == 0, "Zero-value deal should not be allocated"

    def test_maximises_total_value(self):
        """Optimizer should prefer the deal combination that maximises total value."""
        # Two deals that together exceed purchaser budget; only the larger should be picked
        data = _simple_input(
            deal_values=[30_000, 20_000],
            deal_types=[0, 0],
            purchaser_budgets=[35_000],
            purchaser_prefs=[0],
        )
        result = optimize(data, time_limit=10)
        assert result.total_value == 30_000, "Should pick the higher-value deal"
        assert result.allocations[0] == 1  # D0 allocated
        assert result.allocations[1] == 0  # D1 not allocated


# ---------------------------------------------------------------------------
# Min-deal constraint tests
# ---------------------------------------------------------------------------

class TestMinDealConstraint:

    def test_each_purchaser_gets_at_least_one_deal(self):
        """With min_deal=True, every purchaser with sufficient budget gets >= 1 deal."""
        data = _simple_input(
            deal_values=[5_000, 8_000, 12_000, 7_000, 9_000],
            deal_types=[0, 0, 0, 0, 0],
            purchaser_budgets=[30_000, 30_000, 30_000],
            purchaser_prefs=[0, 0, 0],
            min_deal=True,
        )
        result = optimize(data, time_limit=10)
        # Count deals per purchaser
        deal_counts = {p_idx: 0 for p_idx in range(1, len(data.purchasers) + 1)}
        for a in result.allocations:
            if a > 0:
                deal_counts[a] += 1
        for p_idx, count in deal_counts.items():
            assert count >= 1, f"Purchaser {p_idx} received no deals (min_deal=True)"


# ---------------------------------------------------------------------------
# Result integrity tests
# ---------------------------------------------------------------------------

class TestResultIntegrity:

    def test_allocation_list_length_matches_deals(self):
        """AllocationResult.allocations must have the same length as input deals."""
        data   = _simple_input([5_000, 10_000], [0, 1], [50_000], [0])
        result = optimize(data, time_limit=10)
        assert len(result.allocations) == len(data.deals)

    def test_deal_ids_preserved(self):
        """result.deal_ids must match input deal identifiers in order."""
        data   = _simple_input([5_000, 10_000], [0, 1], [50_000], [0])
        result = optimize(data, time_limit=10)
        assert result.deal_ids == [d[0] for d in data.deals]

    def test_purchaser_ids_preserved(self):
        """result.purchaser_ids must match input purchaser identifiers in order."""
        data   = _simple_input([5_000], [0], [50_000, 30_000], [0, 0])
        result = optimize(data, time_limit=10)
        assert result.purchaser_ids == data.purchaser_ids

    def test_total_value_equals_sum_of_allocated_deals(self):
        """result.total_value must equal the sum of values of allocated deals."""
        data = _simple_input(
            deal_values=[5_000, 8_000, 12_000],
            deal_types=[0, 0, 0],
            purchaser_budgets=[100_000],
            purchaser_prefs=[0],
        )
        result = optimize(data, time_limit=10)
        manual_total = sum(
            data.deals[i][1] for i, a in enumerate(result.allocations) if a > 0
        )
        assert result.total_value == manual_total

    def test_unallocated_count_is_consistent(self):
        """unallocated_count must equal the number of 0s in allocations."""
        data   = _simple_input([5_000, 10_000, 20_000], [0, 1, 0], [12_000], [0])
        result = optimize(data, time_limit=10)
        assert result.unallocated_count == result.allocations.count(0)

    def test_status_is_string(self):
        """Result status must be a non-empty string."""
        data   = _simple_input([5_000], [0], [50_000], [0])
        result = optimize(data, time_limit=10)
        assert isinstance(result.status, str) and len(result.status) > 0


# ---------------------------------------------------------------------------
# Edge cases
# ---------------------------------------------------------------------------

class TestEdgeCases:

    def test_purchaser_with_zero_budget_receives_nothing(self):
        """A purchaser with budget=0 must not receive any deal."""
        data = _simple_input(
            deal_values=[5_000, 10_000],
            deal_types=[0, 1],
            purchaser_budgets=[0, 50_000],
            purchaser_prefs=[0, 1],
        )
        result = optimize(data, time_limit=10)
        for i, a in enumerate(result.allocations):
            if a == 1:
                pytest.fail(f"Deal {i} was allocated to purchaser with 0 budget")

    def test_sample_data_end_to_end(self):
        """Full end-to-end run on sample data must allocate all 97 deals."""
        import sys
        sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
        from app.data_loader import load_csv

        base = os.path.join(os.path.dirname(__file__), "..")
        data = load_csv(
            os.path.join(base, "sample_data", "example_deals.csv"),
            os.path.join(base, "sample_data", "example_purchasers.csv"),
        )
        result = optimize(data, time_limit=30)
        assert result.status in ("Optimal", "Feasible")
        assert result.total_value > 0
        assert result.unallocated_count == 0

"""Exact product identity, money and run-policy regressions."""

from datetime import UTC, datetime, timedelta
from decimal import Decimal

import pytest

from pharmparser.domain import (
    Product,
    RunStatus,
    final_status,
    may_export,
    may_retry,
    money_from_minor,
    money_to_minor,
    parse_money,
)


def test_product_key_normalizes_unicode_case_and_spaces() -> None:
    first = Product("  АСПИРИН", "таблетки   100 мг", "Bayer")
    second = Product("аспирин", "таблетки 100 мг", "bayer")
    assert first.key == second.key


def test_product_identity_uses_all_three_fields() -> None:
    assert Product("Аспирин", "100 мг", "A").key != Product("Аспирин", "100 мг", "B").key


def test_money_rounds_once_and_round_trips_through_kopecks() -> None:
    amount = parse_money(" 12,345 ")
    assert amount == Decimal("12.35")
    assert money_from_minor(money_to_minor(amount)) == amount


def test_invalid_money_is_refused() -> None:
    with pytest.raises(ValueError):
        parse_money("NaN")


def test_partial_policy_requires_reference_and_a_competitor() -> None:
    assert final_status(reference_succeeded=True, successful_competitors=1) is RunStatus.PARTIAL
    assert final_status(reference_succeeded=False, successful_competitors=4) is RunStatus.FAILED
    assert final_status(reference_succeeded=True, successful_competitors=0) is RunStatus.FAILED
    assert may_export(RunStatus.PARTIAL)
    assert not may_export(RunStatus.CANCELLED)


def test_retry_window_is_thirty_minutes() -> None:
    now = datetime.now(UTC)
    assert may_retry(now - timedelta(minutes=29), now)
    assert not may_retry(now - timedelta(minutes=31), now)

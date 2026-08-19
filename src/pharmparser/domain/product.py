"""Structured product identity and exact money handling.

This module is deliberately pure.  Product matching must be reproducible across
scrapes, database rows and reports, so it does not depend on locale or fuzzy
matching libraries.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation

_SPACES = re.compile(r"\s+")
_MONEY_QUANTUM = Decimal("0.01")


def normalize_product_part(value: str) -> str:
    """Return the canonical representation used by the strict product key."""
    return _SPACES.sub(" ", unicodedata.normalize("NFKC", value).strip()).casefold()


@dataclass(frozen=True, slots=True)
class Product:
    name: str
    form: str = ""
    manufacturer: str = ""

    @property
    def key(self) -> str:
        return "\x1f".join(
            normalize_product_part(part) for part in (self.name, self.form, self.manufacturer)
        )

    @property
    def label(self) -> str:
        return ", ".join(part for part in (self.name, self.form, self.manufacturer) if part)


@dataclass(frozen=True, slots=True)
class ProductPrice:
    product: Product
    amount: Decimal

    def __post_init__(self) -> None:
        if not self.amount.is_finite() or self.amount < 0:
            raise ValueError("price must be a finite non-negative amount")


def parse_money(value: str | int | float | Decimal) -> Decimal:
    """Parse a BYN value without carrying binary floating-point into storage."""
    text = str(value).strip().replace("\xa0", "").replace(" ", "").replace(",", ".")
    try:
        amount = Decimal(text).quantize(_MONEY_QUANTUM, rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError) as error:
        raise ValueError(f"invalid money amount: {value!r}") from error
    if not amount.is_finite() or amount < 0:
        raise ValueError(f"invalid money amount: {value!r}")
    return amount


def money_to_minor(amount: Decimal) -> int:
    """Convert Decimal BYN to integer kopecks for SQLite."""
    return int((parse_money(amount) * 100).to_integral_exact())


def money_from_minor(value: int) -> Decimal:
    return (Decimal(value) / 100).quantize(_MONEY_QUANTUM)

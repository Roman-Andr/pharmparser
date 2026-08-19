"""PharmParser — price comparison across tabletka.by pharmacies."""

from __future__ import annotations

from importlib.metadata import PackageNotFoundError, version

try:
    __version__ = version("pharmparser")
except PackageNotFoundError:  # pragma: no cover - only when running from a bare tree
    __version__ = "0.0.0+unknown"

__all__ = ["__version__"]

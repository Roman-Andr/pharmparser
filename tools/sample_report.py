"""Build a macro-enabled report from synthetic data — no network, no config, no Excel.

CI runs this to prove the real deliverable is producible on Linux and uploads the
result, so the buttons can be clicked by anyone with Excel to hand.
"""

from __future__ import annotations

import argparse
import random
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from pharmparser.config import ExportSettings
from pharmparser.domain.models import Pharmacy, PriceTable
from pharmparser.export import export_with_macros

PRODUCTS = [
    "Аспирин, таб. 500 мг",
    "Парацетамол, таб. 500 мг",
    "Ибупрофен, капс. 200 мг",
    "Цитрамон, таб.",
    "Анальгин, таб. 500 мг",
    "Но-шпа, таб. 40 мг",
    "Активированный уголь, таб.",
    "Амоксициллин, капс. 500 мг",
]


def sample_table(pharmacies: int, seed: int) -> PriceTable:
    rng = random.Random(seed)
    entries = []
    for index in range(pharmacies):
        # Vary each range so missing items and gaps get exercised.
        stocked = PRODUCTS if index == 0 else rng.sample(PRODUCTS, k=max(2, len(PRODUCTS) - index))
        prices = {product: round(rng.uniform(1.5, 40.0), 2) for product in stocked}
        entries.append((Pharmacy(id=str(index + 1), name=f"Аптека {index + 1}"), prices))
    return PriceTable.build(entries)


def verify(path: Path) -> None:
    parts = zipfile.ZipFile(path).namelist()
    assert "xl/vbaProject.bin" in parts, "the workbook carries no VBA project"
    assert any("vmlDrawing" in part for part in parts), "no buttons were drawn"
    content_types = zipfile.ZipFile(path).read("[Content_Types].xml").decode()
    assert "macroEnabled.main+xml" in content_types, "not marked macro-enabled"


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=Path("build"))
    parser.add_argument("--pharmacies", type=int, default=4)
    parser.add_argument("--seed", type=int, default=0)
    args = parser.parse_args()

    table = sample_table(args.pharmacies, args.seed)
    path = export_with_macros(ExportSettings(), table, args.output / "sample-report.xlsm")
    verify(path)
    print(f"Wrote {path} ({path.stat().st_size} bytes); VBA project and buttons verified.")


if __name__ == "__main__":
    main()

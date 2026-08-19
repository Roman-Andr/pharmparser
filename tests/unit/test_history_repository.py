from decimal import Decimal
from pathlib import Path

import pytest

from pharmparser.application.history import HistoryRepository, ProductCollisionError
from pharmparser.application.models import ProfileRecord
from pharmparser.domain import Product, RunStatus


def profile() -> ProfileRecord:
    return ProfileRecord.model_validate(
        {
            "name": "Минск",
            "reference_pharmacy_id": "1",
            "pharmacies": [
                {"id": "1", "name": "Наша", "url": "https://tabletka.by/pharmacies/1"},
                {"id": "2", "name": "Конкурент", "url": "https://tabletka.by/pharmacies/2"},
            ],
        }
    )


def test_database_enables_wal_foreign_keys_and_schema_version(tmp_path: Path) -> None:
    repository = HistoryRepository(tmp_path / "history.sqlite3")
    with repository.connect() as connection:
        assert connection.execute("PRAGMA journal_mode").fetchone()[0] == "wal"
        assert connection.execute("PRAGMA foreign_keys").fetchone()[0] == 1
        assert connection.execute("PRAGMA user_version").fetchone()[0] == 1


def test_prices_are_stored_as_integer_kopecks(tmp_path: Path) -> None:
    repository = HistoryRepository(tmp_path / "history.sqlite3")
    run_id = repository.create_run(profile())
    repository.start_attempt(run_id, "1", "Наша")
    repository.store_prices(run_id, "1", [(Product("Аспирин", "100 мг", "Bayer"), Decimal("5.17"))])
    with repository.connect() as connection:
        assert connection.execute("SELECT amount_minor FROM prices").fetchone()[0] == 517


def test_normalized_collision_never_overwrites_a_price(tmp_path: Path) -> None:
    repository = HistoryRepository(tmp_path / "history.sqlite3")
    run_id = repository.create_run(profile())
    products = [
        (Product("Аспирин", "100 мг", "Bayer"), Decimal("5")),
        (Product(" АСПИРИН ", "100  мг", "bayer"), Decimal("9")),
    ]
    with pytest.raises(ProductCollisionError):
        repository.store_prices(run_id, "1", products)
    with repository.connect() as connection:
        assert connection.execute("SELECT COUNT(*) FROM prices").fetchone()[0] == 0


def test_retention_keeps_pinned_runs(tmp_path: Path) -> None:
    repository = HistoryRepository(tmp_path / "history.sqlite3")
    selected_profile = profile()
    made = [repository.create_run(selected_profile) for _ in range(12)]
    for run_id in made:
        repository.set_status(run_id, RunStatus.FAILED)
    repository.pin(made[0], True)
    assert repository.enforce_retention(profile().id, 10) == 0  # another profile UUID has no rows
    assert repository.enforce_retention(selected_profile.id, 10) == 1
    assert len(repository.list_runs(selected_profile.id)) == 11
    assert repository.get_run(made[0]).pinned

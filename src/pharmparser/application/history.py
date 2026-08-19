"""Transactional SQLite history for runs, attempts, products and artifacts."""

from __future__ import annotations

import sqlite3
from collections.abc import Iterator, Mapping, Sequence
from contextlib import contextmanager
from datetime import UTC, datetime
from decimal import Decimal
from pathlib import Path
from uuid import UUID, uuid4

from .. import __version__
from ..config.paths import history_path
from ..domain import Product, RunStatus, money_from_minor, money_to_minor
from .models import ProfileRecord, RunSummary

SCHEMA_VERSION = 1

_SCHEMA = """
CREATE TABLE IF NOT EXISTS runs (
    id TEXT PRIMARY KEY,
    profile_id TEXT NOT NULL,
    parent_run_id TEXT REFERENCES runs(id),
    status TEXT NOT NULL CHECK(status IN ('queued','running','partial','completed','failed','cancelled')),
    started_at TEXT NOT NULL,
    finished_at TEXT,
    reference_pharmacy_id TEXT NOT NULL,
    profile_snapshot TEXT NOT NULL,
    parser_version TEXT NOT NULL,
    report_version TEXT NOT NULL,
    pinned INTEGER NOT NULL DEFAULT 0 CHECK(pinned IN (0,1))
);
CREATE TABLE IF NOT EXISTS pharmacy_attempts (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    run_id TEXT NOT NULL REFERENCES runs(id) ON DELETE CASCADE,
    pharmacy_id TEXT NOT NULL,
    pharmacy_name TEXT NOT NULL,
    status TEXT NOT NULL,
    started_at TEXT NOT NULL,
    finished_at TEXT,
    pages INTEGER NOT NULL DEFAULT 0,
    items INTEGER NOT NULL DEFAULT 0,
    error_code TEXT,
    error_message TEXT,
    UNIQUE(run_id, pharmacy_id)
);
CREATE TABLE IF NOT EXISTS products (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    normalized_key TEXT NOT NULL UNIQUE,
    name TEXT NOT NULL,
    form TEXT NOT NULL,
    manufacturer TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS prices (
    run_id TEXT NOT NULL REFERENCES runs(id) ON DELETE CASCADE,
    pharmacy_id TEXT NOT NULL,
    product_id INTEGER NOT NULL REFERENCES products(id),
    amount_minor INTEGER NOT NULL CHECK(amount_minor >= 0),
    PRIMARY KEY(run_id, pharmacy_id, product_id)
);
CREATE TABLE IF NOT EXISTS warnings (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    run_id TEXT NOT NULL REFERENCES runs(id) ON DELETE CASCADE,
    pharmacy_id TEXT,
    code TEXT NOT NULL,
    message TEXT NOT NULL,
    created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS report_artifacts (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    run_id TEXT NOT NULL REFERENCES runs(id) ON DELETE CASCADE,
    format TEXT NOT NULL CHECK(format IN ('xlsm','xlsx')),
    path TEXT NOT NULL,
    created_at TEXT NOT NULL,
    UNIQUE(run_id, format, path)
);
CREATE INDEX IF NOT EXISTS idx_runs_profile_started ON runs(profile_id, started_at DESC);
CREATE INDEX IF NOT EXISTS idx_attempts_run ON pharmacy_attempts(run_id);
CREATE INDEX IF NOT EXISTS idx_prices_run_pharmacy ON prices(run_id, pharmacy_id);
CREATE INDEX IF NOT EXISTS idx_prices_product ON prices(product_id);
"""


def _now() -> str:
    return datetime.now(UTC).isoformat()


class ProductCollisionError(ValueError):
    """Two display variants normalized to one strict key in a single pharmacy."""


class HistoryRepository:
    def __init__(self, path: Path | None = None) -> None:
        self.path = path or history_path()
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.migrate()

    def connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.path, timeout=30, isolation_level=None)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
        connection.execute("PRAGMA journal_mode = WAL")
        connection.execute("PRAGMA busy_timeout = 30000")
        return connection

    @contextmanager
    def transaction(self) -> Iterator[sqlite3.Connection]:
        connection = self.connect()
        try:
            connection.execute("BEGIN IMMEDIATE")
            yield connection
            connection.commit()
        except BaseException:
            connection.rollback()
            raise
        finally:
            connection.close()

    def migrate(self) -> None:
        with self.transaction() as connection:
            version = int(connection.execute("PRAGMA user_version").fetchone()[0])
            if version > SCHEMA_VERSION:
                raise RuntimeError(f"database schema {version} is newer than supported {SCHEMA_VERSION}")
            if version < 1:
                connection.executescript(_SCHEMA)
                connection.execute(f"PRAGMA user_version = {SCHEMA_VERSION}")

    def create_run(self, profile: ProfileRecord, *, parent_run_id: UUID | None = None) -> UUID:
        if profile.reference_pharmacy_id is None:
            raise ValueError("в профиле не выбрана основная аптека")
        run_id = uuid4()
        with self.transaction() as connection:
            connection.execute(
                """INSERT INTO runs
                   (id, profile_id, parent_run_id, status, started_at, reference_pharmacy_id,
                    profile_snapshot, parser_version, report_version)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (
                    str(run_id),
                    str(profile.id),
                    str(parent_run_id) if parent_run_id else None,
                    RunStatus.QUEUED,
                    _now(),
                    profile.reference_pharmacy_id,
                    profile.model_dump_json(),
                    __version__,
                    __version__,
                ),
            )
        return run_id

    def set_status(self, run_id: UUID, status: RunStatus) -> None:
        finished_at = _now() if status.terminal else None
        with self.transaction() as connection:
            result = connection.execute(
                "UPDATE runs SET status = ?, finished_at = COALESCE(?, finished_at) WHERE id = ?",
                (status, finished_at, str(run_id)),
            )
            if result.rowcount != 1:
                raise KeyError(str(run_id))

    def start_attempt(self, run_id: UUID, pharmacy_id: str, pharmacy_name: str) -> None:
        with self.transaction() as connection:
            connection.execute(
                """INSERT INTO pharmacy_attempts
                   (run_id, pharmacy_id, pharmacy_name, status, started_at)
                   VALUES (?, ?, ?, 'running', ?)
                   ON CONFLICT(run_id, pharmacy_id) DO UPDATE SET
                     status='running', started_at=excluded.started_at, finished_at=NULL,
                     error_code=NULL, error_message=NULL""",
                (str(run_id), pharmacy_id, pharmacy_name, _now()),
            )

    def finish_attempt(
        self,
        run_id: UUID,
        pharmacy_id: str,
        *,
        status: str,
        pages: int = 0,
        items: int = 0,
        error_code: str | None = None,
        error_message: str | None = None,
    ) -> None:
        with self.transaction() as connection:
            connection.execute(
                """UPDATE pharmacy_attempts SET status=?, finished_at=?, pages=?, items=?,
                   error_code=?, error_message=? WHERE run_id=? AND pharmacy_id=?""",
                (status, _now(), pages, items, error_code, error_message, str(run_id), pharmacy_id),
            )

    def store_prices(
        self,
        run_id: UUID,
        pharmacy_id: str,
        prices: Sequence[tuple[Product, Decimal]],
    ) -> None:
        seen: dict[str, Product] = {}
        for product, _ in prices:
            if (existing := seen.get(product.key)) and existing != product:
                raise ProductCollisionError(
                    f"коллизия нормализованного ключа: {existing.label!r} и {product.label!r}"
                )
            seen[product.key] = product

        with self.transaction() as connection:
            for product, amount in prices:
                connection.execute(
                    """INSERT INTO products(normalized_key, name, form, manufacturer) VALUES (?, ?, ?, ?)
                       ON CONFLICT(normalized_key) DO NOTHING""",
                    (product.key, product.name, product.form, product.manufacturer),
                )
                product_row = connection.execute(
                    "SELECT id, name, form, manufacturer FROM products WHERE normalized_key = ?",
                    (product.key,),
                ).fetchone()
                assert product_row is not None
                stored = Product(product_row["name"], product_row["form"], product_row["manufacturer"])
                if stored != product:
                    raise ProductCollisionError(
                        f"коллизия нормализованного ключа: {stored.label!r} и {product.label!r}"
                    )
                connection.execute(
                    """INSERT INTO prices(run_id, pharmacy_id, product_id, amount_minor)
                       VALUES (?, ?, ?, ?)
                       ON CONFLICT(run_id, pharmacy_id, product_id) DO UPDATE SET amount_minor=excluded.amount_minor""",
                    (str(run_id), pharmacy_id, product_row["id"], money_to_minor(amount)),
                )

    def add_warning(self, run_id: UUID, code: str, message: str, pharmacy_id: str | None = None) -> None:
        with self.transaction() as connection:
            connection.execute(
                "INSERT INTO warnings(run_id, pharmacy_id, code, message, created_at) VALUES (?, ?, ?, ?, ?)",
                (str(run_id), pharmacy_id, code, message, _now()),
            )

    def add_artifact(self, run_id: UUID, format_: str, path: Path) -> None:
        with self.transaction() as connection:
            connection.execute(
                "INSERT OR IGNORE INTO report_artifacts(run_id, format, path, created_at) VALUES (?, ?, ?, ?)",
                (str(run_id), format_, str(path.absolute()), _now()),
            )

    def copy_pharmacy_result(self, source_run_id: UUID, target_run_id: UUID, pharmacy_id: str) -> None:
        """Reuse a successful result during a retry without touching the network."""
        with self.transaction() as connection:
            attempt = connection.execute(
                """SELECT pharmacy_name, pages, items FROM pharmacy_attempts
                   WHERE run_id=? AND pharmacy_id=? AND status='completed'""",
                (str(source_run_id), pharmacy_id),
            ).fetchone()
            if attempt is None:
                raise KeyError(pharmacy_id)
            connection.execute(
                """INSERT INTO pharmacy_attempts
                   (run_id, pharmacy_id, pharmacy_name, status, started_at, finished_at, pages, items)
                   VALUES (?, ?, ?, 'completed', ?, ?, ?, ?)""",
                (
                    str(target_run_id), pharmacy_id, attempt["pharmacy_name"], _now(), _now(),
                    attempt["pages"], attempt["items"],
                ),
            )
            connection.execute(
                """INSERT INTO prices(run_id, pharmacy_id, product_id, amount_minor)
                   SELECT ?, pharmacy_id, product_id, amount_minor FROM prices
                   WHERE run_id=? AND pharmacy_id=?""",
                (str(target_run_id), str(source_run_id), pharmacy_id),
            )

    def list_runs(self, profile_id: UUID | None = None) -> list[RunSummary]:
        where = "WHERE r.profile_id = ?" if profile_id else ""
        params = (str(profile_id),) if profile_id else ()
        query = f"""SELECT r.*,
                    COUNT(DISTINCT a.pharmacy_id) AS pharmacy_count,
                    COUNT(DISTINCT CASE WHEN a.status='completed' THEN a.pharmacy_id END) AS successful_pharmacies,
                    COUNT(DISTINCT p.product_id) AS product_count,
                    COUNT(DISTINCT w.id) AS warning_count,
                    (SELECT path FROM report_artifacts x WHERE x.run_id=r.id ORDER BY x.id DESC LIMIT 1) report_path
                  FROM runs r
                  LEFT JOIN pharmacy_attempts a ON a.run_id=r.id
                  LEFT JOIN prices p ON p.run_id=r.id
                  LEFT JOIN warnings w ON w.run_id=r.id
                  {where} GROUP BY r.id ORDER BY r.started_at DESC"""
        with self.connect() as connection:
            rows = connection.execute(query, params).fetchall()
        return [self._summary(row) for row in rows]

    def get_run(self, run_id: UUID) -> RunSummary:
        runs = [run for run in self.list_runs() if run.id == run_id]
        if not runs:
            raise KeyError(str(run_id))
        return runs[0]

    @staticmethod
    def _summary(row: sqlite3.Row) -> RunSummary:
        return RunSummary(
            id=UUID(row["id"]),
            profile_id=UUID(row["profile_id"]),
            parent_run_id=UUID(row["parent_run_id"]) if row["parent_run_id"] else None,
            status=RunStatus(row["status"]),
            started_at=datetime.fromisoformat(row["started_at"]),
            finished_at=datetime.fromisoformat(row["finished_at"]) if row["finished_at"] else None,
            reference_pharmacy_id=row["reference_pharmacy_id"],
            pharmacy_count=row["pharmacy_count"],
            successful_pharmacies=row["successful_pharmacies"],
            product_count=row["product_count"],
            pinned=bool(row["pinned"]),
            report_path=row["report_path"],
            warning_count=row["warning_count"],
        )

    def profile_snapshot(self, run_id: UUID) -> ProfileRecord:
        with self.connect() as connection:
            row = connection.execute("SELECT profile_snapshot FROM runs WHERE id=?", (str(run_id),)).fetchone()
        if row is None:
            raise KeyError(str(run_id))
        return ProfileRecord.model_validate_json(row["profile_snapshot"])

    def prices_for_run(self, run_id: UUID) -> dict[str, dict[Product, Decimal]]:
        with self.connect() as connection:
            rows = connection.execute(
                """SELECT p.pharmacy_id, p.amount_minor, d.name, d.form, d.manufacturer
                   FROM prices p JOIN products d ON d.id=p.product_id WHERE p.run_id=?""",
                (str(run_id),),
            ).fetchall()
        result: dict[str, dict[Product, Decimal]] = {}
        for row in rows:
            product = Product(row["name"], row["form"], row["manufacturer"])
            result.setdefault(row["pharmacy_id"], {})[product] = money_from_minor(row["amount_minor"])
        return result

    def attempts_for_run(self, run_id: UUID) -> list[dict[str, object]]:
        with self.connect() as connection:
            rows = connection.execute(
                "SELECT * FROM pharmacy_attempts WHERE run_id=? ORDER BY id", (str(run_id),)
            ).fetchall()
        return [dict(row) for row in rows]

    def warnings_for_run(self, run_id: UUID) -> list[dict[str, object]]:
        with self.connect() as connection:
            rows = connection.execute("SELECT * FROM warnings WHERE run_id=? ORDER BY id", (str(run_id),)).fetchall()
        return [dict(row) for row in rows]

    def pin(self, run_id: UUID, pinned: bool) -> None:
        with self.transaction() as connection:
            connection.execute("UPDATE runs SET pinned=? WHERE id=?", (int(pinned), str(run_id)))

    def delete(self, run_id: UUID) -> None:
        with self.transaction() as connection:
            connection.execute("DELETE FROM runs WHERE id=?", (str(run_id),))

    def previous_completed(self, run_id: UUID) -> UUID | None:
        current = self.get_run(run_id)
        with self.connect() as connection:
            row = connection.execute(
                """SELECT id FROM runs WHERE profile_id=? AND reference_pharmacy_id=?
                   AND status='completed' AND started_at < ? ORDER BY started_at DESC LIMIT 1""",
                (str(current.profile_id), current.reference_pharmacy_id, current.started_at.isoformat()),
            ).fetchone()
        return UUID(row["id"]) if row else None

    def enforce_retention(self, profile_id: UUID, limit: int | None) -> int:
        if limit is None:
            return 0
        if not 10 <= limit <= 500:
            raise ValueError("retention must be between 10 and 500, or unlimited")
        with self.transaction() as connection:
            rows = connection.execute(
                """SELECT id FROM runs WHERE profile_id=? AND pinned=0
                   ORDER BY started_at DESC LIMIT -1 OFFSET ?""",
                (str(profile_id), limit),
            ).fetchall()
            connection.executemany("DELETE FROM runs WHERE id=?", [(row["id"],) for row in rows])
        return len(rows)

    def size_bytes(self) -> int:
        return self.path.stat().st_size if self.path.exists() else 0

    def diagnostics(self) -> Mapping[str, object]:
        with self.connect() as connection:
            counts = {
                table: connection.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]
                for table in ("runs", "pharmacy_attempts", "products", "prices", "warnings")
            }
            version = connection.execute("PRAGMA user_version").fetchone()[0]
        return {"schema_version": version, "size_bytes": self.size_bytes(), "counts": counts}

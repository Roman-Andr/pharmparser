"""Cancellable scrape jobs, partial-result policy and replayable progress events."""

from __future__ import annotations

import asyncio
from collections import defaultdict
from collections.abc import AsyncIterator, Awaitable, Callable, Sequence
from contextlib import suppress
from datetime import UTC, datetime
from uuid import UUID

from ..config import PharmacyEntry, RequestConfig
from ..domain import ProductPrice, RunStatus, final_status, may_retry
from ..scraping import ClientSessionFactory
from ..scraping.parallel import parse_pool
from .history import HistoryRepository, ProductCollisionError
from .models import PharmacyProfileEntry, ProfileRecord, ProgressEvent, RunSummary

FetchFn = Callable[[PharmacyProfileEntry], Awaitable[Sequence[ProductPrice]]]


class ActiveRunError(RuntimeError):
    pass


class RunService:
    def __init__(
        self,
        history: HistoryRepository,
        request_factory: Callable[[], RequestConfig] | None = None,
        *,
        fetch: FetchFn | None = None,
        concurrency: int = 8,
    ) -> None:
        self.history = history
        self.request_factory = request_factory
        self._fetch = fetch
        self.concurrency = concurrency
        self._active_id: UUID | None = None
        self._task: asyncio.Task[None] | None = None
        self._events: dict[UUID, list[ProgressEvent]] = defaultdict(list)
        self._conditions: dict[UUID, asyncio.Condition] = defaultdict(asyncio.Condition)
        self._sequence: dict[UUID, int] = defaultdict(int)

    @property
    def active_run_id(self) -> UUID | None:
        return self._active_id

    async def start(self, profile: ProfileRecord) -> RunSummary:
        if self._task is not None and not self._task.done():
            raise ActiveRunError("уже выполняется другой запуск")
        if profile.reference_pharmacy_id is None:
            raise ValueError("выберите основную аптеку")
        if len(profile.pharmacies) < 2:
            raise ValueError("для сравнения нужны основная аптека и хотя бы один конкурент")
        run_id = self.history.create_run(profile)
        self._active_id = run_id
        self._task = asyncio.create_task(self._execute(run_id, profile))
        return self.history.get_run(run_id)

    async def wait(self, run_id: UUID) -> RunSummary:
        if self._active_id == run_id and self._task is not None:
            with suppress(asyncio.CancelledError):
                await self._task
        return self.history.get_run(run_id)

    async def cancel(self, run_id: UUID) -> RunSummary:
        if self._active_id != run_id or self._task is None or self._task.done():
            return self.history.get_run(run_id)
        self._task.cancel()
        with suppress(asyncio.CancelledError):
            await self._task
        return self.history.get_run(run_id)

    async def retry(self, run_id: UUID) -> RunSummary:
        original = self.history.get_run(run_id)
        if original.status is not RunStatus.PARTIAL or original.finished_at is None:
            raise ValueError("повтор доступен только для неполного запуска")
        if not may_retry(original.finished_at):
            raise ValueError("30-минутное окно повтора истекло; создайте новый запуск")
        if self._task is not None and not self._task.done():
            raise ActiveRunError("уже выполняется другой запуск")
        profile = self.history.profile_snapshot(run_id)
        retry_id = self.history.create_run(profile, parent_run_id=run_id)
        successful = {
            str(attempt["pharmacy_id"])
            for attempt in self.history.attempts_for_run(run_id)
            if attempt["status"] == "completed"
        }
        for pharmacy_id in successful:
            self.history.copy_pharmacy_result(run_id, retry_id, pharmacy_id)
        failed = [entry for entry in profile.pharmacies if entry.id not in successful]
        self._active_id = retry_id
        self._task = asyncio.create_task(self._execute(retry_id, profile, entries=failed, reused=successful))
        return self.history.get_run(retry_id)

    async def _execute(
        self,
        run_id: UUID,
        profile: ProfileRecord,
        *,
        entries: Sequence[PharmacyProfileEntry] | None = None,
        reused: set[str] | None = None,
    ) -> None:
        self.history.set_status(run_id, RunStatus.RUNNING)
        await self._emit(run_id, "run", "Проверка", "Запуск начат")
        selected = list(entries or profile.pharmacies)
        reused = reused or set()
        semaphore = asyncio.Semaphore(max(1, self.concurrency))

        async def scrape(entry: PharmacyProfileEntry, fetch: FetchFn) -> tuple[str, bool]:
            self.history.start_attempt(run_id, entry.id, entry.name)
            await self._emit(run_id, "pharmacy", "Загрузка", f"Загрузка: {entry.name}", entry.id)
            try:
                async with semaphore:
                    prices = await fetch(entry)
                self.history.store_prices(
                    run_id,
                    entry.id,
                    [(item.product, item.amount) for item in prices],
                )
                self.history.finish_attempt(run_id, entry.id, status="completed", items=len(prices))
                await self._emit(
                    run_id, "pharmacy", "Сохранение", f"Готово: {entry.name}", entry.id,
                    current=len(prices), total=len(prices),
                )
                return entry.id, True
            except asyncio.CancelledError:
                self.history.finish_attempt(run_id, entry.id, status="cancelled")
                raise
            except ProductCollisionError as error:
                self.history.finish_attempt(
                    run_id, entry.id, status="failed", error_code="product_collision", error_message=str(error)
                )
                self.history.add_warning(run_id, "product_collision", str(error), entry.id)
            except Exception as error:
                self.history.finish_attempt(
                    run_id, entry.id, status="failed", error_code=type(error).__name__, error_message=str(error)
                )
                self.history.add_warning(run_id, "pharmacy_failed", f"{entry.name}: {error}", entry.id)
            await self._emit(run_id, "pharmacy", "Ошибка", f"Не удалось загрузить: {entry.name}", entry.id)
            return entry.id, False

        try:
            if self._fetch is not None:
                fetch = self._fetch
                results = await asyncio.gather(*(scrape(entry, fetch) for entry in selected))
            else:
                if self.request_factory is None:
                    raise RuntimeError("источник цен не настроен")
                with parse_pool(len(selected)) as pool:
                    async with ClientSessionFactory(self.request_factory(), parse_pool=pool) as client:
                        async def fetch(entry: PharmacyProfileEntry) -> Sequence[ProductPrice]:
                            legacy = PharmacyEntry(name=entry.name, url=str(entry.url))
                            return await client.product_prices_for(legacy)

                        results = await asyncio.gather(*(scrape(entry, fetch) for entry in selected))

            succeeded = reused | {pharmacy_id for pharmacy_id, ok in results if ok}
            reference_ok = profile.reference_pharmacy_id in succeeded
            competitor_ok = len(succeeded - {profile.reference_pharmacy_id})
            status = (
                RunStatus.COMPLETED
                if reference_ok and competitor_ok == len(profile.pharmacies) - 1
                else final_status(reference_succeeded=reference_ok, successful_competitors=competitor_ok)
            )
            self.history.set_status(run_id, status)
            await self._emit(run_id, "run", "Завершено", self._status_message(status))
        except asyncio.CancelledError:
            self.history.set_status(run_id, RunStatus.CANCELLED)
            await self._emit(run_id, "run", "Отменено", "Запуск отменен")
            raise
        except Exception as error:
            self.history.add_warning(run_id, "run_failed", str(error))
            self.history.set_status(run_id, RunStatus.FAILED)
            await self._emit(run_id, "run", "Ошибка", f"Запуск завершился с ошибкой: {error}")
        finally:
            if self._active_id == run_id:
                self._active_id = None

    @staticmethod
    def _status_message(status: RunStatus) -> str:
        if status is RunStatus.COMPLETED:
            return "Все аптеки загружены"
        if status is RunStatus.PARTIAL:
            return "Создан неполный результат; некоторые аптеки пропущены"
        return "Отчет не создан: основная аптека или все конкуренты недоступны"

    async def _emit(
        self,
        run_id: UUID,
        kind: str,
        stage: str,
        message: str,
        pharmacy_id: str | None = None,
        current: int | None = None,
        total: int | None = None,
    ) -> None:
        self._sequence[run_id] += 1
        event = ProgressEvent(
            sequence=self._sequence[run_id], run_id=run_id, kind=kind, pharmacy_id=pharmacy_id,
            stage=stage, message=message, current=current, total=total, timestamp=datetime.now(UTC),
        )
        self._events[run_id].append(event)
        async with self._conditions[run_id]:
            self._conditions[run_id].notify_all()

    async def events(self, run_id: UUID, last_event_id: int = 0) -> AsyncIterator[ProgressEvent]:
        index = 0
        while index < len(self._events[run_id]) and self._events[run_id][index].sequence <= last_event_id:
            index += 1
        while True:
            while index < len(self._events[run_id]):
                event = self._events[run_id][index]
                index += 1
                yield event
            if self.history.get_run(run_id).status.terminal:
                break
            async with self._conditions[run_id]:
                await self._conditions[run_id].wait()

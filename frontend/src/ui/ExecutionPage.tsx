import { useEffect, useState } from "react";
import { useNavigate, useParams } from "react-router-dom";
import { api, streamEvents } from "../api";
import type { ProgressEvent, Run } from "../types";

export function ExecutionPage() {
  const { runId = "" } = useParams(); const navigate = useNavigate(); const [events, setEvents] = useState<ProgressEvent[]>([]); const [done, setDone] = useState(false);
  useEffect(() => { const controller = new AbortController(); streamEvents(runId, (value) => { const event = value as ProgressEvent; setEvents((old) => [...old, event]); if (["Завершено", "Ошибка", "Отменено"].includes(event.stage)) setDone(true); }, controller.signal).catch((error) => { if (!controller.signal.aborted) console.error(error); }); return () => controller.abort(); }, [runId]);
  const pharmacies = new Map(events.filter((event) => event.pharmacy_id).map((event) => [event.pharmacy_id!, event]));
  return <section><header className="page-title"><div><p className="eyebrow">ЗАПУСК</p><h1>{done ? "Выполнение завершено" : "Формируем отчет"}</h1><p>Каждая аптека загружается независимо. Окно можно оставить в фоне.</p></div></header>
    <ol className="steps"><li className="active">Проверка</li><li className={events.length ? "active" : ""}>Загрузка</li><li className={done ? "active" : ""}>Сохранение</li><li className={done ? "active" : ""}>Экспорт</li></ol>
    <div className="pharmacy-grid">{[...pharmacies.values()].map((event) => <article className="card" key={event.pharmacy_id}><span className={`status ${event.stage === "Ошибка" ? "bad" : ""}`}>{event.stage}</span><h3>{event.message.replace(/^[^:]+:\s*/, "")}</h3><p>{event.current != null ? `${event.current} позиций` : "Ожидание данных"}</p></article>)}</div>
    <div className="actions">{!done && <button className="danger" onClick={() => api<Run>(`/runs/${runId}/cancel`, { method: "POST" })}>Отменить</button>}{done && <><button className="primary" onClick={() => navigate("/history")}>Открыть готовый отчет</button><button onClick={() => navigate("/")}>Сформировать снова</button></>}</div>
  </section>;
}

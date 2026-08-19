import { useQueryClient } from "@tanstack/react-query";
import { useNavigate } from "react-router-dom";
import { api } from "../api";
import { useHistory } from "./hooks";

const labels: Record<string, string> = { completed: "Полный", partial: "Неполный", failed: "Ошибка", cancelled: "Отменен", running: "Выполняется", queued: "В очереди" };
export function HistoryPage() {
  const history = useHistory(); const client = useQueryClient(); const navigate = useNavigate();
  const refresh = () => client.invalidateQueries({ queryKey: ["history"] });
  return <section><header className="page-title"><div><p className="eyebrow">ЛОКАЛЬНЫЕ ДАННЫЕ</p><h1>История</h1><p>Повторный экспорт не обращается к сети.</p></div></header>
    <div className="table-wrap"><table><thead><tr><th>Дата</th><th>Статус</th><th>Аптеки</th><th>Товары</th><th>Отчет</th><th><span className="sr-only">Действия</span></th></tr></thead><tbody>
      {history.data?.map((run) => <tr key={run.id}><td>{new Date(run.started_at).toLocaleString("ru-RU")}</td><td><span className={`status ${run.status}`}>{labels[run.status]}</span></td><td>{run.successful_pharmacies} / {run.pharmacy_count}</td><td>{run.product_count}</td><td className="path">{run.report_path || "Не экспортирован"}</td><td className="row-actions">
        {(run.status === "completed" || run.status === "partial") && <button onClick={() => api(`/history/${run.id}/export`, { method: "POST", body: "{}" }).then(refresh)}>Экспорт</button>}
        {run.status === "partial" && <button onClick={() => api<{id: string}>(`/runs/${run.id}/retry`, { method: "POST" }).then((retry) => navigate(`/run/${retry.id}`))}>Повторить ошибки</button>}
        {run.report_path && <><button onClick={() => api("/system/open-report", { method: "POST", body: JSON.stringify({ path: run.report_path }) })}>Открыть</button><button onClick={() => api("/system/open-folder", { method: "POST", body: JSON.stringify({ path: run.report_path }) })}>Папка</button></>}
        <button aria-label={run.pinned ? "Открепить" : "Закрепить"} onClick={() => api(`/history/${run.id}/pin?pinned=${!run.pinned}`, { method: "POST" }).then(refresh)}>{run.pinned ? "★" : "☆"}</button>
      </td></tr>)}
    </tbody></table></div></section>;
}

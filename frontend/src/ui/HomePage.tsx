import { useMutation } from "@tanstack/react-query";
import { useNavigate } from "react-router-dom";
import { api } from "../api";
import { text } from "../strings";
import type { Run } from "../types";
import { useBootstrap, useHistory } from "./hooks";

export function HomePage() {
  const { data } = useBootstrap(); const history = useHistory(); const navigate = useNavigate();
  const settings = data!.settings; const profile = settings.profiles.find((item) => item.id === settings.selected_profile_id) || settings.profiles.find((item) => !item.archived);
  const last = history.data?.find((run) => run.profile_id === profile?.id && run.status === "completed");
  const start = useMutation({ mutationFn: () => api<Run>(`/runs?profile_id=${profile!.id}`, { method: "POST" }), onSuccess: (run) => navigate(`/run/${run.id}`) });
  return <section><header className="page-title"><div><p className="eyebrow">РАБОЧЕЕ ПРОСТРАНСТВО</p><h1>Главная</h1><p>Актуальное сравнение цен аптек — без ручной подготовки файлов.</p></div><span className="version">v{data!.version}</span></header>
    {data!.legacy_config_present && <div className="notice">Найдена старая конфигурация. <button className="link" onClick={() => api("/migration/legacy?remove_secrets=true", { method: "POST" }).then(() => location.reload())}>Импортировать, создать очищенную копию и удалить секреты из старого файла</button></div>}
    {!settings.onboarding_complete && <div className="notice">Первичная настройка не завершена. <button className="link" onClick={() => navigate("/onboarding")}>Продолжить</button></div>}
    <div className="hero card"><div><span className="label">Текущий профиль</span><h2>{profile?.name || "Профиль не создан"}</h2><p>{profile ? `Основная аптека: ${profile.pharmacies.find((p) => p.id === profile.reference_pharmacy_id)?.name || "не выбрана"}` : text.noProfile}</p></div>
      <button className="primary large" disabled={!profile?.reference_pharmacy_id || !data!.credentials.configured || start.isPending} onClick={() => start.mutate()}>{start.isPending ? "Запускаем…" : text.createReport}</button></div>
    {start.error && <p className="error" role="alert">{String(start.error)}</p>}
    <div className="stats"><article className="card"><span className="label">Учетные данные</span><strong>{data!.credentials.configured ? "Настроены" : "Требуют настройки"}</strong><small>{data!.credentials.backend}</small></article>
      <article className="card"><span className="label">Последний полный запуск</span><strong>{last ? new Date(last.started_at).toLocaleString("ru-RU") : "Еще не было"}</strong><small>{last ? `${last.product_count} товаров` : "—"}</small></article>
      <article className="card"><span className="label">Хранилище истории</span><strong>{(data!.history_size_bytes / 1048576).toFixed(1)} МБ</strong><small>{settings.retention ? `до ${settings.retention} запусков на профиль` : "без лимита"}</small></article></div>
  </section>;
}

import { useState } from "react";
import { useForm } from "react-hook-form";
import { useQueryClient } from "@tanstack/react-query";
import { api } from "../api";
import { useBootstrap } from "./hooks";

interface SecretForm { cookie: string; csrf: string }
export function SettingsPage() {
  const { data } = useBootstrap(); const client = useQueryClient(); const settings = data!.settings; const [message, setMessage] = useState("");
  const secrets = useForm<SecretForm>({ defaultValues: { cookie: "", csrf: "" } });
  const patch = async (changes: object) => { await api("/settings", { method: "PATCH", body: JSON.stringify(changes) }); await client.invalidateQueries({ queryKey: ["bootstrap"] }); };
  const theme = (value: "light" | "dark") => { document.documentElement.dataset.theme = value; localStorage.setItem("pharmparser-theme", value); patch({ theme: value }); };
  return <section><header className="page-title"><div><p className="eyebrow">ПАРАМЕТРЫ</p><h1>Настройки</h1><p>Секреты хранятся отдельно от настроек и истории.</p></div></header>
    <div className="settings-grid"><article className="card"><h2>Внешний вид</h2><div className="segmented"><button className={settings.theme === "light" ? "active" : ""} onClick={() => theme("light")}>Светлая</button><button className={settings.theme === "dark" ? "active" : ""} onClick={() => theme("dark")}>Темная</button></div></article>
      <article className="card"><h2>Отчеты</h2><label>Каталог<input defaultValue={settings.output_directory} onBlur={(event) => patch({ output_directory: event.target.value })} /></label><label>Шаблон имени<input defaultValue={settings.file_name_template} onBlur={(event) => patch({ file_name_template: event.target.value })} /></label><label>Формат<select value={settings.report_format} onChange={(event) => patch({ report_format: event.target.value })}><option value="xlsm">XLSM с кнопками</option><option value="xlsx">XLSX без макросов</option></select></label><label>Хранить запусков<input type="number" min="10" max="500" defaultValue={settings.retention || ""} onBlur={(event) => patch({ retention: event.target.value ? Number(event.target.value) : null })} /></label></article>
      <article className="card credentials"><h2>Cookie и CSRF</h2><p className={data!.credentials.configured ? "success" : "error"}>{data!.credentials.configured ? `Настроены (${data!.credentials.backend})` : "Не настроены"}</p>{data!.credentials.warning && <div className="notice">{data!.credentials.warning}</div>}
        <form onSubmit={secrets.handleSubmit(async (values) => { await api("/credentials", { method: "PUT", body: JSON.stringify(values) }); secrets.reset(); setMessage("Учетные данные сохранены"); await client.invalidateQueries({ queryKey: ["bootstrap"] }); })}><label>Cookie<textarea autoComplete="off" {...secrets.register("cookie", { required: true })} /></label><label>CSRF<input type="password" autoComplete="off" {...secrets.register("csrf", { required: true })} /></label><button className="primary" type="submit">Сохранить безопасно</button>{message && <p className="success" role="status">{message}</p>}</form></article>
      <article className="card"><h2>Цвета отчета</h2><div className="colors"><label>Цена ниже<input type="color" defaultValue={`#${settings.red}`} onBlur={(event) => patch({ red: event.target.value.slice(1) })} /></label><label>Цена выше<input type="color" defaultValue={`#${settings.green}`} onBlur={(event) => patch({ green: event.target.value.slice(1) })} /></label></div></article></div>
  </section>;
}

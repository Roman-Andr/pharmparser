import { useEffect } from "react";
import { NavLink, Outlet } from "react-router-dom";
import { api } from "../api";
import { text } from "../strings";
import { useBootstrap } from "./hooks";

export function AppLayout() {
  const bootstrap = useBootstrap();
  useEffect(() => { const beat = () => api("/heartbeat", { method: "POST" }).catch(() => undefined); beat(); const timer = window.setInterval(beat, 30_000); return () => window.clearInterval(timer); }, []);
  if (bootstrap.isLoading) return <main className="center"><div className="spinner" aria-label="Загрузка" /></main>;
  if (bootstrap.error) return <main className="center error">Не удалось запустить приложение: {String(bootstrap.error)}</main>;
  const items = [["/", text.home], ["/profiles", text.profiles], ["/history", text.history], ["/settings", text.settings], ["/help", text.help]];
  return <div className="shell">
    <aside><div className="brand"><span className="brand-mark">P</span><div>PharmParser<small>сравнение цен</small></div></div>
      <nav aria-label="Основная навигация">{items.map(([to, label]) => <NavLink key={to} to={to} end={to === "/"}>{label}</NavLink>)}</nav>
      <button className="ghost exit" onClick={() => api("/exit", { method: "POST" })}>Выйти</button>
    </aside>
    <main><Outlet /></main>
  </div>;
}

import { useState } from "react";
import { useQueryClient } from "@tanstack/react-query";
import { api } from "../api";
import type { Pharmacy, Profile } from "../types";
import { useBootstrap } from "./hooks";

const pharmacyId = (url: string) => url.replace(/\/$/, "").split("/").pop() || "";
export function ProfilesPage() {
  const { data } = useBootstrap(); const client = useQueryClient(); const [selected, setSelected] = useState(data!.settings.profiles[0]?.id || "");
  const original = data!.settings.profiles.find((item) => item.id === selected); const [draft, setDraft] = useState<Profile | null>(original || null);
  const choose = (profile: Profile) => { setSelected(profile.id); setDraft(structuredClone(profile)); api("/settings", { method: "PATCH", body: JSON.stringify({ selected_profile_id: profile.id }) }).then(() => client.invalidateQueries({ queryKey: ["bootstrap"] })); };
  const save = async () => { if (!draft) return; await api(`/profiles/${draft.id}`, { method: "PUT", body: JSON.stringify(draft) }); await api("/settings", { method: "PATCH", body: JSON.stringify({ selected_profile_id: draft.id }) }); await client.invalidateQueries({ queryKey: ["bootstrap"] }); };
  const addPharmacy = () => draft && setDraft({ ...draft, pharmacies: [...draft.pharmacies, { id: "", name: "", url: "" }] });
  const change = (index: number, values: Partial<Pharmacy>) => draft && setDraft({ ...draft, pharmacies: draft.pharmacies.map((item, position) => position === index ? { ...item, ...values } : item) });
  const move = (index: number, direction: -1 | 1) => { if (!draft) return; const next = [...draft.pharmacies]; const target = index + direction; if (target < 0 || target >= next.length) return; [next[index], next[target]] = [next[target], next[index]]; setDraft({ ...draft, pharmacies: next }); };
  return <section><header className="page-title"><div><p className="eyebrow">КОНФИГУРАЦИЯ</p><h1>Профили</h1><p>Основная аптека выбирается явно и не зависит от порядка списка.</p></div></header>
    <div className="split"><div className="profile-list"><button onClick={() => { const profile = { id: crypto.randomUUID(), name: "Новый профиль", pharmacies: [], reference_pharmacy_id: null, archived: false }; setSelected(profile.id); setDraft(profile); }}>＋ Новый профиль</button>{data!.settings.profiles.filter((item) => !item.archived).map((profile) => <button className={profile.id === selected ? "selected" : ""} onClick={() => choose(profile)} key={profile.id}><strong>{profile.name}</strong><small>{profile.pharmacies.length} аптек</small></button>)}</div>
      {draft ? <div className="card editor"><label>Название профиля<input value={draft.name} onChange={(event) => setDraft({ ...draft, name: event.target.value })} /></label><h2>Аптеки</h2>
        <div className="pharmacy-editor">{draft.pharmacies.map((entry, index) => <div className="pharmacy-row" key={`${index}-${entry.id}`}><input aria-label="Основная аптека" type="radio" name="reference" checked={draft.reference_pharmacy_id === entry.id && !!entry.id} onChange={() => setDraft({ ...draft, reference_pharmacy_id: entry.id })} /><label>Название<input value={entry.name} onChange={(event) => change(index, { name: event.target.value })} /></label><label>URL<input value={entry.url} onChange={(event) => change(index, { url: event.target.value, id: pharmacyId(event.target.value) })} /></label><div className="move"><button aria-label="Переместить выше" onClick={() => move(index, -1)}>↑</button><button aria-label="Переместить ниже" onClick={() => move(index, 1)}>↓</button></div></div>)}</div>
        <div className="actions"><button onClick={addPharmacy}>Добавить аптеку</button><button onClick={() => setDraft({ ...structuredClone(draft), id: crypto.randomUUID(), name: `${draft.name} — копия`, archived: false })}>Клонировать</button><button className="primary" onClick={save}>Сохранить профиль</button><button className="danger ghost" onClick={() => api(`/profiles/${draft.id}/archive`, { method: "POST" }).then(() => client.invalidateQueries({ queryKey: ["bootstrap"] }))}>Архивировать</button></div></div> : <div className="card empty">Выберите или создайте профиль</div>}
    </div></section>;
}

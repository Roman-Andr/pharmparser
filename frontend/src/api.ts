const fragmentToken = window.location.hash.slice(1);
const savedToken = sessionStorage.getItem("pharmparser-token");
export const token = fragmentToken || savedToken || "";
if (fragmentToken) {
  sessionStorage.setItem("pharmparser-token", fragmentToken);
  history.replaceState(null, "", window.location.pathname + window.location.search);
}

export async function api<T>(path: string, init: RequestInit = {}): Promise<T> {
  const response = await fetch(`/api${path}`, {
    ...init,
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", ...init.headers },
  });
  if (!response.ok) {
    const text = await response.text();
    throw new Error(text || `Ошибка ${response.status}`);
  }
  if (response.status === 204) return undefined as T;
  return response.json() as Promise<T>;
}

export async function streamEvents(runId: string, onEvent: (event: unknown) => void, signal: AbortSignal) {
  const response = await fetch(`/api/runs/${runId}/events`, {
    headers: { Authorization: `Bearer ${token}` }, signal,
  });
  if (!response.ok || !response.body) throw new Error("Не удалось подключиться к прогрессу");
  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";
  while (true) {
    const { value, done } = await reader.read();
    if (done) break;
    buffer += decoder.decode(value, { stream: true });
    const messages = buffer.split("\n\n");
    buffer = messages.pop() || "";
    for (const message of messages) {
      const data = message.split("\n").find((line) => line.startsWith("data: "));
      if (data) onEvent(JSON.parse(data.slice(6)));
    }
  }
}

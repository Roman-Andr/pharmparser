import { afterEach, describe, expect, it, vi } from "vitest";

describe("локальный API", () => {
  afterEach(() => { vi.restoreAllMocks(); sessionStorage.clear(); });

  it("переносит session token из fragment и не оставляет его в адресной строке", async () => {
    vi.resetModules();
    history.replaceState(null, "", "/#secret-token");
    const module = await import("./api");
    expect(module.token).toBe("secret-token");
    expect(location.hash).toBe("");
    expect(sessionStorage.getItem("pharmparser-token")).toBe("secret-token");
  });

  it("передает token только в Bearer-заголовке", async () => {
    vi.resetModules();
    sessionStorage.setItem("pharmparser-token", "token-123");
    history.replaceState(null, "", "/");
    const fetchMock = vi.spyOn(globalThis, "fetch").mockResolvedValue(
      new Response(JSON.stringify({ ok: true }), { status: 200, headers: { "Content-Type": "application/json" } }),
    );
    const { api } = await import("./api");
    await api("/bootstrap");
    expect(fetchMock).toHaveBeenCalledWith("/api/bootstrap", expect.objectContaining({
      headers: expect.objectContaining({ Authorization: "Bearer token-123" }),
    }));
  });
});

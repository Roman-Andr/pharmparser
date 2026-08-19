import { defineConfig } from "vitest/config";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  build: { outDir: "../src/pharmparser/web/static", emptyOutDir: true },
  server: { host: "127.0.0.1", port: 5173 },
  test: { environment: "happy-dom" },
});

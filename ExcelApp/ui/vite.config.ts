import { defineConfig } from "vitest/config";
import react from "@vitejs/plugin-react";

// ExcelApp UI build. base './' so the built files work from a WebView2 virtual host
// (https://app.excel/index.html) or a plain file server. Vitest reuses this config.
export default defineConfig({
    base: "./",
    plugins: [react()],
    build: {
        outDir: "dist",
        emptyOutDir: true,
        target: "es2020",
    },
    server: {
        port: 5173,
        strictPort: true,
    },
    test: {
        environment: "node",
        include: ["src/**/*.test.ts"],
        globals: false,
    },
});

import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { readFileSync } from "node:fs";

const product = JSON.parse(readFileSync(new URL("../../package.json", import.meta.url), "utf8")) as { version: string; cupolaBuild: string };

export default defineConfig({
  base: "./",
  plugins: [react()],
  define: { __APP_VERSION__: JSON.stringify(product.version), __BUILD_ID__: JSON.stringify(product.cupolaBuild) },
  build: {
    outDir: "dist", emptyOutDir: true, sourcemap: true,
    rollupOptions: { output: { entryFileNames: "assets/workbench.js", chunkFileNames: "assets/[name].js", assetFileNames: "assets/workbench[extname]" } },
  },
});

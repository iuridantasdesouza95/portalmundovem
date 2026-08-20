import { defineConfig } from "@lovable.dev/vite-tanstack-config";
import { resolve } from "node:path";

export default defineConfig({
  build: {
    rollupOptions: {
      input: {
        main: resolve(process.cwd(), "index.html"),
        authPopup: resolve(process.cwd(), "auth-popup.html"),
      },
    },
  },
});

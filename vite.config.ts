// @lovable.dev/vite-tanstack-config already includes the following — do NOT add them manually
// or the app will break with duplicate plugins:
//   - TanStack devtools (dev-only, first), tanstackStart, viteReact, tailwindcss, tsConfigPaths,
//     nitro (build-only using cloudflare as a default target), VITE_* env injection, @ path alias,
//     React/TanStack dedupe, error logger plugins, and sandbox detection (port/host/strictPort).
import { defineConfig } from "@lovable.dev/vite-tanstack-config";
import { resolve } from "node:path";

export default defineConfig({
  tanstackStart: {
    // Redirect TanStack Start's bundled server entry to src/server.ts (our SSR error wrapper).
    server: { entry: "server" },
  },

  // MSAL v5 popup authentication requires a dedicated redirect-bridge page.
  // It must be bundled separately from the React/TanStack application.
  vite: {
    build: {
      rollupOptions: {
        input: {
          authPopup: resolve(process.cwd(), "auth-popup.html"),
        },
      },
    },
  },
});

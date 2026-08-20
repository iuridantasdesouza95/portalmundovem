// @lovable.dev/vite-tanstack-config already includes the following — do NOT add them manually
// or the app will break with duplicate plugins:
//   - TanStack devtools (dev-only, first), tanstackStart, viteReact, tailwindcss, tsConfigPaths,
//     nitro (build-only using cloudflare as a default target), VITE_* env injection, @ path alias,
//     React/TanStack dedupe, error logger plugins, and sandbox detection (port/host/strictPort).
import { defineConfig } from "@lovable.dev/vite-tanstack-config";
import { resolve } from "node:path";

/**
 * MSAL v5 requires a dedicated redirect-bridge HTML entry for popup/silent
 * authentication. We add it after the Lovable/TanStack configuration has
 * resolved its own entries, instead of replacing rollupOptions.input.
 * Replacing the generated Start input breaks the Cloudflare worker bundle.
 */
const msalRedirectBridgeEntry = {
  name: "msal-redirect-bridge-entry",
  configResolved(config: any) {
    // TanStack Start/Lovable uses the normal client build for the browser
    // application. Never alter an SSR/server build input.
    if (config.build?.ssr) return;

    const bridge = resolve(process.cwd(), "auth-popup.html");
    const input = config.build?.rollupOptions?.input;

    if (!input) {
      config.build.rollupOptions.input = {
        authPopup: bridge,
      };
      return;
    }

    if (typeof input === "string") {
      config.build.rollupOptions.input = {
        app: input,
        authPopup: bridge,
      };
      return;
    }

    if (Array.isArray(input)) {
      if (!input.includes(bridge)) input.push(bridge);
      return;
    }

    if (typeof input === "object") {
      input.authPopup = bridge;
    }
  },
};

export default defineConfig({
  tanstackStart: {
    // Keep TanStack Start's normal server entry intact so the Lovable/Cloudflare
    // worker bundle is generated correctly.
    server: { entry: "server" },
  },

  vite: {
    plugins: [msalRedirectBridgeEntry],
  },
});

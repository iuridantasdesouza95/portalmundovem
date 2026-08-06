import { useQuery } from "@tanstack/react-query";
import { supabase } from "@/integrations/supabase/client";

/**
 * Autenticação do Power BI via Microsoft Entra ID (SSO "embed for your organization").
 *
 * O portal usa a própria identidade corporativa do usuário (licença Pro),
 * portanto não é necessário gerar Embed Token por serviço. O token é obtido
 * silenciosamente sempre que possível — o usuário não vê uma segunda tela
 * de login depois do primeiro consentimento.
 */
export const POWERBI_SCOPES = ["https://analysis.windows.net/powerbi/api/Report.Read.All"];

export type EntraConfig = { clientId: string; tenantId: string };

export function useEntraConfig() {
  return useQuery({
    queryKey: ["entra-config"],
    staleTime: 5 * 60_000,
    queryFn: async (): Promise<EntraConfig | null> => {
      const { data } = await supabase
        .from("app_settings")
        .select("key, value")
        .in("key", ["azure_client_id", "azure_tenant_id"]);
      const map = Object.fromEntries((data ?? []).map((s) => [s.key, s.value]));
      const clientId = (map["azure_client_id"] ?? "").trim();
      const tenantId = (map["azure_tenant_id"] ?? "").trim();
      if (!clientId || !tenantId) return null;
      return { clientId, tenantId };
    },
  });
}

let msalPromise: Promise<import("@azure/msal-browser").PublicClientApplication> | null = null;
let msalKey = "";

async function getMsal(config: EntraConfig) {
  const key = `${config.clientId}:${config.tenantId}`;
  if (!msalPromise || msalKey !== key) {
    msalKey = key;
    msalPromise = (async () => {
      const { PublicClientApplication } = await import("@azure/msal-browser");
      const instance = new PublicClientApplication({
        auth: {
          clientId: config.clientId,
          authority: `https://login.microsoftonline.com/${config.tenantId}`,
          redirectUri: window.location.origin,
        },
        cache: { cacheLocation: "localStorage" },
      });
      await instance.initialize();
      await instance.handleRedirectPromise();
      return instance;
    })();
  }
  return msalPromise;
}

/** Retorna um access token do Power BI para o usuário logado. */
export async function getPowerBIToken(config: EntraConfig): Promise<string> {
  const msal = await getMsal(config);
  const account = msal.getAllAccounts()[0];

  if (account) {
    try {
      const result = await msal.acquireTokenSilent({ scopes: POWERBI_SCOPES, account });
      return result.accessToken;
    } catch {
      // cai para o fluxo interativo abaixo
    }
  }

  const result = await msal.acquireTokenPopup({ scopes: POWERBI_SCOPES });
  if (result.account) msal.setActiveAccount(result.account);
  return result.accessToken;
}

export function usePowerBIToken(enabled: boolean) {
  const { data: config } = useEntraConfig();
  return useQuery({
    queryKey: ["powerbi-token", config?.clientId, config?.tenantId],
    enabled: Boolean(enabled && config),
    // tokens do Entra duram ~1h; renovamos com folga
    staleTime: 45 * 60_000,
    retry: false,
    queryFn: async () => getPowerBIToken(config as EntraConfig),
  });
}

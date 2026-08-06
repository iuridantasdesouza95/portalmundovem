import { useEffect } from "react";
import { useQuery, useQueryClient } from "@tanstack/react-query";
import { supabase } from "@/integrations/supabase/client";

/**
 * Autenticação do Power BI via Microsoft Entra ID (MSAL) — "embed for your organization".
 *
 * Estratégia: Access Token delegado do Entra ID (TokenType.Aad).
 * Não usamos Embed Token de serviço: cada usuário acessa com a própria licença Pro,
 * então as permissões do workspace/relatório valem exatamente como no Power BI Service.
 *
 * Login único: a sessão MSAL é criada uma vez (no login do portal, via prewarm) e
 * reaproveitada por todos os relatórios. O visualizador nunca abre nova autenticação.
 */
export const POWERBI_SCOPES = ["https://analysis.windows.net/powerbi/api/Report.Read.All"];

export type EntraConfig = { clientId: string; tenantId: string };
export type PowerBIToken = { accessToken: string; expiresOn: number };

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
        // localStorage mantém a sessão entre abas e reloads: um único login.
        cache: { cacheLocation: "localStorage" },
      });
      await instance.initialize();
      await instance.handleRedirectPromise();
      const account = instance.getAllAccounts()[0];
      if (account) instance.setActiveAccount(account);
      return instance;
    })();
  }
  return msalPromise;
}

/**
 * Obtém o Access Token do Power BI.
 * Ordem: cache/refresh silencioso → SSO silencioso (login hint do portal) → popup (só se tudo falhar).
 */
export async function getPowerBIToken(
  config: EntraConfig,
  options?: { loginHint?: string | undefined; interactive?: boolean | undefined },
): Promise<PowerBIToken> {
  const msal = await getMsal(config);
  const account = msal.getActiveAccount() ?? msal.getAllAccounts()[0];

  const wrap = (r: { accessToken: string; expiresOn: Date | null }): PowerBIToken => ({
    accessToken: r.accessToken,
    expiresOn: r.expiresOn ? r.expiresOn.getTime() : Date.now() + 55 * 60_000,
  });

  if (account) {
    try {
      // Renovação automática: usa o refresh token do MSAL, sem UI e sem reload.
      const result = await msal.acquireTokenSilent({ scopes: POWERBI_SCOPES, account });
      return wrap(result);
    } catch {
      /* segue para SSO silencioso */
    }
  }

  try {
    const result = await msal.ssoSilent({
      scopes: POWERBI_SCOPES,
      ...(options?.loginHint ? { loginHint: options.loginHint } : {}),
    });
    if (result.account) msal.setActiveAccount(result.account);
    return wrap(result);
  } catch {
    /* segue para interativo */
  }

  if (options?.interactive === false) {
    throw new Error("entra_interaction_required");
  }

  // Único ponto interativo: acontece uma vez, no login do portal.
  const result = await msal.acquireTokenPopup({
    scopes: POWERBI_SCOPES,
    ...(options?.loginHint ? { loginHint: options.loginHint } : {}),
  });
  if (result.account) msal.setActiveAccount(result.account);
  return wrap(result);
}

const TOKEN_KEY = (c?: EntraConfig | null) => ["powerbi-token", c?.clientId, c?.tenantId];

/**
 * Token do Power BI com renovação automática antes do vencimento
 * (react-query refetch em background; acquireTokenSilent por baixo).
 */
export function usePowerBIToken(enabled: boolean, loginHint?: string) {
  const { data: config } = useEntraConfig();
  return useQuery({
    queryKey: TOKEN_KEY(config),
    enabled: Boolean(enabled && config),
    retry: false,
    staleTime: 45 * 60_000,
    // renova a cada 40 min (tokens do Entra duram ~60 min), mesmo com a aba em segundo plano
    refetchInterval: 40 * 60_000,
    refetchIntervalInBackground: true,
    refetchOnWindowFocus: true,
    queryFn: async () => getPowerBIToken(config as EntraConfig, { loginHint }),
  });
}

/**
 * Aquece a sessão Entra logo após o login do portal, para que o visualizador
 * de dashboards jamais precise abrir uma segunda tela de autenticação.
 */
export function usePowerBIPrewarm(loginHint?: string) {
  const { data: config } = useEntraConfig();
  const qc = useQueryClient();

  useEffect(() => {
    if (!config) return;
    let cancelled = false;
    void (async () => {
      try {
        const token = await getPowerBIToken(config, { loginHint, interactive: false });
        if (!cancelled) qc.setQueryData(TOKEN_KEY(config), token);
      } catch {
        /* sem sessão Entra ainda: o primeiro acesso a um dashboard fará o login único */
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [config, loginHint, qc]);
}

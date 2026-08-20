import { useEffect } from "react";
import { useQuery, useQueryClient } from "@tanstack/react-query";
import { supabase } from "@/integrations/supabase/client";
import type {
  AccountInfo,
  AuthenticationResult,
  PublicClientApplication,
} from "@azure/msal-browser";

export const LOGIN_SCOPES = ["openid", "profile", "email", "User.Read"];
export const POWERBI_SCOPES = [
  "https://analysis.windows.net/powerbi/api/Report.Read.All",
];
export type EntraConfig = { clientId: string; tenantId: string };
export type PowerBIToken = { accessToken: string; expiresOn: number };

/**
 * Redirect principal do fluxo Microsoft.
 * Deve estar cadastrado no Entra ID como SPA exatamente como retornado aqui.
 */
export function getEntraRedirectUri() {
  return `${window.location.origin}/auth`;
}

/**
 * Redirect silencioso usado exclusivamente para aquisição silenciosa de token.
 */
export function getEntraSilentRedirectUri() {
  return `${window.location.origin}/auth-blank.html`;
}

function parseSettings(
  rows: { key: string; value: unknown }[] | null,
): EntraConfig | null {
  const map = Object.fromEntries(
    (rows ?? []).map((setting) => [setting.key, setting.value]),
  );
  const clientId = String(map["azure_client_id"] ?? "").trim();
  const tenantId = String(map["azure_tenant_id"] ?? "").trim();

  if (!clientId || !tenantId) {
    console.warn("[ENTRA] Client ID ou Tenant ID não configurado.");
    return null;
  }

  return { clientId, tenantId };
}

export async function fetchEntraConfig(): Promise<EntraConfig | null> {
  const { data, error } = await supabase
    .from("app_settings")
    .select("key, value")
    .in("key", ["azure_client_id", "azure_tenant_id"]);

  if (error) throw error;
  return parseSettings(data);
}

export function useEntraConfig() {
  return useQuery({
    queryKey: ["entra-config"],
    staleTime: 5 * 60_000,
    queryFn: fetchEntraConfig,
  });
}

let msalPromise: Promise<PublicClientApplication> | null = null;
let msalKey = "";

export async function getMsal(
  config: EntraConfig,
): Promise<PublicClientApplication> {
  const key = `${config.clientId}:${config.tenantId}`;

  if (!msalPromise || msalKey !== key) {
    msalKey = key;

    msalPromise = (async () => {
      const { PublicClientApplication } = await import("@azure/msal-browser");

      const instance = new PublicClientApplication({
        auth: {
          clientId: config.clientId,
          authority: `https://login.microsoftonline.com/${config.tenantId}`,
          redirectUri: getEntraRedirectUri(),
          postLogoutRedirectUri: getEntraRedirectUri(),
        },
        cache: {
          cacheLocation: "localStorage",
          storeAuthStateInCookie: false,
        },
      });

      await instance.initialize();

      const accounts = instance.getAllAccounts();
      if (accounts.length > 0 && !instance.getActiveAccount()) {
        instance.setActiveAccount(accounts[0]);
      }

      console.info(
        "[ENTRA] MSAL inicializado. Contas em cache:",
        accounts.length,
      );

      return instance;
    })();
  }

  return msalPromise;
}

function resolveAccount(msal: PublicClientApplication): AccountInfo | null {
  const active = msal.getActiveAccount();
  if (active) return active;

  const accounts = msal.getAllAccounts();
  const account: AccountInfo | null =
    accounts.length > 0 ? accounts[0] : null;

  if (account) {
    msal.setActiveAccount(account);
  }

  return account;
}

/**
 * Inicia o login Microsoft usando redirect.
 *
 * IMPORTANTE: esta função NÃO cria a sessão do portal e NÃO retorna
 * AuthenticationResult. O navegador será redirecionado para a Microsoft
 * e, após o retorno para /auth, handleMicrosoftRedirect() deverá processar
 * o resultado.
 */
export async function signInWithMicrosoft(
  config: EntraConfig,
): Promise<void> {
  const msal = await getMsal(config);

  console.info("[ENTRA] Iniciando loginRedirect para Microsoft...");

  await msal.loginRedirect({
    scopes: LOGIN_SCOPES,
    prompt: "select_account",
    redirectUri: getEntraRedirectUri(),
  });
}

/**
 * Processa o retorno do loginRedirect.
 *
 * Deve ser chamado somente na página /auth, que é o redirect cadastrado
 * no Microsoft Entra ID.
 */
export async function handleMicrosoftRedirect(
  config: EntraConfig,
): Promise<AuthenticationResult | null> {
  const msal = await getMsal(config);

  try {
    const result = await msal.handleRedirectPromise();

    if (!result) {
      return null;
    }

    if (!result.account) {
      throw new Error(
        "A Microsoft autenticou o usuário, mas nenhuma conta foi retornada.",
      );
    }

    if (!result.idToken) {
      throw new Error(
        "O Microsoft Entra ID não retornou o token de identidade.",
      );
    }

    msal.setActiveAccount(result.account);

    console.info(
      "[ENTRA] Conta Microsoft autenticada:",
      result.account.username,
    );

    return result;
  } catch (error: any) {
    const code = error?.errorCode ?? error?.code ?? "";

    // Não existe redirect pendente. Isso é normal quando /auth foi aberto
    // diretamente, sem passar primeiro pelo loginRedirect.
    if (code === "no_token_request_cache_error") {
      return null;
    }

    throw error;
  }
}

export async function signOutMsalCache(config?: EntraConfig | null) {
  if (!config) return;

  try {
    const msal = await getMsal(config);
    msal.setActiveAccount(null);
  } catch {
    // Logout do cache não deve impedir o logout principal do portal.
  }
}

const INTERACTION_CODES = new Set([
  "interaction_required",
  "consent_required",
  "login_required",
  "no_tokens_found",
]);

export const POWERBI_LOGIN_REQUIRED = "entra_login_required";
export const POWERBI_CONSENT_REQUIRED =
  "entra_powerbi_interaction_required";

function wrapPowerBIToken(result: {
  accessToken: string;
  expiresOn: Date | null;
}): PowerBIToken {
  return {
    accessToken: result.accessToken,
    expiresOn:
      result.expiresOn?.getTime() ?? Date.now() + 55 * 60_000,
  };
}

export async function getPowerBIToken(
  config: EntraConfig,
  options?: { loginHint?: string },
): Promise<PowerBIToken> {
  const msal = await getMsal(config);
  const account = resolveAccount(msal);

  if (!account) {
    throw new Error(POWERBI_LOGIN_REQUIRED);
  }

  if (
    options?.loginHint &&
    account.username.toLowerCase() !== options.loginHint.toLowerCase()
  ) {
    console.warn(
      "[POWERBI] Conta MSAL difere do usuário do portal.",
    );
  }

  try {
    const result = await msal.acquireTokenSilent({
      scopes: POWERBI_SCOPES,
      account,
      redirectUri: getEntraSilentRedirectUri(),
    });

    if (!result.accessToken) {
      throw new Error(POWERBI_CONSENT_REQUIRED);
    }

    return wrapPowerBIToken(result);
  } catch (error: any) {
    const errorCode = error?.errorCode ?? error?.code ?? "";

    if (INTERACTION_CODES.has(errorCode)) {
      throw new Error(POWERBI_CONSENT_REQUIRED);
    }

    throw error;
  }
}

export async function requestPowerBIConsent(
  config: EntraConfig,
): Promise<PowerBIToken> {
  const msal = await getMsal(config);
  const account = resolveAccount(msal);

  await msal.acquireTokenRedirect({
    scopes: POWERBI_SCOPES,
    ...(account ? { account } : {}),
    redirectUri: getEntraRedirectUri(),
  });

  throw new Error(
    "O consentimento do Power BI foi iniciado. Após concluir, o portal continuará automaticamente.",
  );
}

const TOKEN_KEY = (config?: EntraConfig | null) => [
  "powerbi-token",
  config?.clientId,
  config?.tenantId,
];

export function usePowerBIToken(
  enabled: boolean,
  loginHint?: string,
) {
  const { data: config } = useEntraConfig();
  const queryClient = useQueryClient();

  const query = useQuery({
    queryKey: TOKEN_KEY(config),
    enabled: Boolean(enabled) && Boolean(config),
    retry: false,
    staleTime: 45 * 60_000,
    refetchInterval: 40 * 60_000,
    refetchIntervalInBackground: true,
    refetchOnWindowFocus: true,
    queryFn: async () => {
      if (!config) {
        throw new Error("Entra ID não configurado.");
      }

      return getPowerBIToken(config, { loginHint });
    },
  });

  async function consent() {
    if (!config) {
      throw new Error("Entra ID não configurado.");
    }

    return requestPowerBIConsent(config);
  }

  return { ...query, consent, config };
}

export function usePowerBIPrewarm(loginHint?: string) {
  const { data: config } = useEntraConfig();
  const queryClient = useQueryClient();

  useEffect(() => {
    if (!config) return;

    let cancelled = false;

    void (async () => {
      try {
        const token = await getPowerBIToken(config, { loginHint });

        if (!cancelled) {
          queryClient.setQueryData(TOKEN_KEY(config), token);
        }
      } catch (error: any) {
        if (error?.message === POWERBI_LOGIN_REQUIRED) {
          console.info(
            "[POWERBI] Prewarm aguardando autenticação Microsoft.",
          );
        } else if (error?.message === POWERBI_CONSENT_REQUIRED) {
          console.info(
            "[POWERBI] Prewarm aguardando consentimento do Power BI.",
          );
        }
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [config, loginHint, queryClient]);
}

import { useEffect } from "react";

import {
  useQuery,
  useQueryClient,
} from "@tanstack/react-query";

import { supabase } from "@/integrations/supabase/client";

import type {
  AccountInfo,
  PublicClientApplication,
} from "@azure/msal-browser";

/* -------------------------------------------------------------------------- */
/* CONFIGURAÇÃO                                                               */
/* -------------------------------------------------------------------------- */

export const LOGIN_SCOPES = [
  "openid",
  "profile",
  "email",
  "User.Read",
];

export const POWERBI_SCOPES = [
  "https://analysis.windows.net/powerbi/api/Report.Read.All",
];

export type EntraConfig = {
  clientId: string;
  tenantId: string;
};

export type PowerBIToken = {
  accessToken: string;
  expiresOn: number;
};

function parseSettings(
  rows: { key: string; value: unknown }[] | null,
): EntraConfig | null {
  const map = Object.fromEntries(
    (rows ?? []).map((setting) => [
      setting.key,
      setting.value,
    ]),
  );

  const clientId = String(
    map["azure_client_id"] ?? "",
  ).trim();

  const tenantId = String(
    map["azure_tenant_id"] ?? "",
  ).trim();

  if (!clientId || !tenantId) {
    console.warn(
      "[ENTRA] Client ID ou Tenant ID não configurado.",
    );

    return null;
  }

  return { clientId, tenantId };
}

export async function fetchEntraConfig(): Promise<EntraConfig | null> {
  const { data, error } = await supabase
    .from("app_settings")
    .select("key, value")
    .in("key", ["azure_client_id", "azure_tenant_id"]);

  if (error) {
    console.error(
      "[ENTRA] Erro ao buscar configuração:",
      error,
    );

    throw error;
  }

  return parseSettings(data);
}

export function useEntraConfig() {
  return useQuery({
    queryKey: ["entra-config"],
    staleTime: 5 * 60_000,
    queryFn: fetchEntraConfig,
  });
}

/* -------------------------------------------------------------------------- */
/* MSAL — INSTÂNCIA ÚNICA                                                     */
/* -------------------------------------------------------------------------- */

let msalPromise: Promise<PublicClientApplication> | null = null;
let msalKey = "";

/**
 * Cria (uma única vez) e inicializa a instância MSAL.
 *
 * Fluxo escolhido: POPUP.
 * Por isso NÃO existe handleRedirectPromise() aqui — chamá-lo sem um
 * loginRedirect() correspondente é o que produzia
 * `no_token_request_cache_error`.
 */
export async function getMsal(
  config: EntraConfig,
): Promise<PublicClientApplication> {
  const key = `${config.clientId}:${config.tenantId}`;

  if (!msalPromise || msalKey !== key) {
    msalKey = key;

    msalPromise = (async () => {
      const { PublicClientApplication } = await import(
        "@azure/msal-browser"
      );

      const instance = new PublicClientApplication({
        auth: {
          clientId: config.clientId,
          authority: `https://login.microsoftonline.com/${config.tenantId}`,
          redirectUri: window.location.origin,
          postLogoutRedirectUri: window.location.origin,
        },

        cache: {
          cacheLocation: "localStorage",
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

function resolveAccount(
  msal: PublicClientApplication,
): AccountInfo | null {
  const active = msal.getActiveAccount();

  if (active) {
    return active;
  }

  const account: AccountInfo | null =
    msal.getAllAccounts()[0] ?? null;

  if (account) {
    msal.setActiveAccount(account);
  }

  return account;
}

/* -------------------------------------------------------------------------- */
/* LOGIN MICROSOFT — ÚNICO RESPONSÁVEL                                        */
/* -------------------------------------------------------------------------- */

let loginPromise: Promise<{
  idToken: string;
  account: AccountInfo;
}> | null = null;

/**
 * Único ponto de login Microsoft do portal.
 *
 * 1. Se já existe conta MSAL, tenta obter um ID Token silenciosamente.
 * 2. Só abre `loginPopup()` quando o silencioso não for possível.
 * 3. Nunca usa loginRedirect / handleRedirectPromise.
 */
export async function signInWithMicrosoft(
  config: EntraConfig,
): Promise<{ idToken: string; account: AccountInfo }> {
  if (loginPromise) {
    console.info(
      "[ENTRA] Login Microsoft já em andamento. Reaproveitando.",
    );

    return loginPromise;
  }

  loginPromise = (async () => {
    const msal = await getMsal(config);

    const cached = resolveAccount(msal);

    if (cached) {
      try {
        const silent = await msal.acquireTokenSilent({
          scopes: LOGIN_SCOPES,
          account: cached,
        });

        if (silent.idToken) {
          console.info(
            "[ENTRA] Sessão Microsoft reutilizada silenciosamente:",
            cached.username,
          );

          msal.setActiveAccount(silent.account ?? cached);

          return {
            idToken: silent.idToken,
            account: silent.account ?? cached,
          };
        }
      } catch (error) {
        console.info(
          "[ENTRA] Não foi possível reutilizar a sessão silenciosamente. Abrindo popup.",
        );
      }
    }

    console.info("[ENTRA] Abrindo loginPopup...");

    const response = await msal.loginPopup({
      scopes: LOGIN_SCOPES,
      prompt: "select_account",
    });

    if (!response?.account) {
      throw new Error(
        "A Microsoft autenticou o usuário, mas nenhuma conta foi retornada.",
      );
    }

    if (!response.idToken) {
      throw new Error(
        "O Microsoft Entra ID não retornou o token de identidade.",
      );
    }

    msal.setActiveAccount(response.account);

    console.info(
      "[ENTRA] Login Microsoft concluído:",
      response.account.username,
    );

    return {
      idToken: response.idToken,
      account: response.account,
    };
  })();

  try {
    return await loginPromise;
  } finally {
    loginPromise = null;
  }
}

export async function signOutMsalCache(config?: EntraConfig | null) {
  if (!config) {
    return;
  }

  try {
    const msal = await getMsal(config);
    msal.setActiveAccount(null);
  } catch {
    // noop
  }
}

/* -------------------------------------------------------------------------- */
/* TOKEN POWER BI                                                             */
/* -------------------------------------------------------------------------- */

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

    expiresOn: result.expiresOn
      ? result.expiresOn.getTime()
      : Date.now() + 55 * 60_000,
  };
}

/**
 * Token delegado do Power BI — SEMPRE silencioso.
 * Nunca abre popup durante a navegação.
 */
export async function getPowerBIToken(
  config: EntraConfig,
  options?: { loginHint?: string | undefined },
): Promise<PowerBIToken> {
  const msal = await getMsal(config);

  const account = resolveAccount(msal);

  if (!account) {
    console.info(
      "[POWERBI] Nenhuma sessão Microsoft disponível.",
    );

    throw new Error(POWERBI_LOGIN_REQUIRED);
  }

  if (
    options?.loginHint &&
    account.username.toLowerCase() !==
      options.loginHint.toLowerCase()
  ) {
    console.warn(
      "[POWERBI] Conta MSAL difere do usuário do portal.",
      {
        portal: options.loginHint,
        microsoft: account.username,
      },
    );
  }

  try {
    const result = await msal.acquireTokenSilent({
      scopes: POWERBI_SCOPES,
      account,
    });

    if (!result.accessToken) {
      throw new Error(POWERBI_CONSENT_REQUIRED);
    }

    console.info(
      "[POWERBI] Access Token obtido silenciosamente.",
    );

    return wrapPowerBIToken(result);
  } catch (error: any) {
    const errorCode = error?.errorCode ?? error?.code ?? "";

    if (INTERACTION_CODES.has(errorCode)) {
      console.info(
        "[POWERBI] Consentimento/interação necessária para o Power BI.",
      );

      throw new Error(POWERBI_CONSENT_REQUIRED);
    }

    console.error(
      "[POWERBI] acquireTokenSilent falhou:",
      errorCode || error?.message,
    );

    throw error;
  }
}

/**
 * Consentimento interativo do Power BI.
 * Só deve ser chamado por ação explícita do usuário,
 * quando o fluxo silencioso pedir consentimento.
 */
export async function requestPowerBIConsent(
  config: EntraConfig,
): Promise<PowerBIToken> {
  const msal = await getMsal(config);

  const account = resolveAccount(msal);

  const result = await msal.acquireTokenPopup({
    scopes: POWERBI_SCOPES,
    ...(account ? { account } : {}),
  });

  if (result.account) {
    msal.setActiveAccount(result.account);
  }

  console.info("[POWERBI] Consentimento concluído.");

  return wrapPowerBIToken(result);
}

/* -------------------------------------------------------------------------- */
/* REACT QUERY                                                                */
/* -------------------------------------------------------------------------- */

const TOKEN_KEY = (config?: EntraConfig | null) => [
  "powerbi-token",
  config?.clientId,
  config?.tenantId,
];

export function usePowerBIToken(
  enabled: boolean,
  loginHint?: string | undefined,
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

    const token = await requestPowerBIConsent(config);

    queryClient.setQueryData(TOKEN_KEY(config), token);

    return token;
  }

  return { ...query, consent, config };
}

/* -------------------------------------------------------------------------- */
/* PREWARM — SOMENTE SILENCIOSO                                               */
/* -------------------------------------------------------------------------- */

export function usePowerBIPrewarm(loginHint?: string | undefined) {
  const { data: config } = useEntraConfig();

  const queryClient = useQueryClient();

  useEffect(() => {
    if (!config) {
      return;
    }

    let cancelled = false;

    void (async () => {
      try {
        const token = await getPowerBIToken(config, {
          loginHint,
        });

        if (cancelled) {
          return;
        }

        queryClient.setQueryData(TOKEN_KEY(config), token);

        console.info("[POWERBI] Prewarm concluído.");
      } catch (error: any) {
        if (error?.message === POWERBI_LOGIN_REQUIRED) {
          console.info(
            "[POWERBI] Prewarm aguardando autenticação Microsoft.",
          );
        } else if (
          error?.message === POWERBI_CONSENT_REQUIRED
        ) {
          console.info(
            "[POWERBI] Prewarm aguardando consentimento do Power BI.",
          );
        } else {
          console.warn(
            "[POWERBI] Prewarm silencioso não disponível.",
          );
        }
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [config, loginHint, queryClient]);
}

import { useEffect } from "react";
import { useQuery, useQueryClient } from "@tanstack/react-query";
import { supabase } from "@/integrations/supabase/client";

/**
 * Microsoft Entra ID + Power BI
 *
 * Cenário:
 * Embed for your organization / User owns data
 *
 * O usuário autentica com a própria conta Microsoft Entra ID
 * e o Power BI recebe um Access Token delegado específico
 * para a API do Power BI.
 *
 * IMPORTANTE:
 * - ID Token NÃO é usado pelo Power BI.
 * - Access Token do Microsoft Graph NÃO é usado pelo Power BI.
 * - O Power BI recebe um Access Token específico da API:
 *
 *   https://analysis.windows.net/powerbi/api/
 *
 * Isso permite que o usuário utilize as permissões que já possui
 * no Power BI Service.
 */

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

/* -------------------------------------------------------------------------- */
/* CONFIGURAÇÃO DO ENTRA                                                      */
/* -------------------------------------------------------------------------- */

export function useEntraConfig() {
  return useQuery({
    queryKey: ["entra-config"],

    staleTime: 5 * 60_000,

    queryFn: async (): Promise<EntraConfig | null> => {
      const { data, error } = await supabase
        .from("app_settings")
        .select("key, value")
        .in("key", ["azure_client_id", "azure_tenant_id"]);

      if (error) {
        console.error(
          "[ENTRA] Erro ao carregar configuração:",
          error,
        );

        throw error;
      }

      const map = Object.fromEntries(
        (data ?? []).map((setting) => [
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

      console.log("[ENTRA] Configuração carregada:", {
        clientId,
        tenantId,
      });

      if (!clientId || !tenantId) {
        console.warn(
          "[ENTRA] Client ID ou Tenant ID não configurado.",
        );

        return null;
      }

      return {
        clientId,
        tenantId,
      };
    },
  });
}

/**
 * Versão não-hook da configuração.
 *
 * Usada durante o login do portal.
 */
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

  const map = Object.fromEntries(
    (data ?? []).map((setting) => [
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
    return null;
  }

  return {
    clientId,
    tenantId,
  };
}

/* -------------------------------------------------------------------------- */
/* MSAL                                                                       */
/* -------------------------------------------------------------------------- */

let msalPromise:
  | Promise<
      import("@azure/msal-browser").PublicClientApplication
    >
  | null = null;

let msalKey = "";

/**
 * Cria/reutiliza a instância MSAL.
 *
 * A sessão fica armazenada no localStorage.
 */
export async function getMsal(config: EntraConfig) {
  const key = `${config.clientId}:${config.tenantId}`;

  if (!msalPromise || msalKey !== key) {
    msalKey = key;

    msalPromise = (async () => {
      const {
        PublicClientApplication,
      } = await import("@azure/msal-browser");

      const instance = new PublicClientApplication({
        auth: {
          clientId: config.clientId,

          authority:
            `https://login.microsoftonline.com/${config.tenantId}`,

          redirectUri: window.location.origin,

          postLogoutRedirectUri:
            window.location.origin,
        },

        cache: {
          cacheLocation: "localStorage",

          storeAuthStateInCookie: false,
        },
      });

      await instance.initialize();

      try {
        await instance.handleRedirectPromise();
      } catch (error) {
        console.warn(
          "[ENTRA] handleRedirectPromise:",
          error,
        );
      }

      const accounts = instance.getAllAccounts();

      console.log(
        "[ENTRA] Contas encontradas no cache:",
        accounts.map((account) => account.username),
      );

      if (accounts.length > 0) {
        instance.setActiveAccount(accounts[0]);
      }

      return instance;
    })();
  }

  return msalPromise;
}

/* -------------------------------------------------------------------------- */
/* LOGIN PRINCIPAL                                                            */
/* -------------------------------------------------------------------------- */

/**
 * Login principal do portal.
 *
 * IMPORTANTE:
 *
 * Aqui NÃO tentamos usar o Access Token do Power BI.
 *
 * O login apenas estabelece a sessão Microsoft/Entra.
 *
 * Depois disso, getPowerBIToken() solicita especificamente
 * o Access Token da API do Power BI.
 */
export async function loginWithEntra(
  config: EntraConfig,
) {
  console.log("[ENTRA] Iniciando login Microsoft...");

  const msal = await getMsal(config);

  const result = await msal.loginPopup({
    scopes: [
      "openid",
      "profile",
      "email",
      "User.Read",
    ],

    prompt: "select_account",
  });

  if (!result.account) {
    throw new Error(
      "O Microsoft Entra ID não retornou uma conta.",
    );
  }

  msal.setActiveAccount(result.account);

  console.log(
    "[ENTRA] Login Microsoft concluído:",
    result.account.username,
  );

  if (!result.idToken) {
    throw new Error(
      "Entra ID não retornou o token de identidade.",
    );
  }

  return {
    idToken: result.idToken,
    account: result.account,
  };
}

/* -------------------------------------------------------------------------- */
/* TOKEN DO POWER BI                                                          */
/* -------------------------------------------------------------------------- */

/**
 * Converte o resultado do MSAL para nosso tipo interno.
 */
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
 * Obtém especificamente o Access Token da API do Power BI.
 *
 * Ordem:
 *
 * 1. acquireTokenSilent
 * 2. SSO silent
 * 3. popup
 *
 * O popup só acontece se o token não puder ser obtido
 * silenciosamente.
 */
export async function getPowerBIToken(
  config: EntraConfig,
  options?: {
    loginHint?: string;
    interactive?: boolean;
  },
): Promise<PowerBIToken> {
  console.log(
    "[POWERBI] ========================================",
  );

  console.log(
    "[POWERBI] Iniciando obtenção do token Power BI",
  );

  console.log("[POWERBI] Client ID:", config.clientId);

  console.log("[POWERBI] Tenant ID:", config.tenantId);

  console.log(
    "[POWERBI] Scopes:",
    POWERBI_SCOPES,
  );

  console.log(
    "[POWERBI] Login hint:",
    options?.loginHint,
  );

  const msal = await getMsal(config);

  const accounts = msal.getAllAccounts();

  console.log(
    "[POWERBI] Contas MSAL:",
    accounts.map((account) => account.username),
  );

  const account =
    msal.getActiveAccount() ??
    accounts[0];

  /* ---------------------------------------------------------------------- */
  /* 1. TOKEN SILENCIOSO                                                     */
  /* ---------------------------------------------------------------------- */

  if (account) {
    console.log(
      "[POWERBI] Conta ativa:",
      account.username,
    );

    try {
      console.log(
        "[POWERBI] Tentando acquireTokenSilent...",
      );

      const result =
        await msal.acquireTokenSilent({
          scopes: POWERBI_SCOPES,

          account,
        });

      console.log(
        "[POWERBI] ========================================",
      );

      console.log(
        "[POWERBI] TOKEN POWER BI OBTIDO COM SUCESSO",
      );

      console.log(
        "[POWERBI] Conta:",
        result.account?.username,
      );

      console.log(
        "[POWERBI] Expira em:",
        result.expiresOn,
      );

      console.log(
        "[POWERBI] ========================================",
      );

      return wrapPowerBIToken(result);
    } catch (error) {
      console.error(
        "[POWERBI] acquireTokenSilent FALHOU:",
        error,
      );
    }
  } else {
    console.warn(
      "[POWERBI] Nenhuma conta MSAL encontrada.",
    );
  }

  /* ---------------------------------------------------------------------- */
  /* 2. SSO SILENCIOSO                                                       */
  /* ---------------------------------------------------------------------- */

  try {
    console.log(
      "[POWERBI] Tentando ssoSilent...",
    );

    const result = await msal.ssoSilent({
      scopes: POWERBI_SCOPES,

      ...(options?.loginHint
        ? {
            loginHint: options.loginHint,
          }
        : {}),
    });

    if (result.account) {
      msal.setActiveAccount(result.account);
    }

    console.log(
      "[POWERBI] ssoSilent conseguiu o token.",
    );

    console.log(
      "[POWERBI] Conta:",
      result.account?.username,
    );

    console.log(
      "[POWERBI] Expira em:",
      result.expiresOn,
    );

    return wrapPowerBIToken(result);
  } catch (error) {
    console.error(
      "[POWERBI] ssoSilent FALHOU:",
      error,
    );
  }

  /* ---------------------------------------------------------------------- */
  /* 3. INTERAÇÃO                                                            */
  /* ---------------------------------------------------------------------- */

  if (options?.interactive === false) {
    console.error(
      "[POWERBI] Token requer interação, mas interactive=false.",
    );

    throw new Error(
      "entra_interaction_required",
    );
  }

  console.log(
    "[POWERBI] Abrindo popup para consentimento/token Power BI...",
  );

  try {
    const result =
      await msal.acquireTokenPopup({
        scopes: POWERBI_SCOPES,

        ...(options?.loginHint
          ? {
              loginHint: options.loginHint,
            }
          : {}),
      });

    if (result.account) {
      msal.setActiveAccount(result.account);
    }

    console.log(
      "[POWERBI] Token Power BI obtido via popup.",
    );

    console.log(
      "[POWERBI] Conta:",
      result.account?.username,
    );

    console.log(
      "[POWERBI] Expira em:",
      result.expiresOn,
    );

    return wrapPowerBIToken(result);
  } catch (error) {
    console.error(
      "[POWERBI] acquireTokenPopup FALHOU:",
      error,
    );

    throw error;
  }
}

/* -------------------------------------------------------------------------- */
/* REACT QUERY                                                                */
/* -------------------------------------------------------------------------- */

const TOKEN_KEY = (
  config?: EntraConfig | null,
) => [
  "powerbi-token",
  config?.clientId,
  config?.tenantId,
];

/**
 * Token Power BI com cache e renovação automática.
 */
export function usePowerBIToken(
  enabled: boolean,
  loginHint?: string,
) {
  const { data: config } =
    useEntraConfig();

  return useQuery({
    queryKey: TOKEN_KEY(config),

    enabled:
      Boolean(enabled) &&
      Boolean(config),

    retry: false,

    staleTime: 45 * 60_000,

    refetchInterval: 40 * 60_000,

    refetchIntervalInBackground: true,

    refetchOnWindowFocus: true,

    queryFn: async () => {
      if (!config) {
        throw new Error(
          "Entra ID não configurado.",
        );
      }

      return getPowerBIToken(
        config,
        {
          loginHint,
        },
      );
    },
  });
}

/* -------------------------------------------------------------------------- */
/* PREWARM                                                                    */
/* -------------------------------------------------------------------------- */

/**
 * Tenta aquecer o token Power BI após o login.
 *
 * Não abre popup.
 *
 * Se não existir consentimento/sessão suficiente,
 * simplesmente deixa para o primeiro acesso ao dashboard.
 */
export function usePowerBIPrewarm(
  loginHint?: string,
) {
  const { data: config } =
    useEntraConfig();

  const queryClient =
    useQueryClient();

  useEffect(() => {
    if (!config) {
      return;
    }

    let cancelled = false;

    void (async () => {
      try {
        console.log(
          "[POWERBI] Prewarm iniciado.",
        );

        const token =
          await getPowerBIToken(
            config,
            {
              loginHint,
              interactive: false,
            },
          );

        if (!cancelled) {
          queryClient.setQueryData(
            TOKEN_KEY(config),
            token,
          );

          console.log(
            "[POWERBI] Prewarm concluído.",
          );
        }
      } catch (error) {
        console.warn(
          "[POWERBI] Prewarm não conseguiu obter token:",
          error,
        );
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [
    config,
    loginHint,
    queryClient,
  ]);
}

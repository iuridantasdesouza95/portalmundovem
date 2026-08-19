import { useEffect } from "react";
import {
  useQuery,
  useQueryClient,
} from "@tanstack/react-query";
import { supabase } from "@/integrations/supabase/client";

/* -------------------------------------------------------------------------- */
/* CONFIGURAÇÃO                                                               */
/* -------------------------------------------------------------------------- */

/**
 * Permissão necessária para solicitar o token do Power BI.
 */
export const POWERBI_SCOPES = [
  "https://analysis.windows.net/powerbi/api/Report.Read.All",
];

/**
 * Permissões utilizadas durante o login Microsoft.
 *
 * User.Read é do Microsoft Graph.
 * POWERBI_SCOPES é do recurso Power BI.
 *
 * Recursos diferentes devem ser solicitados em chamadas separadas.
 */
const LOGIN_SCOPES = [
  "openid",
  "profile",
  "email",
  "User.Read",
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
/* CONFIGURAÇÃO ENTRA                                                         */
/* -------------------------------------------------------------------------- */

export function useEntraConfig() {
  return useQuery({
    queryKey: ["entra-config"],

    staleTime: 5 * 60_000,

    retry: false,

    queryFn: async (): Promise<EntraConfig | null> => {
      const { data, error } = await supabase
        .from("app_settings")
        .select("key, value")
        .in("key", [
          "azure_client_id",
          "azure_tenant_id",
        ]);

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

/* -------------------------------------------------------------------------- */
/* CONFIGURAÇÃO NÃO-HOOK                                                      */
/* -------------------------------------------------------------------------- */

export async function fetchEntraConfig(): Promise<EntraConfig | null> {
  console.log("[ENTRA] Buscando configuração...");

  const { data, error } = await supabase
    .from("app_settings")
    .select("key, value")
    .in("key", [
      "azure_client_id",
      "azure_tenant_id",
    ]);

  console.log("[ENTRA] Resultado da consulta:", {
    data,
    error,
  });

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

  console.log("[ENTRA] Configuração encontrada:", {
    hasClientId: Boolean(clientId),
    hasTenantId: Boolean(tenantId),
  });

  if (!clientId || !tenantId) {
    console.error(
      "[ENTRA] Client ID ou Tenant ID não foi encontrado.",
    );

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

type MsalInstance =
  import("@azure/msal-browser").PublicClientApplication;

let msalPromise: Promise<MsalInstance> | null = null;

let msalKey = "";

let redirectHandled = false;

/**
 * Evita duas inicializações simultâneas do MSAL.
 */
export async function getMsal(
  config: EntraConfig,
): Promise<MsalInstance> {
  const key =
    `${config.clientId}:${config.tenantId}`;

  if (msalPromise && msalKey === key) {
    return msalPromise;
  }

  msalKey = key;

  msalPromise = (async () => {
    const {
      PublicClientApplication,
    } = await import("@azure/msal-browser");

    const instance =
      new PublicClientApplication({
        auth: {
          clientId: config.clientId,

          authority:
            `https://login.microsoftonline.com/${config.tenantId}`,

          redirectUri:
            `${window.location.origin}/auth`,

          postLogoutRedirectUri:
            `${window.location.origin}/auth`,
        },

        cache: {
          cacheLocation:
            "localStorage",

          storeAuthStateInCookie:
            false,
        },

        system: {
          allowRedirectInIframe: false,
        },
      });

    /*
     * O initialize precisa terminar antes de qualquer operação MSAL.
     */
    await instance.initialize();

    console.log(
      "[ENTRA] MSAL inicializado.",
    );

    /*
     * IMPORTANTE:
     *
     * Se o usuário acabou de voltar da Microsoft,
     * handleRedirectPromise precisa ser concluído
     * antes de consultar contas ou iniciar outro login.
     */
    if (!redirectHandled) {
      try {
        const redirectResult =
          await instance.handleRedirectPromise();

        redirectHandled = true;

        if (redirectResult?.account) {
          instance.setActiveAccount(
            redirectResult.account,
          );

          console.log(
            "[ENTRA] Login recebido pelo redirect:",
            redirectResult.account.username,
          );
        } else {
          console.log(
            "[ENTRA] Nenhum resultado de redirect pendente.",
          );
        }
      } catch (error) {
        redirectHandled = true;

        console.error(
          "[ENTRA] Erro ao processar retorno do Microsoft:",
          error,
        );

        throw error;
      }
    }

    /*
     * Recupera contas já armazenadas no cache.
     */
    const accounts =
      instance.getAllAccounts();

    console.info(
      "[ENTRA] Contas MSAL disponíveis:",
      accounts.map(
        (account) =>
          account.username,
      ),
    );

    /*
     * Define uma conta ativa caso exista apenas uma.
     */
    if (accounts.length === 1) {
      instance.setActiveAccount(
        accounts[0],
      );

      console.info(
        "[ENTRA] Conta Microsoft definida como ativa:",
        accounts[0].username,
      );
    }

    /*
     * Se não houver conta ativa, tenta recuperar a primeira.
     */
    if (
      !instance.getActiveAccount() &&
      accounts.length > 0
    ) {
      instance.setActiveAccount(
        accounts[0],
      );
    }

    return instance;
  })();

  try {
    return await msalPromise;
  } catch (error) {
    /*
     * Permite nova tentativa caso a inicialização tenha falhado.
     */
    msalPromise = null;
    msalKey = "";

    throw error;
  }
}

/* -------------------------------------------------------------------------- */
/* LOGIN MICROSOFT                                                            */
/* -------------------------------------------------------------------------- */

/**
 * Login principal do Portal.
 *
 * IMPORTANTE:
 *
 * Este método NÃO tenta:
 *
 * - acquireTokenPopup
 * - ssoSilent
 * - acquireTokenSilent
 *
 * Primeiro precisamos autenticar o usuário.
 *
 * Depois do retorno da Microsoft, o /auth será carregado novamente
 * e getMsal() processará handleRedirectPromise().
 */
export async function loginWithEntra(
  config: EntraConfig,
): Promise<never> {
  console.log(
    "[ENTRA] Iniciando login Microsoft...",
  );

  const msal =
    await getMsal(config);

  const currentAccount =
    msal.getActiveAccount();

  if (currentAccount) {
    console.log(
      "[ENTRA] Já existe uma conta Microsoft ativa:",
      currentAccount.username,
    );

    /*
     * A conta já está autenticada no MSAL.
     *
     * Não iniciamos outro login.
     *
     * O auth.tsx poderá continuar o processo.
     */
    throw new Error(
      "entra_already_authenticated",
    );
  }

  console.log(
    "[ENTRA] Nenhuma conta autenticada.",
  );

  console.log(
    "[ENTRA] Redirecionando para Microsoft Entra ID...",
  );

  /*
   * loginRedirect navega para a Microsoft.
   *
   * O código abaixo não será executado nesta página.
   *
   * Ao retornar para /auth, getMsal() executará
   * handleRedirectPromise() e recuperará a conta.
   */
  await msal.loginRedirect({
    scopes: LOGIN_SCOPES,

    prompt: "select_account",

    redirectUri:
      `${window.location.origin}/auth`,
  });

  /*
   * Segurança de TypeScript.
   */
  throw new Error(
    "entra_redirect_started",
  );
}

/* -------------------------------------------------------------------------- */
/* CONTA MICROSOFT                                                            */
/* -------------------------------------------------------------------------- */

/**
 * Retorna a conta Microsoft atualmente autenticada.
 */
export async function getEntraAccount(
  config: EntraConfig,
) {
  const msal =
    await getMsal(config);

  let account =
    msal.getActiveAccount();

  if (account) {
    return account;
  }

  const accounts =
    msal.getAllAccounts();

  if (accounts.length > 0) {
    account =
      accounts[0];

    msal.setActiveAccount(
      account,
    );

    return account;
  }

  return null;
}

/* -------------------------------------------------------------------------- */
/* TOKEN POWER BI                                                             */
/* -------------------------------------------------------------------------- */

function wrapPowerBIToken(result: {
  accessToken: string;
  expiresOn: Date | null;
}): PowerBIToken {
  return {
    accessToken:
      result.accessToken,

    expiresOn:
      result.expiresOn
        ? result.expiresOn.getTime()
        : Date.now() +
          55 * 60_000,
  };
}

/**
 * Obtém o Access Token específico do Power BI.
 *
 * NÃO faz login.
 *
 * NÃO abre popup.
 *
 * NÃO executa ssoSilent.
 *
 * O usuário precisa estar autenticado no Microsoft Entra.
 */
export async function getPowerBIToken(
  config: EntraConfig,
  options?: {
    loginHint?: string;
    interactive?: boolean;
  },
): Promise<PowerBIToken> {
  console.info(
    "[POWERBI] Obtendo Access Token...",
  );

  const msal =
    await getMsal(config);

  let account =
    msal.getActiveAccount();

  /*
   * Recupera a conta do cache caso necessário.
   */
  if (!account) {
    const accounts =
      msal.getAllAccounts();

    if (accounts.length > 0) {
      account =
        accounts[0];

      msal.setActiveAccount(
        account,
      );

      console.info(
        "[POWERBI] Conta encontrada no cache:",
        account.username,
      );
    }
  }

  /*
   * Sem conta Microsoft não existe como obter
   * silenciosamente o token do Power BI.
   */
  if (!account) {
    console.warn(
      "[POWERBI] Nenhuma sessão Microsoft encontrada.",
    );

    throw new Error(
      "entra_login_required",
    );
  }

  console.info(
    "[POWERBI] Conta Microsoft utilizada:",
    account.username,
  );

  /*
   * Apenas diagnóstico.
   */
  if (
    options?.loginHint &&
    account.username.toLowerCase() !==
      options.loginHint.toLowerCase()
  ) {
    console.warn(
      "[POWERBI] Login hint diferente da conta MSAL ativa.",
      {
        loginHint:
          options.loginHint,

        account:
          account.username,
      },
    );
  }

  /*
   * O Dashboard não abre popup.
   *
   * Primeiro tenta sempre o cache/token silencioso.
   */
  try {
    const result =
      await msal.acquireTokenSilent({
        scopes:
          POWERBI_SCOPES,

        account,

        /*
         * Não força refresh.
         * O MSAL pode utilizar o token em cache
         * ou renovar silenciosamente.
         */
        forceRefresh: false,
      });

    if (!result.accessToken) {
      throw new Error(
        "O Microsoft Entra ID não retornou um Access Token do Power BI.",
      );
    }

    console.info(
      "[POWERBI] Access Token obtido silenciosamente.",
    );

    console.info(
      "[POWERBI] Conta:",
      result.account?.username ??
        account.username,
    );

    console.info(
      "[POWERBI] Expira em:",
      result.expiresOn,
    );

    return wrapPowerBIToken(
      result,
    );
  } catch (error: any) {
    console.error(
      "[POWERBI] acquireTokenSilent falhou:",
      error,
    );

    const errorCode =
      error?.errorCode ??
      error?.code ??
      "";

    /*
     * Esses erros significam que o usuário precisa
     * de interação para autorizar o recurso Power BI.
     */
    if (
      errorCode ===
        "interaction_required" ||
      errorCode ===
        "consent_required" ||
      errorCode ===
        "login_required" ||
      errorCode ===
        "no_tokens_found"
    ) {
      throw new Error(
        "entra_powerbi_interaction_required",
      );
    }

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
 * Token utilizado pelos dashboards.
 *
 * O Dashboard:
 *
 * 1. Não faz login.
 * 2. Não abre popup.
 * 3. Não usa ssoSilent.
 * 4. Reutiliza a conta Microsoft autenticada no Portal.
 */
export function usePowerBIToken(
  enabled: boolean,
  loginHint?: string,
) {
  const {
    data: config,
  } = useEntraConfig();

  return useQuery({
    queryKey:
      TOKEN_KEY(config),

    enabled:
      Boolean(enabled) &&
      Boolean(config),

    retry: false,

    staleTime:
      45 * 60_000,

    refetchInterval:
      40 * 60_000,

    refetchIntervalInBackground:
      true,

    refetchOnWindowFocus:
      true,

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

          /*
           * Explicitamente não interativo.
           */
          interactive: false,
        },
      );
    },
  });
}

/* -------------------------------------------------------------------------- */
/* PREWARM                                                                    */
/* -------------------------------------------------------------------------- */

/**
 * Tenta obter antecipadamente o token do Power BI.
 *
 * Nunca abre popup.
 *
 * Nunca inicia login.
 */
export function usePowerBIPrewarm(
  loginHint?: string,
) {
  const {
    data: config,
  } = useEntraConfig();

  const queryClient =
    useQueryClient();

  useEffect(() => {
    if (!config) {
      return;
    }

    let cancelled = false;

    void (async () => {
      try {
        console.info(
          "[POWERBI] Prewarm iniciado.",
        );

        const token =
          await getPowerBIToken(
            config,
            {
              loginHint,

              interactive:
                false,
            },
          );

        if (cancelled) {
          return;
        }

        queryClient.setQueryData(
          TOKEN_KEY(config),
          token,
        );

        console.info(
          "[POWERBI] Prewarm concluído.",
        );
      } catch (error: any) {
        if (cancelled) {
          return;
        }

        const message =
          error?.message ?? "";

        if (
          message ===
          "entra_login_required"
        ) {
          console.info(
            "[POWERBI] Prewarm aguardando login Microsoft.",
          );

          return;
        }

        if (
          message ===
          "entra_powerbi_interaction_required"
        ) {
          console.info(
            "[POWERBI] Prewarm aguardando autorização do Power BI.",
          );

          return;
        }

        console.warn(
          "[POWERBI] Prewarm silencioso não disponível:",
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

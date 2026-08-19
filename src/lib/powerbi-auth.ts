import { useEffect } from "react";
import {
  useQuery,
  useQueryClient,
} from "@tanstack/react-query";
import { supabase } from "@/integrations/supabase/client";

/* -------------------------------------------------------------------------- */
/* CONFIGURAÇÃO                                                               */
/* -------------------------------------------------------------------------- */

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
/* CONFIGURAÇÃO ENTRA                                                         */
/* -------------------------------------------------------------------------- */

export function useEntraConfig() {
  return useQuery({
    queryKey: ["entra-config"],

    staleTime: 5 * 60_000,

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

let msalPromise:
  | Promise<
      import("@azure/msal-browser").PublicClientApplication
    >
  | null = null;

let msalKey = "";

/**
 * Cria ou reutiliza a mesma instância MSAL durante toda a sessão
 * do Portal.
 *
 * IMPORTANTE:
 * O login Microsoft acontece no /auth.
 *
 * O Dashboard NÃO deve criar uma nova sessão.
 */
export async function getMsal(
  config: EntraConfig,
) {
  const key =
    `${config.clientId}:${config.tenantId}`;

  if (!msalPromise || msalKey !== key) {
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
              window.location.origin,

            postLogoutRedirectUri:
              window.location.origin,
          },

          cache: {
            cacheLocation:
              "localStorage",

            storeAuthStateInCookie:
              false,
          },
        });

      await instance.initialize();

      /*
       * Processa eventual retorno de autenticação.
       */
      try {
        await instance.handleRedirectPromise();
      } catch (error) {
        console.warn(
          "[ENTRA] handleRedirectPromise:",
          error,
        );
      }

      /*
       * Recupera a conta Microsoft que já entrou no Portal.
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

      if (accounts.length > 0) {
        const active =
          instance.getActiveAccount();

        if (!active) {
          instance.setActiveAccount(
            accounts[0],
          );

          console.info(
            "[ENTRA] Conta MSAL definida como ativa:",
            accounts[0].username,
          );
        }
      }

      return instance;
    })();
  }

  return msalPromise;
}

/* -------------------------------------------------------------------------- */
/* LOGIN MICROSOFT                                                            */
/* -------------------------------------------------------------------------- */

/**
 * Login principal do Portal.
 *
 * Este login acontece somente na tela /auth.
 *
 * Depois que o usuário entra:
 *
 * Microsoft Entra
 *      ↓
 * MSAL mantém a conta
 *      ↓
 * Supabase cria a sessão do Portal
 *
 * O Dashboard reutiliza essa mesma sessão.
 */
export async function loginWithEntra(
  config: EntraConfig,
) {
  console.info(
    "[ENTRA] Iniciando login Microsoft...",
  );

  const msal =
    await getMsal(config);

  const result =
    await msal.loginPopup({
      scopes: [
        "openid",
        "profile",
        "email",
        "User.Read",
      ],

      prompt:
        "select_account",
    });

  if (!result.account) {
    throw new Error(
      "O Microsoft Entra ID não retornou uma conta.",
    );
  }

  /*
   * MUITO IMPORTANTE:
   * mantém exatamente a mesma conta que acabou
   * de fazer login como conta ativa do MSAL.
   */
  msal.setActiveAccount(
    result.account,
  );

  console.info(
    "[ENTRA] Login Microsoft concluído:",
    result.account.username,
  );

  return {
    idToken:
      result.idToken,

    account:
      result.account,
  };
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
 * IMPORTANTE:
 *
 * NÃO faz login.
 * NÃO abre popup.
 * NÃO usa ssoSilent.
 *
 * O usuário já deve estar autenticado no Portal.
 *
 * Fluxo:
 *
 * sessão Microsoft existente
 *          ↓
 * getActiveAccount()
 *          ↓
 * acquireTokenSilent()
 *          ↓
 * Access Token Power BI
 */
export async function getPowerBIToken(
  config: EntraConfig,
  options?: {
    loginHint?: string;
    interactive?: boolean;
  },
): Promise<PowerBIToken> {
  console.info(
    "[POWERBI] Obtendo Access Token silenciosamente...",
  );

  const msal =
    await getMsal(config);

  /*
   * Primeiro tenta a conta ativa.
   */
  let account =
    msal.getActiveAccount();

  /*
   * Se não houver conta ativa, procura no cache.
   */
  if (!account) {
    const accounts =
      msal.getAllAccounts();

    account =
      accounts[0] ?? null;

    if (account) {
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
   * O Dashboard não deve fazer login.
   */
  if (!account) {
    console.warn(
      "[POWERBI] Nenhuma sessão Microsoft disponível.",
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
   * O loginHint é apenas diagnóstico neste ponto.
   * Não abrimos popup com ele.
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
   * ÚNICA tentativa para o Dashboard:
   * acquireTokenSilent.
   */
  try {
    const result =
      await msal.acquireTokenSilent({
        scopes:
          POWERBI_SCOPES,

        account,
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

    /*
     * Não abrimos popup aqui.
     *
     * Se houver necessidade de consentimento/interação,
     * o fluxo deverá ser resolvido no login do Portal.
     */
    const errorCode =
      error?.errorCode ??
      error?.code ??
      "";

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
 * Token Power BI utilizado pelo Dashboard.
 *
 * IMPORTANTE:
 * interactive NÃO é utilizado aqui.
 *
 * O Dashboard apenas reutiliza a sessão Microsoft
 * criada no /auth.
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
           * O Dashboard nunca abre popup.
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
 * Pré-aquece o token do Power BI.
 *
 * O Prewarm:
 *
 * - NÃO faz login;
 * - NÃO abre popup;
 * - NÃO usa ssoSilent;
 * - apenas tenta aproveitar uma sessão Microsoft
 *   que já existe.
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

        /*
         * interactive=false:
         * nunca abre popup.
         */
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
        /*
         * É normal o Prewarm não conseguir token
         * se o consentimento ainda não tiver sido
         * realizado.
         *
         * Não tratamos isso como erro fatal.
         */
        if (
          error?.message ===
          "entra_login_required"
        ) {
          console.info(
            "[POWERBI] Prewarm aguardando sessão Microsoft.",
          );
        } else if (
          error?.message ===
          "entra_powerbi_interaction_required"
        ) {
          console.info(
            "[POWERBI] Prewarm aguardando consentimento do Power BI.",
          );
        } else {
          console.warn(
            "[POWERBI] Prewarm silencioso não disponível.",
            error,
          );
        }
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

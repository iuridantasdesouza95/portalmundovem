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

  console.log(
    "[ENTRA] Configuração encontrada:",
    {
      hasClientId: Boolean(clientId),
      hasTenantId: Boolean(tenantId),
    },
  );

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

/*
 * Protege contra duas chamadas simultâneas de loginPopup().
 *
 * Esse é justamente um dos problemas que gera:
 *
 * interaction_in_progress
 */
let loginPromise:
  | Promise<
      import("@azure/msal-browser").AuthenticationResult
    >
  | null = null;

/* -------------------------------------------------------------------------- */
/* GET MSAL                                                                    */
/* -------------------------------------------------------------------------- */

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
              `${window.location.origin}/auth`,

            postLogoutRedirectUri:
              `${window.location.origin}/auth`,
          },

          cache: {
            cacheLocation: "localStorage",

            storeAuthStateInCookie: true,
          },

          system: {
            allowNativeBroker: false,
          },
        });

      await instance.initialize();

      /*
       * Muito importante:
       *
       * processa qualquer resposta pendente do MSAL
       * antes de tentar uma nova interação.
       */
      try {
        const redirectResponse =
          await instance.handleRedirectPromise();

        if (redirectResponse?.account) {
          instance.setActiveAccount(
            redirectResponse.account,
          );

          console.info(
            "[ENTRA] Conta recuperada do redirect:",
            redirectResponse.account.username,
          );
        }
      } catch (error) {
        console.warn(
          "[ENTRA] Não foi possível processar redirect pendente:",
          error,
        );
      }

      const accounts =
        instance.getAllAccounts();

      console.info(
        "[ENTRA] Contas MSAL disponíveis:",
        accounts.map(
          (account) => account.username,
        ),
      );

      if (
        accounts.length > 0 &&
        !instance.getActiveAccount()
      ) {
        instance.setActiveAccount(
          accounts[0],
        );

        console.info(
          "[ENTRA] Conta MSAL definida como ativa:",
          accounts[0].username,
        );
      }

      console.info(
        "[ENTRA] MSAL inicializado.",
      );

      return instance;
    })();
  }

  return msalPromise;
}

/* -------------------------------------------------------------------------- */
/* LOGIN MICROSOFT                                                            */
/* -------------------------------------------------------------------------- */

export async function loginWithEntra(
  config: EntraConfig,
) {
  /*
   * Se já existe um login em andamento,
   * reutilizamos a mesma Promise.
   *
   * Isso elimina o interaction_in_progress causado
   * por chamadas duplicadas.
   */
  if (loginPromise) {
    console.info(
      "[ENTRA] Login Microsoft já está em andamento. Aguardando...",
    );

    return loginPromise;
  }

  loginPromise = (async () => {
    const msal =
      await getMsal(config);

    const activeAccount =
      msal.getActiveAccount();

    /*
     * Se já existe uma conta Microsoft válida,
     * não abrimos outro popup.
     */
    if (activeAccount) {
      console.info(
        "[ENTRA] Sessão Microsoft já existente:",
        activeAccount.username,
      );

      return {
        idToken: "",
        account: activeAccount,
      };
    }

    console.log(
      "[ENTRA] Abrindo loginPopup...",
    );

    const response =
      await msal.loginPopup({
        /*
         * O Power BI já é solicitado durante o
         * login Microsoft.
         *
         * Assim o token fica disponível no cache
         * para o dashboard.
         */
        scopes: [
          "openid",
          "profile",
          "email",
          "User.Read",
          ...POWERBI_SCOPES,
        ],

        prompt: "select_account",

        redirectUri:
          `${window.location.origin}/auth`,
      });

    if (!response) {
      throw new Error(
        "A Microsoft não retornou uma resposta de autenticação.",
      );
    }

    if (!response.account) {
      throw new Error(
        "A Microsoft autenticou o usuário, mas nenhuma conta foi retornada.",
      );
    }

    if (!response.idToken) {
      throw new Error(
        "O Microsoft Entra ID não retornou o token de identidade.",
      );
    }

    msal.setActiveAccount(
      response.account,
    );

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
    /*
     * Libera o lock somente depois que
     * toda a interação terminou.
     */
    loginPromise = null;
  }
}

/* -------------------------------------------------------------------------- */
/* TOKEN                                                                       */
/* -------------------------------------------------------------------------- */

function wrapPowerBIToken(
  result: {
    accessToken: string;
    expiresOn: Date | null;
  },
): PowerBIToken {
  return {
    accessToken:
      result.accessToken,

    expiresOn:
      result.expiresOn
        ? result.expiresOn.getTime()
        : Date.now() + 55 * 60_000,
  };
}

/* -------------------------------------------------------------------------- */
/* TOKEN POWER BI                                                             */
/* -------------------------------------------------------------------------- */

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
   * Sem conta Microsoft não existe token
   * delegado do Power BI.
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

  try {
    const result =
      await msal.acquireTokenSilent({
        scopes:
          POWERBI_SCOPES,

        account,

        redirectUri:
          `${window.location.origin}/auth`,
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

/* -------------------------------------------------------------------------- */
/* HOOK TOKEN                                                                 */
/* -------------------------------------------------------------------------- */

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
          interactive: false,
        },
      );
    },
  });
}

/* -------------------------------------------------------------------------- */
/* PREWARM                                                                    */
/* -------------------------------------------------------------------------- */

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
              interactive: false,
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

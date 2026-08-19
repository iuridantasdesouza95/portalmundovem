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
 *
 * O redirect utilizado pelo MSAL é uma página dedicada,
 * evitando que o roteador principal do Portal interfira
 * no retorno da autenticação.
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

      const redirectUri =
        `${window.location.origin}/auth-redirect.html`;

      const instance =
        new PublicClientApplication({
          auth: {
            clientId:
              config.clientId,

            authority:
              `https://login.microsoftonline.com/${config.tenantId}`,

            redirectUri,

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

      try {
        await instance.handleRedirectPromise();
      } catch (error) {
        console.warn(
          "[ENTRA] handleRedirectPromise:",
          error,
        );
      }

      const accounts =
        instance.getAllAccounts();

      console.log(
        "[ENTRA] Contas encontradas no cache:",
        accounts.map(
          (account) =>
            account.username,
        ),
      );

      if (accounts.length > 0) {
        instance.setActiveAccount(
          accounts[0],
        );
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
 * Esta função deve ser chamada por uma ação explícita
 * do usuário, como o botão "Entrar com Microsoft".
 */
export async function loginWithEntra(
  config: EntraConfig,
) {
  console.log(
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

  msal.setActiveAccount(
    result.account,
  );

  console.log(
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
/* TOKEN                                                                      */
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

/* -------------------------------------------------------------------------- */
/* TOKEN POWER BI                                                             */
/* -------------------------------------------------------------------------- */

/**
 * Obtém o Access Token delegado da API do Power BI.
 *
 * Fluxo:
 *
 * 1. Procura uma conta Microsoft já autenticada.
 * 2. Tenta acquireTokenSilent.
 * 3. Se não houver conta, informa que o login é necessário.
 * 4. Se o token exigir interação:
 *      - interactive=false → não abre popup
 *      - interactive=true  → permite acquireTokenPopup
 *
 * IMPORTANTE:
 *
 * O Dashboard não deve iniciar login Microsoft automaticamente.
 * O login deve acontecer pelo fluxo de autenticação do Portal.
 */
export async function getPowerBIToken(
  config: EntraConfig,
  options?: {
    loginHint?: string;
    interactive?: boolean;
  },
): Promise<PowerBIToken> {
  console.log(
    "[POWERBI] Obtendo Access Token...",
  );

  const msal =
    await getMsal(config);

  /* ---------------------------------------------------------------------- */
  /* LOCALIZA CONTA                                                         */
  /* ---------------------------------------------------------------------- */

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
    }
  }

  /* ---------------------------------------------------------------------- */
  /* SEM CONTA                                                              */
  /* ---------------------------------------------------------------------- */

  if (!account) {
    console.warn(
      "[POWERBI] Nenhuma sessão Microsoft encontrada.",
    );

    throw new Error(
      "entra_login_required",
    );
  }

  console.log(
    "[POWERBI] Conta Microsoft encontrada:",
    account.username,
  );

  /* ---------------------------------------------------------------------- */
  /* TOKEN SILENCIOSO                                                       */
  /* ---------------------------------------------------------------------- */

  try {
    console.log(
      "[POWERBI] Tentando acquireTokenSilent...",
    );

    const result =
      await msal.acquireTokenSilent({
        scopes:
          POWERBI_SCOPES,

        account,

        ...(options?.loginHint
          ? {
              loginHint:
                options.loginHint,
            }
          : {}),
      });

    console.log(
      "[POWERBI] Access Token Power BI obtido silenciosamente.",
    );

    console.log(
      "[POWERBI] Conta:",
      result.account?.username,
    );

    console.log(
      "[POWERBI] Expira em:",
      result.expiresOn,
    );

    return wrapPowerBIToken(
      result,
    );
  } catch (error) {
    console.warn(
      "[POWERBI] acquireTokenSilent não conseguiu obter o token:",
      error,
    );
  }

  /* ---------------------------------------------------------------------- */
  /* PREWARM / NÃO INTERATIVO                                               */
  /* ---------------------------------------------------------------------- */

  if (
    options?.interactive === false
  ) {
    console.info(
      "[POWERBI] Token Power BI requer interação.",
    );

    throw new Error(
      "entra_interaction_required",
    );
  }

  /* ---------------------------------------------------------------------- */
  /* INTERAÇÃO                                                              */
  /* ---------------------------------------------------------------------- */

  console.log(
    "[POWERBI] Token Power BI requer autenticação interativa.",
  );

  try {
    const result =
      await msal.acquireTokenPopup({
        scopes:
          POWERBI_SCOPES,

        account,

        ...(options?.loginHint
          ? {
              loginHint:
                options.loginHint,
            }
          : {}),
      });

    if (result.account) {
      msal.setActiveAccount(
        result.account,
      );
    }

    console.log(
      "[POWERBI] Access Token Power BI obtido via popup.",
    );

    console.log(
      "[POWERBI] Conta:",
      result.account?.username,
    );

    console.log(
      "[POWERBI] Expira em:",
      result.expiresOn,
    );

    return wrapPowerBIToken(
      result,
    );
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
 *
 * IMPORTANTE:
 * O hook não deve iniciar o login Microsoft.
 *
 * Se não houver sessão ou consentimento,
 * o erro será retornado para a interface.
 */
export function usePowerBIToken(
  enabled: boolean,
  loginHint?: string,
) {
  const {
    data: config,
  } =
    useEntraConfig();

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
           * IMPORTANTE:
           *
           * O Dashboard não deve abrir
           * popup automaticamente.
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
 * Tenta obter silenciosamente o token Power BI
 * depois que o usuário já estiver autenticado.
 *
 * NUNCA abre popup.
 */
export function usePowerBIPrewarm(
  loginHint?: string,
) {
  const {
    data: config,
  } =
    useEntraConfig();

  const queryClient =
    useQueryClient();

  useEffect(() => {
    if (!config) {
      return;
    }

    let cancelled =
      false;

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

              /*
               * Prewarm nunca pode abrir
               * uma janela de autenticação.
               */
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
        /*
         * Isso NÃO deve ser tratado como erro
         * fatal do Portal.
         *
         * Significa apenas que ainda não
         * existe token Power BI silencioso.
         */
        console.info(
          "[POWERBI] Prewarm silencioso não disponível.",
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

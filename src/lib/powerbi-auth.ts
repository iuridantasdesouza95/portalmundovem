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

export async function loginWithEntra(
  config: EntraConfig,
) {
  const msal = await getMsal(config);

  const result =
    await msal.loginPopup({
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

  msal.setActiveAccount(
    result.account,
  );

  return {
    idToken: result.idToken,
    account: result.account,
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

/**
 * Obtém o Access Token delegado da API do Power BI.
 *
 * Fluxo:
 *
 * 1. acquireTokenSilent
 * 2. se interação for necessária:
 *    acquireTokenPopup
 *
 * Não usamos ssoSilent porque ele pode ficar aguardando
 * a sessão Microsoft em iframe e resultar em timeout.
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

  let account =
    msal.getActiveAccount();

  if (!account) {
    const accounts =
      msal.getAllAccounts();

    account = accounts[0] ?? null;

    if (account) {
      msal.setActiveAccount(
        account,
      );
    }
  }

  /*
   * 1. Tenta silenciosamente.
   */
  if (account) {
    try {
      console.log(
        "[POWERBI] Tentando acquireTokenSilent...",
      );

      const result =
        await msal.acquireTokenSilent({
          scopes:
            POWERBI_SCOPES,

          account,
        });

      console.log(
        "[POWERBI] Access Token Power BI obtido silenciosamente.",
      );

      return wrapPowerBIToken(
        result,
      );
    } catch (error) {
      console.warn(
        "[POWERBI] Token silencioso não disponível:",
        error,
      );
    }
  }

  /*
   * 2. Se não houver conta, precisamos
   * autenticar o usuário.
   */
  if (!account) {
    if (
      options?.interactive === false
    ) {
      throw new Error(
        "entra_login_required",
      );
    }

    console.log(
      "[POWERBI] Nenhuma sessão Microsoft encontrada. Abrindo login...",
    );

    const login =
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

    if (!login.account) {
      throw new Error(
        "Não foi possível identificar a conta Microsoft.",
      );
    }

    account =
      login.account;

    msal.setActiveAccount(
      account,
    );
  }

  /*
   * 3. Token interativo do Power BI.
   */
  if (
    options?.interactive === false
  ) {
    throw new Error(
      "entra_interaction_required",
    );
  }

  console.log(
    "[POWERBI] Solicitando permissão para a API do Power BI...",
  );

  const result =
    await msal.acquireTokenPopup({
      scopes:
        POWERBI_SCOPES,

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
    "[POWERBI] Access Token Power BI obtido.",
  );

  return wrapPowerBIToken(
    result,
  );
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
          interactive: true,
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
        console.log(
          "[POWERBI] Prewarm iniciado.",
        );

        /*
         * Prewarm NÃO abre popup.
         * Só aproveita uma sessão/token existente.
         */
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
        }
      } catch (error) {
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

import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { useEffect, useState } from "react";
import { toast } from "sonner";

import { supabase } from "@/integrations/supabase/client";
import {
  fetchEntraConfig,
  getMsal,
} from "@/lib/powerbi-auth";
import { entraSignIn as entraSignInFn } from "@/lib/entra-session.functions";
import { lovable } from "@/integrations/lovable/index";

import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";

export const Route = createFileRoute("/auth")({
  ssr: false,

  head: () => ({
    meta: [
      {
        title: "Entrar — Portal Corporativo BI Vem",
      },
      {
        name: "description",
        content:
          "Acesse o Portal Corporativo de BI da Vem para consultar os dashboards oficiais do Power BI.",
      },
      {
        property: "og:title",
        content: "Entrar — Portal Corporativo BI Vem",
      },
      {
        property: "og:description",
        content:
          "Autenticação do Portal Corporativo de BI da Vem.",
      },
      {
        property: "og:type",
        content: "website",
      },
      {
        name: "twitter:card",
        content: "summary_large_image",
      },
      {
        name: "robots",
        content: "noindex",
      },
    ],
  }),

  component: AuthPage,
});

function AuthPage() {
  const navigate = useNavigate();

  const [mode, setMode] =
    useState<"login" | "signup">("login");

  const [showEmail, setShowEmail] =
    useState(false);

  const [email, setEmail] =
    useState("");

  const [password, setPassword] =
    useState("");

  const [name, setName] =
    useState("");

  const [loading, setLoading] =
    useState(false);

  const [msLoading, setMsLoading] =
    useState(false);

  /* ---------------------------------------------------------------------- */
  /* INICIALIZAÇÃO DA AUTENTICAÇÃO                                         */
  /* ---------------------------------------------------------------------- */

  useEffect(() => {
    let mounted = true;

    async function initializeAuth() {
      try {
        console.log(
          "[AUTH] ========================================",
        );

        console.log(
          "[AUTH] Inicializando autenticação...",
        );

        /* ---------------------------------------------------------------- */
        /* 1. VERIFICAR SESSÃO SUPABASE                                     */
        /* ---------------------------------------------------------------- */

        const {
          data: sessionData,
          error: sessionError,
        } =
          await supabase.auth.getSession();

        if (sessionError) {
          console.error(
            "[AUTH] Erro ao verificar sessão Supabase:",
            sessionError,
          );
        }

        if (sessionData.session) {
          console.log(
            "[AUTH] Sessão Supabase existente encontrada.",
          );

          if (mounted) {
            navigate({
              to: "/inicio",
              replace: true,
            });
          }

          return;
        }

        console.log(
          "[AUTH] Nenhuma sessão Supabase encontrada.",
        );

        /* ---------------------------------------------------------------- */
        /* 2. BUSCAR CONFIGURAÇÃO ENTRA                                    */
        /* ---------------------------------------------------------------- */

        console.log(
          "[AUTH] Verificando configuração Microsoft Entra...",
        );

        const config =
          await fetchEntraConfig();

        if (!config) {
          console.warn(
            "[AUTH] Microsoft Entra ID não configurado.",
          );

          return;
        }

        console.log(
          "[AUTH] Configuração Entra encontrada.",
          {
            clientId:
              `${config.clientId.slice(0, 8)}...`,

            tenantId:
              `${config.tenantId.slice(0, 8)}...`,
          },
        );

        /* ---------------------------------------------------------------- */
        /* 3. INICIALIZAR MSAL                                             */
        /* ---------------------------------------------------------------- */

        const msal =
          await getMsal(config);

        console.log(
          "[AUTH] MSAL inicializado.",
        );

        /*
         * IMPORTANTE:
         *
         * Não usamos loginRedirect().
         *
         * O Portal pode estar sendo executado dentro
         * de um iframe do Lovable.
         *
         * Portanto, o fluxo Microsoft é iniciado
         * exclusivamente pelo loginPopup().
         */

        const accounts =
          msal.getAllAccounts();

        console.log(
          "[AUTH] Contas Microsoft disponíveis:",
          accounts.map(
            (account) =>
              account.username,
          ),
        );

        /*
         * Se já existir uma conta Microsoft no cache,
         * podemos aproveitar essa autenticação.
         */

        const account =
          msal.getActiveAccount() ??
          accounts[0] ??
          null;

        if (!account) {
          console.log(
            "[AUTH] Nenhuma conta Microsoft após inicialização.",
          );

          return;
        }

        console.log(
          "[AUTH] Conta Microsoft encontrada:",
          {
            username:
              account.username,

            name:
              account.name,
          },
        );

        msal.setActiveAccount(
          account,
        );

        /* ---------------------------------------------------------------- */
        /* 4. OBTER TOKEN DA MICROSOFT                                    */
        /* ---------------------------------------------------------------- */

        console.log(
          "[AUTH] Obtendo token Microsoft silenciosamente...",
        );

        let tokenResult;

        try {
          tokenResult =
            await msal.acquireTokenSilent({
              scopes: [
                "openid",
                "profile",
                "email",
                "User.Read",
              ],

              account,
            });
        } catch (silentError) {
          console.warn(
            "[AUTH] Não foi possível obter token silenciosamente.",
            silentError,
          );

          /*
           * Não abrimos popup automaticamente durante
           * a inicialização da página.
           *
           * O usuário poderá clicar novamente no botão
           * "Entrar com Microsoft".
           */

          return;
        }

        if (!tokenResult.idToken) {
          console.warn(
            "[AUTH] Token Microsoft não possui idToken.",
          );

          return;
        }

        console.log(
          "[AUTH] Token Microsoft obtido.",
        );

        /* ---------------------------------------------------------------- */
        /* 5. CONVERTER MICROSOFT EM SESSÃO DO PORTAL                     */
        /* ---------------------------------------------------------------- */

        console.log(
          "[AUTH] Criando sessão do portal...",
        );

        const {
          email: portalEmail,
          tokenHash,
        } =
          await entraSignInFn({
            data: {
              idToken:
                tokenResult.idToken,
            },
          });

        if (!portalEmail) {
          throw new Error(
            "O servidor não retornou o e-mail da conta Microsoft.",
          );
        }

        if (!tokenHash) {
          throw new Error(
            "O servidor não retornou o token necessário para criar a sessão.",
          );
        }

        console.log(
          "[AUTH] Identidade Microsoft validada:",
          portalEmail,
        );

        /* ---------------------------------------------------------------- */
        /* 6. CRIAR SESSÃO SUPABASE                                       */
        /* ---------------------------------------------------------------- */

        const {
          error: verifyError,
        } =
          await supabase.auth.verifyOtp({
            email:
              portalEmail,

            token_hash:
              tokenHash,

            type:
              "email",
          });

        if (verifyError) {
          throw verifyError;
        }

        console.log(
          "[AUTH] Sessão Supabase criada.",
        );

        /* ---------------------------------------------------------------- */
        /* 7. CONFIRMAR SESSÃO                                            */
        /* ---------------------------------------------------------------- */

        const {
          data:
            confirmedSession,
        } =
          await supabase.auth.getSession();

        if (
          !confirmedSession.session
        ) {
          throw new Error(
            "O login Microsoft foi concluído, mas a sessão do portal não foi criada.",
          );
        }

        console.log(
          "[AUTH] Sessão do portal confirmada.",
        );

        console.log(
          "[AUTH] Redirecionando para /inicio...",
        );

        console.log(
          "[AUTH] ========================================",
        );

        if (mounted) {
          navigate({
            to: "/inicio",
            replace: true,
          });
        }
      } catch (error) {
        console.error(
          "[AUTH] Erro ao inicializar autenticação:",
          error,
        );

        if (mounted) {
          toast.error(
            error instanceof Error
              ? error.message
              : "Não foi possível autenticar com a Microsoft.",
          );
        }
      }
    }

    void initializeAuth();

    return () => {
      mounted = false;
    };
  }, [navigate]);

  /* ---------------------------------------------------------------------- */
  /* LOGIN / CADASTRO POR E-MAIL                                            */
  /* ---------------------------------------------------------------------- */

  async function submit(
    e: React.FormEvent<HTMLFormElement>,
  ) {
    e.preventDefault();

    if (loading || msLoading) {
      return;
    }

    setLoading(true);

    try {
      /* ---------------------------------------------------------------- */
      /* LOGIN                                                             */
      /* ---------------------------------------------------------------- */

      if (mode === "login") {
        console.log(
          "[AUTH] Iniciando login por e-mail...",
        );

        const {
          error,
        } =
          await supabase.auth.signInWithPassword({
            email,
            password,
          });

        if (error) {
          throw error;
        }

        console.log(
          "[AUTH] Login por e-mail concluído.",
        );

        navigate({
          to: "/inicio",
          replace: true,
        });

        return;
      }

      /* ---------------------------------------------------------------- */
      /* CADASTRO                                                          */
      /* ---------------------------------------------------------------- */

      console.log(
        "[AUTH] Criando conta por e-mail...",
      );

      const {
        data,
        error,
      } =
        await supabase.auth.signUp({
          email,
          password,

          options: {
            emailRedirectTo:
              window.location.origin,

            data: {
              full_name:
                name,
            },
          },
        });

      if (error) {
        throw error;
      }

      if (data.session) {
        console.log(
          "[AUTH] Conta criada e sessão iniciada.",
        );

        navigate({
          to: "/inicio",
          replace: true,
        });
      } else {
        toast.success(
          "Conta criada. Confirme o e-mail para acessar o portal.",
        );
      }
    } catch (error) {
      console.error(
        "[AUTH] Erro no login/cadastro:",
        error,
      );

      toast.error(
        error instanceof Error
          ? error.message
          : "Não foi possível autenticar.",
      );
    } finally {
      setLoading(false);
    }
  }

  /* ---------------------------------------------------------------------- */
  /* LOGIN MICROSOFT ENTRA ID                                              */
  /* ---------------------------------------------------------------------- */

  async function microsoft() {
    if (msLoading || loading) {
      return;
    }

    setMsLoading(true);

    try {
      console.log(
        "[AUTH] ========================================",
      );

      console.log(
        "[AUTH] Iniciando login Microsoft...",
      );

      /* ---------------------------------------------------------------- */
      /* 1. CONFIGURAÇÃO ENTRA                                            */
      /* ---------------------------------------------------------------- */

      console.log(
        "[AUTH] Buscando configuração Entra...",
      );

      const config =
        await fetchEntraConfig();

      console.log(
        "[AUTH] Configuração Entra:",
        {
          hasConfig:
            Boolean(config),

          clientId:
            config?.clientId
              ? `${config.clientId.slice(
                  0,
                  8,
                )}...`
              : null,

          tenantId:
            config?.tenantId
              ? `${config.tenantId.slice(
                  0,
                  8,
                )}...`
              : null,
        },
      );

      if (!config) {
        throw new Error(
          "Microsoft Entra ID não está configurado no portal. Verifique azure_client_id e azure_tenant_id na tabela app_settings.",
        );
      }

      /* ---------------------------------------------------------------- */
      /* 2. INICIALIZAR MSAL                                              */
      /* ---------------------------------------------------------------- */

      const msal =
        await getMsal(config);

      console.log(
        "[AUTH] MSAL inicializado.",
      );

      /* ---------------------------------------------------------------- */
      /* 3. VERIFICAR CONTA EXISTENTE                                    */
      /* ---------------------------------------------------------------- */

      let account =
        msal.getActiveAccount() ??
        msal.getAllAccounts()[0] ??
        null;

      if (account) {
        console.log(
          "[AUTH] Conta Microsoft já disponível:",
          account.username,
        );

        msal.setActiveAccount(
          account,
        );

        /*
         * Tenta obter o token silenciosamente.
         */

        try {
          const tokenResult =
            await msal.acquireTokenSilent({
              scopes: [
                "openid",
                "profile",
                "email",
                "User.Read",
              ],

              account,
            });

          if (tokenResult.idToken) {
            console.log(
              "[AUTH] Token Microsoft reutilizado.",
            );

            await completeMicrosoftLogin(
              tokenResult.idToken,
            );

            return;
          }
        } catch (error) {
          console.warn(
            "[AUTH] Não foi possível reutilizar a sessão Microsoft. Será aberto o login.",
            error,
          );
        }
      }

      /* ---------------------------------------------------------------- */
      /* 4. LOGIN POPUP                                                  */
      /* ---------------------------------------------------------------- */

      console.log(
        "[AUTH] Nenhuma sessão Microsoft válida encontrada.",
      );

      console.log(
        "[AUTH] Abrindo loginPopup do Microsoft Entra ID...",
      );

      /*
       * IMPORTANTE:
       *
       * NÃO usar loginRedirect().
       *
       * O Portal pode estar dentro de iframe.
       *
       * loginPopup() funciona como uma janela de
       * autenticação iniciada diretamente pelo clique
       * do usuário.
       */

      const loginResult =
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

      console.log(
        "[AUTH] Popup Microsoft concluído.",
      );

      account =
        loginResult.account ??
        msal.getActiveAccount() ??
        msal.getAllAccounts()[0] ??
        null;

      if (!account) {
        throw new Error(
          "A Microsoft autenticou o usuário, mas nenhuma conta foi retornada.",
        );
      }

      msal.setActiveAccount(
        account,
      );

      console.log(
        "[AUTH] Conta Microsoft autenticada:",
        account.username,
      );

      /* ---------------------------------------------------------------- */
      /* 5. OBTER ID TOKEN                                               */
      /* ---------------------------------------------------------------- */

      let idToken =
        loginResult.idToken;

      /*
       * Se o loginResult não possuir idToken,
       * tenta adquirir silenciosamente.
       */

      if (!idToken) {
        console.log(
          "[AUTH] Obtendo idToken após login...",
        );

        const tokenResult =
          await msal.acquireTokenSilent({
            scopes: [
              "openid",
              "profile",
              "email",
              "User.Read",
            ],

            account,
          });

        idToken =
          tokenResult.idToken;
      }

      if (!idToken) {
        throw new Error(
          "O Microsoft Entra ID não retornou o token de identidade.",
        );
      }

      /* ---------------------------------------------------------------- */
      /* 6. FINALIZAR LOGIN DO PORTAL                                    */
      /* ---------------------------------------------------------------- */

      await completeMicrosoftLogin(
        idToken,
      );
    } catch (error: any) {
      console.error(
        "[AUTH] ========================================",
      );

      console.error(
        "[AUTH] Erro no login Microsoft:",
        error,
      );

      console.error(
        "[AUTH] Código:",
        error?.errorCode,
      );

      console.error(
        "[AUTH] ========================================",
      );

      const errorCode =
        error?.errorCode ??
        error?.code ??
        "";

      const message =
        error instanceof Error
          ? error.message
          : "Falha no login com Microsoft.";

      /*
       * Cancelamento normal do popup.
       */

      const cancelled =
        /user_cancelled|user_cancelled_error|popup_window_error|popup_window_closed/i.test(
          errorCode ||
            message,
        );

      if (!cancelled) {
        toast.error(
          message,
        );
      }
    } finally {
      setMsLoading(false);
    }
  }

  /* ---------------------------------------------------------------------- */
  /* FINALIZA LOGIN MICROSOFT                                              */
  /* ---------------------------------------------------------------------- */

  async function completeMicrosoftLogin(
    idToken: string,
  ) {
    console.log(
      "[AUTH] Criando sessão do portal...",
    );

    if (!idToken) {
      throw new Error(
        "Token de identidade Microsoft não informado.",
      );
    }

    /* ------------------------------------------------------------------ */
    /* ENVIAR TOKEN PARA O SERVIDOR                                       */
    /* ------------------------------------------------------------------ */

    const {
      email: portalEmail,
      tokenHash,
    } =
      await entraSignInFn({
        data: {
          idToken,
        },
      });

    if (!portalEmail) {
      throw new Error(
        "O servidor não retornou o e-mail da conta Microsoft.",
      );
    }

    if (!tokenHash) {
      throw new Error(
        "O servidor não retornou o token necessário para criar a sessão.",
      );
    }

    console.log(
      "[AUTH] Identidade Microsoft validada:",
      portalEmail,
    );

    /* ------------------------------------------------------------------ */
    /* CRIAR SESSÃO SUPABASE                                             */
    /* ------------------------------------------------------------------ */

    const {
      error: verifyError,
    } =
      await supabase.auth.verifyOtp({
        email:
          portalEmail,

        token_hash:
          tokenHash,

        type:
          "email",
      });

    if (verifyError) {
      throw verifyError;
    }

    console.log(
      "[AUTH] Sessão Supabase criada.",
    );

    /* ------------------------------------------------------------------ */
    /* CONFIRMAR SESSÃO                                                  */
    /* ------------------------------------------------------------------ */

    const {
      data:
        sessionData,
    } =
      await supabase.auth.getSession();

    if (!sessionData.session) {
      throw new Error(
        "O login Microsoft foi concluído, mas a sessão do portal não foi criada.",
      );
    }

    console.log(
      "[AUTH] Sessão do portal confirmada.",
    );

    console.log(
      "[AUTH] Login Microsoft concluído com sucesso.",
    );

    console.log(
      "[AUTH] Redirecionando para /inicio...",
    );

    console.log(
      "[AUTH] ========================================",
    );

    navigate({
      to: "/inicio",
      replace: true,
    });
  }

  /* ---------------------------------------------------------------------- */
  /* LOGIN GOOGLE                                                          */
  /* ---------------------------------------------------------------------- */

  async function google() {
    if (loading || msLoading) {
      return;
    }

    try {
      console.log(
        "[AUTH] Iniciando login Google...",
      );

      const result =
        await lovable.auth.signInWithOAuth(
          "google",
          {
            redirect_uri:
              window.location.origin,
          },
        );

      if (result.error) {
        toast.error(
          "Falha no login com Google.",
        );

        return;
      }

      if (result.redirected) {
        return;
      }

      navigate({
        to: "/inicio",
        replace: true,
      });
    } catch (error) {
      console.error(
        "[AUTH] Erro no login Google:",
        error,
      );

      toast.error(
        error instanceof Error
          ? error.message
          : "Falha no login com Google.",
      );
    }
  }

  /* ---------------------------------------------------------------------- */
  /* INTERFACE                                                             */
  /* ---------------------------------------------------------------------- */

  return (
    <div className="grid min-h-screen lg:grid-cols-[1.05fr_1fr]">

      {/* ---------------------------------------------------------------- */}
      {/* LADO ESQUERDO                                                    */}
      {/* ---------------------------------------------------------------- */}

      <div className="relative hidden flex-col justify-between overflow-hidden bg-primary p-12 lg:flex">

        <div className="absolute inset-0 opacity-[0.14] grid-fade" />

        <span className="relative font-display text-lg font-semibold text-primary-foreground">
          Vem · Portal BI
        </span>

        <div className="relative max-w-md">

          <h2 className="font-display text-4xl font-semibold leading-tight text-primary-foreground">
            Todos os indicadores da companhia em um só lugar.
          </h2>

          <p className="mt-4 text-sm leading-relaxed text-primary-foreground/80">
            Dashboards publicados no Power BI
            Service, organizados por área, com
            acesso controlado por perfil.
          </p>

        </div>

        <p className="relative text-xs text-primary-foreground/60">
          Ambiente corporativo · Acesso monitorado
        </p>

      </div>

      {/* ---------------------------------------------------------------- */}
      {/* LADO DIREITO                                                     */}
      {/* ---------------------------------------------------------------- */}

      <div className="flex items-center justify-center px-6 py-16">

        <div className="w-full max-w-sm">

          <h1 className="font-display text-2xl font-semibold">

            {mode === "login"
              ? "Acessar o portal"
              : "Criar acesso"}

          </h1>

          <p className="mt-1.5 text-sm text-muted-foreground">
            Entre com sua conta corporativa
            Microsoft — a mesma identidade é
            utilizada para acessar os dashboards
            do Power BI.
          </p>

          {/* ------------------------------------------------------------ */}
          {/* MICROSOFT                                                     */}
          {/* ------------------------------------------------------------ */}

          <Button
            className="mt-8 w-full"
            onClick={microsoft}
            disabled={
              msLoading ||
              loading
            }
          >
            {msLoading
              ? "Conectando à Microsoft…"
              : "Entrar com Microsoft"}
          </Button>

          <p className="mt-3 text-center text-[11px] text-muted-foreground">
            Login único com Microsoft Entra ID.
          </p>

          {/* ------------------------------------------------------------ */}
          {/* SEPARADOR                                                     */}
          {/* ------------------------------------------------------------ */}

          <div className="my-6 flex items-center gap-3 text-[11px] uppercase tracking-wider text-muted-foreground">

            <span className="h-px flex-1 bg-border" />

            ou

            <span className="h-px flex-1 bg-border" />

          </div>

          {/* ------------------------------------------------------------ */}
          {/* GOOGLE                                                        */}
          {/* ------------------------------------------------------------ */}

          <Button
            variant="outline"
            className="w-full"
            onClick={google}
            disabled={
              msLoading ||
              loading
            }
          >
            Continuar com Google
          </Button>

          {/* ------------------------------------------------------------ */}
          {/* E-MAIL                                                        */}
          {/* ------------------------------------------------------------ */}

          {!showEmail ? (
            <button
              type="button"
              onClick={() =>
                setShowEmail(true)
              }
              className="mt-6 w-full text-center text-sm text-muted-foreground underline-offset-4 hover:underline"
            >
              Entrar com e-mail e senha
            </button>
          ) : (
            <>

              <form
                onSubmit={submit}
                className="mt-6 space-y-4"
              >

                {mode === "signup" && (
                  <div className="space-y-1.5">

                    <Label htmlFor="name">
                      Nome completo
                    </Label>

                    <Input
                      id="name"
                      value={name}
                      onChange={(e) =>
                        setName(
                          e.target.value,
                        )
                      }
                      required
                    />

                  </div>
                )}

                <div className="space-y-1.5">

                  <Label htmlFor="email">
                    E-mail corporativo
                  </Label>

                  <Input
                    id="email"
                    type="email"
                    value={email}
                    onChange={(e) =>
                      setEmail(
                        e.target.value,
                      )
                    }
                    required
                  />

                </div>

                <div className="space-y-1.5">

                  <Label htmlFor="password">
                    Senha
                  </Label>

                  <Input
                    id="password"
                    type="password"
                    value={password}
                    onChange={(e) =>
                      setPassword(
                        e.target.value,
                      )
                    }
                    required
                    minLength={6}
                  />

                </div>

                <Button
                  type="submit"
                  variant="secondary"
                  className="w-full"
                  disabled={
                    loading ||
                    msLoading
                  }
                >
                  {loading
                    ? "Aguarde…"
                    : mode === "login"
                      ? "Entrar"
                      : "Criar acesso"}
                </Button>

              </form>

              <button
                type="button"
                onClick={() =>
                  setMode(
                    mode === "login"
                      ? "signup"
                      : "login",
                  )
                }
                className="mt-6 w-full text-center text-sm text-muted-foreground underline-offset-4 hover:underline"
              >
                {mode === "login"
                  ? "Não tem acesso? Criar conta"
                  : "Já tenho acesso"}
              </button>

            </>
          )}

        </div>

      </div>

    </div>
  );
}

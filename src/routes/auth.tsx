import {
  createFileRoute,
  useNavigate,
} from "@tanstack/react-router";

import {
  useEffect,
  useState,
  type FormEvent,
} from "react";

import { toast } from "sonner";

import { supabase } from "@/integrations/supabase/client";

import {
  fetchEntraConfig,
  loginWithEntra,
} from "@/lib/powerbi-auth";

import {
  entraSignIn as entraSignInFn,
} from "@/lib/entra-session.functions";

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
        content:
          "Entrar — Portal Corporativo BI Vem",
      },
      {
        property: "og:description",
        content:
          "Autenticação do Portal Corporativo BI da Vem.",
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

  /* ------------------------------------------------------------------ */
  /* INICIALIZAÇÃO                                                      */
  /* ------------------------------------------------------------------ */

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

        const {
          data,
          error,
        } = await supabase.auth.getSession();

        if (error) {
          console.error(
            "[AUTH] Erro ao verificar sessão Supabase:",
            error,
          );

          return;
        }

        if (!mounted) {
          return;
        }

        /*
         * Se já existe sessão do portal,
         * não mostramos a tela de login.
         */
        if (data.session) {
          console.log(
            "[AUTH] Sessão Supabase existente encontrada:",
            data.session.user.email,
          );

          await navigate({
            to: "/inicio",
            replace: true,
          });

          return;
        }

        console.log(
          "[AUTH] Nenhuma sessão Supabase encontrada.",
        );

        console.log(
          "[AUTH] Aguardando ação do usuário.",
        );

        console.log(
          "[AUTH] Login Microsoft será iniciado somente pelo botão.",
        );

        console.log(
          "[AUTH] ========================================",
        );
      } catch (error) {
        console.error(
          "[AUTH] Erro ao inicializar autenticação:",
          error,
        );

        if (!mounted) {
          return;
        }

        toast.error(
          error instanceof Error
            ? error.message
            : "Não foi possível verificar a sessão.",
        );
      }
    }

    void initializeAuth();

    return () => {
      mounted = false;
    };
  }, [navigate]);

  /* ------------------------------------------------------------------ */
  /* LOGIN / CADASTRO POR E-MAIL                                        */
  /* ------------------------------------------------------------------ */

  async function submit(
    e: FormEvent<HTMLFormElement>,
  ) {
    e.preventDefault();

    if (loading || msLoading) {
      return;
    }

    setLoading(true);

    try {
      /* -------------------------------------------------------------- */
      /* LOGIN                                                           */
      /* -------------------------------------------------------------- */

      if (mode === "login") {
        console.log(
          "[AUTH] Iniciando login por e-mail...",
        );

        const {
          error,
        } =
          await supabase.auth.signInWithPassword({
            email: email.trim(),
            password,
          });

        if (error) {
          throw error;
        }

        console.log(
          "[AUTH] Login por e-mail concluído.",
        );

        await navigate({
          to: "/inicio",
          replace: true,
        });

        return;
      }

      /* -------------------------------------------------------------- */
      /* CADASTRO                                                        */
      /* -------------------------------------------------------------- */

      console.log(
        "[AUTH] Criando conta por e-mail...",
      );

      const {
        data,
        error,
      } =
        await supabase.auth.signUp({
          email: email.trim(),
          password,

          options: {
            emailRedirectTo:
              window.location.origin,

            data: {
              full_name:
                name.trim(),
            },
          },
        });

      if (error) {
        throw error;
      }

      if (data.session) {
        await navigate({
          to: "/inicio",
          replace: true,
        });

        return;
      }

      toast.success(
        "Conta criada. Confirme o e-mail para acessar o portal.",
      );
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

  /* ------------------------------------------------------------------ */
  /* LOGIN MICROSOFT                                                    */
  /* ------------------------------------------------------------------ */

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

      /* -------------------------------------------------------------- */
      /* CONFIGURAÇÃO                                                    */
      /* -------------------------------------------------------------- */

      const config =
        await fetchEntraConfig();

      if (!config) {
        throw new Error(
          "Microsoft Entra ID não está configurado no portal.",
        );
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

      /* -------------------------------------------------------------- */
      /* LOGIN MICROSOFT                                                 */
      /* -------------------------------------------------------------- */

      /*
       * loginWithEntra() é responsável por:
       *
       * 1. Inicializar MSAL;
       * 2. Recuperar conta existente;
       * 3. Evitar loginPopup duplicado;
       * 4. Usar acquireTokenSilent quando possível;
       * 5. Abrir popup somente quando realmente necessário.
       */

      const loginResult =
        await loginWithEntra(config);

      if (!loginResult.account) {
        throw new Error(
          "A Microsoft autenticou o usuário, mas não retornou a conta.",
        );
      }

      console.log(
        "[AUTH] Conta Microsoft autenticada:",
        {
          username:
            loginResult.account.username,

          name:
            loginResult.account.name,
        },
      );

      /* -------------------------------------------------------------- */
      /* CRIAÇÃO DA SESSÃO DO PORTAL                                    */
      /* -------------------------------------------------------------- */

      /*
       * Mesmo que a conta Microsoft já estivesse
       * autenticada no navegador, ainda precisamos
       * garantir que o Supabase tenha uma sessão.
       *
       * O loginWithEntra() sempre tenta devolver
       * um ID Token válido para essa finalidade.
       */

      if (!loginResult.idToken) {
        throw new Error(
          "A Microsoft não retornou um ID Token válido para criar a sessão do portal.",
        );
      }

      console.log(
        "[AUTH] ID Token Microsoft disponível.",
      );

      await createPortalSession(
        loginResult.idToken,
      );
    } catch (error: any) {
      console.error(
        "[AUTH] ========================================",
      );

      console.error(
        "[AUTH] Erro no login Microsoft:",
        error,
      );

      const errorCode =
        error?.errorCode ??
        error?.code ??
        "";

      console.error(
        "[AUTH] Código:",
        errorCode ||
          "sem código",
      );

      console.error(
        "[AUTH] Mensagem:",
        error?.message ??
          "sem mensagem",
      );

      console.error(
        "[AUTH] ========================================",
      );

      /* -------------------------------------------------------------- */
      /* CANCELAMENTO                                                    */
      /* -------------------------------------------------------------- */

      const cancelledCodes = [
        "user_cancelled",
        "user_canceled",
        "popup_window_error",
      ];

      if (
        cancelledCodes.includes(
          errorCode,
        )
      ) {
        console.log(
          "[AUTH] Login Microsoft cancelado pelo usuário.",
        );

        return;
      }

      /* -------------------------------------------------------------- */
      /* LOGIN JÁ EM ANDAMENTO                                           */
      /* -------------------------------------------------------------- */

      if (
        errorCode ===
        "interaction_in_progress"
      ) {
        toast.error(
          "A autenticação Microsoft ainda está sendo processada. Aguarde alguns segundos e tente novamente.",
        );

        return;
      }

      /* -------------------------------------------------------------- */
      /* FALHA MICROSOFT                                                 */
      /* -------------------------------------------------------------- */

      toast.error(
        "Não foi possível entrar com a Microsoft. Se necessário, utilize o login por e-mail.",
      );
    } finally {
      setMsLoading(false);
    }
  }

  /* ------------------------------------------------------------------ */
  /* CRIAR SESSÃO DO PORTAL                                             */
  /* ------------------------------------------------------------------ */

  async function createPortalSession(
    idToken: string,
  ) {
    console.log(
      "[AUTH] Criando sessão do portal...",
    );

    /*
     * O ID Token Microsoft é enviado para
     * a Server Function.
     *
     * A validação real acontece no servidor.
     */

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

    /* -------------------------------------------------------------- */
    /* CRIAR SESSÃO SUPABASE                                            */
    /* -------------------------------------------------------------- */

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
      console.error(
        "[AUTH] Erro ao criar sessão Supabase:",
        verifyError,
      );

      throw verifyError;
    }

    console.log(
      "[AUTH] Sessão Supabase criada.",
    );

    /* -------------------------------------------------------------- */
    /* CONFIRMAR SESSÃO                                                 */
    /* -------------------------------------------------------------- */

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
      "[AUTH] Sessão Supabase confirmada:",
      sessionData.session.user.email,
    );

    /* -------------------------------------------------------------- */
    /* REDIRECIONAR                                                     */
    /* -------------------------------------------------------------- */

    console.log(
      "[AUTH] Redirecionando para /inicio...",
    );

    await navigate({
      to: "/inicio",
      replace: true,
    });
  }

  /* ------------------------------------------------------------------ */
  /* GOOGLE                                                             */
  /* ------------------------------------------------------------------ */

  async function google() {
    if (loading || msLoading) {
      return;
    }

    setMsLoading(true);

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
        console.error(
          "[AUTH] Erro no login Google:",
          result.error,
        );

        toast.error(
          "Falha no login com Google.",
        );

        return;
      }

      if (result.redirected) {
        console.log(
          "[AUTH] Redirecionando para autenticação Google...",
        );

        return;
      }

      await navigate({
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
    } finally {
      setMsLoading(false);
    }
  }

  /* ------------------------------------------------------------------ */
  /* INTERFACE                                                          */
  /* ------------------------------------------------------------------ */

  return (
    <div className="grid min-h-screen lg:grid-cols-[1.05fr_1fr]">

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

          {/* -------------------------------------------------------- */}
          {/* MICROSOFT                                                  */}
          {/* -------------------------------------------------------- */}

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

          <div className="my-6 flex items-center gap-3 text-[11px] uppercase tracking-wider text-muted-foreground">
            <span className="h-px flex-1 bg-border" />

            ou

            <span className="h-px flex-1 bg-border" />
          </div>

          {/* -------------------------------------------------------- */}
          {/* GOOGLE                                                     */}
          {/* -------------------------------------------------------- */}

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

          {/* -------------------------------------------------------- */}
          {/* E-MAIL                                                     */}
          {/* -------------------------------------------------------- */}

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
                      type="text"
                      value={name}
                      onChange={(e) =>
                        setName(
                          e.target.value,
                        )
                      }
                      autoComplete="name"
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
                    autoComplete="email"
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
                    autoComplete={
                      mode === "login"
                        ? "current-password"
                        : "new-password"
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

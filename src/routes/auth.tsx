import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { useEffect, useState } from "react";
import { toast } from "sonner";

import { supabase } from "@/integrations/supabase/client";

import {
  fetchEntraConfig,
  loginWithEntra,
  getPowerBIToken,
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

  /* ---------------------------------------------------------------------- */
  /* VERIFICA SESSÃO SUPABASE                                               */
  /* ---------------------------------------------------------------------- */

  useEffect(() => {
    let mounted = true;

    void supabase.auth
      .getSession()
      .then(({ data }) => {
        if (
          mounted &&
          data.session
        ) {
          navigate({
            to: "/inicio",
            replace: true,
          });
        }
      });

    return () => {
      mounted = false;
    };
  }, [navigate]);

  /* ---------------------------------------------------------------------- */
  /* LOGIN E-MAIL                                                            */
  /* ---------------------------------------------------------------------- */

  async function submit(
    e: React.FormEvent,
  ) {
    e.preventDefault();

    setLoading(true);

    try {
      if (mode === "login") {
        const { error } =
          await supabase.auth.signInWithPassword({
            email,
            password,
          });

        if (error) {
          throw error;
        }

        navigate({
          to: "/inicio",
          replace: true,
        });

        return;
      }

      const {
        data,
        error,
      } = await supabase.auth.signUp({
        email,
        password,

        options: {
          emailRedirectTo:
            window.location.origin,

          data: {
            full_name: name,
          },
        },
      });

      if (error) {
        throw error;
      }

      if (data.session) {
        navigate({
          to: "/inicio",
          replace: true,
        });
      } else {
        toast.success(
          "Conta criada. Confirme o e-mail para acessar o portal.",
        );
      }
    } catch (err) {
      toast.error(
        err instanceof Error
          ? err.message
          : "Não foi possível autenticar.",
      );
    } finally {
      setLoading(false);
    }
  }

  /* ---------------------------------------------------------------------- */
  /* LOGIN MICROSOFT                                                        */
  /* ---------------------------------------------------------------------- */

  async function microsoft() {
    if (msLoading) {
      return;
    }

    setMsLoading(true);

    try {
      console.log(
        "[AUTH] Iniciando login Microsoft...",
      );

      const config =
        await fetchEntraConfig();

      /*
       * Se não houver configuração do Entra
       * no banco, utiliza o OAuth anterior
       * do Lovable.
       */
      if (!config) {
        console.warn(
          "[AUTH] Entra não configurado. Utilizando OAuth legado.",
        );

        await microsoftLegacy();

        return;
      }

      console.log(
        "[AUTH] Configuração Entra encontrada.",
      );

      /* ------------------------------------------------------------------ */
      /* 1. LOGIN MICROSOFT                                                 */
      /* ------------------------------------------------------------------ */

      const {
        idToken,
        account,
      } = await loginWithEntra(
        config,
      );

      if (!account) {
        throw new Error(
          "Não foi possível identificar a conta Microsoft.",
        );
      }

      const microsoftEmail =
        account.username;

      console.log(
        "[AUTH] Microsoft autenticado:",
        microsoftEmail,
      );

      /* ------------------------------------------------------------------ */
      /* 2. CRIA / SINCRONIZA SESSÃO DO PORTAL                              */
      /* ------------------------------------------------------------------ */

      const {
        email: portalEmail,
        tokenHash,
      } = await entraSignInFn({
        data: {
          idToken,
        },
      });

      console.log(
        "[AUTH] Sessão do portal recebida:",
        portalEmail,
      );

      const {
        error: verifyError,
      } = await supabase.auth.verifyOtp({
        email: portalEmail,
        token_hash: tokenHash,
        type: "email",
      });

      if (verifyError) {
        throw verifyError;
      }

      console.log(
        "[AUTH] Sessão Supabase criada.",
      );

      /* ------------------------------------------------------------------ */
      /* 3. OBTÉM O TOKEN DO POWER BI                                       */
      /* ------------------------------------------------------------------ */

      /*
       * IMPORTANTE:
       *
       * O usuário acabou de autenticar no Microsoft.
       *
       * Portanto agora tentamos obter o token
       * específico do Power BI silenciosamente.
       *
       * Não usamos ID Token para o Power BI.
       */
      try {
        console.log(
          "[AUTH] Obtendo token do Power BI...",
        );

        const powerBIToken =
          await getPowerBIToken(
            config,
            {
              loginHint:
                microsoftEmail,

              /*
               * Não abrimos outro popup aqui.
               *
               * O login Microsoft já aconteceu.
               */
              interactive: false,
            },
          );

        if (
          powerBIToken?.accessToken
        ) {
          console.log(
            "[AUTH] Token Power BI obtido com sucesso.",
          );
        }
      } catch (powerBIError) {
        /*
         * Não derrubamos o login do portal.
         *
         * Se o usuário entrou corretamente no Microsoft,
         * ele deve conseguir chegar ao portal.
         *
         * O dashboard tentará novamente quando for aberto.
         */
        console.warn(
          "[AUTH] Token Power BI ainda não disponível:",
          powerBIError,
        );
      }

      /* ------------------------------------------------------------------ */
      /* 4. ENTRA NO PORTAL                                                 */
      /* ------------------------------------------------------------------ */

      console.log(
        "[AUTH] Login concluído. Redirecionando para o portal...",
      );

      navigate({
        to: "/inicio",
        replace: true,
      });
    } catch (err) {
      console.error(
        "[AUTH] Erro no login Microsoft:",
        err,
      );

      const message =
        err instanceof Error
          ? err.message
          : "Falha no login com Microsoft.";

      /*
       * Não mostramos erros técnicos de popup
       * como mensagem para o usuário.
       */
      if (
        !/user_cancelled|popup_window_error|window_closed/i.test(
          message,
        )
      ) {
        toast.error(message);
      }
    } finally {
      setMsLoading(false);
    }
  }

  /* ---------------------------------------------------------------------- */
  /* LOGIN MICROSOFT LEGADO                                                 */
  /* ---------------------------------------------------------------------- */

  async function microsoftLegacy() {
    const result =
      await lovable.auth.signInWithOAuth(
        "microsoft",
        {
          redirect_uri:
            window.location.origin,

          extraParams: {
            prompt: "select_account",
          },
        },
      );

    if (result.error) {
      toast.error(
        "Falha no login com Microsoft.",
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
  }

  /* ---------------------------------------------------------------------- */
  /* GOOGLE                                                                  */
  /* ---------------------------------------------------------------------- */

  async function google() {
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
  }

  /* ---------------------------------------------------------------------- */
  /* INTERFACE                                                               */
  /* ---------------------------------------------------------------------- */

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
            Dashboards publicados no Power BI Service,
            organizados por área, com acesso controlado
            por perfil.
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
            Entre com sua conta corporativa Microsoft —
            a mesma identidade será utilizada para acessar
            os dashboards do Power BI.
          </p>

          <Button
            className="mt-8 w-full"
            onClick={microsoft}
            disabled={msLoading}
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

          <Button
            variant="outline"
            className="w-full"
            onClick={google}
          >
            Continuar com Google
          </Button>

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
                  disabled={loading}
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

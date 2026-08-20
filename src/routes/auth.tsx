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
  signInWithMicrosoft,
} from "@/lib/powerbi-auth";

import { entraSignIn as entraSignInFn } from "@/lib/entra-session.functions";

import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";

export const Route = createFileRoute("/auth")({
  ssr: false,

  head: () => ({
    meta: [
      { title: "Entrar — Portal Corporativo BI Vem" },
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
          "Autenticação corporativa Microsoft do Portal Corporativo BI da Vem.",
      },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary_large_image" },
      { name: "robots", content: "noindex" },
    ],
  }),

  component: AuthPage,
});

function AuthPage() {
  const navigate = useNavigate();

  const [showEmail, setShowEmail] = useState(false);
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [loading, setLoading] = useState(false);
  const [msLoading, setMsLoading] = useState(false);

  const busy = loading || msLoading;

  useEffect(() => {
    let mounted = true;

    void (async () => {
      const { data, error } = await supabase.auth.getSession();

      if (error) {
        console.error("[AUTH] Erro ao recuperar sessão:", error);
        return;
      }

      if (!mounted || !data.session) {
        console.log(
          "[AUTH] Nenhuma sessão do portal. Aguardando ação do usuário.",
        );
        return;
      }

      console.log(
        "[AUTH] Sessão do portal já existente:",
        data.session.user.email,
      );

      await navigate({ to: "/inicio", replace: true });
    })();

    return () => {
      mounted = false;
    };
  }, [navigate]);

  async function createPortalSession(idToken: string) {
    console.log("[AUTH] Validando identidade Microsoft no servidor...");

    const { email: portalEmail, tokenHash } = await entraSignInFn({
      data: { idToken },
    });

    if (!portalEmail || !tokenHash) {
      throw new Error(
        "O servidor não retornou os dados necessários para criar a sessão do portal.",
      );
    }

    console.log(
      "[AUTH] Identidade Microsoft validada pelo servidor:",
      portalEmail,
    );

    /*
     * entra-session.server.ts usa admin.generateLink({ type: "magiclink" }).
     * Portanto, o verifyOtp precisa receber o tipo "magiclink".
     * O tipo "email" não corresponde ao token emitido por generateLink
     * e fazia o Supabase responder 400 validation_failed.
     */
    const { data: verifyData, error: verifyError } =
      await supabase.auth.verifyOtp({
        email: portalEmail,
        token_hash: tokenHash,
        type: "magiclink",
      });

    if (verifyError) {
      console.error(
        "[AUTH] Erro ao criar sessão Supabase:",
        verifyError,
      );
      throw verifyError;
    }

    if (!verifyData.session) {
      const { data: currentSession } =
        await supabase.auth.getSession();

      if (!currentSession.session) {
        throw new Error(
          "O Supabase validou o acesso, mas não retornou uma sessão do portal.",
        );
      }
    }

    console.log("[AUTH] Sessão do portal criada.");

    await navigate({ to: "/inicio", replace: true });
  }

  async function microsoft() {
    if (busy) {
      return;
    }

    setMsLoading(true);

    try {
      console.log("[AUTH] Iniciando login Microsoft...");

      const config = await fetchEntraConfig();

      if (!config) {
        throw new Error(
          "Microsoft Entra ID não está configurado no portal.",
        );
      }

      const { idToken, account } = await signInWithMicrosoft(config);

      console.log(
        "[AUTH] Conta Microsoft autenticada:",
        account.username,
      );

      await createPortalSession(idToken);
    } catch (error: any) {
      const errorCode = error?.errorCode ?? error?.code ?? "";

      console.error(
        "[AUTH] Falha no login Microsoft:",
        errorCode || error?.message,
      );

      if (
        errorCode === "user_cancelled" ||
        errorCode === "popup_window_error" ||
        errorCode === "empty_window_error"
      ) {
        toast.error(
          "Login Microsoft cancelado. Verifique se o navegador permite pop-ups deste site.",
        );
        return;
      }

      if (errorCode === "interaction_in_progress") {
        toast.error(
          "A autenticação Microsoft ainda está em andamento. Aguarde alguns segundos.",
        );
        return;
      }

      toast.error(
        error instanceof Error && error.message
          ? error.message
          : "Não foi possível entrar com a Microsoft.",
      );
    } finally {
      setMsLoading(false);
    }
  }

  async function submit(e: FormEvent<HTMLFormElement>) {
    e.preventDefault();

    if (busy) {
      return;
    }

    setLoading(true);

    try {
      const { error } = await supabase.auth.signInWithPassword({
        email: email.trim(),
        password,
      });

      if (error) {
        throw error;
      }

      console.log("[AUTH] Login por e-mail concluído.");

      await navigate({ to: "/inicio", replace: true });
    } catch (error) {
      console.error("[AUTH] Falha no login por e-mail:", error);

      toast.error(
        error instanceof Error
          ? error.message
          : "Não foi possível autenticar.",
      );
    } finally {
      setLoading(false);
    }
  }

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
            organizados por área, com acesso controlado por
            perfil.
          </p>
        </div>

        <p className="relative text-xs text-primary-foreground/60">
          Ambiente corporativo · Acesso monitorado
        </p>
      </div>

      <div className="flex items-center justify-center px-6 py-16">
        <div className="w-full max-w-sm">
          <h1 className="font-display text-2xl font-semibold">
            Acessar o portal
          </h1>

          <p className="mt-1.5 text-sm text-muted-foreground">
            Entre com sua conta corporativa Microsoft — a
            mesma identidade é utilizada para abrir os
            dashboards do Power BI.
          </p>

          <Button
            className="mt-8 w-full"
            onClick={microsoft}
            disabled={busy}
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

          {!showEmail ? (
            <button
              type="button"
              onClick={() => setShowEmail(true)}
              className="w-full text-center text-sm text-muted-foreground underline-offset-4 hover:underline"
            >
              Entrar com e-mail e senha
            </button>
          ) : (
            <form onSubmit={submit} className="space-y-4">
              <p className="text-xs text-muted-foreground">
                Use apenas quando o Microsoft Entra ID não estiver disponível.
              </p>

              <div className="space-y-1.5">
                <Label htmlFor="email">E-mail corporativo</Label>
                <Input
                  id="email"
                  type="email"
                  value={email}
                  onChange={(e) => setEmail(e.target.value)}
                  autoComplete="email"
                  required
                />
              </div>

              <div className="space-y-1.5">
                <Label htmlFor="password">Senha</Label>
                <Input
                  id="password"
                  type="password"
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  autoComplete="current-password"
                  required
                  minLength={6}
                />
              </div>

              <Button
                type="submit"
                variant="secondary"
                className="w-full"
                disabled={busy}
              >
                {loading ? "Aguarde…" : "Entrar"}
              </Button>
            </form>
          )}
        </div>
      </div>
    </div>
  );
}

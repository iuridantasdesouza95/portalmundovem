import { useEffect, useRef, useState } from "react";
import { Loader2, ShieldAlert } from "lucide-react";
import { usePowerBIToken } from "@/lib/powerbi-auth";
import { useProfile } from "@/lib/portal-data";

type Props = {
  reportUrl: string;
  reportId?: string | null;
  name: string;
};

type PowerBIReportInstance = {
  setAccessToken: (token: string) => Promise<void>;
  on: (eventName: string, handler: (event: any) => void) => void;
  off: (eventName: string, handler?: (event: any) => void) => void;
};

/**
 * Exibição de relatório Power BI usando:
 *
 * Microsoft Entra ID
 *       ↓
 * Access Token delegado
 *       ↓
 * Power BI SDK
 *       ↓
 * TokenType.Aad
 *
 * Não utiliza Embed Token.
 * Não inicia login próprio do Power BI.
 * Não utiliza iframe como fallback silencioso.
 */
export function PowerBIReport({
  reportUrl,
  reportId,
  name,
}: Props) {
  const containerRef = useRef<HTMLDivElement>(null);
  const embedRef = useRef<PowerBIReportInstance | null>(null);

  const [sdkError, setSdkError] = useState<string | null>(null);
  const [isEmbedding, setIsEmbedding] = useState(false);

  const { data: profile } = useProfile();

  const wantsSdk = Boolean(reportUrl && reportId);

  const {
    data: token,
    isLoading: tokenLoading,
    isError: tokenError,
    error: tokenErrorObject,
  } = usePowerBIToken(wantsSdk, profile?.email);

  const accessToken = token?.accessToken;

  /**
   * Não exibimos o token.
   * Apenas sabemos se ele existe ou não.
   */
  const hasAccessToken = Boolean(accessToken);

  /**
   * Diagnóstico do token.
   */
  useEffect(() => {
    if (!wantsSdk) {
      console.warn(
        "[PowerBI] SDK não será utilizado: reportUrl ou reportId ausente.",
      );
      return;
    }

    console.info("[PowerBI] Configuração do relatório:", {
      name,
      reportId,
      reportUrl,
      userEmail: profile?.email,
      hasAccessToken,
      tokenLoading,
      tokenError,
    });

    if (tokenError) {
      console.error(
        "[PowerBI] Erro ao obter Access Token:",
        tokenErrorObject,
      );
    }
  }, [
    wantsSdk,
    name,
    reportId,
    reportUrl,
    profile?.email,
    hasAccessToken,
    tokenLoading,
    tokenError,
    tokenErrorObject,
  ]);

  /**
   * Cria o embed somente quando:
   *
   * - existe URL
   * - existe Report ID
   * - existe Access Token
   */
  useEffect(() => {
    if (!wantsSdk) return;

    if (!accessToken) {
      return;
    }

    if (!containerRef.current) {
      console.warn(
        "[PowerBI] Container do relatório ainda não está disponível.",
      );
      return;
    }

    let disposed = false;

    const node = containerRef.current;

    async function embedReport() {
      try {
        setSdkError(null);
        setIsEmbedding(true);

        console.info("[PowerBI] Inicializando SDK...");

        const pbi = await import("powerbi-client");

        if (disposed) return;

        /**
         * Instância oficial do Power BI SDK.
         */
        const service = new pbi.service.Service(
          pbi.factories.hpmFactory,
          pbi.factories.wpmpFactory,
          pbi.factories.routerFactory,
        );

        /**
         * Limpa qualquer embed anterior.
         */
        service.reset(node);

        console.info("[PowerBI] Iniciando embed:", {
          reportId,
          reportUrl,
          tokenType: "Aad",
        });

        /**
         * IMPORTANTE:
         *
         * TokenType.Aad significa:
         * "este token pertence ao usuário autenticado
         * no Microsoft Entra ID".
         */
        const report = service.embed(node, {
          type: "report",

          id: reportId as string,

          embedUrl: reportUrl,

          accessToken: accessToken as string,

          tokenType: pbi.models.TokenType.Aad,

          settings: {
            panes: {
              filters: {
                visible: false,
                expanded: false,
              },

              pageNavigation: {
                visible: false,
              },
            },

            bars: {
              statusBar: {
                visible: false,
              },
            },

            navContentPaneEnabled: false,

            filterPaneEnabled: false,

            background:
              pbi.models.BackgroundType.Transparent,
          },
        });

        embedRef.current =
          report as unknown as PowerBIReportInstance;

        /**
         * Evento de erro do Power BI.
         *
         * NÃO fazemos fallback automático para iframe.
         *
         * Isso é proposital:
         * precisamos descobrir o erro real.
         */
        const handleError = (event: any) => {
          console.error(
            "[PowerBI SDK ERROR]",
            event?.detail,
          );

          const detail = event?.detail;

          let message = "O Power BI não conseguiu carregar o relatório.";

          if (detail) {
            if (typeof detail === "string") {
              message = detail;
            } else if (detail.message) {
              message = detail.message;
            } else if (detail.error) {
              message =
                typeof detail.error === "string"
                  ? detail.error
                  : detail.error.message ??
                    message;
            }
          }

          if (!disposed) {
            setSdkError(message);
            setIsEmbedding(false);
          }
        };

        report.off("error");

        report.on("error", handleError);

        /**
         * Eventos úteis para diagnóstico.
         */
        report.on("loaded", () => {
          console.info(
            "[PowerBI] Relatório carregado.",
          );
        });

        report.on("rendered", () => {
          console.info(
            "[PowerBI] Relatório renderizado.",
          );

          if (!disposed) {
            setIsEmbedding(false);
          }
        });

        setIsEmbedding(false);

        console.info(
          "[PowerBI] Embed criado com sucesso.",
        );
      } catch (error) {
        console.error(
          "[POWER BI EMBED ERROR]",
          error,
        );

        if (!disposed) {
          setIsEmbedding(false);

          setSdkError(
            error instanceof Error
              ? error.message
              : "Erro desconhecido ao inicializar o Power BI.",
          );
        }
      }
    }

    void embedReport();

    return () => {
      disposed = true;

      if (embedRef.current) {
        try {
          embedRef.current.off("error");
          embedRef.current.off("loaded");
          embedRef.current.off("rendered");
        } catch {
          // noop
        }
      }

      embedRef.current = null;

      try {
        node.replaceChildren();
      } catch {
        // noop
      }
    };
  }, [
    wantsSdk,
    reportId,
    reportUrl,
    accessToken,
  ]);

  /**
   * Renovação automática do Access Token.
   *
   * Quando o MSAL renovar o token:
   *
   * token antigo
   *      ↓
   * novo token
   *      ↓
   * setAccessToken()
   *
   * O relatório não precisa ser recarregado.
   */
  useEffect(() => {
    if (!accessToken) return;

    if (!embedRef.current) return;

    void embedRef.current
      .setAccessToken(accessToken)
      .then(() => {
        console.info(
          "[PowerBI] Access Token atualizado silenciosamente.",
        );
      })
      .catch((error) => {
        console.error(
          "[PowerBI] Erro ao atualizar Access Token:",
          error,
        );
      });
  }, [accessToken]);

  /**
   * Sem URL.
   */
  if (!reportUrl) {
    return (
      <div className="flex h-full w-full flex-col items-center justify-center gap-3 overflow-hidden rounded-xl border bg-card p-10 text-center">
        <ShieldAlert className="size-7 text-muted-foreground" />

        <p className="max-w-md text-sm text-muted-foreground">
          Nenhuma URL de relatório cadastrada para este
          dashboard.
        </p>

        <p className="max-w-md text-xs text-muted-foreground">
          Informe a URL de incorporação na área de
          Administração.
        </p>
      </div>
    );
  }

  /**
   * SDK precisa de Report ID.
   */
  if (!reportId) {
    return (
      <div className="flex h-full w-full flex-col items-center justify-center gap-3 overflow-hidden rounded-xl border bg-card p-10 text-center">
        <ShieldAlert className="size-7 text-muted-foreground" />

        <p className="max-w-md text-sm text-muted-foreground">
          O Report ID do Power BI não está configurado.
        </p>

        <p className="max-w-md text-xs text-muted-foreground">
          Informe o Report ID utilizado pelo SDK na
          configuração do dashboard.
        </p>
      </div>
    );
  }

  /**
   * Aguardando token.
   */
  if (tokenLoading && !hasAccessToken) {
    return (
      <div className="flex h-full w-full items-center justify-center gap-2 rounded-xl border bg-card text-sm text-muted-foreground">
        <Loader2 className="size-4 animate-spin" />

        Autenticando no Microsoft Entra ID…
      </div>
    );
  }

  /**
   * Falha na obtenção do token.
   */
  if (tokenError || !hasAccessToken) {
    return (
      <div className="flex h-full w-full flex-col items-center justify-center gap-3 rounded-xl border bg-card p-10 text-center">
        <ShieldAlert className="size-7 text-destructive" />

        <p className="max-w-lg text-sm font-medium">
          Não foi possível obter a autenticação do
          Microsoft Entra ID para o Power BI.
        </p>

        <p className="max-w-lg text-xs text-muted-foreground">
          Verifique se o login Microsoft foi concluído e
          se o aplicativo possui permissão para acessar
          o Power BI.
        </p>

        <p className="max-w-lg text-[11px] text-muted-foreground">
          Abra o Console do navegador (F12) para consultar
          o diagnóstico.
        </p>
      </div>
    );
  }

  /**
   * Erro real do SDK.
   */
  if (sdkError) {
    return (
      <div className="flex h-full w-full flex-col items-center justify-center gap-4 rounded-xl border bg-card p-10 text-center">
        <ShieldAlert className="size-8 text-destructive" />

        <div>
          <p className="text-sm font-medium">
            Não foi possível carregar o Dashboard
          </p>

          <p className="mt-2 max-w-2xl text-xs text-muted-foreground">
            O Microsoft Entra ID autenticou o usuário,
            mas o Power BI retornou um erro ao carregar
            o relatório.
          </p>
        </div>

        <div className="max-w-2xl rounded-md border bg-muted/40 p-3 text-left">
          <p className="break-all text-xs text-muted-foreground">
            {sdkError}
          </p>
        </div>

        <p className="text-[11px] text-muted-foreground">
          Consulte o Console do navegador (F12) para obter
          detalhes técnicos.
        </p>
      </div>
    );
  }

  /**
   * Embed em andamento.
   */
  return (
    <div className="relative h-full w-full overflow-hidden rounded-xl border bg-card">
      {isEmbedding && (
        <div className="absolute inset-0 z-10 flex items-center justify-center gap-2 bg-background/80 text-sm text-muted-foreground">
          <Loader2 className="size-4 animate-spin" />

          Carregando relatório…
        </div>
      )}

      <div
        ref={containerRef}
        className="h-full w-full [&_iframe]:border-0"
      />
    </div>
  );
}

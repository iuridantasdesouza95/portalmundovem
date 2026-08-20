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
  on: (
    eventName: string,
    handler: (event: any) => void,
  ) => void;
  off: (
    eventName: string,
    handler?: (event: any) => void,
  ) => void;
};

export function PowerBIReport({
  reportUrl,
  reportId,
  name,
}: Props) {
  const containerRef = useRef<HTMLDivElement>(null);
  const embedRef =
    useRef<PowerBIReportInstance | null>(null);

  const [sdkError, setSdkError] =
    useState<string | null>(null);

  const [isEmbedding, setIsEmbedding] =
    useState(false);

  const { data: profile } = useProfile();

  const wantsSdk = Boolean(reportUrl && reportId);

  const {
    data: token,
    isLoading: tokenLoading,
    isError: tokenError,
    error: tokenErrorObject,
  } = usePowerBIToken(
    wantsSdk,
    profile?.email,
  );

  const accessToken = token?.accessToken;

  const hasAccessToken = Boolean(accessToken);

  /*
   * Diagnóstico.
   *
   * IMPORTANTE:
   * Não registrar o token no console.
   */
  useEffect(() => {
    console.log("[POWERBI] Estado:", {
      wantsSdk,
      name,
      reportId,
      reportUrl,
      userEmail: profile?.email,
      tokenLoading,
      tokenError,
      hasAccessToken,
      expiresOn: token?.expiresOn,
    });

    if (tokenError) {
      console.error(
        "[POWERBI] Erro ao obter token:",
        tokenErrorObject,
      );
    }
  }, [
    wantsSdk,
    name,
    reportId,
    reportUrl,
    profile?.email,
    tokenLoading,
    tokenError,
    tokenErrorObject,
    hasAccessToken,
    token?.expiresOn,
  ]);

  /*
   * Embed do relatório.
   */
  useEffect(() => {
    if (!wantsSdk) {
      return;
    }

    if (!accessToken) {
      return;
    }

    if (!containerRef.current) {
      return;
    }

    let disposed = false;

    const node = containerRef.current;

    async function embedReport() {
      try {
        setSdkError(null);
        setIsEmbedding(true);

        console.log(
          "[POWERBI] Inicializando Power BI SDK...",
        );

        const pbi =
          await import("powerbi-client");

        if (disposed) {
          return;
        }

        const service =
          new pbi.service.Service(
            pbi.factories.hpmFactory,
            pbi.factories.wpmpFactory,
            pbi.factories.routerFactory,
          );

        service.reset(node);

        console.log(
          "[POWERBI] Criando embed:",
          {
            reportId,
            reportUrl,
            tokenType: "Aad",
          },
        );

        const embedConfig = {
          type: "report",

          id: reportId as string,

          embedUrl: reportUrl,

          accessToken,

          tokenType:
            pbi.models.TokenType.Aad,

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
        } as unknown as pbi.IEmbedConfiguration;

        const report = service.embed(node, embedConfig);

        embedRef.current =
          report as unknown as PowerBIReportInstance;

        const handleError = (event: any) => {
          console.error(
            "[POWERBI SDK ERROR]",
            event?.detail,
          );

          const detail = event?.detail;

          let message =
            "O Power BI não conseguiu carregar o relatório.";

          if (detail) {
            if (typeof detail === "string") {
              message = detail;
            } else if (detail.message) {
              message = detail.message;
            } else if (detail.error) {
              message =
                typeof detail.error === "string"
                  ? detail.error
                  : detail.error?.message ??
                    message;
            }
          }

          if (!disposed) {
            setSdkError(message);
            setIsEmbedding(false);
          }
        };

        const handleLoaded = () => {
          console.log(
            "[POWERBI] Relatório carregado.",
          );
        };

        const handleRendered = () => {
          console.log(
            "[POWERBI] Relatório renderizado.",
          );

          if (!disposed) {
            setIsEmbedding(false);
          }
        };

        report.off("error");
        report.off("loaded");
        report.off("rendered");

        report.on("error", handleError);
        report.on("loaded", handleLoaded);
        report.on("rendered", handleRendered);

        console.log(
          "[POWERBI] Embed criado com sucesso.",
        );
      } catch (error) {
        console.error(
          "[POWERBI] Erro ao inicializar SDK:",
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

  /*
   * Atualização do token sem recarregar o relatório.
   */
  useEffect(() => {
    if (!accessToken) {
      return;
    }

    if (!embedRef.current) {
      return;
    }

    void embedRef.current
      .setAccessToken(accessToken)
      .then(() => {
        console.log(
          "[POWERBI] Access Token atualizado.",
        );
      })
      .catch((error) => {
        console.error(
          "[POWERBI] Erro ao atualizar token:",
          error,
        );
      });
  }, [accessToken]);

  /*
   * URL ausente.
   */
  if (!reportUrl) {
    return (
      <div className="flex h-full w-full flex-col items-center justify-center gap-3 overflow-hidden rounded-xl border bg-card p-10 text-center">
        <ShieldAlert className="size-7 text-muted-foreground" />

        <p className="max-w-md text-sm text-muted-foreground">
          Nenhuma URL de relatório cadastrada para este dashboard.
        </p>
      </div>
    );
  }

  /*
   * Report ID ausente.
   */
  if (!reportId) {
    return (
      <div className="flex h-full w-full flex-col items-center justify-center gap-3 overflow-hidden rounded-xl border bg-card p-10 text-center">
        <ShieldAlert className="size-7 text-muted-foreground" />

        <p className="max-w-md text-sm text-muted-foreground">
          O Report ID do Power BI não está configurado.
        </p>
      </div>
    );
  }

  /*
   * Obtendo autenticação.
   */
  if (tokenLoading && !hasAccessToken) {
    return (
      <div className="flex h-full w-full items-center justify-center gap-2 rounded-xl border bg-card text-sm text-muted-foreground">
        <Loader2 className="size-4 animate-spin" />

        Autenticando no Microsoft 365...
      </div>
    );
  }

  /*
   * Token não obtido.
   */
  if (tokenError || !hasAccessToken) {
    return (
      <div className="flex h-full w-full flex-col items-center justify-center gap-4 rounded-xl border bg-card p-10 text-center">
        <ShieldAlert className="size-8 text-destructive" />

        <div>
          <p className="text-sm font-medium">
            Não foi possível autenticar no Power BI
          </p>

          <p className="mt-2 max-w-lg text-xs text-muted-foreground">
            Faça login com sua conta Microsoft corporativa
            para acessar este dashboard.
          </p>
        </div>

        {tokenErrorObject instanceof Error && (
          <div className="max-w-2xl rounded-md border bg-muted/40 p-3 text-left">
            <p className="break-all text-xs text-muted-foreground">
              {tokenErrorObject.message}
            </p>
          </div>
        )}
      </div>
    );
  }

  /*
   * Erro do SDK.
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
            A autenticação foi realizada, mas o Power BI
            retornou um erro ao carregar o relatório.
          </p>
        </div>

        <div className="max-w-2xl rounded-md border bg-muted/40 p-3 text-left">
          <p className="break-all text-xs text-muted-foreground">
            {sdkError}
          </p>
        </div>
      </div>
    );
  }

  return (
    <div className="relative h-full w-full overflow-hidden rounded-xl border bg-card">
      {isEmbedding && (
        <div className="absolute inset-0 z-10 flex items-center justify-center gap-2 bg-background/80 text-sm text-muted-foreground">
          <Loader2 className="size-4 animate-spin" />

          Carregando relatório...
        </div>
      )}

      <div
        ref={containerRef}
        className="h-full w-full [&_iframe]:border-0"
      />
    </div>
  );
}

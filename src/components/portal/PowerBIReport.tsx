import { useEffect, useRef, useState } from "react";
import { Loader2, ShieldAlert } from "lucide-react";
import {
  POWERBI_CONSENT_REQUIRED,
  POWERBI_LOGIN_REQUIRED,
  usePowerBIToken,
} from "@/lib/powerbi-auth";
import { Button } from "@/components/ui/button";
import { toast } from "sonner";
import { useProfile } from "@/lib/portal-data";

type Props = {
  reportUrl: string;
  reportId?: string | null;
  pageName?: string | null;
  name: string;
};

type PowerBIPage = {
  name: string;
  displayName: string;
  order: number;
};

type PowerBIReportInstance = {
  setAccessToken: (token: string) => Promise<void>;
  setPage?: (pageName: string) => Promise<void>;
  getPages?: () => Promise<PowerBIPage[]>;
  on: (eventName: string, handler: (event: any) => void) => void;
  off: (eventName: string, handler?: (event: any) => void) => void;
};

export function PowerBIReport({ reportUrl, reportId, pageName, name }: Props) {
  const containerRef = useRef<HTMLDivElement>(null);
  const embedRef = useRef<PowerBIReportInstance | null>(null);
  const [sdkError, setSdkError] = useState<string | null>(null);
  const [isEmbedding, setIsEmbedding] = useState(false);
  const [consenting, setConsenting] = useState(false);
  const { data: profile } = useProfile();

  const wantsSdk = Boolean(reportUrl && reportId);
  const {
    data: token,
    isLoading: tokenLoading,
    isError: tokenError,
    error: tokenErrorObject,
    consent,
  } = usePowerBIToken(wantsSdk, profile?.email);

  const accessToken = token?.accessToken;
  const hasAccessToken = Boolean(accessToken);

  useEffect(() => {
    console.log("[POWERBI] Estado:", {
      wantsSdk,
      name,
      reportId,
      reportUrl,
      pageName,
      userEmail: profile?.email,
      tokenLoading,
      tokenError,
      hasAccessToken,
      expiresOn: token?.expiresOn,
    });

    if (tokenError) {
      console.error("[POWERBI] Erro ao obter token:", tokenErrorObject);
    }
  }, [
    wantsSdk,
    name,
    reportId,
    reportUrl,
    pageName,
    profile?.email,
    tokenLoading,
    tokenError,
    tokenErrorObject,
    hasAccessToken,
    token?.expiresOn,
  ]);

  useEffect(() => {
    if (!wantsSdk || !accessToken || !containerRef.current) return;

    let disposed = false;
    const node = containerRef.current;

    async function embedReport() {
      try {
        setSdkError(null);
        setIsEmbedding(true);

        const pbi = await import("powerbi-client");
        if (disposed) return;

        const service = new pbi.service.Service(
          pbi.factories.hpmFactory,
          pbi.factories.wpmpFactory,
          pbi.factories.routerFactory,
        );

        service.reset(node);

        console.log("[POWERBI] Criando embed:", {
          reportId,
          reportUrl,
          pageName,
          tokenType: "Aad",
        });

        const embedConfig = {
          type: "report",
          id: reportId as string,
          embedUrl: reportUrl,
          accessToken,
          tokenType: pbi.models.TokenType.Aad,
          ...(pageName ? { pageName } : {}),
          settings: {
            panes: {
              filters: { visible: false, expanded: false },
              pageNavigation: { visible: false },
            },
            bars: { statusBar: { visible: false } },
            navContentPaneEnabled: false,
            filterPaneEnabled: false,
            background: pbi.models.BackgroundType.Transparent,
          },
        } as unknown as Parameters<typeof service.embed>[1];

        const report = service.embed(node, embedConfig) as unknown as PowerBIReportInstance;
        embedRef.current = report;

        // Exposto temporariamente para diagnóstico no navegador.
        // Permite executar window.__POWERBI_REPORT__.getPages() no Console.
        (window as typeof window & { __POWERBI_REPORT__?: PowerBIReportInstance }).__POWERBI_REPORT__ = report;

        const handleError = (event: any) => {
          console.error("[POWERBI SDK ERROR]", event?.detail);
          const detail = event?.detail;
          let message = "O Power BI não conseguiu carregar o relatório.";
          if (detail) {
            if (typeof detail === "string") message = detail;
            else if (detail.message) message = detail.message;
            else if (detail.error) {
              message = typeof detail.error === "string" ? detail.error : detail.error?.message ?? message;
            }
          }
          if (!disposed) {
            setSdkError(message);
            setIsEmbedding(false);
          }
        };

        const handleLoaded = async () => {
          console.log("[POWERBI] Relatório carregado.");

          if (typeof report.getPages !== "function") {
            console.warn("[POWERBI] getPages() não está disponível nesta versão/instância do SDK.");
          } else {
            try {
              const pages = await report.getPages();
              const normalizedPages = pages.map((page) => ({
                name: page.name,
                displayName: page.displayName,
                order: page.order,
              }));

              console.group("[POWERBI] Páginas disponíveis");
              console.table(normalizedPages);
              console.info(
                "[POWERBI] Para configurar uma página específica, use o valor da coluna 'name' no campo Page Name (Power BI).",
              );
              console.groupEnd();
            } catch (error) {
              console.warn("[POWERBI] Não foi possível obter as páginas:", error);
            }
          }

          if (pageName && report.setPage) {
            try {
              await report.setPage(pageName);
              console.log("[POWERBI] Página inicial aplicada:", pageName);
            } catch (error) {
              console.warn("[POWERBI] Não foi possível aplicar pageName:", error);
            }
          }
        };

        const handleRendered = () => {
          console.log("[POWERBI] Relatório renderizado.");
          if (!disposed) setIsEmbedding(false);
        };

        report.off("error");
        report.off("loaded");
        report.off("rendered");
        report.on("error", handleError);
        report.on("loaded", handleLoaded);
        report.on("rendered", handleRendered);
      } catch (error) {
        console.error("[POWERBI] Erro ao inicializar SDK:", error);
        if (!disposed) {
          setIsEmbedding(false);
          setSdkError(error instanceof Error ? error.message : "Erro desconhecido ao inicializar o Power BI.");
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
      const globalWindow = window as typeof window & { __POWERBI_REPORT__?: PowerBIReportInstance };
      delete globalWindow.__POWERBI_REPORT__;
      try {
        node.replaceChildren();
      } catch {
        // noop
      }
    };
  }, [wantsSdk, reportId, reportUrl, pageName, accessToken]);

  useEffect(() => {
    if (!accessToken || !embedRef.current) return;
    void embedRef.current.setAccessToken(accessToken).catch((error) => {
      console.error("[POWERBI] Erro ao atualizar token:", error);
    });
  }, [accessToken]);

  if (!reportUrl) {
    return <EmptyState message="Nenhuma URL de relatório cadastrada para este dashboard." />;
  }

  if (!reportId) {
    return <EmptyState message="O Report ID do Power BI não está configurado." />;
  }

  if (tokenLoading && !hasAccessToken) {
    return (
      <div className="flex h-full w-full items-center justify-center gap-2 rounded-xl border bg-card text-sm text-muted-foreground">
        <Loader2 className="size-4 animate-spin" />
        Autenticando no Microsoft 365...
      </div>
    );
  }

  if (tokenError || !hasAccessToken) {
    const reason = tokenErrorObject instanceof Error ? tokenErrorObject.message : "";
    const needsConsent = reason === POWERBI_CONSENT_REQUIRED;
    const needsLogin = reason === POWERBI_LOGIN_REQUIRED;

    return (
      <div className="flex h-full w-full flex-col items-center justify-center gap-4 rounded-xl border bg-card p-10 text-center">
        <ShieldAlert className="size-8 text-destructive" />
        <div>
          <p className="text-sm font-medium">
            {needsConsent ? "Autorização do Power BI necessária" : "Não foi possível autenticar no Power BI"}
          </p>
          <p className="mt-2 max-w-lg text-xs text-muted-foreground">
            {needsConsent
              ? "Na primeira utilização, a Microsoft precisa do seu consentimento para o portal ler seus relatórios do Power BI."
              : needsLogin
                ? "Sua sessão Microsoft não está disponível neste navegador. Saia e entre novamente com a conta corporativa Microsoft."
                : "Não foi possível obter o acesso ao Power BI com a sua conta Microsoft."}
          </p>
        </div>
        {needsConsent && (
          <Button
            disabled={consenting}
            onClick={async () => {
              setConsenting(true);
              try {
                await consent();
              } catch (error) {
                console.error("[POWERBI] Consentimento não concluído:", error);
                toast.error("Não foi possível concluir a autorização do Power BI.");
              } finally {
                setConsenting(false);
              }
            }}
          >
            {consenting ? "Aguardando Microsoft…" : "Autorizar Power BI"}
          </Button>
        )}
        {!needsConsent && !needsLogin && tokenErrorObject instanceof Error && (
          <div className="max-w-2xl rounded-md border bg-muted/40 p-3 text-left">
            <p className="break-all text-xs text-muted-foreground">{tokenErrorObject.message}</p>
          </div>
        )}
      </div>
    );
  }

  if (sdkError) {
    return (
      <div className="flex h-full w-full flex-col items-center justify-center gap-4 rounded-xl border bg-card p-10 text-center">
        <ShieldAlert className="size-8 text-destructive" />
        <div>
          <p className="text-sm font-medium">Não foi possível carregar o Dashboard</p>
          <p className="mt-2 max-w-2xl text-xs text-muted-foreground">
            A autenticação foi realizada, mas o Power BI retornou um erro ao carregar o relatório.
          </p>
        </div>
        <div className="max-w-2xl rounded-md border bg-muted/40 p-3 text-left">
          <p className="break-all text-xs text-muted-foreground">{sdkError}</p>
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
      <div ref={containerRef} className="h-full w-full [&_iframe]:border-0" />
    </div>
  );
}

function EmptyState({ message }: { message: string }) {
  return (
    <div className="flex h-full w-full flex-col items-center justify-center gap-3 overflow-hidden rounded-xl border bg-card p-10 text-center">
      <ShieldAlert className="size-7 text-muted-foreground" />
      <p className="max-w-md text-sm text-muted-foreground">{message}</p>
    </div>
  );
}

import { useEffect, useRef, useState } from "react";
import { Loader2, ShieldAlert } from "lucide-react";
import { usePowerBIToken } from "@/lib/powerbi-auth";

type Props = {
  reportUrl: string;
  reportId?: string | null;
  name: string;
};

/**
 * Camada de exibição do Power BI.
 *
 * Usa sempre o SDK oficial (powerbi-client) com o token do próprio usuário
 * (SSO Microsoft Entra ID). O iframe permanece apenas como fallback quando
 * o SDK falha ou quando o Entra ID ainda não está configurado.
 * Qualquer republicação feita no Power BI Desktop aparece aqui automaticamente.
 */
export function PowerBIReport({ reportUrl, reportId, name }: Props) {
  const containerRef = useRef<HTMLDivElement>(null);
  const [sdkFailed, setSdkFailed] = useState(false);

  const wantsSdk = Boolean(reportUrl && reportId);
  const { data: accessToken, isLoading, isError } = usePowerBIToken(wantsSdk);
  const useSdk = wantsSdk && Boolean(accessToken) && !sdkFailed;

  useEffect(() => {
    if (!useSdk || !containerRef.current) return;
    let disposed = false;
    const node = containerRef.current;

    (async () => {
      try {
        const pbi = await import("powerbi-client");
        if (disposed) return;
        const service = new pbi.service.Service(
          pbi.factories.hpmFactory,
          pbi.factories.wpmpFactory,
          pbi.factories.routerFactory,
        );
        service.reset(node);
        service.embed(node, {
          type: "report",
          id: reportId as string,
          embedUrl: reportUrl,
          accessToken: accessToken as string,
          tokenType: pbi.models.TokenType.Aad,
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
        });
      } catch {
        if (!disposed) setSdkFailed(true);
      }
    })();

    return () => {
      disposed = true;
      try {
        node.replaceChildren();
      } catch {
        /* noop */
      }
    };
  }, [useSdk, reportId, reportUrl, accessToken]);

  if (!reportUrl) {
    return (
      <div className="flex h-full w-full flex-col items-center justify-center gap-3 overflow-hidden rounded-xl border bg-card p-10 text-center">
        <ShieldAlert className="size-7 text-muted-foreground" />
        <p className="max-w-md text-sm text-muted-foreground">
          Nenhuma URL de relatório cadastrada para este dashboard. Informe a URL de incorporação na
          área de Administração.
        </p>
      </div>
    );
  }

  if (wantsSdk && isLoading && !isError) {
    return (
      <div className="flex h-full w-full items-center justify-center gap-2 rounded-xl border bg-card text-sm text-muted-foreground">
        <Loader2 className="size-4 animate-spin" /> Carregando relatório…
      </div>
    );
  }

  if (useSdk) {
    return (
      <div className="h-full w-full overflow-hidden rounded-xl border bg-card">
        <div ref={containerRef} className="h-full w-full [&_iframe]:border-0" />
      </div>
    );
  }

  return (
    <div className="h-full w-full overflow-hidden rounded-xl border bg-card">
      <iframe title={name} src={reportUrl} className="h-full w-full border-0" allowFullScreen />
    </div>
  );
}

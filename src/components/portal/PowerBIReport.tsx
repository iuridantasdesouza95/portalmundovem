import { useEffect, useRef, useState } from "react";
import { ExternalLink, ShieldAlert } from "lucide-react";
import { Button } from "@/components/ui/button";

type Props = {
  reportUrl: string;
  reportId?: string | null;
  accessToken?: string | null;
  name: string;
};

/**
 * Camada de exibição do Power BI.
 *
 * Usa o SDK oficial (powerbi-client) sempre que houver um token de embed
 * disponível — carregado dinamicamente no browser. Sem token, mantém o
 * carregamento seguro do relatório publicado, sem recriar nenhum visual.
 * Qualquer republicação feita no Power BI Desktop/Service aparece aqui
 * automaticamente, sem alteração de código.
 */
export function PowerBIReport({ reportUrl, reportId, accessToken, name }: Props) {
  const containerRef = useRef<HTMLDivElement>(null);
  const [sdkError, setSdkError] = useState<string | null>(null);
  const embedded = Boolean(accessToken && reportId);

  useEffect(() => {
    if (!embedded || !containerRef.current) return;
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
          tokenType: pbi.models.TokenType.Embed,
          settings: {
            panes: { filters: { visible: false }, pageNavigation: { visible: true } },
            background: pbi.models.BackgroundType.Transparent,
          },
        });
      } catch (err) {
        if (!disposed) setSdkError(err instanceof Error ? err.message : "Falha ao carregar o SDK");
      }
    })();

    return () => {
      disposed = true;
    };
  }, [embedded, reportId, reportUrl, accessToken]);

  if (embedded) {
    return (
      <div className="h-full w-full overflow-hidden rounded-xl border bg-card">
        <div ref={containerRef} className="h-full w-full [&_iframe]:border-0" />
        {sdkError ? (
          <p className="p-4 text-sm text-muted-foreground">{sdkError}</p>
        ) : null}
      </div>
    );
  }

  return (
    <div className="flex h-full w-full flex-col overflow-hidden rounded-xl border bg-card">
      {reportUrl ? (
        <iframe
          title={name}
          src={reportUrl}
          className="h-full w-full flex-1 border-0"
          allowFullScreen
        />
      ) : (
        <div className="flex flex-1 flex-col items-center justify-center gap-3 p-10 text-center">
          <ShieldAlert className="size-7 text-muted-foreground" />
          <p className="max-w-md text-sm text-muted-foreground">
            Nenhuma URL de relatório cadastrada para este dashboard. Informe a URL do Power BI na
            área de Administração.
          </p>
        </div>
      )}
      <div className="flex items-center justify-between border-t px-4 py-2 text-xs text-muted-foreground">
        <span>Conteúdo servido diretamente do Power BI Service.</span>
        {reportUrl ? (
          <Button asChild variant="ghost" size="sm">
            <a href={reportUrl} target="_blank" rel="noreferrer">
              Abrir em nova aba <ExternalLink className="ml-1 size-3.5" />
            </a>
          </Button>
        ) : null}
      </div>
    </div>
  );
}

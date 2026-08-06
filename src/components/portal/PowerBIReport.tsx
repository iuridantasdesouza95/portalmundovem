import { useEffect, useRef, useState } from "react";
import { Loader2, ShieldAlert } from "lucide-react";
import { usePowerBIToken } from "@/lib/powerbi-auth";
import { useProfile } from "@/lib/portal-data";

type Props = {
  reportUrl: string;
  reportId?: string | null;
  name: string;
};

/**
 * Camada de exibição do Power BI.
 *
 * Modo principal: SDK oficial (powerbi-client) com Access Token do Microsoft Entra ID
 * (TokenType.Aad) do próprio usuário — a mesma sessão criada no login do portal.
 * Este componente NUNCA inicia uma nova autenticação: apenas consome o token do cache.
 * O iframe é apenas fallback (SDK falhou, Power BI indisponível ou Entra não configurado).
 */
export function PowerBIReport({ reportUrl, reportId, name }: Props) {
  const containerRef = useRef<HTMLDivElement>(null);
  const embedRef = useRef<{ setAccessToken: (t: string) => Promise<void> } | null>(null);
  const [sdkFailed, setSdkFailed] = useState(false);

  const { data: profile } = useProfile();
  const wantsSdk = Boolean(reportUrl && reportId);
  const { data: token, isLoading, isError } = usePowerBIToken(wantsSdk, profile?.email);
  const accessToken = token?.accessToken;
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
        const report = service.embed(node, {
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
        embedRef.current = report as unknown as { setAccessToken: (t: string) => Promise<void> };
        report.off("error");
        report.on("error", () => {
          if (!disposed) setSdkFailed(true);
        });
      } catch {
        if (!disposed) setSdkFailed(true);
      }
    })();

    return () => {
      disposed = true;
      embedRef.current = null;
      try {
        node.replaceChildren();
      } catch {
        /* noop */
      }
    };
    // o embed não é recriado quando o token é renovado — ver efeito abaixo
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [useSdk, reportId, reportUrl]);

  // Renovação automática: injeta o novo Access Token no relatório já carregado,
  // sem recarregar a página e sem pedir login novamente.
  useEffect(() => {
    if (!accessToken || !embedRef.current) return;
    void embedRef.current.setAccessToken(accessToken).catch(() => {
      /* mantém o token anterior até a próxima renovação */
    });
  }, [accessToken]);

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

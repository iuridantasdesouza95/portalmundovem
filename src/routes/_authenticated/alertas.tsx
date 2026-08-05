import { createFileRoute } from "@tanstack/react-router";
import { useAlerts } from "@/lib/portal-data";

export const Route = createFileRoute("/_authenticated/alertas")({
  head: () => ({
    meta: [
      { title: "Alertas — Portal BI Vem" },
      {
        name: "description",
        content: "Alertas e notificações dos indicadores corporativos monitorados pelo portal.",
      },
      { property: "og:title", content: "Alertas — Portal BI Vem" },
      { property: "og:description", content: "Notificações dos indicadores corporativos." },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary_large_image" },
    ],
  }),
  component: AlertsPage,
});

const levelStyles: Record<string, string> = {
  critico: "bg-destructive",
  atencao: "bg-warning",
  sucesso: "bg-success",
  info: "bg-info",
};

function AlertsPage() {
  const { data: alerts = [] } = useAlerts();

  return (
    <div className="mx-auto w-full max-w-4xl px-6 py-8 xl:px-10">
      <h1 className="font-display text-3xl font-semibold">Alertas</h1>
      <p className="mt-1.5 text-sm text-muted-foreground">
        Notificações geradas a partir dos indicadores monitorados.
      </p>

      <div className="mt-8 space-y-3">
        {alerts.map((a) => (
          <div key={a.id} className="surface-panel flex gap-4 p-5">
            <span
              className={`mt-1.5 size-2 shrink-0 rounded-full ${levelStyles[a.level] ?? "bg-info"}`}
            />
            <div className="min-w-0">
              <p className="text-sm font-medium">{a.title}</p>
              <p className="mt-1 text-sm leading-relaxed text-muted-foreground">{a.message}</p>
              <p className="mt-2 text-[11px] text-muted-foreground">
                {new Date(a.created_at).toLocaleString("pt-BR")}
              </p>
            </div>
          </div>
        ))}
        {!alerts.length && <p className="text-sm text-muted-foreground">Nenhum alerta ativo.</p>}
      </div>
    </div>
  );
}

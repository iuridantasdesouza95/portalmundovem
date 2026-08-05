import { createFileRoute, Link } from "@tanstack/react-router";
import { useEffect } from "react";
import { ArrowLeft, RefreshCw, Star } from "lucide-react";
import { getIcon } from "@/lib/icons";
import {
  registerView,
  useCategories,
  useDashboards,
  useFavorites,
  useToggleFavorite,
} from "@/lib/portal-data";
import { PowerBIReport } from "@/components/portal/PowerBIReport";
import { Button } from "@/components/ui/button";
import { cn } from "@/lib/utils";

export const Route = createFileRoute("/_authenticated/dashboards/$id")({
  head: () => ({
    meta: [
      { title: "Dashboard — Portal BI Vem" },
      {
        name: "description",
        content: "Visualização do dashboard publicado no Power BI Service.",
      },
      { property: "og:title", content: "Dashboard — Portal BI Vem" },
      {
        property: "og:description",
        content: "Relatório corporativo exibido no Portal BI da Vem.",
      },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary_large_image" },
    ],
  }),
  component: DashboardViewer,
});

function DashboardViewer() {
  const { id } = Route.useParams();
  const { data: dashboards = [], isLoading } = useDashboards();
  const { data: categories = [] } = useCategories();
  const { data: favIds = [] } = useFavorites();
  const toggleFav = useToggleFavorite();

  const dashboard = dashboards.find((d) => d.id === id);
  const category = categories.find((c) => c.id === dashboard?.category_id);
  const isFav = favIds.includes(id);

  useEffect(() => {
    if (dashboard) void registerView(dashboard.id);
  }, [dashboard]);

  if (isLoading) {
    return <div className="p-10 text-sm text-muted-foreground">Carregando dashboard…</div>;
  }

  if (!dashboard) {
    return (
      <div className="flex flex-col items-center justify-center gap-4 p-20 text-center">
        <p className="text-sm text-muted-foreground">
          Dashboard não encontrado ou não liberado para o seu perfil.
        </p>
        <Button asChild variant="outline">
          <Link to="/dashboards">Voltar à central</Link>
        </Button>
      </div>
    );
  }

  const Icon = getIcon(dashboard.icon);

  return (
    <div className="flex h-[calc(100vh-3.5rem)] flex-col">
      <header className="flex flex-wrap items-center gap-4 border-b px-6 py-4 xl:px-8">
        <Button asChild variant="ghost" size="icon" aria-label="Voltar">
          <Link to="/dashboards">
            <ArrowLeft className="size-4" />
          </Link>
        </Button>
        <span className="grid size-10 place-items-center rounded-xl bg-primary-soft text-primary">
          <Icon className="size-5" />
        </span>
        <div className="min-w-0 flex-1">
          <div className="flex items-center gap-2">
            <h1 className="truncate font-display text-lg font-semibold">{dashboard.name}</h1>
            {category && (
              <span className="rounded-full bg-muted px-2 py-0.5 text-[11px] text-muted-foreground">
                {category.name}
              </span>
            )}
          </div>
          <p className="truncate text-xs text-muted-foreground">{dashboard.description}</p>
        </div>
        <div className="flex items-center gap-2 text-xs text-muted-foreground">
          <RefreshCw className="size-3.5 text-success" />
          {new Date(dashboard.last_published_at).toLocaleString("pt-BR", {
            dateStyle: "short",
            timeStyle: "short",
          })}
        </div>
        <Button
          variant="outline"
          size="sm"
          onClick={() => toggleFav.mutate({ dashboardId: dashboard.id, isFav })}
        >
          <Star className={cn("mr-1.5 size-3.5", isFav && "fill-primary text-primary")} />
          {isFav ? "Favorito" : "Favoritar"}
        </Button>
      </header>

      <div className="flex-1 p-4 xl:p-6">
        <PowerBIReport
          reportUrl={dashboard.report_url}
          reportId={dashboard.report_id}
          name={dashboard.name}
        />
      </div>
    </div>
  );
}

import { createFileRoute, Link } from "@tanstack/react-router";
import { useEffect, useMemo, useState } from "react";
import { ArrowRight, Bell, Clock, RefreshCw, Star } from "lucide-react";
import { getIcon } from "@/lib/icons";
import {
  useAlerts,
  useCategories,
  useDashboards,
  useFavorites,
  useProfile,
  useRecentViews,
} from "@/lib/portal-data";
import { Button } from "@/components/ui/button";

export const Route = createFileRoute("/_authenticated/inicio")({
  head: () => ({
    meta: [
      { title: "Início — Portal Corporativo BI Vem" },
      {
        name: "description",
        content: "Cockpit executivo com favoritos, acessos recentes e alertas dos dashboards.",
      },
      { property: "og:title", content: "Início — Portal Corporativo BI Vem" },
      { property: "og:description", content: "Cockpit executivo do Portal Corporativo de BI." },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary_large_image" },
    ],
  }),
  component: Home,
});

function useClock() {
  const [now, setNow] = useState(() => new Date());
  useEffect(() => {
    const t = setInterval(() => setNow(new Date()), 30_000);
    return () => clearInterval(t);
  }, []);
  return now;
}

function Home() {
  const now = useClock();
  const { data: profile } = useProfile();
  const { data: dashboards = [] } = useDashboards();
  const { data: categories = [] } = useCategories();
  const { data: favIds = [] } = useFavorites();
  const { data: recents = [] } = useRecentViews();
  const { data: alerts = [] } = useAlerts();

  const hour = now.getHours();
  const greeting = hour < 12 ? "Bom dia" : hour < 18 ? "Boa tarde" : "Boa noite";
  const firstName = (profile?.full_name || "").split(" ")[0] || "";

  const favorites = dashboards.filter((d) => favIds.includes(d.id));
  const recentDashboards = recents
    .map((r) => dashboards.find((d) => d.id === r.dashboard_id))
    .filter(Boolean)
    .slice(0, 4);

  const lastUpdate = useMemo(() => {
    if (!dashboards.length) return null;
    return dashboards
      .map((d) => new Date(d.last_published_at))
      .sort((a, b) => b.getTime() - a.getTime())[0];
  }, [dashboards]);

  const stats = [
    { label: "Dashboards disponíveis", value: dashboards.length },
    { label: "Áreas monitoradas", value: categories.length },
    { label: "Favoritos", value: favorites.length },
    { label: "Alertas ativos", value: alerts.length },
  ];

  return (
    <div className="mx-auto w-full max-w-[1600px] px-6 py-8 xl:px-10">
      <header className="flex flex-wrap items-end justify-between gap-4">
        <div>
          <p className="text-sm capitalize text-muted-foreground">
            {now.toLocaleDateString("pt-BR", {
              weekday: "long",
              day: "2-digit",
              month: "long",
              year: "numeric",
            })}{" "}
            ·{" "}
            {now.toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" })}
          </p>
          <h1 className="mt-1 font-display text-3xl font-semibold">
            {greeting}
            {firstName ? `, ${firstName}` : ""}.
          </h1>
        </div>
        {lastUpdate && (
          <div className="flex items-center gap-2 rounded-lg border bg-card px-3 py-2 text-xs text-muted-foreground">
            <RefreshCw className="size-3.5 text-success" />
            Última publicação no Power BI:{" "}
            {lastUpdate.toLocaleString("pt-BR", { dateStyle: "short", timeStyle: "short" })}
          </div>
        )}
      </header>

      <section className="mt-8 grid gap-4 sm:grid-cols-2 xl:grid-cols-4">
        {stats.map((s) => (
          <div key={s.label} className="surface-panel p-5">
            <p className="font-display text-3xl font-semibold">{s.value}</p>
            <p className="mt-1 text-sm text-muted-foreground">{s.label}</p>
          </div>
        ))}
      </section>

      <div className="mt-10 grid gap-10 xl:grid-cols-[1.6fr_1fr]">
        <div className="space-y-10">
          <Block
            title="Favoritos"
            icon={<Star className="size-4" />}
            empty="Marque dashboards como favoritos para acessá-los aqui."
            items={favorites.slice(0, 4)}
          />
          <Block
            title="Acessados recentemente"
            icon={<Clock className="size-4" />}
            empty="Seus últimos dashboards abertos aparecerão aqui."
            items={recentDashboards as typeof dashboards}
          />
        </div>

        <aside className="space-y-4">
          <div className="flex items-center justify-between">
            <h2 className="flex items-center gap-2 font-display text-base font-semibold">
              <Bell className="size-4" /> Alertas
            </h2>
            <Button asChild variant="ghost" size="sm">
              <Link to="/alertas">Ver todos</Link>
            </Button>
          </div>
          <div className="space-y-3">
            {alerts.slice(0, 4).map((a) => (
              <div key={a.id} className="surface-panel p-4">
                <div className="flex items-center gap-2">
                  <span
                    className={
                      "size-1.5 rounded-full " +
                      (a.level === "critico"
                        ? "bg-destructive"
                        : a.level === "atencao"
                          ? "bg-warning"
                          : a.level === "sucesso"
                            ? "bg-success"
                            : "bg-info")
                    }
                  />
                  <p className="text-sm font-medium">{a.title}</p>
                </div>
                <p className="mt-1.5 text-xs leading-relaxed text-muted-foreground">{a.message}</p>
              </div>
            ))}
            {!alerts.length && (
              <p className="text-sm text-muted-foreground">Nenhum alerta no momento.</p>
            )}
          </div>
        </aside>
      </div>
    </div>
  );
}

function Block({
  title,
  icon,
  items,
  empty,
}: {
  title: string;
  icon: React.ReactNode;
  items: { id: string; name: string; description: string; icon: string }[];
  empty: string;
}) {
  return (
    <section>
      <div className="flex items-center justify-between">
        <h2 className="flex items-center gap-2 font-display text-base font-semibold">
          {icon} {title}
        </h2>
        <Button asChild variant="ghost" size="sm">
          <Link to="/dashboards">
            Central de dashboards <ArrowRight className="ml-1 size-3.5" />
          </Link>
        </Button>
      </div>
      {items.length ? (
        <div className="mt-4 grid gap-4 sm:grid-cols-2">
          {items.map((d) => {
            const Icon = getIcon(d.icon);
            return (
              <Link
                key={d.id}
                to="/dashboards/$id"
                params={{ id: d.id }}
                className="surface-panel group flex items-start gap-3 p-4 transition-shadow hover:shadow-lift"
              >
                <span className="grid size-9 shrink-0 place-items-center rounded-lg bg-primary-soft text-primary">
                  <Icon className="size-4" />
                </span>
                <span className="min-w-0">
                  <span className="block truncate text-sm font-medium">{d.name}</span>
                  <span className="mt-0.5 line-clamp-2 block text-xs text-muted-foreground">
                    {d.description}
                  </span>
                </span>
              </Link>
            );
          })}
        </div>
      ) : (
        <p className="mt-4 text-sm text-muted-foreground">{empty}</p>
      )}
    </section>
  );
}

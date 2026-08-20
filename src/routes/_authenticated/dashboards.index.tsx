import { createFileRoute, Link } from "@tanstack/react-router";
import { useState } from "react";
import { Search, Star } from "lucide-react";
import { getIcon } from "@/lib/icons";
import { useCategories, useDashboards, useFavorites, useToggleFavorite } from "@/lib/portal-data";
import { Button } from "@/components/ui/button";
import { cn } from "@/lib/utils";

type DashboardSearch = { cat?: string };

export const Route = createFileRoute("/_authenticated/dashboards/")({
  validateSearch: (search: Record<string, unknown>): DashboardSearch => typeof search["cat"] === "string" ? { cat: search["cat"] } : {},
  head: () => ({ meta: [
    { title: "Central de Dashboards — Portal BI Vem" },
    { name: "description", content: "Todos os dashboards do Power BI liberados para o seu perfil, organizados por área." },
    { property: "og:title", content: "Central de Dashboards — Portal BI Vem" },
    { property: "og:description", content: "Dashboards corporativos publicados no Power BI Service." },
    { property: "og:type", content: "website" },
    { name: "twitter:card", content: "summary_large_image" },
  ] }),
  component: DashboardCenter,
});

function DashboardCenter() {
  const { cat } = Route.useSearch();
  const [term, setTerm] = useState("");
  const { data: categories = [] } = useCategories();
  const { data: dashboards = [] } = useDashboards();
  const { data: favIds = [] } = useFavorites();
  const toggleFav = useToggleFavorite();
  const activeCategory = categories.find((c) => c.slug === cat);
  const list = dashboards
    .filter((d) => { if (!activeCategory) return true; const ids = d.category_ids?.length ? d.category_ids : d.category_id ? [d.category_id] : []; return ids.includes(activeCategory.id); })
    .filter((d) => d.name.toLowerCase().includes(term.toLowerCase().trim()));

  return <div className="mx-auto w-full max-w-[1600px] px-6 py-8 xl:px-10">
    <header><h1 className="font-display text-3xl font-semibold">{activeCategory ? activeCategory.name : "Central de dashboards"}</h1><p className="mt-1.5 text-sm text-muted-foreground">{list.length} dashboard{list.length === 1 ? "" : "s"} disponível{list.length === 1 ? "" : "eis"} para o seu perfil.</p></header>
    <div className="mt-6 flex flex-wrap items-center gap-3">
      <div className="relative w-full max-w-xs"><Search className="pointer-events-none absolute left-3 top-1/2 size-4 -translate-y-1/2 text-muted-foreground" /><input value={term} onChange={(e) => setTerm(e.target.value)} placeholder="Buscar por nome" className="h-9 w-full rounded-lg border bg-card pl-9 pr-3 text-sm outline-none focus:border-ring" /></div>
      <div className="flex flex-wrap gap-2"><Chip to={{}} label="Todos" active={!cat} />{categories.map((c) => <Chip key={c.id} to={{ cat: c.slug }} label={c.name} active={cat === c.slug} />)}</div>
    </div>
    <div className="mt-8 grid gap-5 sm:grid-cols-2 xl:grid-cols-3 2xl:grid-cols-4">
      {list.map((d) => { const Icon = getIcon(d.icon); const isFav = favIds.includes(d.id); const ids = d.category_ids?.length ? d.category_ids : d.category_id ? [d.category_id] : []; const dashboardCategories = categories.filter((c) => ids.includes(c.id)); return <article key={d.id} className="surface-panel flex flex-col p-5 transition-shadow hover:shadow-lift">
        <div className="flex items-start justify-between"><span className="grid size-10 place-items-center rounded-xl bg-primary-soft text-primary"><Icon className="size-5" /></span><button onClick={() => toggleFav.mutate({ dashboardId: d.id, isFav })} aria-label="Favoritar" className="rounded-md p-1.5 text-muted-foreground transition-colors hover:text-primary"><Star className={cn("size-4", isFav && "fill-primary text-primary")} /></button></div>
        <h2 className="mt-4 font-display text-base font-semibold">{d.name}</h2><p className="mt-1.5 line-clamp-2 flex-1 text-sm leading-relaxed text-muted-foreground">{d.description}</p>
        <div className="mt-4 flex flex-wrap items-center gap-2 text-[11px] text-muted-foreground">{dashboardCategories.map((c) => <span key={c.id} className="rounded-full bg-muted px-2 py-0.5">{c.name}</span>)}<span>Atualizado {new Date(d.last_published_at).toLocaleDateString("pt-BR", { day: "2-digit", month: "short" })}</span></div>
        <Button asChild className="mt-4 w-full"><Link to="/dashboards/$id" params={{ id: d.id }}>Abrir dashboard</Link></Button>
      </article>; })}
    </div>
    {!list.length && <p className="mt-16 text-center text-sm text-muted-foreground">Nenhum dashboard encontrado para este filtro ou perfil de acesso.</p>}
  </div>;
}

function Chip({ to, label, active }: { to: DashboardSearch; label: string; active: boolean }) { return <Link to="/dashboards" search={to} className={cn("rounded-full border px-3 py-1.5 text-xs transition-colors", active ? "border-primary bg-primary-soft font-medium text-primary" : "bg-card text-muted-foreground hover:bg-muted")}>{label}</Link>; }

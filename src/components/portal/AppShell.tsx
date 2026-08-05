import { Link, useNavigate, useRouterState } from "@tanstack/react-router";
import { useQueryClient } from "@tanstack/react-query";
import { useEffect, useState } from "react";
import {
  Bell,
  ChevronLeft,
  LayoutGrid,
  LogOut,
  Menu,
  Search,
  Settings,
  ShieldCheck,
  UserRound,
} from "lucide-react";
import { supabase } from "@/integrations/supabase/client";
import { getIcon } from "@/lib/icons";
import {
  useAlerts,
  useCategories,
  useDashboards,
  useMyRoles,
  useProfile,
  roleLabel,
} from "@/lib/portal-data";
import { Button } from "@/components/ui/button";
import { Avatar, AvatarFallback } from "@/components/ui/avatar";
import { cn } from "@/lib/utils";

function Brand({ collapsed }: { collapsed: boolean }) {
  return (
    <Link to="/inicio" className="flex items-center gap-2.5 px-4 py-5">
      <span className="grid size-8 shrink-0 place-items-center rounded-lg bg-primary font-display text-sm font-bold text-primary-foreground">
        V
      </span>
      {!collapsed && (
        <span className="font-display text-sm font-semibold tracking-tight">
          Portal BI
          <span className="block text-[11px] font-normal text-muted-foreground">Vem Corporativo</span>
        </span>
      )}
    </Link>
  );
}

export function AppShell({ children }: { children: React.ReactNode }) {
  const [collapsed, setCollapsed] = useState(false);
  const [mobileOpen, setMobileOpen] = useState(false);
  const navigate = useNavigate();
  const qc = useQueryClient();
  const pathname = useRouterState({ select: (s) => s.location.pathname });
  const { data: categories = [] } = useCategories();
  const { data: dashboards = [] } = useDashboards();
  const { data: roles = [] } = useMyRoles();
  const { data: profile } = useProfile();
  const { data: alerts = [] } = useAlerts();
  const isAdmin = roles.includes("administrador");

  useEffect(() => setMobileOpen(false), [pathname]);

  const visibleCategories = categories.filter(
    (c) => isAdmin || dashboards.some((d) => d.category_id === c.id),
  );

  async function signOut() {
    await qc.cancelQueries();
    qc.clear();
    await supabase.auth.signOut();
    navigate({ to: "/auth", replace: true });
  }

  const initials = (profile?.full_name || profile?.email || "U")
    .split(" ")
    .map((p: string) => p[0])
    .slice(0, 2)
    .join("")
    .toUpperCase();

  const nav = (
    <nav className="flex h-full flex-col">
      <Brand collapsed={collapsed} />
      <div className="flex-1 space-y-6 overflow-y-auto px-3 pb-4">
        <Section collapsed={collapsed}>
          <Item to="/inicio" icon={getIcon("Gauge")} label="Início" collapsed={collapsed} />
          <Item
            to="/dashboards"
            icon={LayoutGrid}
            label="Todos os dashboards"
            collapsed={collapsed}
          />
        </Section>

        <Section title="Áreas" collapsed={collapsed}>
          {visibleCategories.map((c) => (
            <Item
              key={c.id}
              to="/dashboards"
              search={{ cat: c.slug }}
              icon={getIcon(c.icon)}
              label={c.name}
              collapsed={collapsed}
              active={pathname === "/dashboards"}
            />
          ))}
        </Section>

        <Section title="Portal" collapsed={collapsed}>
          <Item
            to="/alertas"
            icon={Bell}
            label="Alertas"
            collapsed={collapsed}
            badge={alerts.length}
          />
          {isAdmin && (
            <Item to="/admin" icon={Settings} label="Administração" collapsed={collapsed} />
          )}
          <Item to="/perfil" icon={UserRound} label="Meu perfil" collapsed={collapsed} />
        </Section>
      </div>

      <div className="border-t p-3">
        <div className={cn("flex items-center gap-3 rounded-lg px-2 py-2", collapsed && "justify-center")}>
          <Avatar className="size-8">
            <AvatarFallback className="bg-primary-soft text-xs font-semibold text-primary">
              {initials}
            </AvatarFallback>
          </Avatar>
          {!collapsed && (
            <div className="min-w-0 flex-1">
              <p className="truncate text-sm font-medium">{profile?.full_name || "Usuário"}</p>
              <p className="truncate text-[11px] text-muted-foreground">
                {roles.map(roleLabel).join(", ") || "Sem perfil"}
              </p>
            </div>
          )}
          <Button variant="ghost" size="icon" className="size-8" onClick={signOut} aria-label="Sair">
            <LogOut className="size-4" />
          </Button>
        </div>
      </div>
    </nav>
  );

  return (
    <div className="flex min-h-screen w-full bg-background">
      <aside
        className={cn(
          "sticky top-0 hidden h-screen shrink-0 border-r bg-sidebar transition-all duration-300 lg:block",
          collapsed ? "w-[76px]" : "w-[264px]",
        )}
      >
        {nav}
      </aside>

      {mobileOpen && (
        <div className="fixed inset-0 z-50 lg:hidden">
          <button
            aria-label="Fechar menu"
            className="absolute inset-0 bg-foreground/30"
            onClick={() => setMobileOpen(false)}
          />
          <div className="absolute left-0 top-0 h-full w-[264px] border-r bg-sidebar">{nav}</div>
        </div>
      )}

      <div className="flex min-w-0 flex-1 flex-col">
        <header className="sticky top-0 z-40 flex h-14 items-center gap-2 border-b bg-background/85 px-4 backdrop-blur">
          <Button
            variant="ghost"
            size="icon"
            className="lg:hidden"
            onClick={() => setMobileOpen(true)}
            aria-label="Abrir menu"
          >
            <Menu className="size-4" />
          </Button>
          <Button
            variant="ghost"
            size="icon"
            className="hidden lg:inline-flex"
            onClick={() => setCollapsed((v) => !v)}
            aria-label="Recolher menu"
          >
            <ChevronLeft className={cn("size-4 transition-transform", collapsed && "rotate-180")} />
          </Button>

          <div className="relative ml-1 hidden max-w-md flex-1 md:block">
            <Search className="pointer-events-none absolute left-3 top-1/2 size-4 -translate-y-1/2 text-muted-foreground" />
            <input
              readOnly
              onFocus={() => navigate({ to: "/dashboards" })}
              placeholder="Buscar dashboards"
              className="h-9 w-full rounded-lg border bg-muted/60 pl-9 pr-3 text-sm outline-none transition focus:border-ring focus:bg-card"
            />
          </div>

          <div className="ml-auto flex items-center gap-1">
            {isAdmin && (
              <span className="mr-1 hidden items-center gap-1.5 rounded-full bg-primary-soft px-2.5 py-1 text-[11px] font-medium text-primary sm:inline-flex">
                <ShieldCheck className="size-3" /> Administrador
              </span>
            )}
            <Button variant="ghost" size="icon" asChild aria-label="Alertas">
              <Link to="/alertas">
                <Bell className="size-4" />
              </Link>
            </Button>
          </div>
        </header>
        <main className="min-w-0 flex-1">{children}</main>
      </div>
    </div>
  );
}

function Section({
  title,
  collapsed,
  children,
}: {
  title?: string;
  collapsed: boolean;
  children: React.ReactNode;
}) {
  return (
    <div className="space-y-1">
      {title && !collapsed && (
        <p className="px-3 pb-1 text-[11px] font-semibold uppercase tracking-wider text-muted-foreground">
          {title}
        </p>
      )}
      {children}
    </div>
  );
}

function Item({
  to,
  search,
  icon: Icon,
  label,
  collapsed,
  badge,
}: {
  to: string;
  search?: Record<string, string>;
  icon: React.ComponentType<{ className?: string }>;
  label: string;
  collapsed: boolean;
  active?: boolean;
  badge?: number;
}) {
  return (
    <Link
      to={to}
      search={search as never}
      title={collapsed ? label : undefined}
      activeOptions={{ exact: !search, includeSearch: Boolean(search) }}
      className={cn(
        "flex items-center gap-3 rounded-lg px-3 py-2 text-sm text-sidebar-foreground transition-colors hover:bg-sidebar-accent",
        collapsed && "justify-center px-0",
      )}
      activeProps={{ className: "bg-primary-soft text-primary font-medium hover:bg-primary-soft" }}
    >
      <Icon className="size-4 shrink-0" />
      {!collapsed && <span className="truncate">{label}</span>}
      {!collapsed && badge ? (
        <span className="ml-auto rounded-full bg-muted px-1.5 text-[11px] text-muted-foreground">
          {badge}
        </span>
      ) : null}
    </Link>
  );
}

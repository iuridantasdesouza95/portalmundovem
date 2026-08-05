import { createFileRoute } from "@tanstack/react-router";
import { useEffect, useState } from "react";
import { useQuery, useQueryClient } from "@tanstack/react-query";
import { toast } from "sonner";
import { Pencil, Plus, Trash2 } from "lucide-react";
import { supabase } from "@/integrations/supabase/client";
import { iconNames, getIcon } from "@/lib/icons";
import {
  ROLES,
  logAction,
  roleLabel,
  useCategories,
  useDashboardRoles,
  useDashboards,
  type AppRole,
  type Dashboard,
} from "@/lib/portal-data";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Textarea } from "@/components/ui/textarea";
import { Checkbox } from "@/components/ui/checkbox";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";

export const Route = createFileRoute("/_authenticated/admin")({
  head: () => ({
    meta: [
      { title: "Administração — Portal BI Vem" },
      {
        name: "description",
        content:
          "Gestão de dashboards, categorias, usuários, perfis, permissões, logs e configurações do portal.",
      },
      { property: "og:title", content: "Administração — Portal BI Vem" },
      { property: "og:description", content: "Área administrativa do Portal Corporativo de BI." },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary_large_image" },
    ],
  }),
  component: AdminPage,
});

function AdminPage() {
  return (
    <div className="mx-auto w-full max-w-[1400px] px-6 py-8 xl:px-10">
      <h1 className="font-display text-3xl font-semibold">Administração</h1>
      <p className="mt-1.5 text-sm text-muted-foreground">
        Cadastre dashboards, defina permissões e gerencie o portal sem alterar código.
      </p>

      <Tabs defaultValue="dashboards" className="mt-8">
        <TabsList className="flex-wrap">
          <TabsTrigger value="dashboards">Dashboards</TabsTrigger>
          <TabsTrigger value="categorias">Categorias</TabsTrigger>
          <TabsTrigger value="usuarios">Usuários e perfis</TabsTrigger>
          <TabsTrigger value="logs">Logs</TabsTrigger>
          <TabsTrigger value="config">Configurações</TabsTrigger>
        </TabsList>

        <TabsContent value="dashboards" className="mt-6">
          <DashboardsAdmin />
        </TabsContent>
        <TabsContent value="categorias" className="mt-6">
          <CategoriesAdmin />
        </TabsContent>
        <TabsContent value="usuarios" className="mt-6">
          <UsersAdmin />
        </TabsContent>
        <TabsContent value="logs" className="mt-6">
          <LogsAdmin />
        </TabsContent>
        <TabsContent value="config" className="mt-6">
          <SettingsAdmin />
        </TabsContent>
      </Tabs>
    </div>
  );
}

const emptyForm = {
  name: "",
  description: "",
  category_id: "",
  icon: "BarChart3",
  report_url: "",
  workspace: "",
  report_id: "",
  sort_order: 0,
  status: "ativo",
};

function DashboardsAdmin() {
  const qc = useQueryClient();
  const { data: dashboards = [] } = useDashboards();
  const { data: categories = [] } = useCategories();
  const { data: dashRoles = [] } = useDashboardRoles();
  const [open, setOpen] = useState(false);
  const [editing, setEditing] = useState<Dashboard | null>(null);
  const [form, setForm] = useState({ ...emptyForm });
  const [roles, setRoles] = useState<AppRole[]>([]);

  useEffect(() => {
    if (!open) return;
    if (editing) {
      setForm({
        name: editing.name,
        description: editing.description,
        category_id: editing.category_id ?? "",
        icon: editing.icon,
        report_url: editing.report_url,
        workspace: editing.workspace,
        report_id: editing.report_id ?? "",
        sort_order: editing.sort_order,
        status: editing.status,
      });
      setRoles(dashRoles.filter((r) => r.dashboard_id === editing.id).map((r) => r.role));
    } else {
      setForm({ ...emptyForm });
      setRoles([]);
    }
  }, [open, editing, dashRoles]);

  async function save(e: React.FormEvent) {
    e.preventDefault();
    const payload = {
      ...form,
      category_id: form.category_id || null,
      report_id: form.report_id || null,
      sort_order: Number(form.sort_order) || 0,
    };
    let dashboardId = editing?.id;
    if (editing) {
      const { error } = await supabase.from("dashboards").update(payload).eq("id", editing.id);
      if (error) return toast.error(error.message);
    } else {
      const { data, error } = await supabase.from("dashboards").insert(payload).select("id").single();
      if (error) return toast.error(error.message);
      dashboardId = data.id;
    }
    if (dashboardId) {
      await supabase.from("dashboard_roles").delete().eq("dashboard_id", dashboardId);
      if (roles.length) {
        await supabase
          .from("dashboard_roles")
          .insert(roles.map((role) => ({ dashboard_id: dashboardId as string, role })));
      }
    }
    await logAction(editing ? "atualizou" : "cadastrou", "dashboard", form.name);
    qc.invalidateQueries({ queryKey: ["dashboards"] });
    qc.invalidateQueries({ queryKey: ["dashboard-roles"] });
    qc.invalidateQueries({ queryKey: ["audit-logs"] });
    setOpen(false);
    setEditing(null);
    toast.success("Dashboard salvo");
  }

  async function remove(d: Dashboard) {
    const { error } = await supabase.from("dashboards").delete().eq("id", d.id);
    if (error) return toast.error(error.message);
    await logAction("removeu", "dashboard", d.name);
    qc.invalidateQueries({ queryKey: ["dashboards"] });
    toast.success("Dashboard removido");
  }

  return (
    <div className="surface-panel overflow-hidden">
      <div className="flex items-center justify-between border-b p-4">
        <p className="text-sm font-medium">{dashboards.length} dashboards cadastrados</p>
        <Button
          size="sm"
          onClick={() => {
            setEditing(null);
            setOpen(true);
          }}
        >
          <Plus className="mr-1.5 size-4" /> Novo dashboard
        </Button>
      </div>
      <Table>
        <TableHeader>
          <TableRow>
            <TableHead>Nome</TableHead>
            <TableHead>Categoria</TableHead>
            <TableHead>Workspace</TableHead>
            <TableHead>Perfis</TableHead>
            <TableHead>Status</TableHead>
            <TableHead className="w-24" />
          </TableRow>
        </TableHeader>
        <TableBody>
          {dashboards.map((d) => {
            const Icon = getIcon(d.icon);
            return (
              <TableRow key={d.id}>
                <TableCell className="font-medium">
                  <span className="flex items-center gap-2">
                    <Icon className="size-4 text-primary" /> {d.name}
                  </span>
                </TableCell>
                <TableCell className="text-muted-foreground">
                  {categories.find((c) => c.id === d.category_id)?.name ?? "—"}
                </TableCell>
                <TableCell className="text-muted-foreground">{d.workspace || "—"}</TableCell>
                <TableCell className="text-xs text-muted-foreground">
                  {dashRoles
                    .filter((r) => r.dashboard_id === d.id)
                    .map((r) => roleLabel(r.role))
                    .join(", ") || "—"}
                </TableCell>
                <TableCell className="text-muted-foreground">{d.status}</TableCell>
                <TableCell>
                  <div className="flex justify-end gap-1">
                    <Button
                      variant="ghost"
                      size="icon"
                      aria-label="Editar"
                      onClick={() => {
                        setEditing(d);
                        setOpen(true);
                      }}
                    >
                      <Pencil className="size-4" />
                    </Button>
                    <Button
                      variant="ghost"
                      size="icon"
                      aria-label="Remover"
                      onClick={() => remove(d)}
                    >
                      <Trash2 className="size-4" />
                    </Button>
                  </div>
                </TableCell>
              </TableRow>
            );
          })}
        </TableBody>
      </Table>

      <Dialog open={open} onOpenChange={setOpen}>
        <DialogContent className="max-h-[90vh] overflow-y-auto sm:max-w-lg">
          <DialogHeader>
            <DialogTitle>{editing ? "Editar dashboard" : "Cadastrar dashboard"}</DialogTitle>
          </DialogHeader>
          <form onSubmit={save} className="space-y-4">
            <Field label="Nome">
              <Input
                value={form.name}
                onChange={(e) => setForm({ ...form, name: e.target.value })}
                required
              />
            </Field>
            <Field label="Descrição">
              <Textarea
                value={form.description}
                onChange={(e) => setForm({ ...form, description: e.target.value })}
              />
            </Field>
            <div className="grid gap-4 sm:grid-cols-2">
              <Field label="Categoria">
                <select
                  value={form.category_id}
                  onChange={(e) => setForm({ ...form, category_id: e.target.value })}
                  className="h-9 w-full rounded-md border bg-card px-3 text-sm"
                >
                  <option value="">Sem categoria</option>
                  {categories.map((c) => (
                    <option key={c.id} value={c.id}>
                      {c.name}
                    </option>
                  ))}
                </select>
              </Field>
              <Field label="Ícone">
                <select
                  value={form.icon}
                  onChange={(e) => setForm({ ...form, icon: e.target.value })}
                  className="h-9 w-full rounded-md border bg-card px-3 text-sm"
                >
                  {iconNames.map((n) => (
                    <option key={n} value={n}>
                      {n}
                    </option>
                  ))}
                </select>
              </Field>
            </div>
            <Field label="URL do relatório (Power BI)">
              <Input
                value={form.report_url}
                onChange={(e) => setForm({ ...form, report_url: e.target.value })}
                placeholder="https://app.powerbi.com/reportEmbed?reportId=…"
              />
            </Field>
            <div className="grid gap-4 sm:grid-cols-2">
              <Field label="Workspace">
                <Input
                  value={form.workspace}
                  onChange={(e) => setForm({ ...form, workspace: e.target.value })}
                />
              </Field>
              <Field label="Report ID (SDK)">
                <Input
                  value={form.report_id}
                  onChange={(e) => setForm({ ...form, report_id: e.target.value })}
                />
              </Field>
            </div>
            <div className="grid gap-4 sm:grid-cols-2">
              <Field label="Ordem de exibição">
                <Input
                  type="number"
                  value={form.sort_order}
                  onChange={(e) => setForm({ ...form, sort_order: Number(e.target.value) })}
                />
              </Field>
              <Field label="Status">
                <select
                  value={form.status}
                  onChange={(e) => setForm({ ...form, status: e.target.value })}
                  className="h-9 w-full rounded-md border bg-card px-3 text-sm"
                >
                  <option value="ativo">Ativo</option>
                  <option value="inativo">Inativo</option>
                </select>
              </Field>
            </div>
            <Field label="Perfis com acesso">
              <div className="grid grid-cols-2 gap-2">
                {ROLES.map((r) => (
                  <label key={r.value} className="flex items-center gap-2 text-sm">
                    <Checkbox
                      checked={roles.includes(r.value)}
                      onCheckedChange={(v) =>
                        setRoles(
                          v ? [...roles, r.value] : roles.filter((x) => x !== r.value),
                        )
                      }
                    />
                    {r.label}
                  </label>
                ))}
              </div>
            </Field>
            <Button type="submit" className="w-full">
              Salvar
            </Button>
          </form>
        </DialogContent>
      </Dialog>
    </div>
  );
}

function CategoriesAdmin() {
  const qc = useQueryClient();
  const { data: categories = [] } = useCategories();
  const [name, setName] = useState("");
  const [icon, setIcon] = useState("LayoutGrid");

  async function add(e: React.FormEvent) {
    e.preventDefault();
    const slug = name
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-|-$/g, "");
    const { error } = await supabase
      .from("categories")
      .insert({ name, slug, icon, sort_order: categories.length + 1 });
    if (error) return toast.error(error.message);
    setName("");
    qc.invalidateQueries({ queryKey: ["categories"] });
    toast.success("Categoria criada");
  }

  async function remove(id: string) {
    const { error } = await supabase.from("categories").delete().eq("id", id);
    if (error) return toast.error(error.message);
    qc.invalidateQueries({ queryKey: ["categories"] });
  }

  return (
    <div className="grid gap-6 lg:grid-cols-[1fr_320px]">
      <div className="surface-panel overflow-hidden">
        <Table>
          <TableHeader>
            <TableRow>
              <TableHead>Categoria</TableHead>
              <TableHead>Slug</TableHead>
              <TableHead className="w-16" />
            </TableRow>
          </TableHeader>
          <TableBody>
            {categories.map((c) => {
              const Icon = getIcon(c.icon);
              return (
                <TableRow key={c.id}>
                  <TableCell className="flex items-center gap-2 font-medium">
                    <Icon className="size-4 text-primary" /> {c.name}
                  </TableCell>
                  <TableCell className="text-muted-foreground">{c.slug}</TableCell>
                  <TableCell>
                    <Button
                      variant="ghost"
                      size="icon"
                      aria-label="Remover"
                      onClick={() => remove(c.id)}
                    >
                      <Trash2 className="size-4" />
                    </Button>
                  </TableCell>
                </TableRow>
              );
            })}
          </TableBody>
        </Table>
      </div>

      <form onSubmit={add} className="surface-panel h-fit space-y-4 p-5">
        <p className="font-display text-sm font-semibold">Nova categoria</p>
        <Field label="Nome">
          <Input value={name} onChange={(e) => setName(e.target.value)} required />
        </Field>
        <Field label="Ícone">
          <select
            value={icon}
            onChange={(e) => setIcon(e.target.value)}
            className="h-9 w-full rounded-md border bg-card px-3 text-sm"
          >
            {iconNames.map((n) => (
              <option key={n} value={n}>
                {n}
              </option>
            ))}
          </select>
        </Field>
        <Button type="submit" className="w-full">
          Adicionar
        </Button>
      </form>
    </div>
  );
}

function UsersAdmin() {
  const qc = useQueryClient();
  const { data: users = [] } = useQuery({
    queryKey: ["admin-users"],
    queryFn: async () => {
      const { data: profiles } = await supabase
        .from("profiles")
        .select("*")
        .order("created_at", { ascending: true });
      const { data: userRoles } = await supabase.from("user_roles").select("*");
      return (profiles ?? []).map((p) => ({
        ...p,
        roles: (userRoles ?? []).filter((r) => r.user_id === p.id).map((r) => r.role as AppRole),
      }));
    },
  });

  async function toggleRole(userId: string, role: AppRole, has: boolean) {
    if (has) {
      await supabase.from("user_roles").delete().eq("user_id", userId).eq("role", role);
    } else {
      await supabase.from("user_roles").insert({ user_id: userId, role });
    }
    await logAction(has ? "removeu perfil" : "concedeu perfil", "usuário", roleLabel(role));
    qc.invalidateQueries({ queryKey: ["admin-users"] });
  }

  return (
    <div className="surface-panel overflow-x-auto">
      <Table>
        <TableHeader>
          <TableRow>
            <TableHead>Usuário</TableHead>
            {ROLES.map((r) => (
              <TableHead key={r.value} className="text-center">
                {r.label}
              </TableHead>
            ))}
          </TableRow>
        </TableHeader>
        <TableBody>
          {users.map((u) => (
            <TableRow key={u.id}>
              <TableCell>
                <p className="font-medium">{u.full_name || "—"}</p>
                <p className="text-xs text-muted-foreground">{u.email}</p>
              </TableCell>
              {ROLES.map((r) => {
                const has = u.roles.includes(r.value);
                return (
                  <TableCell key={r.value} className="text-center">
                    <Checkbox
                      checked={has}
                      onCheckedChange={() => toggleRole(u.id, r.value, has)}
                      aria-label={`${r.label} para ${u.email}`}
                    />
                  </TableCell>
                );
              })}
            </TableRow>
          ))}
        </TableBody>
      </Table>
    </div>
  );
}

function LogsAdmin() {
  const { data: logs = [] } = useQuery({
    queryKey: ["audit-logs"],
    queryFn: async () => {
      const { data } = await supabase
        .from("audit_logs")
        .select("*")
        .order("created_at", { ascending: false })
        .limit(100);
      return data ?? [];
    },
  });

  return (
    <div className="surface-panel overflow-hidden">
      <Table>
        <TableHeader>
          <TableRow>
            <TableHead>Data</TableHead>
            <TableHead>Usuário</TableHead>
            <TableHead>Ação</TableHead>
            <TableHead>Detalhe</TableHead>
          </TableRow>
        </TableHeader>
        <TableBody>
          {logs.map((l) => (
            <TableRow key={l.id}>
              <TableCell className="text-muted-foreground">
                {new Date(l.created_at).toLocaleString("pt-BR")}
              </TableCell>
              <TableCell>{l.actor_name}</TableCell>
              <TableCell>
                {l.action} {l.entity}
              </TableCell>
              <TableCell className="text-muted-foreground">{l.detail}</TableCell>
            </TableRow>
          ))}
        </TableBody>
      </Table>
      {!logs.length && (
        <p className="p-6 text-sm text-muted-foreground">Nenhuma ação registrada ainda.</p>
      )}
    </div>
  );
}

function SettingsAdmin() {
  const qc = useQueryClient();
  const { data: settings = [] } = useQuery({
    queryKey: ["settings"],
    queryFn: async () => {
      const { data } = await supabase.from("app_settings").select("*").order("key");
      return data ?? [];
    },
  });
  const [values, setValues] = useState<Record<string, string>>({});

  useEffect(() => {
    setValues(Object.fromEntries(settings.map((s) => [s.key, s.value])));
  }, [settings]);

  async function save() {
    for (const [key, value] of Object.entries(values)) {
      await supabase.from("app_settings").update({ value }).eq("key", key);
    }
    qc.invalidateQueries({ queryKey: ["settings"] });
    toast.success("Configurações salvas");
  }

  return (
    <div className="surface-panel max-w-2xl space-y-5 p-6">
      {settings.map((s) => (
        <Field key={s.key} label={s.key}>
          <Input
            value={values[s.key] ?? ""}
            onChange={(e) => setValues({ ...values, [s.key]: e.target.value })}
          />
        </Field>
      ))}
      <Button onClick={save}>Salvar configurações</Button>
    </div>
  );
}

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div className="space-y-1.5">
      <Label>{label}</Label>
      {children}
    </div>
  );
}

import { createFileRoute } from "@tanstack/react-router";
import { useEffect, useState } from "react";
import { useQueryClient } from "@tanstack/react-query";
import { toast } from "sonner";
import { supabase } from "@/integrations/supabase/client";
import { roleLabel, useMyRoles, useProfile } from "@/lib/portal-data";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";

export const Route = createFileRoute("/_authenticated/perfil")({
  head: () => ({
    meta: [
      { title: "Meu perfil — Portal BI Vem" },
      {
        name: "description",
        content: "Dados do usuário, perfis de acesso e preferências no Portal Corporativo de BI.",
      },
      { property: "og:title", content: "Meu perfil — Portal BI Vem" },
      { property: "og:description", content: "Dados e permissões do usuário no portal." },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary_large_image" },
    ],
  }),
  component: ProfilePage,
});

function ProfilePage() {
  const { data: profile } = useProfile();
  const { data: roles = [] } = useMyRoles();
  const qc = useQueryClient();
  const [fullName, setFullName] = useState("");
  const [jobTitle, setJobTitle] = useState("");
  const [department, setDepartment] = useState("");
  const [saving, setSaving] = useState(false);

  useEffect(() => {
    if (profile) {
      setFullName(profile.full_name ?? "");
      setJobTitle(profile.job_title ?? "");
      setDepartment(profile.department ?? "");
    }
  }, [profile]);

  async function save(e: React.FormEvent) {
    e.preventDefault();
    if (!profile) return;
    setSaving(true);
    const { error } = await supabase
      .from("profiles")
      .update({ full_name: fullName, job_title: jobTitle, department })
      .eq("id", profile.id);
    setSaving(false);
    if (error) toast.error("Não foi possível salvar");
    else {
      toast.success("Perfil atualizado");
      qc.invalidateQueries({ queryKey: ["profile"] });
    }
  }

  return (
    <div className="mx-auto w-full max-w-3xl px-6 py-8 xl:px-10">
      <h1 className="font-display text-3xl font-semibold">Meu perfil</h1>
      <p className="mt-1.5 text-sm text-muted-foreground">{profile?.email}</p>

      <div className="mt-6 flex flex-wrap gap-2">
        {roles.map((r) => (
          <span
            key={r}
            className="rounded-full bg-primary-soft px-3 py-1 text-xs font-medium text-primary"
          >
            {roleLabel(r)}
          </span>
        ))}
      </div>

      <form onSubmit={save} className="surface-panel mt-8 space-y-5 p-6">
        <div className="space-y-1.5">
          <Label htmlFor="fullName">Nome completo</Label>
          <Input id="fullName" value={fullName} onChange={(e) => setFullName(e.target.value)} />
        </div>
        <div className="grid gap-5 sm:grid-cols-2">
          <div className="space-y-1.5">
            <Label htmlFor="jobTitle">Cargo</Label>
            <Input id="jobTitle" value={jobTitle} onChange={(e) => setJobTitle(e.target.value)} />
          </div>
          <div className="space-y-1.5">
            <Label htmlFor="department">Área</Label>
            <Input
              id="department"
              value={department}
              onChange={(e) => setDepartment(e.target.value)}
            />
          </div>
        </div>
        <Button type="submit" disabled={saving}>
          {saving ? "Salvando…" : "Salvar alterações"}
        </Button>
      </form>
    </div>
  );
}

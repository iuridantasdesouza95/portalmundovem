import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { supabase } from "@/integrations/supabase/client";

export type AppRole = "administrador" | "diretor" | "gerente" | "supervisor" | "analista";
export const ROLES: { value: AppRole; label: string }[] = [
  { value: "administrador", label: "Administrador" },
  { value: "diretor", label: "Diretor" },
  { value: "gerente", label: "Gerente" },
  { value: "supervisor", label: "Supervisor" },
  { value: "analista", label: "Analista" },
];
export const roleLabel = (r: string) => ROLES.find((x) => x.value === r)?.label ?? r;
export type Category = { id: string; name: string; slug: string; icon: string; sort_order: number };
export type Dashboard = {
  id: string;
  name: string;
  description: string;
  category_id: string | null;
  icon: string;
  report_url: string;
  workspace: string;
  report_id: string | null;
  page_name: string | null;
  sort_order: number;
  status: string;
  last_published_at: string;
  updated_at: string;
};
export function useSession() { return useQuery({ queryKey: ["session"], queryFn: async () => (await supabase.auth.getUser()).data.user, staleTime: 60_000 }); }
export function useProfile() { return useQuery({ queryKey: ["profile"], queryFn: async () => { const { data: userData } = await supabase.auth.getUser(); const uid = userData.user?.id; if (!uid) return null; const { data } = await supabase.from("profiles").select("*").eq("id", uid).maybeSingle(); return data; } }); }
export function useMyRoles() { return useQuery({ queryKey: ["my-roles"], queryFn: async () => { const { data: userData } = await supabase.auth.getUser(); const uid = userData.user?.id; if (!uid) return [] as AppRole[]; const { data } = await supabase.from("user_roles").select("role").eq("user_id", uid); return (data ?? []).map((r) => r.role as AppRole); } }); }
export function useCategories() { return useQuery({ queryKey: ["categories"], queryFn: async () => { const { data, error } = await supabase.from("categories").select("*").order("sort_order", { ascending: true }); if (error) throw error; return (data ?? []) as Category[]; } }); }
export function useDashboards() { return useQuery({ queryKey: ["dashboards"], queryFn: async () => { const { data, error } = await supabase.from("dashboards").select("*").order("sort_order", { ascending: true }); if (error) throw error; return (data ?? []) as Dashboard[]; } }); }
export function useDashboardRoles() { return useQuery({ queryKey: ["dashboard-roles"], queryFn: async () => { const { data, error } = await supabase.from("dashboard_roles").select("*"); if (error) throw error; return (data ?? []) as { id: string; dashboard_id: string; role: AppRole }[]; } }); }
export function useFavorites() { return useQuery({ queryKey: ["favorites"], queryFn: async () => { const { data } = await supabase.from("favorites").select("dashboard_id"); return (data ?? []).map((f) => f.dashboard_id); } }); }
export function useToggleFavorite() { const qc = useQueryClient(); return useMutation({ mutationFn: async ({ dashboardId, isFav }: { dashboardId: string; isFav: boolean }) => { const { data: userData } = await supabase.auth.getUser(); const uid = userData.user?.id; if (!uid) throw new Error("Sessão expirada"); if (isFav) await supabase.from("favorites").delete().eq("dashboard_id", dashboardId).eq("user_id", uid); else await supabase.from("favorites").insert({ dashboard_id: dashboardId, user_id: uid }); }, onSuccess: () => qc.invalidateQueries({ queryKey: ["favorites"] }) }); }
export function useRecentViews() { return useQuery({ queryKey: ["recent-views"], queryFn: async () => { const { data } = await supabase.from("recent_views").select("dashboard_id, viewed_at").order("viewed_at", { ascending: false }).limit(6); return data ?? []; } }); }
export async function registerView(dashboardId: string) { const { data: userData } = await supabase.auth.getUser(); const uid = userData.user?.id; if (!uid) return; await supabase.from("recent_views").upsert({ user_id: uid, dashboard_id: dashboardId, viewed_at: new Date().toISOString() }, { onConflict: "user_id,dashboard_id" }); }
export function useAlerts() { return useQuery({ queryKey: ["alerts"], queryFn: async () => { const { data } = await supabase.from("alerts").select("*").order("created_at", { ascending: false }); return data ?? []; } }); }
export async function logAction(action: string, entity: string, detail: string) { const { data: userData } = await supabase.auth.getUser(); const uid = userData.user?.id; if (!uid) return; await supabase.from("audit_logs").insert({ user_id: uid, actor_name: userData.user?.email ?? "", action, entity, detail }); }

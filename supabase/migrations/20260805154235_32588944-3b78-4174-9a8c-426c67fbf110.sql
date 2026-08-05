-- ENUM de perfis
CREATE TYPE public.app_role AS ENUM ('administrador','diretor','gerente','supervisor','analista');

-- updated_at helper
CREATE OR REPLACE FUNCTION public.set_updated_at()
RETURNS TRIGGER LANGUAGE plpgsql SET search_path = public AS $$
BEGIN NEW.updated_at = now(); RETURN NEW; END; $$;

-- PROFILES
CREATE TABLE public.profiles (
  id uuid PRIMARY KEY REFERENCES auth.users(id) ON DELETE CASCADE,
  full_name text NOT NULL DEFAULT '',
  email text NOT NULL DEFAULT '',
  job_title text,
  department text,
  avatar_url text,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now()
);
GRANT SELECT, INSERT, UPDATE, DELETE ON public.profiles TO authenticated;
GRANT ALL ON public.profiles TO service_role;
ALTER TABLE public.profiles ENABLE ROW LEVEL SECURITY;

-- USER ROLES
CREATE TABLE public.user_roles (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  user_id uuid NOT NULL REFERENCES auth.users(id) ON DELETE CASCADE,
  role public.app_role NOT NULL,
  created_at timestamptz NOT NULL DEFAULT now(),
  UNIQUE (user_id, role)
);
GRANT SELECT, INSERT, UPDATE, DELETE ON public.user_roles TO authenticated;
GRANT ALL ON public.user_roles TO service_role;
ALTER TABLE public.user_roles ENABLE ROW LEVEL SECURITY;

CREATE OR REPLACE FUNCTION public.has_role(_user_id uuid, _role public.app_role)
RETURNS boolean LANGUAGE sql STABLE SECURITY DEFINER SET search_path = public AS $$
  SELECT EXISTS (SELECT 1 FROM public.user_roles WHERE user_id = _user_id AND role = _role);
$$;

CREATE OR REPLACE FUNCTION public.is_admin()
RETURNS boolean LANGUAGE sql STABLE SECURITY DEFINER SET search_path = public AS $$
  SELECT EXISTS (SELECT 1 FROM public.user_roles WHERE user_id = auth.uid() AND role = 'administrador');
$$;

CREATE OR REPLACE FUNCTION public.my_roles()
RETURNS SETOF public.app_role LANGUAGE sql STABLE SECURITY DEFINER SET search_path = public AS $$
  SELECT role FROM public.user_roles WHERE user_id = auth.uid();
$$;

-- profiles policies
CREATE POLICY "profiles_select_own_or_admin" ON public.profiles FOR SELECT TO authenticated
  USING (id = auth.uid() OR public.is_admin());
CREATE POLICY "profiles_update_own_or_admin" ON public.profiles FOR UPDATE TO authenticated
  USING (id = auth.uid() OR public.is_admin()) WITH CHECK (id = auth.uid() OR public.is_admin());
CREATE POLICY "profiles_insert_self" ON public.profiles FOR INSERT TO authenticated
  WITH CHECK (id = auth.uid() OR public.is_admin());
CREATE POLICY "profiles_delete_admin" ON public.profiles FOR DELETE TO authenticated
  USING (public.is_admin());
CREATE TRIGGER profiles_updated_at BEFORE UPDATE ON public.profiles
  FOR EACH ROW EXECUTE FUNCTION public.set_updated_at();

-- user_roles policies
CREATE POLICY "user_roles_select_own_or_admin" ON public.user_roles FOR SELECT TO authenticated
  USING (user_id = auth.uid() OR public.is_admin());
CREATE POLICY "user_roles_admin_write" ON public.user_roles FOR ALL TO authenticated
  USING (public.is_admin()) WITH CHECK (public.is_admin());

-- CATEGORIES
CREATE TABLE public.categories (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  name text NOT NULL,
  slug text NOT NULL UNIQUE,
  icon text NOT NULL DEFAULT 'LayoutGrid',
  sort_order int NOT NULL DEFAULT 0,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now()
);
GRANT SELECT, INSERT, UPDATE, DELETE ON public.categories TO authenticated;
GRANT ALL ON public.categories TO service_role;
ALTER TABLE public.categories ENABLE ROW LEVEL SECURITY;
CREATE POLICY "categories_read" ON public.categories FOR SELECT TO authenticated USING (true);
CREATE POLICY "categories_admin_write" ON public.categories FOR ALL TO authenticated
  USING (public.is_admin()) WITH CHECK (public.is_admin());
CREATE TRIGGER categories_updated_at BEFORE UPDATE ON public.categories
  FOR EACH ROW EXECUTE FUNCTION public.set_updated_at();

-- DASHBOARDS
CREATE TABLE public.dashboards (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  name text NOT NULL,
  description text NOT NULL DEFAULT '',
  category_id uuid REFERENCES public.categories(id) ON DELETE SET NULL,
  icon text NOT NULL DEFAULT 'BarChart3',
  report_url text NOT NULL DEFAULT '',
  workspace text NOT NULL DEFAULT '',
  report_id text,
  sort_order int NOT NULL DEFAULT 0,
  status text NOT NULL DEFAULT 'ativo',
  last_published_at timestamptz NOT NULL DEFAULT now(),
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now()
);
GRANT SELECT, INSERT, UPDATE, DELETE ON public.dashboards TO authenticated;
GRANT ALL ON public.dashboards TO service_role;
ALTER TABLE public.dashboards ENABLE ROW LEVEL SECURITY;

CREATE TABLE public.dashboard_roles (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  dashboard_id uuid NOT NULL REFERENCES public.dashboards(id) ON DELETE CASCADE,
  role public.app_role NOT NULL,
  UNIQUE (dashboard_id, role)
);
GRANT SELECT, INSERT, UPDATE, DELETE ON public.dashboard_roles TO authenticated;
GRANT ALL ON public.dashboard_roles TO service_role;
ALTER TABLE public.dashboard_roles ENABLE ROW LEVEL SECURITY;

CREATE OR REPLACE FUNCTION public.can_view_dashboard(_dashboard_id uuid)
RETURNS boolean LANGUAGE sql STABLE SECURITY DEFINER SET search_path = public AS $$
  SELECT public.is_admin() OR EXISTS (
    SELECT 1 FROM public.dashboard_roles dr
    JOIN public.user_roles ur ON ur.role = dr.role
    WHERE dr.dashboard_id = _dashboard_id AND ur.user_id = auth.uid()
  );
$$;

CREATE POLICY "dashboards_read_permitted" ON public.dashboards FOR SELECT TO authenticated
  USING (public.is_admin() OR (status = 'ativo' AND public.can_view_dashboard(id)));
CREATE POLICY "dashboards_admin_write" ON public.dashboards FOR ALL TO authenticated
  USING (public.is_admin()) WITH CHECK (public.is_admin());
CREATE TRIGGER dashboards_updated_at BEFORE UPDATE ON public.dashboards
  FOR EACH ROW EXECUTE FUNCTION public.set_updated_at();

CREATE POLICY "dashboard_roles_read" ON public.dashboard_roles FOR SELECT TO authenticated USING (true);
CREATE POLICY "dashboard_roles_admin_write" ON public.dashboard_roles FOR ALL TO authenticated
  USING (public.is_admin()) WITH CHECK (public.is_admin());

-- FAVORITES
CREATE TABLE public.favorites (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  user_id uuid NOT NULL REFERENCES auth.users(id) ON DELETE CASCADE,
  dashboard_id uuid NOT NULL REFERENCES public.dashboards(id) ON DELETE CASCADE,
  created_at timestamptz NOT NULL DEFAULT now(),
  UNIQUE (user_id, dashboard_id)
);
GRANT SELECT, INSERT, UPDATE, DELETE ON public.favorites TO authenticated;
GRANT ALL ON public.favorites TO service_role;
ALTER TABLE public.favorites ENABLE ROW LEVEL SECURITY;
CREATE POLICY "favorites_own" ON public.favorites FOR ALL TO authenticated
  USING (user_id = auth.uid()) WITH CHECK (user_id = auth.uid());

-- RECENT VIEWS
CREATE TABLE public.recent_views (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  user_id uuid NOT NULL REFERENCES auth.users(id) ON DELETE CASCADE,
  dashboard_id uuid NOT NULL REFERENCES public.dashboards(id) ON DELETE CASCADE,
  viewed_at timestamptz NOT NULL DEFAULT now(),
  UNIQUE (user_id, dashboard_id)
);
GRANT SELECT, INSERT, UPDATE, DELETE ON public.recent_views TO authenticated;
GRANT ALL ON public.recent_views TO service_role;
ALTER TABLE public.recent_views ENABLE ROW LEVEL SECURITY;
CREATE POLICY "recent_views_own" ON public.recent_views FOR ALL TO authenticated
  USING (user_id = auth.uid()) WITH CHECK (user_id = auth.uid());

-- ALERTS
CREATE TABLE public.alerts (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  title text NOT NULL,
  message text NOT NULL DEFAULT '',
  level text NOT NULL DEFAULT 'info',
  dashboard_id uuid REFERENCES public.dashboards(id) ON DELETE SET NULL,
  created_at timestamptz NOT NULL DEFAULT now()
);
GRANT SELECT, INSERT, UPDATE, DELETE ON public.alerts TO authenticated;
GRANT ALL ON public.alerts TO service_role;
ALTER TABLE public.alerts ENABLE ROW LEVEL SECURITY;
CREATE POLICY "alerts_read" ON public.alerts FOR SELECT TO authenticated USING (true);
CREATE POLICY "alerts_admin_write" ON public.alerts FOR ALL TO authenticated
  USING (public.is_admin()) WITH CHECK (public.is_admin());

-- AUDIT LOGS
CREATE TABLE public.audit_logs (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  user_id uuid REFERENCES auth.users(id) ON DELETE SET NULL,
  actor_name text NOT NULL DEFAULT '',
  action text NOT NULL,
  entity text NOT NULL DEFAULT '',
  detail text NOT NULL DEFAULT '',
  created_at timestamptz NOT NULL DEFAULT now()
);
GRANT SELECT, INSERT ON public.audit_logs TO authenticated;
GRANT ALL ON public.audit_logs TO service_role;
ALTER TABLE public.audit_logs ENABLE ROW LEVEL SECURITY;
CREATE POLICY "audit_insert_self" ON public.audit_logs FOR INSERT TO authenticated
  WITH CHECK (user_id = auth.uid());
CREATE POLICY "audit_read_admin" ON public.audit_logs FOR SELECT TO authenticated
  USING (public.is_admin());

-- APP SETTINGS
CREATE TABLE public.app_settings (
  key text PRIMARY KEY,
  value text NOT NULL DEFAULT '',
  updated_at timestamptz NOT NULL DEFAULT now()
);
GRANT SELECT, INSERT, UPDATE, DELETE ON public.app_settings TO authenticated;
GRANT ALL ON public.app_settings TO service_role;
ALTER TABLE public.app_settings ENABLE ROW LEVEL SECURITY;
CREATE POLICY "settings_read" ON public.app_settings FOR SELECT TO authenticated USING (true);
CREATE POLICY "settings_admin_write" ON public.app_settings FOR ALL TO authenticated
  USING (public.is_admin()) WITH CHECK (public.is_admin());

-- NEW USER TRIGGER
CREATE OR REPLACE FUNCTION public.handle_new_user()
RETURNS TRIGGER LANGUAGE plpgsql SECURITY DEFINER SET search_path = public AS $$
BEGIN
  INSERT INTO public.profiles (id, full_name, email, avatar_url)
  VALUES (
    NEW.id,
    COALESCE(NEW.raw_user_meta_data->>'full_name', NEW.raw_user_meta_data->>'name', split_part(NEW.email,'@',1)),
    COALESCE(NEW.email,''),
    NEW.raw_user_meta_data->>'avatar_url'
  ) ON CONFLICT (id) DO NOTHING;

  IF NOT EXISTS (SELECT 1 FROM public.user_roles) THEN
    INSERT INTO public.user_roles (user_id, role) VALUES (NEW.id, 'administrador');
  ELSE
    INSERT INTO public.user_roles (user_id, role) VALUES (NEW.id, 'analista') ON CONFLICT DO NOTHING;
  END IF;
  RETURN NEW;
END; $$;

CREATE TRIGGER on_auth_user_created
AFTER INSERT ON auth.users FOR EACH ROW EXECUTE FUNCTION public.handle_new_user();

-- SEED
INSERT INTO public.categories (name, slug, icon, sort_order) VALUES
  ('Executivo','executivo','Gauge',1),
  ('Comercial','comercial','TrendingUp',2),
  ('Financeiro','financeiro','Wallet',3),
  ('Produção','producao','Factory',4),
  ('PCP','pcp','CalendarRange',5),
  ('Logística','logistica','Truck',6),
  ('RH','rh','Users',7),
  ('Indicadores','indicadores','Target',8),
  ('Relatórios','relatorios','FileText',9);

INSERT INTO public.dashboards (name, description, category_id, icon, report_url, workspace, sort_order, status)
SELECT d.name, d.description, c.id, d.icon, d.report_url, d.workspace, d.sort_order, 'ativo'
FROM (VALUES
  ('Dashboard Executivo','Visão consolidada dos principais indicadores da companhia.','executivo','Gauge','https://app.powerbi.com/reportEmbed?reportId=00000000-0000-0000-0000-000000000001','Vem BI Corporativo',1),
  ('Performance Comercial','Vendas, funil, carteira e desempenho por região.','comercial','TrendingUp','https://app.powerbi.com/reportEmbed?reportId=00000000-0000-0000-0000-000000000002','Vem BI Comercial',2),
  ('Resultado Financeiro','DRE, fluxo de caixa, margens e contas a receber.','financeiro','Wallet','https://app.powerbi.com/reportEmbed?reportId=00000000-0000-0000-0000-000000000003','Vem BI Financeiro',3),
  ('Produção Industrial','OEE, refugo, paradas e produtividade por linha.','producao','Factory','https://app.powerbi.com/reportEmbed?reportId=00000000-0000-0000-0000-000000000004','Vem BI Operações',4),
  ('PCP e Programação','Aderência ao plano, capacidade e sequenciamento.','pcp','CalendarRange','https://app.powerbi.com/reportEmbed?reportId=00000000-0000-0000-0000-000000000005','Vem BI Operações',5),
  ('Logística e Expedição','OTIF, frete, ocupação de frota e entregas.','logistica','Truck','https://app.powerbi.com/reportEmbed?reportId=00000000-0000-0000-0000-000000000006','Vem BI Supply',6),
  ('Pessoas e RH','Headcount, turnover, absenteísmo e clima.','rh','Users','https://app.powerbi.com/reportEmbed?reportId=00000000-0000-0000-0000-000000000007','Vem BI Pessoas',7)
) AS d(name, description, cat, icon, report_url, workspace, sort_order)
JOIN public.categories c ON c.slug = d.cat;

INSERT INTO public.dashboard_roles (dashboard_id, role)
SELECT d.id, r.role FROM public.dashboards d
CROSS JOIN LATERAL (
  SELECT unnest(
    CASE d.name
      WHEN 'Dashboard Executivo' THEN ARRAY['diretor','gerente']::public.app_role[]
      WHEN 'Performance Comercial' THEN ARRAY['diretor','gerente','analista']::public.app_role[]
      WHEN 'Resultado Financeiro' THEN ARRAY['diretor','gerente']::public.app_role[]
      WHEN 'Produção Industrial' THEN ARRAY['gerente','supervisor','analista']::public.app_role[]
      WHEN 'PCP e Programação' THEN ARRAY['gerente','supervisor','analista']::public.app_role[]
      WHEN 'Logística e Expedição' THEN ARRAY['gerente','supervisor']::public.app_role[]
      ELSE ARRAY['diretor','gerente']::public.app_role[]
    END
  ) AS role
) r;

INSERT INTO public.alerts (title, message, level) VALUES
  ('Atualização concluída','Os conjuntos de dados do Power BI foram atualizados com sucesso.','sucesso'),
  ('Margem abaixo da meta','A margem bruta consolidada ficou 1,8 p.p. abaixo da meta do mês.','atencao'),
  ('OTIF em queda','Indicador de entregas no prazo caiu para 92% na última semana.','critico');

INSERT INTO public.app_settings (key, value) VALUES
  ('portal_name','Portal Corporativo BI'),
  ('company_name','Vem'),
  ('powerbi_tenant',''),
  ('refresh_note','Dashboards publicados no Power BI Service refletem automaticamente no portal.');

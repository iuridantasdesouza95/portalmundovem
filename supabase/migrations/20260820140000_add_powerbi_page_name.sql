-- Permite que cada dashboard do portal aponte para uma página específica do relatório Power BI.
-- O valor deve ser o `name` interno retornado pelo Power BI SDK (ex.: ReportSection...).
ALTER TABLE public.dashboards
ADD COLUMN IF NOT EXISTS page_name text;

COMMENT ON COLUMN public.dashboards.page_name IS
  'Nome interno da página do relatório Power BI usado pelo SDK para abrir a página inicial.';

-- Campo opcional: dashboards existentes continuam funcionando sem page_name.

-- Automação Conexa (faturas de serviço) -> NIBO (tela de conferência) -> baixa no ClickUp
-- Reaproveita o mesmo projeto/tabela `clientes` já usados no fluxo OneFlow -> Nibo.
-- Segue a mesma convenção já validada em produção: chave natural `cnpj`, não um FK de
-- id (o schema real de `clientes`/`entregas` não foi inspecionado diretamente nesta
-- sessão, então evitamos assumir uma coluna `id` que pode não existir).
-- Não altera a tabela `entregas` (em produção para o fluxo OneFlow) -- a tabela nova
-- abaixo é isolada para não afetar o que já está rodando.

alter table clientes
  add column if not exists conexa_customer_id text,
  add column if not exists clickup_list_id text;

create unique index if not exists clientes_conexa_customer_id_key
  on clientes (conexa_customer_id)
  where conexa_customer_id is not null;

create table if not exists envios_nibo_faturas (
  id uuid primary key default gen_random_uuid(),
  cnpj text not null,
  conexa_invoice_id text not null,
  nibo_file_id text,
  clickup_task_id text,
  status text not null default 'enviado' check (status in ('enviado', 'confirmado', 'erro')),
  error_message text,
  sent_at timestamptz,
  confirmed_at timestamptz,
  created_at timestamptz not null default now(),
  unique (cnpj, conexa_invoice_id)
);

create index if not exists envios_nibo_faturas_status_idx
  on envios_nibo_faturas (status);

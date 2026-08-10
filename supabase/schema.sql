-- ═══════════════════════════════════════════════════════════════════════
-- Gestor Financeiro de Condomínios — Schema Supabase (Fase 2)
-- Projeto Supabase NOVO e dedicado a este cliente.
--
-- Como aplicar: Supabase Dashboard → SQL Editor → colar tudo → Run
-- ═══════════════════════════════════════════════════════════════════════

-- ─── Tabela de histórico de relatórios gerados ──────────────────────────
create table public.relatorios (
  id                uuid primary key default gen_random_uuid(),
  user_id           uuid not null default auth.uid()
                    references auth.users (id) on delete cascade,
  nome_condominio   text,
  data_assembleia   text,
  ano_previsao      integer,
  num_unidades      integer,
  total_despesas    numeric(14,2),
  total_rateado     numeric(14,2),
  taxa_ideal_mensal numeric(12,2),
  taxa_atual        numeric(12,2),
  reajuste_pct      numeric(7,2),
  nome_arquivo      text,
  criado_em         timestamptz not null default now()
);

comment on table public.relatorios is
  'Histórico de apresentações geradas pelo Gestor Financeiro de Condomínios';

-- Índice para a listagem do histórico (por usuário, mais recentes primeiro)
create index relatorios_user_criado_idx
  on public.relatorios (user_id, criado_em desc);

-- ─── Row Level Security ─────────────────────────────────────────────────
alter table public.relatorios enable row level security;

-- Cada usuário enxerga apenas os próprios relatórios
create policy "usuario_le_proprios_relatorios"
  on public.relatorios
  for select
  to authenticated
  using (user_id = auth.uid());

-- Cada usuário insere apenas em nome próprio
-- (user_id tem default auth.uid(), então o insert nem precisa enviá-lo)
create policy "usuario_insere_proprios_relatorios"
  on public.relatorios
  for insert
  to authenticated
  with check (user_id = auth.uid());

-- Cada usuário pode remover entradas do próprio histórico
create policy "usuario_apaga_proprios_relatorios"
  on public.relatorios
  for delete
  to authenticated
  using (user_id = auth.uid());

-- Sem policy de UPDATE: histórico é imutável por design.
-- Sem acesso anon: apenas usuários autenticados.

-- ─── Observações operacionais (não são SQL) ─────────────────────────────
-- 1. USUÁRIOS: não há signup público. Crie os acessos do cliente em
--    Authentication → Users → "Add user" (e-mail + senha), com
--    "Auto Confirm User" marcado.
--
-- 2. DESATIVAR SIGNUP PÚBLICO: Authentication → Sign In / Up →
--    desmarque "Allow new users to sign up".
--
-- 3. CHAVES: use a chave "publishable" (sb_publishable_...) no frontend
--    e no backend (variável SUPABASE_ANON_KEY). Nunca exponha a
--    service_role / secret key.

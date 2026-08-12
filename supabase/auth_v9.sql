-- ═══════════════════════════════════════════════════════════════════════
-- Gestor Financeiro de Condomínios — Autocadastro com Chave Corporativa (v9)
--
-- O que este script faz:
--   1. Cria a tabela de chaves corporativas (com validade)
--   2. Cria o gatilho que BLOQUEIA qualquer cadastro sem chave válida
--   3. Insere a chave inicial do cliente, válida por 2 ANOS
--
-- Como aplicar: Supabase Dashboard → SQL Editor → colar tudo → Run
--
-- ⚠ PASSO MANUAL OBRIGATÓRIO após rodar o SQL:
--   Authentication → Sign In / Up:
--     • MARCAR   "Allow new users to sign up"  (o gatilho protege o cadastro)
--     • DESMARCAR "Confirm email"              (acesso imediato após cadastro)
--
-- ⚠ IMPORTANTE: a partir desta versão, TODO cadastro passa pelo site com a
--   chave corporativa. O botão "Add user" do dashboard deixa de funcionar
--   (o gatilho bloqueia cadastros sem chave) — e é assim que deve ser.
-- ═══════════════════════════════════════════════════════════════════════

-- ─── 1. Tabela de chaves corporativas ────────────────────────────────────
create table if not exists public.chaves_corporativas (
  chave      text primary key,
  descricao  text,
  valida_ate timestamptz not null,
  ativa      boolean not null default true,
  criada_em  timestamptz not null default now()
);

-- Ninguém acessa a tabela diretamente (nem authenticated, nem anon).
-- A validação acontece apenas dentro do gatilho (security definer).
alter table public.chaves_corporativas enable row level security;

-- ─── 2. Gatilho de validação no cadastro ─────────────────────────────────
create or replace function public.validar_chave_no_cadastro()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
declare
  chave_informada text;
begin
  chave_informada := new.raw_user_meta_data ->> 'chave_corporativa';

  if chave_informada is null or chave_informada = '' then
    raise exception 'Cadastro requer uma chave de acesso corporativa.';
  end if;

  if not exists (
    select 1 from public.chaves_corporativas
    where chave = chave_informada
      and ativa
      and valida_ate > now()
  ) then
    raise exception 'Chave de acesso corporativa inválida ou expirada.';
  end if;

  return new;
end;
$$;

drop trigger if exists trg_validar_chave_cadastro on auth.users;
create trigger trg_validar_chave_cadastro
  before insert on auth.users
  for each row execute function public.validar_chave_no_cadastro();

-- ─── 3. Chave inicial do cliente — VÁLIDA POR 2 ANOS ─────────────────────
-- (Pode trocar o texto da chave antes de rodar, se preferir outro código.)
insert into public.chaves_corporativas (chave, descricao, valida_ate)
values (
  'SUPORT-2026-K7M4-XR9D-QZ2B',
  'Chave corporativa — Suport Condomínios',
  now() + interval '2 years'
)
on conflict (chave) do nothing;

-- ─── Consultas úteis (não precisam rodar agora) ──────────────────────────
-- Ver chaves e validades:
--   select chave, descricao, valida_ate, ativa from public.chaves_corporativas;
-- Revogar uma chave imediatamente:
--   update public.chaves_corporativas set ativa = false where chave = '...';
-- Renovar por mais 2 anos (na renovação do contrato):
--   update public.chaves_corporativas
--     set valida_ate = now() + interval '2 years' where chave = '...';

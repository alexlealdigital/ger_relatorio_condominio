-- ═══════════════════════════════════════════════════════════════════════
-- Gestor Financeiro de Condomínios — Storage dos relatórios (v6)
-- Permite baixar os PPTX gerados a partir do histórico.
--
-- Como aplicar: Supabase Dashboard → SQL Editor → colar tudo → Run
-- (Execute UMA vez. Seguro rodar após o schema.sql original.)
-- ═══════════════════════════════════════════════════════════════════════

-- 1) Coluna com o caminho do arquivo no Storage
alter table public.relatorios
  add column if not exists arquivo_path text;

-- 2) Bucket privado para os PPTX
insert into storage.buckets (id, name, public)
values ('relatorios', 'relatorios', false)
on conflict (id) do nothing;

-- 3) Políticas: cada usuário lê/grava apenas a própria pasta
--    (os arquivos são salvos como {user_id}/{nome_do_arquivo}.pptx)
create policy "usuario_le_proprios_arquivos"
  on storage.objects for select to authenticated
  using (
    bucket_id = 'relatorios'
    and (storage.foldername(name))[1] = auth.uid()::text
  );

create policy "usuario_grava_proprios_arquivos"
  on storage.objects for insert to authenticated
  with check (
    bucket_id = 'relatorios'
    and (storage.foldername(name))[1] = auth.uid()::text
  );

create policy "usuario_apaga_proprios_arquivos"
  on storage.objects for delete to authenticated
  using (
    bucket_id = 'relatorios'
    and (storage.foldername(name))[1] = auth.uid()::text
  );

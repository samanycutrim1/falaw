-- Foto interna dos membros da equipe (advogados/equipe do escritório).
-- Usada SOMENTE no admin: Pauta de Audiências (advogado designado) e painel de
-- menções. Nunca é publicada no site: o admin não inclui este campo no
-- team-data.js público. Guarda um data URL JPEG já reduzido (~256px).
-- Rode UMA VEZ no Supabase: Dashboard → SQL Editor → New query.
alter table public.equipe
  add column if not exists foto text;

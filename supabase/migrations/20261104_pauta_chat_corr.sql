-- Chat entre o escritório (admin.html) e a Administradora de Correspondentes,
-- por audiência da pauta. A administradora (cliente/dashboard.html) só enxerga
-- a conversa de uma audiência a partir da primeira mensagem do escritório que
-- a marcou com @ (marca_corr = true); nada anterior a isso é exibido a ela.
CREATE TABLE IF NOT EXISTS pauta_chat_corr (
  id BIGSERIAL PRIMARY KEY,
  audiencia_id TEXT NOT NULL,
  origem TEXT NOT NULL CHECK (origem IN ('falaw', 'corr_admin')),
  autor TEXT,
  texto TEXT NOT NULL,
  marca_corr BOOLEAN NOT NULL DEFAULT FALSE,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_pauta_chat_corr_aud ON pauta_chat_corr (audiencia_id, created_at);

ALTER TABLE pauta_chat_corr ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "anon full access" ON pauta_chat_corr;
CREATE POLICY "anon full access" ON pauta_chat_corr FOR ALL TO anon USING (true) WITH CHECK (true);

-- Realtime (para o chat atualizar sozinho nas duas telas)
DO $$
BEGIN
  ALTER PUBLICATION supabase_realtime ADD TABLE public.pauta_chat_corr;
EXCEPTION WHEN duplicate_object THEN
  NULL;
END $$;

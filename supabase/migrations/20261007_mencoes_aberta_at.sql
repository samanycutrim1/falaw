-- Menções: registra quando a menção foi aberta (clicada no painel flutuante)
-- sem ainda ter recebido baixa no ✓ (lida_at). Permite destacar em cores
-- diferentes as não abertas (vermelho) e as abertas aguardando baixa (âmbar).
ALTER TABLE mencoes ADD COLUMN IF NOT EXISTS aberta_at TIMESTAMPTZ;

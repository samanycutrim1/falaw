-- Acesso individual da Equipe (login com e-mail e senha pelo Supabase Auth).
--
-- * equipe_acessos: um registro por pessoa da Equipe com login ativo ou não,
--   tipo de acesso (admin = painel admin com abas liberadas; dashboard = só o
--   dashboard do escritório) e a marca de "trocar a senha no primeiro login".
-- * Funções falaw_*: criar usuário, redefinir senha, atualizar acesso. Ficam
--   concentradas aqui e são chamadas só pelo helper acesso-api.js; para migrar
--   para uma Edge Function basta trocar o miolo desse helper.
-- * Superadmin de contingência: NÃO é e-mail escrito no código. É a conta do
--   Supabase Auth com app_metadata {"superadmin": true}. Para marcar (uma vez,
--   no SQL Editor, trocando pelo e-mail da conta criada em Authentication):
--
--     UPDATE auth.users
--        SET raw_app_meta_data = coalesce(raw_app_meta_data, '{}'::jsonb) || '{"superadmin": true}'::jsonb
--      WHERE email = '<EMAIL-DO-SUPERADMIN>';
--
--   (a pessoa precisa sair e entrar de novo para o token trazer a marca).
--
-- Segurança de linha (RLS) e chaves ficam fora desta migração.
-- As senhas nunca são gravadas em arquivo: chegam como parâmetro e viram hash bcrypt.

CREATE TABLE IF NOT EXISTS equipe_acessos (
  equipe_id   TEXT PRIMARY KEY,                       -- equipe.id
  user_id     UUID UNIQUE,                            -- auth.users.id
  email       TEXT NOT NULL,
  ativo       BOOLEAN NOT NULL DEFAULT true,
  tipo        TEXT NOT NULL DEFAULT 'dashboard' CHECK (tipo IN ('admin', 'dashboard')),
  abas        TEXT[] NOT NULL DEFAULT '{}',           -- abas liberadas (só tipo admin)
  must_change BOOLEAN NOT NULL DEFAULT true,          -- troca de senha pendente
  created_at  TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at  TIMESTAMPTZ NOT NULL DEFAULT now()
);
CREATE UNIQUE INDEX IF NOT EXISTS uq_equipe_acessos_email ON equipe_acessos (lower(email));

-- ── Auxiliares de permissão (dentro das funções) ─────────────────────────────
CREATE OR REPLACE FUNCTION falaw_is_superadmin() RETURNS boolean
LANGUAGE sql STABLE AS $$
  SELECT coalesce((auth.jwt() -> 'app_metadata' ->> 'superadmin')::boolean, false)
$$;

-- Quem pode gerenciar acessos: superadmin ou admin ativo com a aba "equipe".
CREATE OR REPLACE FUNCTION falaw_pode_gerenciar_acessos() RETURNS boolean
LANGUAGE sql STABLE SECURITY DEFINER SET search_path = public, auth AS $$
  SELECT falaw_is_superadmin() OR EXISTS (
    SELECT 1 FROM equipe_acessos a
     WHERE a.user_id = auth.uid() AND a.ativo AND a.tipo = 'admin' AND 'equipe' = ANY (a.abas)
  )
$$;

-- ── Meu acesso: um retorno único para o login decidir o que mostrar ─────────
CREATE OR REPLACE FUNCTION falaw_meu_acesso() RETURNS jsonb
LANGUAGE plpgsql STABLE SECURITY DEFINER SET search_path = public, auth AS $$
DECLARE r equipe_acessos;
BEGIN
  IF auth.uid() IS NULL THEN RETURN NULL; END IF;
  IF falaw_is_superadmin() THEN
    RETURN jsonb_build_object('superadmin', true, 'ativo', true, 'tipo', 'admin',
                              'abas', '[]'::jsonb, 'must_change', false);
  END IF;
  SELECT * INTO r FROM equipe_acessos WHERE user_id = auth.uid();
  IF NOT FOUND THEN RETURN NULL; END IF;
  RETURN jsonb_build_object('superadmin', false, 'equipe_id', r.equipe_id, 'email', r.email,
                            'ativo', r.ativo, 'tipo', r.tipo, 'abas', to_jsonb(r.abas),
                            'must_change', r.must_change);
END $$;

-- ── Criar usuário (senha padrão definida por quem cria) ──────────────────────
CREATE OR REPLACE FUNCTION falaw_criar_usuario(
  p_equipe_id TEXT, p_email TEXT, p_senha TEXT, p_tipo TEXT, p_abas TEXT[]
) RETURNS jsonb
LANGUAGE plpgsql SECURITY DEFINER SET search_path = public, auth, extensions AS $$
DECLARE
  v_email TEXT := lower(btrim(p_email));
  v_uid UUID := gen_random_uuid();
BEGIN
  IF NOT falaw_pode_gerenciar_acessos() THEN RAISE EXCEPTION 'sem_permissao'; END IF;
  IF v_email !~ '^[^\s@]+@[^\s@]+\.[^\s@]+$' THEN RAISE EXCEPTION 'email_invalido'; END IF;
  IF p_senha IS NULL OR length(p_senha) < 8 THEN RAISE EXCEPTION 'senha_curta'; END IF;
  IF p_tipo NOT IN ('admin', 'dashboard') THEN RAISE EXCEPTION 'tipo_invalido'; END IF;
  IF EXISTS (SELECT 1 FROM auth.users WHERE lower(email) = v_email) THEN RAISE EXCEPTION 'email_ja_existe'; END IF;
  IF EXISTS (SELECT 1 FROM equipe_acessos WHERE equipe_id = p_equipe_id) THEN RAISE EXCEPTION 'pessoa_ja_tem_acesso'; END IF;

  INSERT INTO auth.users (
    instance_id, id, aud, role, email, encrypted_password, email_confirmed_at,
    raw_app_meta_data, raw_user_meta_data, created_at, updated_at,
    confirmation_token, recovery_token, email_change_token_new, email_change
  ) VALUES (
    '00000000-0000-0000-0000-000000000000', v_uid, 'authenticated', 'authenticated', v_email,
    crypt(p_senha, gen_salt('bf')), now(),
    '{"provider":"email","providers":["email"]}'::jsonb, '{}'::jsonb, now(), now(),
    '', '', '', ''
  );
  INSERT INTO auth.identities (id, user_id, provider_id, identity_data, provider,
                               last_sign_in_at, created_at, updated_at)
  VALUES (gen_random_uuid(), v_uid, v_uid::text,
          jsonb_build_object('sub', v_uid::text, 'email', v_email, 'email_verified', true),
          'email', now(), now(), now());

  INSERT INTO equipe_acessos (equipe_id, user_id, email, ativo, tipo, abas, must_change)
  VALUES (p_equipe_id, v_uid, v_email, true, p_tipo,
          CASE WHEN p_tipo = 'admin' THEN coalesce(p_abas, '{}') ELSE '{}' END, true);

  RETURN jsonb_build_object('user_id', v_uid, 'email', v_email);
END $$;

-- ── Redefinir para a senha padrão (volta a exigir troca no próximo login) ───
CREATE OR REPLACE FUNCTION falaw_redefinir_senha(p_equipe_id TEXT, p_senha TEXT) RETURNS void
LANGUAGE plpgsql SECURITY DEFINER SET search_path = public, auth, extensions AS $$
DECLARE v_uid UUID;
BEGIN
  IF NOT falaw_pode_gerenciar_acessos() THEN RAISE EXCEPTION 'sem_permissao'; END IF;
  IF p_senha IS NULL OR length(p_senha) < 8 THEN RAISE EXCEPTION 'senha_curta'; END IF;
  SELECT user_id INTO v_uid FROM equipe_acessos WHERE equipe_id = p_equipe_id;
  IF v_uid IS NULL THEN RAISE EXCEPTION 'acesso_inexistente'; END IF;
  UPDATE auth.users SET encrypted_password = crypt(p_senha, gen_salt('bf')), updated_at = now() WHERE id = v_uid;
  UPDATE equipe_acessos SET must_change = true, updated_at = now() WHERE equipe_id = p_equipe_id;
END $$;

-- ── Ativar/desativar, tipo e abas ────────────────────────────────────────────
-- Desativar também bloqueia o usuário no Auth (não consegue nem renovar a sessão).
CREATE OR REPLACE FUNCTION falaw_atualizar_acesso(
  p_equipe_id TEXT, p_ativo BOOLEAN, p_tipo TEXT, p_abas TEXT[]
) RETURNS void
LANGUAGE plpgsql SECURITY DEFINER SET search_path = public, auth AS $$
DECLARE v_uid UUID;
BEGIN
  IF NOT falaw_pode_gerenciar_acessos() THEN RAISE EXCEPTION 'sem_permissao'; END IF;
  IF p_tipo NOT IN ('admin', 'dashboard') THEN RAISE EXCEPTION 'tipo_invalido'; END IF;
  UPDATE equipe_acessos
     SET ativo = p_ativo, tipo = p_tipo,
         abas = CASE WHEN p_tipo = 'admin' THEN coalesce(p_abas, '{}') ELSE '{}' END,
         updated_at = now()
   WHERE equipe_id = p_equipe_id
   RETURNING user_id INTO v_uid;
  IF v_uid IS NULL THEN RAISE EXCEPTION 'acesso_inexistente'; END IF;
  UPDATE auth.users SET banned_until = CASE WHEN p_ativo THEN NULL ELSE now() + interval '100 years' END
   WHERE id = v_uid;
END $$;

-- ── Primeiro login: pessoa conclui a troca (a senha nova vai por auth.updateUser)
CREATE OR REPLACE FUNCTION falaw_concluir_troca_senha() RETURNS void
LANGUAGE sql SECURITY DEFINER SET search_path = public, auth AS $$
  UPDATE equipe_acessos SET must_change = false, updated_at = now() WHERE user_id = auth.uid()
$$;

-- Só usuários autenticados chamam (o anon não).
REVOKE ALL ON FUNCTION falaw_is_superadmin(), falaw_pode_gerenciar_acessos(), falaw_meu_acesso(),
  falaw_criar_usuario(TEXT, TEXT, TEXT, TEXT, TEXT[]), falaw_redefinir_senha(TEXT, TEXT),
  falaw_atualizar_acesso(TEXT, BOOLEAN, TEXT, TEXT[]), falaw_concluir_troca_senha() FROM PUBLIC, anon;
GRANT EXECUTE ON FUNCTION falaw_is_superadmin(), falaw_pode_gerenciar_acessos(), falaw_meu_acesso(),
  falaw_criar_usuario(TEXT, TEXT, TEXT, TEXT, TEXT[]), falaw_redefinir_senha(TEXT, TEXT),
  falaw_atualizar_acesso(TEXT, BOOLEAN, TEXT, TEXT[]), falaw_concluir_troca_senha() TO authenticated;

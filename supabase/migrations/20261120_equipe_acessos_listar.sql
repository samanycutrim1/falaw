-- Lista os acessos da Equipe por função (SECURITY DEFINER), sem depender da regra de
-- leitura (RLS) da tabela equipe_acessos. Em projetos em que a tabela nasce com RLS ativa
-- e sem regra, o SELECT direto devolve vazio e a página Equipe mostraria "sem login"
-- mesmo com o acesso criado. Só superadmin ou admin com a aba Equipe consegue listar.
CREATE OR REPLACE FUNCTION falaw_listar_acessos()
RETURNS TABLE (equipe_id TEXT, email TEXT, ativo BOOLEAN, tipo TEXT, abas TEXT[], must_change BOOLEAN)
LANGUAGE plpgsql STABLE SECURITY DEFINER SET search_path = public, auth AS $$
BEGIN
  IF NOT falaw_pode_gerenciar_acessos() THEN RAISE EXCEPTION 'sem_permissao'; END IF;
  RETURN QUERY
    SELECT a.equipe_id, a.email, a.ativo, a.tipo, a.abas, a.must_change FROM equipe_acessos a;
END $$;

REVOKE ALL ON FUNCTION falaw_listar_acessos() FROM PUBLIC, anon;
GRANT EXECUTE ON FUNCTION falaw_listar_acessos() TO authenticated;

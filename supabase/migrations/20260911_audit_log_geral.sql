-- Auditoria completa de alterações na pauta de audiências: quem alterou
-- (e-mail quando logado via Supabase Auth, ou "escritório"/"cliente" pelo
-- app de origem), IP, navegador, dia/hora, e o antes/depois de cada
-- alteração — permitindo responder "quem mudou o quê, quando e quantas
-- vezes" para qualquer registro.

create table if not exists audit_log (
  id           bigint generated always as identity primary key,
  tabela       text        not null,
  registro_id  text        not null,
  operacao     text        not null check (operacao in ('INSERT','UPDATE','DELETE')),
  dados_antes  jsonb,
  dados_depois jsonb,
  autor_email  text,
  autor_tipo   text        not null default 'desconhecido'
                 check (autor_tipo in ('escritorio','cliente','sistema','desconhecido')),
  ip           text,
  user_agent   text,
  criado_em    timestamptz not null default now()
);

create index if not exists idx_audit_log_tabela_registro
  on audit_log (tabela, registro_id, criado_em desc);

alter table audit_log enable row level security;

-- Só usuários com login individual real (Supabase Auth) podem ler o
-- audit_log. NUNCA liberar para "anon" aqui: o dashboard do cliente usa a
-- mesma anon key publicamente, e isso vazaria o histórico de todos os
-- casos (de todos os clientes) para quem abrir o console do navegador.
drop policy if exists "authenticated read audit_log" on audit_log;
create policy "authenticated read audit_log" on audit_log
  for select to authenticated using (true);

-- Função genérica: anexável em qualquer tabela. Grava o estado completo
-- da linha antes/depois, quem fez (e-mail do JWT quando autenticado,
-- senão o app de origem via header X-Client-App), IP e user-agent.
create or replace function trg_generic_audit_log()
returns trigger
language plpgsql
security definer
as $$
declare
  v_headers json := nullif(current_setting('request.headers', true), '')::json;
  v_ip      text;
  v_ua      text;
  v_app     text;
  v_email   text;
  v_tipo    text;
  v_id      text;
begin
  v_ip    := coalesce(split_part(v_headers->>'x-forwarded-for', ',', 1), inet_client_addr()::text);
  v_ua    := v_headers->>'user-agent';
  v_app   := lower(coalesce(v_headers->>'x-client-app', ''));

  begin
    v_email := auth.email();
  exception when others then
    v_email := null;
  end;

  if v_email is not null then
    v_tipo := 'escritorio';
  elsif v_app = 'admin' then
    v_tipo := 'escritorio';
  elsif v_app = 'cliente' then
    v_tipo := 'cliente';
  elsif coalesce(current_setting('request.jwt.claims', true), '')::json->>'role' = 'service_role' then
    v_tipo := 'sistema';
  else
    v_tipo := 'desconhecido';
  end if;

  v_id := coalesce((case when tg_op = 'DELETE' then old.id else new.id end)::text, '');

  insert into audit_log (tabela, registro_id, operacao, dados_antes, dados_depois, autor_email, autor_tipo, ip, user_agent)
  values (
    tg_table_name,
    v_id,
    tg_op,
    case when tg_op = 'INSERT' then null else to_jsonb(old) end,
    case when tg_op = 'DELETE' then null else to_jsonb(new) end,
    v_email,
    v_tipo,
    v_ip,
    v_ua
  );

  return coalesce(new, old);
end;
$$;

drop trigger if exists pauta_audiencias_audit on pauta_audiencias;
create trigger pauta_audiencias_audit
after insert or update or delete on pauta_audiencias
for each row execute function trg_generic_audit_log();

-- Corrige acesso: a conta audiencias@falaw.com.br (e qualquer outra) que
-- loga via Supabase Auth (JWT "authenticated") não tinha nenhuma policy
-- nessa tabela — só existia "anon full access". Sem isso, quem usa login
-- individual não conseguia sequer salvar uma audiência.
drop policy if exists "authenticated full access" on pauta_audiencias;
create policy "authenticated full access" on pauta_audiencias
  for all to authenticated using (true) with check (true);

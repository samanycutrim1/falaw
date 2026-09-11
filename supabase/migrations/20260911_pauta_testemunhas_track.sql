-- Registra automaticamente IP, user-agent e timestamp de quando o campo
-- `testemunhas` de uma audiência é alterado, para permitir auditar se a
-- informação foi inserida pelo escritório ou pelo cliente.

alter table pauta_audiencias
  add column if not exists testemunhas_ip          text,
  add column if not exists testemunhas_user_agent   text,
  add column if not exists testemunhas_alterado_em  timestamptz;

create or replace function trg_pauta_testemunhas_track()
returns trigger
language plpgsql
security definer
as $$
begin
  if new.testemunhas is distinct from old.testemunhas then
    new.testemunhas_ip := coalesce(
      split_part(current_setting('request.headers', true)::json->>'x-forwarded-for', ',', 1),
      inet_client_addr()::text
    );
    new.testemunhas_user_agent := current_setting('request.headers', true)::json->>'user-agent';
    new.testemunhas_alterado_em := now();
  end if;
  return new;
end;
$$;

drop trigger if exists pauta_testemunhas_track on pauta_audiencias;
create trigger pauta_testemunhas_track
before update on pauta_audiencias
for each row execute function trg_pauta_testemunhas_track();

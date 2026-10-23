-- Schéma du workflow de demandes d'attestations administratives.
-- À exécuter une seule fois dans Supabase > SQL Editor.

create table if not exists public.attestation_requests (
    id uuid primary key,
    created_at timestamptz not null default now(),
    updated_at timestamptz not null default now(),
    department text not null,
    nom text not null,
    prenom text not null,
    date_naissance text not null,
    lieu_naissance text not null,
    email text not null default '',
    telephone text not null default '',
    purpose text not null default '',
    language_mode text not null default 'Bilingue (arabe puis français)',
    status text not null default 'En attente',
    final_data jsonb,
    attestation_number integer,
    issued_date text,
    issued_year integer,
    constraint attestation_requests_issued_number_unique unique (issued_year, attestation_number)
);

create index if not exists attestation_requests_dept_status_created_idx
on public.attestation_requests (department, status, created_at desc);

create table if not exists public.attestation_sequences (
    issue_year integer primary key,
    last_number integer not null
);

alter table public.attestation_requests enable row level security;
alter table public.attestation_sequences enable row level security;
-- Ne créez aucune politique pour anon/authenticated. L'application utilise la clé service role côté serveur.
-- Gardez SUPABASE_SERVICE_ROLE_KEY dans Streamlit Secrets uniquement, jamais dans GitHub.

grant all on public.attestation_requests to service_role;
grant all on public.attestation_sequences to service_role;

create or replace function public.issue_attestation(p_request_id uuid, p_final_data jsonb)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
    req public.attestation_requests%rowtype;
    v_year integer;
    v_number integer;
    v_date date := (now() at time zone 'Africa/Algiers')::date;
begin
    select * into req
      from public.attestation_requests
     where id = p_request_id
     for update;

    if not found then
        raise exception 'Demande introuvable';
    end if;

    -- Un re-téléchargement conserve numéro/date et ne consomme pas un nouveau numéro.
    if req.status = 'Générée' and req.attestation_number is not null then
        return to_jsonb(req);
    end if;

    v_year := extract(year from v_date)::integer;
    insert into public.attestation_sequences(issue_year, last_number)
    values (v_year, 1)
    on conflict (issue_year)
    do update set last_number = public.attestation_sequences.last_number + 1
    returning last_number into v_number;

    update public.attestation_requests
       set final_data = p_final_data,
           status = 'Générée',
           attestation_number = v_number,
           issued_date = to_char(v_date, 'DD/MM/YYYY'),
           issued_year = v_year,
           updated_at = now()
     where id = p_request_id
     returning * into req;

    return to_jsonb(req);
end;
$$;

revoke all on function public.issue_attestation(uuid, jsonb) from public, anon, authenticated;
grant execute on function public.issue_attestation(uuid, jsonb) to service_role;

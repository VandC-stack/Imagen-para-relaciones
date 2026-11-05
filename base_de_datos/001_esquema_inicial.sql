-- Esquema inicial de Supabase para el Sistema Generador de Documentos V&C.
--
-- Las contraseñas no están aquí: las administra Supabase Auth (auth.users).
-- Los PDF y los JSON completos de cada dictamen se quedan en Dropbox.
-- La app se conecta con la clave publicable, así que todo el control de acceso
-- está en este archivo. Cualquier tabla nueva debe llevar RLS.

begin;

-- ---------------------------------------------------------------- Usuarios
-- Un perfil por cuenta de Auth. Se escribe solo desde el servidor (función
-- Edge con clave secreta), porque crear cuentas requiere esa clave.
create table public.perfiles (
    id         uuid primary key references auth.users (id) on delete cascade,
    username   text not null check (username ~ '^[A-Za-z0-9_.-]+$'),
    nombre     text not null,
    rol        text not null check (rol in ('admin', 'ejecutivo')),
    accesos    text[] not null check (accesos <@ array[
                   'principal', 'historial', 'clientes', 'inspectores',
                   'reportes', 'reporte_ema', 'usuarios'
               ]::text[]),
    activo     boolean not null default true,
    creado_en  timestamptz not null default now()
);
create unique index perfiles_username_unico on public.perfiles (lower(username));

-- ---------------------------------------------------------------- Catálogos
create table public.normas (
    id       bigint generated always as identity primary key,
    nom      text not null unique,
    nombre   text,
    capitulo text
);

-- El RFC no es único: un mismo RFC puede tener más de un registro de cliente.
create table public.clientes (
    id              bigint generated always as identity primary key,
    rfc             text not null,
    nombre          text not null,
    numero_contrato text,
    fecha_contrato  date,
    actividad       text,
    curp            text,
    creado_en       timestamptz not null default now()
);
create index clientes_rfc on public.clientes (rfc);

create table public.cliente_direcciones (
    id           bigint generated always as identity primary key,
    cliente_id   bigint not null references public.clientes (id) on delete cascade,
    calle_numero text,
    colonia      text,
    municipio    text,
    estado       text,
    cp           text,
    servicio     text
);
create index cliente_direcciones_cliente on public.cliente_direcciones (cliente_id);

create table public.inspectores (
    id                 bigint generated always as identity primary key,
    nombre             text not null,
    correo             text,
    firma              text not null,
    puesto             text,
    vigencia           text,
    fecha_acreditacion date,
    referencia         text,
    imagen             text,   -- ruta relativa, p. ej. Firmas/ASANCHEZ.png
    creado_en          timestamptz not null default now()
);
create unique index inspectores_firma_unica on public.inspectores (upper(firma));

create table public.inspector_normas (
    inspector_id bigint not null references public.inspectores (id) on delete cascade,
    norma_id     bigint not null references public.normas (id) on delete restrict,
    primary key (inspector_id, norma_id)
);
create index inspector_normas_norma on public.inspector_normas (norma_id);

-- ---------------------------------------------------------------- Folios
-- Solo se modifica con reservar_folios y liberar_folios.
create table public.contadores_folio (
    clave  text primary key check (clave in ('visita', 'documento')),
    ultimo bigint not null default 0 check (ultimo >= 0)
);
insert into public.contadores_folio (clave) values ('visita'), ('documento');

-- ---------------------------------------------------------------- Reporte EMA
-- Una fila por documento emitido. Cliente e inspector se guardan como texto a
-- propósito: el reporte debe mostrar lo que decía el documento al emitirse,
-- aunque el catálogo cambie después. SOL EMA, FOLIO EMA e INSP EMA se derivan
-- de estas columnas al armar el Excel.
create table public.ema_documentos (
    tipo_documento        text not null check (tipo_documento in (
                              'Dictamen', 'Constancia',
                              'Negación de Dictamen', 'Negación de Constancia'
                          )),
    folio                 bigint not null check (folio > 0),
    solicitud             text not null,
    folio_visita          text,
    estatus               text,
    documento_emitido     text,
    fecha_documento       date,
    cliente               text,
    numero_contrato       text,
    rfc                   text,
    producto              text,
    marcas                text,
    modelos               text,
    norma                 text,
    verificador           text,
    pedimento             text,
    fecha_desaduanamiento date,
    fecha_visita          date,
    registrado_por        uuid references public.perfiles (id) on delete set null,
    registrado_en         timestamptz not null default now(),
    primary key (tipo_documento, folio)
);
create index ema_documentos_fecha on public.ema_documentos (fecha_documento);
create index ema_documentos_visita on public.ema_documentos (folio_visita);

-- ---------------------------------------------------------------- Permisos
create function public.tiene_acceso(p_acceso text)
returns boolean language sql stable security definer set search_path = ''
as $$
    select exists (
        select 1 from public.perfiles
        where id = (select auth.uid()) and activo and p_acceso = any (accesos)
    );
$$;

create function public.usuario_activo()
returns boolean language sql stable security definer set search_path = ''
as $$
    select exists (
        select 1 from public.perfiles where id = (select auth.uid()) and activo
    );
$$;

-- Captura y edición de visitas: las dos secciones que emiten o eliminan documentos.
create function public.puede_operar_documentos()
returns boolean language sql stable security definer set search_path = ''
as $$
    select public.tiene_acceso('principal') or public.tiene_acceso('historial');
$$;

-- ---------------------------------------------------------------- Operaciones
-- UPDATE ... RETURNING toma el bloqueo de la fila: dos PCs que reservan al mismo
-- tiempo se atienden una tras otra y nunca reciben el mismo folio.
create function public.reservar_folios(p_clave text, p_cantidad integer default 1)
returns bigint language plpgsql security definer set search_path = ''
as $$
declare
    v_ultimo bigint;
begin
    if not public.puede_operar_documentos() then
        raise exception 'Sin acceso para reservar folios' using errcode = '42501';
    end if;
    if p_cantidad is null or p_cantidad < 1 then
        raise exception 'La cantidad debe ser mayor a cero' using errcode = '22023';
    end if;
    update public.contadores_folio set ultimo = ultimo + p_cantidad
    where clave = p_clave returning ultimo into v_ultimo;
    if v_ultimo is null then
        raise exception 'Contador desconocido: %', p_clave using errcode = '22023';
    end if;
    return v_ultimo - p_cantidad + 1;   -- primer folio del bloque
end;
$$;

-- Regresa el contador solo si el bloque es lo último emitido. Si ya hay folios
-- posteriores devuelve false y el contador no cambia: bajarlo reasignaría folios
-- que otra persona ya usó.
create function public.liberar_folios(p_clave text, p_desde bigint, p_hasta bigint)
returns boolean language plpgsql security definer set search_path = ''
as $$
begin
    if not public.puede_operar_documentos() then
        raise exception 'Sin acceso para liberar folios' using errcode = '42501';
    end if;
    if p_desde is null or p_hasta is null or p_desde < 1 or p_hasta < p_desde then
        raise exception 'Rango de folios inválido' using errcode = '22023';
    end if;
    update public.contadores_folio set ultimo = p_desde - 1
    where clave = p_clave and ultimo = p_hasta;
    return found;
end;
$$;

-- Recibe todos los documentos de una generación en una sola llamada. Al regenerar
-- un documento se reemplaza su fila.
create function public.registrar_documentos(p_documentos jsonb)
returns integer language plpgsql security definer set search_path = ''
as $$
declare
    v_total integer;
begin
    if not public.puede_operar_documentos() then
        raise exception 'Sin acceso para registrar documentos' using errcode = '42501';
    end if;
    insert into public.ema_documentos (
        tipo_documento, folio, solicitud, folio_visita, estatus, documento_emitido,
        fecha_documento, cliente, numero_contrato, rfc, producto, marcas, modelos,
        norma, verificador, pedimento, fecha_desaduanamiento, fecha_visita, registrado_por
    )
    select
        d.tipo_documento, d.folio, d.solicitud, d.folio_visita, d.estatus, d.documento_emitido,
        d.fecha_documento, d.cliente, d.numero_contrato, d.rfc, d.producto, d.marcas, d.modelos,
        d.norma, d.verificador, d.pedimento, d.fecha_desaduanamiento, d.fecha_visita, (select auth.uid())
    from jsonb_to_recordset(p_documentos) as d (
        tipo_documento text, folio bigint, solicitud text, folio_visita text, estatus text,
        documento_emitido text, fecha_documento date, cliente text, numero_contrato text,
        rfc text, producto text, marcas text, modelos text, norma text, verificador text,
        pedimento text, fecha_desaduanamiento date, fecha_visita date
    )
    on conflict (tipo_documento, folio) do update set
        solicitud = excluded.solicitud,
        folio_visita = excluded.folio_visita,
        estatus = excluded.estatus,
        documento_emitido = excluded.documento_emitido,
        fecha_documento = excluded.fecha_documento,
        cliente = excluded.cliente,
        numero_contrato = excluded.numero_contrato,
        rfc = excluded.rfc,
        producto = excluded.producto,
        marcas = excluded.marcas,
        modelos = excluded.modelos,
        norma = excluded.norma,
        verificador = excluded.verificador,
        pedimento = excluded.pedimento,
        fecha_desaduanamiento = excluded.fecha_desaduanamiento,
        fecha_visita = excluded.fecha_visita,
        registrado_por = excluded.registrado_por,
        registrado_en = now();
    get diagnostics v_total = row_count;
    return v_total;
end;
$$;

create function public.actualizar_estatus_visita(p_folio_visita text, p_estatus text)
returns integer language plpgsql security definer set search_path = ''
as $$
declare
    v_total integer;
begin
    if not public.puede_operar_documentos() then
        raise exception 'Sin acceso para actualizar documentos' using errcode = '42501';
    end if;
    update public.ema_documentos set estatus = p_estatus where folio_visita = p_folio_visita;
    get diagnostics v_total = row_count;
    return v_total;
end;
$$;

create function public.eliminar_documentos_visita(p_folio_visita text)
returns integer language plpgsql security definer set search_path = ''
as $$
declare
    v_total integer;
begin
    if not public.tiene_acceso('historial') then
        raise exception 'Sin acceso para eliminar documentos' using errcode = '42501';
    end if;
    delete from public.ema_documentos where folio_visita = p_folio_visita;
    get diagnostics v_total = row_count;
    return v_total;
end;
$$;

-- ---------------------------------------------------------------- RLS
alter table public.perfiles            enable row level security;
alter table public.normas              enable row level security;
alter table public.clientes            enable row level security;
alter table public.cliente_direcciones enable row level security;
alter table public.inspectores         enable row level security;
alter table public.inspector_normas    enable row level security;
alter table public.contadores_folio    enable row level security;
alter table public.ema_documentos      enable row level security;

create policy ver_perfiles on public.perfiles for select to authenticated
    using (id = (select auth.uid()) or (select public.tiene_acceso('usuarios')));

-- Todos leen los catálogos porque se necesitan para generar documentos, aunque
-- la sección de administración de cada uno esté restringida.
create policy leer on public.normas              for select to authenticated using ((select public.usuario_activo()));
create policy leer on public.clientes            for select to authenticated using ((select public.usuario_activo()));
create policy leer on public.cliente_direcciones for select to authenticated using ((select public.usuario_activo()));
create policy leer on public.inspectores         for select to authenticated using ((select public.usuario_activo()));
create policy leer on public.inspector_normas    for select to authenticated using ((select public.usuario_activo()));
create policy leer on public.contadores_folio    for select to authenticated using ((select public.usuario_activo()));

create policy administrar on public.clientes for all to authenticated
    using ((select public.tiene_acceso('clientes'))) with check ((select public.tiene_acceso('clientes')));
create policy administrar on public.cliente_direcciones for all to authenticated
    using ((select public.tiene_acceso('clientes'))) with check ((select public.tiene_acceso('clientes')));
create policy administrar on public.inspectores for all to authenticated
    using ((select public.tiene_acceso('inspectores'))) with check ((select public.tiene_acceso('inspectores')));
create policy administrar on public.inspector_normas for all to authenticated
    using ((select public.tiene_acceso('inspectores'))) with check ((select public.tiene_acceso('inspectores')));

create policy consultar on public.ema_documentos for select to authenticated
    using ((select public.tiene_acceso('reporte_ema')));

-- ---------------------------------------------------------------- Privilegios
-- Supabase da privilegios por defecto a anon y authenticated; aquí se dejan solo
-- los necesarios. Las escrituras de folios y documentos pasan por las funciones.
revoke all on all tables in schema public from anon, authenticated;
revoke all on all functions in schema public from public, anon, authenticated;

grant select on public.perfiles, public.normas, public.contadores_folio, public.ema_documentos to authenticated;
grant select, insert, update, delete on
    public.clientes, public.cliente_direcciones, public.inspectores, public.inspector_normas
    to authenticated;
grant execute on function
    public.tiene_acceso(text),
    public.usuario_activo(),
    public.puede_operar_documentos(),
    public.reservar_folios(text, integer),
    public.liberar_folios(text, bigint, bigint),
    public.registrar_documentos(jsonb),
    public.actualizar_estatus_visita(text, text),
    public.eliminar_documentos_visita(text)
    to authenticated;

commit;

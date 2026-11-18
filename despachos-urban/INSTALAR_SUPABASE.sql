-- DESPACHOS URBAN - Ejecutar en Supabase SQL Editor ANTES de publicar el sitio.
-- No elimina ni modifica registros existentes de transportes u otros módulos.
CREATE TABLE IF NOT EXISTS public.ct_urban_permissions (
 user_id uuid PRIMARY KEY REFERENCES public.usuarios(id) ON DELETE CASCADE,
 can_view boolean NOT NULL DEFAULT false,
 can_add boolean NOT NULL DEFAULT false,
 can_dispatch boolean NOT NULL DEFAULT false
);
CREATE TABLE IF NOT EXISTS public.ct_urban_dispatches (
 id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
 fecha date NOT NULL,
 pedido text NOT NULL,
 entrega text NOT NULL,
 transporte text NOT NULL,
 placa_documento text NOT NULL,
 placa_fisica text NOT NULL,
 estado text NOT NULL DEFAULT 'PENDIENTE' CHECK(estado IN ('PENDIENTE','DESPACHADO')),
 creado_por uuid REFERENCES public.usuarios(id),
 creado_en timestamptz NOT NULL DEFAULT now(),
 despachado_por uuid REFERENCES public.usuarios(id),
 despachado_en timestamptz
);
CREATE INDEX IF NOT EXISTS ct_urban_dispatches_estado_idx ON public.ct_urban_dispatches(estado,fecha DESC,creado_en DESC);
ALTER TABLE public.ct_urban_permissions ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.ct_urban_dispatches ENABLE ROW LEVEL SECURITY;
REVOKE ALL ON public.ct_urban_permissions, public.ct_urban_dispatches FROM PUBLIC, anon, authenticated;
-- Los accesos a tablas son únicamente por funciones SECURITY DEFINER que comprueban PIN.
CREATE OR REPLACE FUNCTION public.ct_urban_rights(p_pin text) RETURNS jsonb
LANGUAGE plpgsql STABLE SECURITY DEFINER SET search_path=public,extensions AS $$
DECLARE a public.usuarios%ROWTYPE; p public.ct_urban_permissions%ROWTYPE;
BEGIN
 SELECT * INTO a FROM public._transport_user_by_pin(p_pin);
 IF a.id IS NULL THEN RETURN jsonb_build_object('can_view',false,'can_add',false,'can_dispatch',false,'is_admin',false); END IF;
 SELECT * INTO p FROM public.ct_urban_permissions WHERE user_id=a.id;
 RETURN jsonb_build_object('can_view',a.is_admin OR coalesce(p.can_view,false) OR coalesce(p.can_add,false) OR coalesce(p.can_dispatch,false),
 'can_add',a.is_admin OR coalesce(p.can_add,false), 'can_dispatch',a.is_admin OR coalesce(p.can_dispatch,false),
 'is_admin',a.is_admin,'user_name',a.name);
END;$$;
CREATE OR REPLACE FUNCTION public.ct_urban_list(p_pin text) RETURNS jsonb
LANGUAGE plpgsql STABLE SECURITY DEFINER SET search_path=public,extensions AS $$
DECLARE permission jsonb; output jsonb;
BEGIN
 permission:=public.ct_urban_rights(p_pin);
 IF NOT (permission->>'can_view')::boolean THEN RAISE EXCEPTION 'Sin autorización para consultar Despachos Urban'; END IF;
 SELECT coalesce(jsonb_agg(to_jsonb(d) || jsonb_build_object('despachado_por',u.name) ORDER BY d.fecha DESC,d.creado_en DESC),'[]'::jsonb) INTO output
 FROM public.ct_urban_dispatches d LEFT JOIN public.usuarios u ON u.id=d.despachado_por;
 RETURN output;
END;$$;
CREATE OR REPLACE FUNCTION public.ct_urban_add(p_pin text,p_row jsonb) RETURNS uuid
LANGUAGE plpgsql VOLATILE SECURITY DEFINER SET search_path=public,extensions AS $$
DECLARE a public.usuarios%ROWTYPE; permit jsonb; new_id uuid; k text; val text;
BEGIN
 SELECT * INTO a FROM public._transport_user_by_pin(p_pin);permit:=public.ct_urban_rights(p_pin);
 IF a.id IS NULL OR NOT (permit->>'can_add')::boolean THEN RAISE EXCEPTION 'Sin permiso para agregar despachos'; END IF;
 IF jsonb_typeof(p_row) IS DISTINCT FROM 'object' THEN RAISE EXCEPTION 'Registro inválido'; END IF;
 FOREACH k IN ARRAY ARRAY['fecha','pedido','entrega','transporte','placa_documento','placa_fisica'] LOOP
  val:=btrim(coalesce(p_row->>k,''));
  IF length(val)=0 OR length(val)>100 THEN RAISE EXCEPTION 'Campo obligatorio o demasiado largo: %',k; END IF;
 END LOOP;
 INSERT INTO public.ct_urban_dispatches(fecha,pedido,entrega,transporte,placa_documento,placa_fisica,creado_por)
 VALUES ((p_row->>'fecha')::date,btrim(p_row->>'pedido'),btrim(p_row->>'entrega'),btrim(p_row->>'transporte'),upper(btrim(p_row->>'placa_documento')),upper(btrim(p_row->>'placa_fisica')),a.id)
 RETURNING id INTO new_id;RETURN new_id;
END;$$;
CREATE OR REPLACE FUNCTION public.ct_urban_dispatch(p_pin text,p_ids uuid[]) RETURNS jsonb
LANGUAGE plpgsql VOLATILE SECURITY DEFINER SET search_path=public,extensions AS $$
DECLARE a public.usuarios%ROWTYPE; permit jsonb; n integer;
BEGIN
 SELECT * INTO a FROM public._transport_user_by_pin(p_pin);permit:=public.ct_urban_rights(p_pin);
 IF a.id IS NULL OR NOT (permit->>'can_dispatch')::boolean THEN RAISE EXCEPTION 'Sin permiso para confirmar despachos'; END IF;
 IF coalesce(array_length(p_ids,1),0)>500 THEN RAISE EXCEPTION 'Máximo 500 registros por operación'; END IF;
 UPDATE public.ct_urban_dispatches SET estado='DESPACHADO',despachado_por=a.id,despachado_en=now()
 WHERE id=ANY(p_ids) AND estado='PENDIENTE';GET DIAGNOSTICS n=ROW_COUNT;
 RETURN jsonb_build_object('updated',n);
END;$$;
CREATE OR REPLACE FUNCTION public.ct_urban_users(p_admin_pin text) RETURNS jsonb
LANGUAGE plpgsql STABLE SECURITY DEFINER SET search_path=public,extensions AS $$
DECLARE a public.usuarios%ROWTYPE; output jsonb;
BEGIN
 SELECT * INTO a FROM public._transport_user_by_pin(p_admin_pin);
 IF a.id IS NULL OR NOT a.is_admin THEN RAISE EXCEPTION 'Solo el administrador puede ver permisos'; END IF;
 SELECT coalesce(jsonb_agg(jsonb_build_object('user_id',u.id,'user_name',u.name,'active',u.active,'is_admin',u.is_admin,
 'can_view',u.is_admin OR coalesce(p.can_view,false), 'can_add',u.is_admin OR coalesce(p.can_add,false),
 'can_dispatch',u.is_admin OR coalesce(p.can_dispatch,false)) ORDER BY u.name),'[]'::jsonb) INTO output
 FROM public.usuarios u LEFT JOIN public.ct_urban_permissions p ON p.user_id=u.id;
 RETURN output;
END;$$;
CREATE OR REPLACE FUNCTION public.ct_urban_set_permission(p_admin_pin text,p_user_id uuid,p_can_view boolean,p_can_add boolean,p_can_dispatch boolean) RETURNS void
LANGUAGE plpgsql VOLATILE SECURITY DEFINER SET search_path=public,extensions AS $$
DECLARE a public.usuarios%ROWTYPE; target public.usuarios%ROWTYPE;
BEGIN
 SELECT * INTO a FROM public._transport_user_by_pin(p_admin_pin);
 IF a.id IS NULL OR NOT a.is_admin THEN RAISE EXCEPTION 'Solo el administrador puede gestionar permisos'; END IF;
 SELECT * INTO target FROM public.usuarios WHERE id=p_user_id;
 IF target.id IS NULL OR target.is_admin THEN RAISE EXCEPTION 'Usuario no permitido'; END IF;
 INSERT INTO public.ct_urban_permissions(user_id,can_view,can_add,can_dispatch)
 VALUES(p_user_id,coalesce(p_can_view,false) OR coalesce(p_can_add,false) OR coalesce(p_can_dispatch,false),coalesce(p_can_add,false),coalesce(p_can_dispatch,false))
 ON CONFLICT(user_id) DO UPDATE SET can_view=excluded.can_view,can_add=excluded.can_add,can_dispatch=excluded.can_dispatch;
END;$$;
GRANT EXECUTE ON FUNCTION public.ct_urban_rights(text) TO anon, authenticated;
GRANT EXECUTE ON FUNCTION public.ct_urban_list(text) TO anon, authenticated;
GRANT EXECUTE ON FUNCTION public.ct_urban_add(text,jsonb) TO anon, authenticated;
GRANT EXECUTE ON FUNCTION public.ct_urban_dispatch(text,uuid[]) TO anon, authenticated;
GRANT EXECUTE ON FUNCTION public.ct_urban_users(text) TO anon, authenticated;
GRANT EXECUTE ON FUNCTION public.ct_urban_set_permission(text,uuid,boolean,boolean,boolean) TO anon, authenticated;

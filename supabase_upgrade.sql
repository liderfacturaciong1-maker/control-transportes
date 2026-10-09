-- ACTUALIZACIÓN SEGURA DE PERMISOS Y ROLES
-- Ejecutar completo en Supabase > SQL Editor ANTES de publicar los archivos.
-- Hace copia lógica de los datos existentes: no elimina transportes al importar.

ALTER TABLE public.usuarios
  ADD COLUMN IF NOT EXISTS role text NOT NULL DEFAULT 'Operativo',
  ADD COLUMN IF NOT EXISTS can_upload boolean NOT NULL DEFAULT false,
  ADD COLUMN IF NOT EXISTS can_create_users boolean NOT NULL DEFAULT false;

-- Normalizar roles existentes. Los administradores conservan ambos permisos.
UPDATE public.usuarios
SET role = CASE WHEN is_admin THEN 'Administrador' ELSE COALESCE(NULLIF(role,''),'Operativo') END,
    can_upload = CASE WHEN is_admin THEN true ELSE can_upload END,
    can_create_users = CASE WHEN is_admin THEN true ELSE can_create_users END;

-- Verifica el PIN y devuelve únicamente un usuario activo.
CREATE OR REPLACE FUNCTION public._transport_user_by_pin(p_pin text)
RETURNS SETOF public.usuarios
LANGUAGE sql STABLE SECURITY DEFINER
SET search_path = public, extensions
AS $$
  SELECT u.* FROM public.usuarios u
  WHERE u.active = true
    AND u.pin_hash = extensions.crypt(p_pin, u.pin_hash)
  LIMIT 1;
$$;
REVOKE ALL ON FUNCTION public._transport_user_by_pin(text) FROM PUBLIC, anon, authenticated;

CREATE OR REPLACE FUNCTION public.get_user_permissions_by_pin(p_pin text)
RETURNS TABLE(is_admin boolean, can_upload boolean, can_create_users boolean, role text)
LANGUAGE plpgsql STABLE SECURITY DEFINER
SET search_path = public
AS $$
DECLARE actor public.usuarios%ROWTYPE;
BEGIN
  SELECT * INTO actor FROM public._transport_user_by_pin(p_pin);
  IF actor.id IS NULL THEN RAISE EXCEPTION 'PIN inválido o usuario inactivo'; END IF;
  RETURN QUERY SELECT actor.is_admin, (actor.is_admin OR actor.can_upload),
                      (actor.is_admin OR actor.can_create_users),
                      CASE WHEN actor.is_admin THEN 'Administrador' ELSE actor.role END;
END;
$$;
GRANT EXECUTE ON FUNCTION public.get_user_permissions_by_pin(text) TO anon, authenticated;

-- Importación segura: agrega solamente transportes que todavía no existen.
-- Devuelve cantidades para la vista previa final; nunca ejecuta DELETE.
CREATE OR REPLACE FUNCTION public.replace_transports_secure(p_rows jsonb, p_pin text)
RETURNS jsonb
LANGUAGE plpgsql SECURITY DEFINER
SET search_path = public, extensions
AS $$
DECLARE actor public.usuarios%ROWTYPE; added integer := 0; total_rows integer := 0;
BEGIN
  SELECT * INTO actor FROM public._transport_user_by_pin(p_pin);
  IF actor.id IS NULL THEN RAISE EXCEPTION 'PIN inválido o usuario inactivo'; END IF;
  IF NOT (actor.is_admin OR actor.can_upload) THEN RAISE EXCEPTION 'No tienes permiso para cargar archivos'; END IF;
  IF jsonb_typeof(p_rows) <> 'array' THEN RAISE EXCEPTION 'El archivo no contiene una lista válida'; END IF;
  SELECT count(*) INTO total_rows FROM jsonb_array_elements(p_rows);

  INSERT INTO public.transportes(transporte, viaje, placa)
  SELECT DISTINCT trim(x->>'transporte'), trim(x->>'viaje'), trim(x->>'placa')
  FROM jsonb_array_elements(p_rows) x
  WHERE (COALESCE(trim(x->>'transporte'),'') <> ''
     OR COALESCE(trim(x->>'viaje'),'') <> ''
     OR COALESCE(trim(x->>'placa'),'') <> '')
    AND NOT EXISTS (
      SELECT 1 FROM public.transportes t
      WHERE lower(trim(COALESCE(t.transporte,''))) = lower(trim(COALESCE(x->>'transporte','')))
        AND lower(trim(COALESCE(t.viaje,''))) = lower(trim(COALESCE(x->>'viaje','')))
        AND lower(trim(COALESCE(t.placa,''))) = lower(trim(COALESCE(x->>'placa','')))
    );
  GET DIAGNOSTICS added = ROW_COUNT;
  RETURN jsonb_build_object('added', added, 'existing', GREATEST(total_rows-added,0));
END;
$$;
GRANT EXECUTE ON FUNCTION public.replace_transports_secure(jsonb,text) TO anon, authenticated;

-- Crear usuario. Un creador autorizado puede crear Operativos/Supervisores,
-- pero solo un administrador puede crear administradores o conceder permisos.
CREATE OR REPLACE FUNCTION public.admin_create_user_secure(
  p_name text, p_pin text, p_creator_pin text, p_role text,
  p_can_upload boolean DEFAULT false, p_can_create_users boolean DEFAULT false
) RETURNS uuid
LANGUAGE plpgsql SECURITY DEFINER
SET search_path = public, extensions
AS $$
DECLARE actor public.usuarios%ROWTYPE; new_id uuid; admin_role boolean;
BEGIN
  SELECT * INTO actor FROM public._transport_user_by_pin(p_creator_pin);
  IF actor.id IS NULL THEN RAISE EXCEPTION 'PIN del creador inválido o usuario inactivo'; END IF;
  IF NOT (actor.is_admin OR actor.can_create_users) THEN RAISE EXCEPTION 'No tienes permiso para crear usuarios'; END IF;
  IF length(trim(COALESCE(p_name,''))) = 0 THEN RAISE EXCEPTION 'El nombre es obligatorio'; END IF;
  IF p_pin !~ '^\d{4,8}$' THEN RAISE EXCEPTION 'El PIN debe tener entre 4 y 8 dígitos'; END IF;
  IF p_role NOT IN ('Administrador','Operativo','Supervisor') THEN RAISE EXCEPTION 'Rol no válido'; END IF;
  IF p_role = 'Administrador' AND NOT actor.is_admin THEN RAISE EXCEPTION 'Solo un administrador puede crear administradores'; END IF;
  IF (p_can_upload OR p_can_create_users) AND NOT actor.is_admin THEN RAISE EXCEPTION 'Solo un administrador puede conceder permisos'; END IF;
  IF EXISTS(SELECT 1 FROM public.usuarios u WHERE u.pin_hash = extensions.crypt(p_pin,u.pin_hash)) THEN
    RAISE EXCEPTION 'Ese PIN ya existe';
  END IF;
  admin_role := p_role = 'Administrador';
  INSERT INTO public.usuarios(name,pin_hash,is_admin,active,role,can_upload,can_create_users)
  VALUES(trim(p_name),extensions.crypt(p_pin,extensions.gen_salt('bf')),admin_role,true,p_role,
         (admin_role OR p_can_upload),(admin_role OR p_can_create_users))
  RETURNING id INTO new_id;
  RETURN new_id;
END;
$$;
GRANT EXECUTE ON FUNCTION public.admin_create_user_secure(text,text,text,text,boolean,boolean) TO anon, authenticated;

CREATE OR REPLACE FUNCTION public.admin_users_with_permissions(p_admin_pin text)
RETURNS TABLE(id uuid,name text,is_admin boolean,active boolean,role text,can_upload boolean,can_create_users boolean)
LANGUAGE plpgsql STABLE SECURITY DEFINER
SET search_path = public
AS $$
DECLARE actor public.usuarios%ROWTYPE;
BEGIN
  SELECT * INTO actor FROM public._transport_user_by_pin(p_admin_pin);
  IF actor.id IS NULL OR NOT actor.is_admin THEN RAISE EXCEPTION 'Solo un administrador puede consultar usuarios'; END IF;
  RETURN QUERY SELECT u.id,u.name,u.is_admin,u.active,
    CASE WHEN u.is_admin THEN 'Administrador' ELSE u.role END,
    (u.is_admin OR u.can_upload),(u.is_admin OR u.can_create_users)
  FROM public.usuarios u ORDER BY u.name;
END;
$$;
GRANT EXECUTE ON FUNCTION public.admin_users_with_permissions(text) TO anon, authenticated;

CREATE OR REPLACE FUNCTION public.admin_set_user_permissions_secure(
  p_admin_pin text,p_user_id uuid,p_role text,p_can_upload boolean,p_can_create_users boolean
) RETURNS void
LANGUAGE plpgsql SECURITY DEFINER
SET search_path = public
AS $$
DECLARE actor public.usuarios%ROWTYPE;
BEGIN
  SELECT * INTO actor FROM public._transport_user_by_pin(p_admin_pin);
  IF actor.id IS NULL OR NOT actor.is_admin THEN RAISE EXCEPTION 'Solo un administrador puede modificar permisos'; END IF;
  IF p_role NOT IN ('Administrador','Operativo','Supervisor') THEN RAISE EXCEPTION 'Rol no válido'; END IF;
  UPDATE public.usuarios SET role=p_role,is_admin=(p_role='Administrador'),
    can_upload=(p_role='Administrador' OR p_can_upload),
    can_create_users=(p_role='Administrador' OR p_can_create_users)
  WHERE id=p_user_id;
  IF NOT FOUND THEN RAISE EXCEPTION 'Usuario no encontrado'; END IF;
END;
$$;
GRANT EXECUTE ON FUNCTION public.admin_set_user_permissions_secure(text,uuid,text,boolean,boolean) TO anon, authenticated;

CREATE OR REPLACE FUNCTION public.admin_set_user_active_secure(p_admin_pin text,p_user_id uuid,p_active boolean)
RETURNS void
LANGUAGE plpgsql SECURITY DEFINER
SET search_path = public
AS $$
DECLARE actor public.usuarios%ROWTYPE;
BEGIN
  SELECT * INTO actor FROM public._transport_user_by_pin(p_admin_pin);
  IF actor.id IS NULL OR NOT actor.is_admin THEN RAISE EXCEPTION 'Solo un administrador puede activar o desactivar usuarios'; END IF;
  IF actor.id=p_user_id AND NOT p_active THEN RAISE EXCEPTION 'No puedes desactivar tu propio usuario'; END IF;
  UPDATE public.usuarios SET active=p_active WHERE id=p_user_id;
  IF NOT FOUND THEN RAISE EXCEPTION 'Usuario no encontrado'; END IF;
END;
$$;
GRANT EXECUTE ON FUNCTION public.admin_set_user_active_secure(text,uuid,boolean) TO anon, authenticated;

-- Bloquear los RPC antiguos que no validan PIN/permisos para estas acciones.
DO $$ BEGIN
  IF to_regprocedure('public.replace_transports(jsonb)') IS NOT NULL THEN
    EXECUTE 'REVOKE EXECUTE ON FUNCTION public.replace_transports(jsonb) FROM PUBLIC, anon, authenticated';
  END IF;
  IF to_regprocedure('public.admin_create_user(text,text)') IS NOT NULL THEN
    EXECUTE 'REVOKE EXECUTE ON FUNCTION public.admin_create_user(text,text) FROM PUBLIC, anon, authenticated';
  END IF;
  IF to_regprocedure('public.admin_set_user_active(uuid,boolean)') IS NOT NULL THEN
    EXECUTE 'REVOKE EXECUTE ON FUNCTION public.admin_set_user_active(uuid,boolean) FROM PUBLIC, anon, authenticated';
  END IF;
  IF to_regprocedure('public.admin_users()') IS NOT NULL THEN
    EXECUTE 'REVOKE EXECUTE ON FUNCTION public.admin_users() FROM PUBLIC, anon, authenticated';
  END IF;
END $$;

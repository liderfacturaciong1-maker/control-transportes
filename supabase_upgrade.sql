-- ACTUALIZACIÓN V3: corregir creación de usuarios y carga de Excel.
-- Ejecutar en Supabase SQL Editor.

-- La versión anterior de admin_create_user devolvía void.
-- La eliminamos para poder recrearla con un retorno UUID.
DROP FUNCTION IF EXISTS public.admin_create_user(text, text);

CREATE FUNCTION public.admin_create_user(
  p_name text,
  p_pin text
)
RETURNS uuid
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = public, extensions
AS $$
DECLARE
  new_id uuid;
BEGIN
  IF length(trim(p_name)) = 0 THEN
    RAISE EXCEPTION 'El nombre es obligatorio';
  END IF;
  IF p_pin !~ '^\d{4,8}$' THEN
    RAISE EXCEPTION 'El PIN debe tener entre 4 y 8 dígitos';
  END IF;
  IF EXISTS(
    SELECT 1 FROM public.usuarios
    WHERE pin_hash = extensions.crypt(p_pin, pin_hash)
  ) THEN
    RAISE EXCEPTION 'Ese PIN ya existe';
  END IF;

  INSERT INTO public.usuarios(name, pin_hash, is_admin, active)
  VALUES(trim(p_name), extensions.crypt(p_pin, extensions.gen_salt('bf')), false, true)
  RETURNING id INTO new_id;

  RETURN new_id;
END;
$$;

GRANT EXECUTE ON FUNCTION public.admin_create_user(text,text) TO anon, authenticated;

-- Reemplazo seguro de la lista de transportes.
CREATE OR REPLACE FUNCTION public.replace_transports(p_rows jsonb)
RETURNS void
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = public
AS $$
BEGIN
  DELETE FROM public.transportes WHERE true;

  INSERT INTO public.transportes(transporte, viaje, placa)
  SELECT trim(x->>'transporte'), trim(x->>'viaje'), trim(x->>'placa')
  FROM jsonb_array_elements(p_rows) x;
END;
$$;

GRANT EXECUTE ON FUNCTION public.replace_transports(jsonb) TO anon, authenticated;

-- Permisos explícitos de las funciones utilizadas por el navegador.
GRANT EXECUTE ON FUNCTION public.login_by_pin(text) TO anon, authenticated;
GRANT EXECUTE ON FUNCTION public.list_transports() TO anon, authenticated;
GRANT EXECUTE ON FUNCTION public.advance_transport(bigint) TO anon, authenticated;
GRANT EXECUTE ON FUNCTION public.recent_history(integer) TO anon, authenticated;
GRANT EXECUTE ON FUNCTION public.admin_users() TO anon, authenticated;
GRANT EXECUTE ON FUNCTION public.admin_set_user_active(uuid,boolean) TO anon, authenticated;

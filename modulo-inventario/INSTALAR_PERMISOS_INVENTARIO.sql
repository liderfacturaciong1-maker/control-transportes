-- INSTALAR UNA VEZ en Supabase SQL Editor.
-- Solo controla el permiso para la pantalla inicial de Inventario; no modifica otros módulos.
CREATE EXTENSION IF NOT EXISTS pgcrypto WITH SCHEMA extensions;
CREATE TABLE IF NOT EXISTS public.ct_inventory_authorized_pins (
    pin_digest text PRIMARY KEY,
    allowed boolean NOT NULL DEFAULT true,
    updated_at timestamptz NOT NULL DEFAULT now()
);
ALTER TABLE public.ct_inventory_authorized_pins ENABLE ROW LEVEL SECURITY;
REVOKE ALL ON public.ct_inventory_authorized_pins FROM PUBLIC, anon, authenticated;

CREATE OR REPLACE FUNCTION public.ct_inventory_can_access(p_pin text)
RETURNS boolean LANGUAGE plpgsql SECURITY DEFINER SET search_path TO 'public','extensions','pg_temp'
AS $function$
DECLARE valid_user boolean; admin_user boolean; permitted boolean;
BEGIN
    IF p_pin IS NULL OR p_pin !~ '^[0-9]{1,8}$' THEN RETURN false; END IF;
    SELECT EXISTS(SELECT 1 FROM public.login_by_pin(p_pin)) INTO valid_user;
    IF NOT valid_user THEN RETURN false; END IF;
    SELECT COALESCE(bool_or(COALESCE((to_jsonb(r)->>'is_admin')::boolean,false)),false)
      INTO admin_user FROM public.get_user_permissions_by_pin(p_pin) AS r;
    IF admin_user THEN RETURN true; END IF;
    SELECT COALESCE(allowed,false) INTO permitted
      FROM public.ct_inventory_authorized_pins
     WHERE pin_digest=encode(extensions.digest(p_pin,'sha256'),'hex');
    RETURN COALESCE(permitted,false);
END;
$function$;
REVOKE ALL ON FUNCTION public.ct_inventory_can_access(text) FROM PUBLIC;
GRANT EXECUTE ON FUNCTION public.ct_inventory_can_access(text) TO anon, authenticated;

-- Dar acceso a un usuario EXISTENTE (ejecutar únicamente desde SQL Editor autorizado):
-- INSERT INTO public.ct_inventory_authorized_pins(pin_digest,allowed)
-- VALUES (encode(extensions.digest('PIN_DEL_USUARIO','sha256'),'hex'),true)
-- ON CONFLICT(pin_digest) DO UPDATE SET allowed=true,updated_at=now();
-- Revocar acceso (si el usuario no es administrador):
-- UPDATE public.ct_inventory_authorized_pins SET allowed=false,updated_at=now()
-- WHERE pin_digest=encode(extensions.digest('PIN_DEL_USUARIO','sha256'),'hex');
-- Nota: administradores con is_admin=true conservan acceso; gestión de permisos
-- por ID de usuario puede integrarse después sin recopilar PIN personales.

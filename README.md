# Control de Transportes — actualización de permisos y roles

Esta versión incorpora:
- Prefijo automático `CO` en **Placa Física**, sin duplicarlo.
- Roles: Administrador, Operativo y Supervisor.
- Permisos individuales para cargar/descargar archivos y crear usuarios.
- Validación de PIN y permisos en funciones de Supabase.
- La importación agrega solo transportes nuevos y no borra los existentes.

## Orden obligatorio para actualizar

1. Entra a Supabase del proyecto y abre **SQL Editor**.
2. Abre `supabase_upgrade.sql`, copia todo el contenido y ejecútalo.
3. Si Supabase muestra un error, **detente y no publiques todavía los archivos**; comparte una captura del error para corregirlo.
4. Cuando el SQL termine correctamente, publica los archivos de esta carpeta en el repositorio GitHub reemplazando los existentes. Vercel debería desplegar automáticamente.
5. Entra con un PIN de administrador y prueba la carga de un Excel pequeño con registros nuevos y existentes.

## Permisos

- Administrador: todos los permisos. Puede crear usuarios, elegir el rol y conceder los permisos de carga/creación.
- Operativo y Supervisor: solo las capacidades que el administrador les conceda. Si tienen permiso de creación, pueden crear usuarios Operativos o Supervisores; no pueden conceder permisos ni crear administradores.
- La asignación de permisos y la carga se verifican en Supabase, no solo en la interfaz.

**Importante:** conserva una copia del ZIP anterior. No compartas PIN de usuarios ni claves secretas de Supabase. La clave `publishable` de `config.js` está destinada al navegador.

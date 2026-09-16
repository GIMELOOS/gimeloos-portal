-- ============================================================
-- RLS para app_settings
-- Ejecutar en: Supabase Dashboard → SQL Editor
-- ============================================================
-- La tabla app_settings guarda el token OAuth de Google Drive.
-- Solo los admins deben poder leerla o modificarla.
-- Requiere que la función is_admin() del script principal ya exista.
-- ============================================================

ALTER TABLE app_settings ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "app_settings_admin_only_select" ON app_settings;
DROP POLICY IF EXISTS "app_settings_admin_only_insert" ON app_settings;
DROP POLICY IF EXISTS "app_settings_admin_only_update" ON app_settings;
DROP POLICY IF EXISTS "app_settings_admin_only_delete" ON app_settings;

CREATE POLICY "app_settings_admin_only_select" ON app_settings
  FOR SELECT USING (is_admin());

CREATE POLICY "app_settings_admin_only_insert" ON app_settings
  FOR INSERT WITH CHECK (is_admin());

CREATE POLICY "app_settings_admin_only_update" ON app_settings
  FOR UPDATE USING (is_admin());

CREATE POLICY "app_settings_admin_only_delete" ON app_settings
  FOR DELETE USING (is_admin());

-- NOTA: Las API routes (tracking-sheet, proxy-sheet, google-callback)
-- usan supabaseAdmin con el service role key, que bypasa RLS por diseño.
-- Este RLS solo protege el acceso directo desde el cliente (anon key).

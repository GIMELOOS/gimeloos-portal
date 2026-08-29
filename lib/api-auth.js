import { createClient } from "@supabase/supabase-js";
import { NextResponse } from "next/server";

const supabaseAdmin = createClient(
  process.env.NEXT_PUBLIC_SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY
);

/**
 * Verifica el header Authorization y devuelve el usuario autenticado.
 * Returns { user } si ok, o { error: NextResponse } si falla.
 */
export async function requireAuth(request) {
  const authHeader = request.headers.get("Authorization");
  if (!authHeader?.startsWith("Bearer ")) {
    return {
      error: NextResponse.json({ error: "No autorizado" }, { status: 401 }),
    };
  }

  const token = authHeader.replace("Bearer ", "");
  const {
    data: { user },
    error,
  } = await supabaseAdmin.auth.getUser(token);

  if (error || !user) {
    return {
      error: NextResponse.json({ error: "Token inválido o expirado" }, { status: 401 }),
    };
  }

  return { user };
}

/**
 * Verifica que el usuario sea administrador.
 * Returns { user } si ok, o { error: NextResponse } si falla.
 */
export async function requireAdmin(request) {
  const { user, error } = await requireAuth(request);
  if (error) return { error };

  const { data: participant } = await supabaseAdmin
    .from("participants")
    .select("role")
    .eq("auth_uid", user.id)
    .single();

  if (participant?.role !== "admin") {
    return {
      error: NextResponse.json({ error: "Acceso denegado: se requiere rol admin" }, { status: 403 }),
    };
  }

  return { user };
}

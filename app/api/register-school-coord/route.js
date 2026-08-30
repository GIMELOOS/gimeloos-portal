import { NextResponse } from "next/server";
import { createClient } from "@supabase/supabase-js";

export const runtime = "nodejs";

const supabaseAdmin = createClient(
  process.env.NEXT_PUBLIC_SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY
);

export async function GET(request) {
  const { searchParams } = new URL(request.url);
  const token = searchParams.get("token");
  if (!token) return NextResponse.json({ error: "Token requerido" }, { status: 400 });

  const { data: school, error } = await supabaseAdmin
    .from("schools")
    .select("id, name, email, contact_name, invite_token, invite_token_expires_at")
    .eq("invite_token", token)
    .maybeSingle();

  if (error || !school) return NextResponse.json({ error: "Invitación no válida o expirada" }, { status: 404 });

  const now = new Date();
  const expires = school.invite_token_expires_at ? new Date(school.invite_token_expires_at) : null;
  if (!expires || now > expires) {
    return NextResponse.json({ error: "El enlace de invitación ha expirado. Solicita uno nuevo al administrador." }, { status: 410 });
  }

  return NextResponse.json({
    schoolId: school.id,
    schoolName: school.name,
    email: school.email,
    contactName: school.contact_name,
  });
}

export async function POST(request) {
  const { token, password, nombre } = await request.json();

  if (!token || !password || !nombre) {
    return NextResponse.json({ error: "Faltan campos obligatorios" }, { status: 400 });
  }
  if (password.length < 8) {
    return NextResponse.json({ error: "La contraseña debe tener al menos 8 caracteres" }, { status: 400 });
  }

  // Validar token
  const { data: school, error: schoolErr } = await supabaseAdmin
    .from("schools")
    .select("id, name, email, invite_token_expires_at")
    .eq("invite_token", token)
    .maybeSingle();

  if (schoolErr || !school) return NextResponse.json({ error: "Invitación no válida" }, { status: 404 });

  const now = new Date();
  const expires = school.invite_token_expires_at ? new Date(school.invite_token_expires_at) : null;
  if (!expires || now > expires) {
    return NextResponse.json({ error: "El enlace de invitación ha expirado" }, { status: 410 });
  }

  // Crear usuario en Supabase Auth
  const { data: authData, error: authErr } = await supabaseAdmin.auth.admin.createUser({
    email: school.email,
    password,
    email_confirm: true,
    user_metadata: { nombre, school_id: school.id, role: "school" },
  });

  if (authErr) {
    if (authErr.message?.includes("already been registered") || authErr.message?.includes("already exists")) {
      return NextResponse.json({ error: "Este email ya tiene una cuenta. Inicia sesión directamente." }, { status: 409 });
    }
    return NextResponse.json({ error: "Error al crear la cuenta: " + authErr.message }, { status: 500 });
  }

  const userId = authData.user.id;

  // Crear entrada en participants si no existe
  await supabaseAdmin.from("participants").upsert(
    { auth_uid: userId, email: school.email, nombre, role: "school" },
    { onConflict: "auth_uid" }
  );

  // Vincular el auth_uid al colegio y limpiar el token
  await supabaseAdmin
    .from("schools")
    .update({ auth_uid: userId, invite_token: null, invite_token_expires_at: null, contact_name: nombre })
    .eq("id", school.id);

  return NextResponse.json({ ok: true });
}

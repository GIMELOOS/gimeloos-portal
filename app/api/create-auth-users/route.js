import { NextResponse } from "next/server";
import { createClient } from "@supabase/supabase-js";
import { requireAdmin } from "@/lib/api-auth";

export const runtime = "nodejs";

const supabaseAdmin = createClient(
  process.env.NEXT_PUBLIC_SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY
);

// Recibe: [{ participantId, email, password }]
// Crea o actualiza la cuenta Auth y vincula auth_uid en participants
export async function POST(request) {
  const { error: authError } = await requireAdmin(request);
  if (authError) return authError;

  if (!process.env.SUPABASE_SERVICE_ROLE_KEY) {
    return NextResponse.json({ error: "Falta SUPABASE_SERVICE_ROLE_KEY" }, { status: 500 });
  }

  const { participants } = await request.json();
  if (!Array.isArray(participants) || !participants.length) {
    return NextResponse.json({ error: "Lista de participantes vacía" }, { status: 400 });
  }

  const results = { created: 0, updated: 0, skipped: 0, errors: [] };

  for (const { participantId, email, password } of participants) {
    if (!email || !password || !participantId) {
      results.skipped++;
      continue;
    }

    try {
      // Buscar si ya existe en Auth por email (evita límite de paginación de listUsers)
      const normalizedEmail = email.toLowerCase().trim();
      const { data: byEmail } = await supabaseAdmin.auth.admin.getUserByEmail(normalizedEmail);
      const existingUser = byEmail?.user ?? null;

      let authUid;

      if (existingUser) {
        // Actualizar contraseña si cambió
        const { error: updErr } = await supabaseAdmin.auth.admin.updateUserById(existingUser.id, { password });
        if (updErr) { results.errors.push({ email: normalizedEmail, error: updErr.message }); continue; }
        authUid = existingUser.id;
        results.updated++;
      } else {
        // Crear cuenta nueva — email_confirm: true para acceso inmediato
        const { data: created, error: createErr } = await supabaseAdmin.auth.admin.createUser({
          email: normalizedEmail,
          password,
          email_confirm: true,
        });
        if (createErr) {
          results.errors.push({ participantId, error: createErr.message });
          continue;
        }
        authUid = created.user.id;
        results.created++;
      }

      // Vincular auth_uid en participants
      const { error: linkErr } = await supabaseAdmin
        .from("participants")
        .update({ auth_uid: authUid })
        .eq("id", participantId);
      if (linkErr) { results.errors.push({ participantId, error: "auth_uid link: " + linkErr.message }); continue; }

    } catch (err) {
      results.errors.push({ participantId, error: err.message });
    }
  }

  return NextResponse.json({ ...results, total: participants.length });
}

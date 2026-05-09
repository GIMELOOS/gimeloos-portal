import { NextResponse } from "next/server";
import { createClient } from "@supabase/supabase-js";

export const runtime = "nodejs";

const supabaseAdmin = createClient(
  process.env.NEXT_PUBLIC_SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY
);

// Recibe: [{ participantId, email, password }]
// Crea o actualiza la cuenta Auth y vincula auth_uid en participants
export async function POST(request) {
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
      // Buscar si ya existe en Auth
      const { data: existing } = await supabaseAdmin.auth.admin.listUsers();
      const existingUser = existing?.users?.find((u) => u.email === email.toLowerCase().trim());

      let authUid;

      if (existingUser) {
        // Actualizar contraseña si cambió
        await supabaseAdmin.auth.admin.updateUserById(existingUser.id, { password });
        authUid = existingUser.id;
        results.updated++;
      } else {
        // Crear cuenta nueva — email_confirm: true para acceso inmediato
        const { data: created, error: createErr } = await supabaseAdmin.auth.admin.createUser({
          email: email.toLowerCase().trim(),
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
      await supabaseAdmin
        .from("participants")
        .update({ auth_uid: authUid })
        .eq("id", participantId);

    } catch (err) {
      results.errors.push({ participantId, error: err.message });
    }
  }

  return NextResponse.json({ ...results, total: participants.length });
}

import { NextResponse } from "next/server";
import { createClient } from "@supabase/supabase-js";
import { requireAdmin } from "@/lib/api-auth";

export const runtime = "nodejs";

const supabaseAdmin = createClient(
  process.env.NEXT_PUBLIC_SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY
);

export async function POST(request) {
  const { error: authError } = await requireAdmin(request);
  if (authError) return authError;

  if (!process.env.SUPABASE_SERVICE_ROLE_KEY) {
    return NextResponse.json({ error: "Falta SUPABASE_SERVICE_ROLE_KEY" }, { status: 500 });
  }

  const { participantId } = await request.json();
  if (!participantId) {
    return NextResponse.json({ error: "Falta participantId" }, { status: 400 });
  }

  const { data: participant, error: fetchErr } = await supabaseAdmin
    .from("participants")
    .select("id, username, email, participant_name, auth_uid")
    .eq("id", participantId)
    .maybeSingle();

  if (fetchErr || !participant) {
    return NextResponse.json({ error: "Participante no encontrado" }, { status: 404 });
  }

  if (!participant.email) {
    return NextResponse.json({ error: "El participante no tiene email registrado" }, { status: 400 });
  }

  // Si ya tiene cuenta Auth, reenviar email de acceso
  if (participant.auth_uid) {
    const { error: resetErr } = await supabaseAdmin.auth.admin.generateLink({
      type: "recovery",
      email: participant.email,
    });
    if (resetErr) {
      return NextResponse.json({ error: resetErr.message }, { status: 500 });
    }
    return NextResponse.json({ message: "Email de acceso reenviado.", alreadyExisted: true });
  }

  // Crear cuenta nueva en Supabase Auth via invitación por email
  const { data: invited, error: inviteErr } = await supabaseAdmin.auth.admin.inviteUserByEmail(
    participant.email,
    { data: { role: "client", participant_id: participant.id } }
  );

  if (inviteErr) {
    return NextResponse.json({ error: inviteErr.message }, { status: 500 });
  }

  // Vincular auth_uid en la tabla participants
  const { error: updateErr } = await supabaseAdmin
    .from("participants")
    .update({ auth_uid: invited.user.id })
    .eq("id", participant.id);

  if (updateErr) {
    return NextResponse.json({ error: updateErr.message }, { status: 500 });
  }

  return NextResponse.json({
    message: `Invitación enviada a ${participant.email}`,
    authUid: invited.user.id,
  });
}

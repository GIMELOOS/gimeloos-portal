import { NextResponse } from "next/server";
import { createClient } from "@supabase/supabase-js";
import { Resend } from "resend";
import { requireAdmin } from "@/lib/api-auth";
import crypto from "crypto";

export const runtime = "nodejs";

const supabaseAdmin = createClient(
  process.env.NEXT_PUBLIC_SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY
);

const resend = new Resend(process.env.RESEND_API_KEY);
const FROM = process.env.NOTIFY_FROM || "onboarding@resend.dev";

export async function POST(request) {
  const { error: authError } = await requireAdmin(request);
  if (authError) return authError;

  const { school_id } = await request.json();
  if (!school_id) return NextResponse.json({ error: "Falta school_id" }, { status: 400 });

  const { data: school, error: schoolErr } = await supabaseAdmin
    .from("schools")
    .select("id, name, email, contact_name")
    .eq("id", school_id)
    .single();

  if (schoolErr || !school) return NextResponse.json({ error: "Colegio no encontrado" }, { status: 404 });
  if (!school.email) return NextResponse.json({ error: "El colegio no tiene email configurado" }, { status: 400 });

  // Generar token único de 32 bytes
  const token = crypto.randomBytes(32).toString("hex");
  const expires = new Date(Date.now() + 7 * 24 * 60 * 60 * 1000); // 7 días

  const { error: updateErr } = await supabaseAdmin
    .from("schools")
    .update({ invite_token: token, invite_token_expires_at: expires.toISOString() })
    .eq("id", school_id);

  if (updateErr) return NextResponse.json({ error: "Error guardando token: " + updateErr.message }, { status: 500 });

  const appUrl = process.env.NEXT_PUBLIC_APP_URL || "https://gimeloos-portal.vercel.app";
  const inviteUrl = `${appUrl}/?invite=${token}`;
  const contactName = school.contact_name || "Coordinador/a";

  const html = `
    <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;background:#fff;border:1px solid #e4e4e7">
      <div style="background:#18181b;padding:28px 40px">
        <div style="color:white;font-size:20px;font-weight:700;letter-spacing:-0.3px">GIMELOOS</div>
        <div style="color:#a1a1aa;font-size:11px;letter-spacing:0.15em;margin-top:4px;text-transform:uppercase">Área privada de clientes</div>
      </div>
      <div style="padding:40px;background:#fff">
        <div style="color:#FF3131;font-size:11px;font-weight:700;letter-spacing:0.15em;text-transform:uppercase;margin-bottom:12px">PORTAL DEL COLEGIO</div>
        <h1 style="color:#18181b;font-size:24px;font-weight:700;margin:0 0 24px;line-height:1.3">Accede al Portal del Colegio</h1>
        <p style="color:#52525b;font-size:15px;line-height:1.7;margin:0 0 16px">Hola, <strong>${contactName}</strong>.</p>
        <p style="color:#52525b;font-size:15px;line-height:1.7;margin:0 0 24px">
          Te invitamos a activar tu acceso al <strong>Portal del Colegio GIMELOOS</strong> de <strong>${school.name}</strong>.
          Desde aquí podrás gestionar la documentación del viaje, ver el listado de alumnos y mucho más.
        </p>
        <a href="${inviteUrl}" style="display:inline-block;background:#FF3131;color:white;font-weight:600;font-size:15px;padding:14px 28px;border-radius:10px;text-decoration:none;margin-bottom:24px">
          Crear mi cuenta →
        </a>
        <p style="color:#71717a;font-size:13px;margin:24px 0 0">
          Este enlace es válido durante <strong>7 días</strong>. Si el botón no funciona, copia este enlace en tu navegador:<br>
          <a href="${inviteUrl}" style="color:#FF3131;word-break:break-all;font-size:12px">${inviteUrl}</a>
        </p>
      </div>
      <div style="background:#f4f4f5;padding:24px 40px;text-align:center;font-size:12px;color:#71717a;border-top:1px solid #e4e4e7">
        <p style="margin:0 0 8px">Recibes este correo porque eres coordinador/a de ${school.name} con GIMELOOS.</p>
        <p style="margin:0">© 2026 GIMELOOS EVENTOS Y ACTIVIDADES SL</p>
      </div>
    </div>
  `;

  const { error: emailError } = await resend.emails.send({
    from: FROM,
    to: school.email,
    subject: `Invitación al Portal del Colegio — ${school.name}`,
    html,
  });

  if (emailError) {
    console.error("Resend error:", emailError);
    return NextResponse.json({ error: "No se pudo enviar el email: " + (emailError.message || JSON.stringify(emailError)) }, { status: 500 });
  }

  return NextResponse.json({ ok: true, sentTo: school.email });
}

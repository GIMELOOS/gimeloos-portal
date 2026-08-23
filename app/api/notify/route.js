import { NextResponse } from "next/server";
import { Resend } from "resend";
import { createClient } from "@supabase/supabase-js";
import { requireAuth } from "@/lib/api-auth";

export const runtime = "nodejs";

const resend = new Resend(process.env.RESEND_API_KEY);
const FROM = process.env.NOTIFY_FROM || "onboarding@resend.dev";
const ADMIN_EMAIL = process.env.ADMIN_EMAIL || "info@gimeloos.com";

const supabaseAdmin = createClient(
  process.env.NEXT_PUBLIC_SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY
);

// ─── Plantillas de email ──────────────────────────────────────────────────────

const baseStyle = `font-family:Arial,sans-serif;max-width:560px;margin:0 auto;background:#fff;border-radius:16px;overflow:hidden;border:1px solid #e4e4e7`;
const headerStyle = `background:#C0392B;padding:32px 32px 24px;text-align:center`;
const bodyStyle = `padding:32px`;
const footerStyle = `padding:16px 32px;background:#fafafa;border-top:1px solid #e4e4e7;text-align:center;font-size:12px;color:#71717a`;

function emailWrapper(title, body) {
  return `
  <div style="${baseStyle}">
    <div style="${headerStyle}">
      <div style="color:white;font-size:22px;font-weight:700;letter-spacing:-0.5px">GIMELOOS</div>
      <div style="color:rgba(255,255,255,0.85);font-size:13px;margin-top:4px">${title}</div>
    </div>
    <div style="${bodyStyle}">${body}</div>
    <div style="${footerStyle}">
      © GIMELOOS · Este es un mensaje automático, no respondas a este correo.
    </div>
  </div>`;
}

const templates = {
  // Participante: documento confirmado
  doc_confirmed: ({ participantName, docName, tripName }) => ({
    subject: `✅ Documento confirmado — ${docName}`,
    html: emailWrapper("Documento confirmado", `
      <p style="color:#18181b;font-size:16px">Hola, <strong>${participantName}</strong></p>
      <p style="color:#52525b;font-size:14px;line-height:1.6">Tu documento <strong>${docName}</strong> para <strong>${tripName}</strong> ha sido revisado y <strong style="color:#16a34a">confirmado</strong> por el equipo de GIMELOOS. ¡Perfecto!</p>
      <div style="margin:24px 0;padding:16px;background:#f0fdf4;border-radius:12px;border-left:4px solid #16a34a">
        <div style="color:#15803d;font-weight:600;font-size:14px">✓ ${docName}</div>
        <div style="color:#52525b;font-size:13px;margin-top:4px">Estado: Confirmado</div>
      </div>
      <p style="color:#52525b;font-size:14px">Accede al portal para ver el estado de todos tus documentos.</p>
    `),
  }),

  // Participante: documento rechazado
  doc_rejected: ({ participantName, docName, tripName }) => ({
    subject: `❌ Documento rechazado — ${docName}`,
    html: emailWrapper("Documento rechazado", `
      <p style="color:#18181b;font-size:16px">Hola, <strong>${participantName}</strong></p>
      <p style="color:#52525b;font-size:14px;line-height:1.6">Tu documento <strong>${docName}</strong> para <strong>${tripName}</strong> ha sido revisado pero necesita correcciones.</p>
      <div style="margin:24px 0;padding:16px;background:#fef2f2;border-radius:12px;border-left:4px solid #dc2626">
        <div style="color:#dc2626;font-weight:600;font-size:14px">✗ ${docName}</div>
        <div style="color:#52525b;font-size:13px;margin-top:4px">Estado: Rechazado — por favor sube una versión corregida</div>
      </div>
      <p style="color:#52525b;font-size:14px">Accede al portal para subir el documento corregido.</p>
    `),
  }),

  // Participante: pago confirmado
  payment_confirmed: ({ participantName, paymentName, amount, tripName }) => ({
    subject: `✅ Pago confirmado — ${paymentName}`,
    html: emailWrapper("Pago confirmado", `
      <p style="color:#18181b;font-size:16px">Hola, <strong>${participantName}</strong></p>
      <p style="color:#52525b;font-size:14px;line-height:1.6">Hemos confirmado tu pago de <strong>${paymentName}</strong> para <strong>${tripName}</strong>.</p>
      <div style="margin:24px 0;padding:16px;background:#f0fdf4;border-radius:12px;border-left:4px solid #16a34a">
        <div style="color:#15803d;font-weight:600;font-size:14px">✓ ${paymentName}</div>
        <div style="color:#52525b;font-size:13px;margin-top:4px">Importe: ${amount} · Estado: Confirmado</div>
      </div>
    `),
  }),

  // Participante: respuesta a pregunta
  question_replied: ({ participantName, question, reply }) => ({
    subject: `💬 Tienes una respuesta de GIMELOOS`,
    html: emailWrapper("Respuesta a tu consulta", `
      <p style="color:#18181b;font-size:16px">Hola, <strong>${participantName}</strong></p>
      <p style="color:#52525b;font-size:14px">El equipo de GIMELOOS ha respondido a tu consulta:</p>
      <div style="margin:16px 0;padding:16px;background:#f4f4f5;border-radius:12px">
        <div style="color:#71717a;font-size:12px;margin-bottom:6px">Tu pregunta</div>
        <p style="color:#3f3f46;font-size:14px;margin:0">${question}</p>
      </div>
      <div style="margin:16px 0;padding:16px;background:#fef3c7;border-radius:12px;border-left:4px solid #d97706">
        <div style="color:#92400e;font-size:12px;margin-bottom:6px;font-weight:600">Respuesta del equipo</div>
        <p style="color:#3f3f46;font-size:14px;margin:0">${reply}</p>
      </div>
      <p style="color:#52525b;font-size:14px">Si tienes más dudas, accede al portal y envíanos otro mensaje.</p>
    `),
  }),

  // Participante: recordatorio pago próximo
  payment_reminder: ({ participantName, paymentName, amount, dueDate, daysLeft, tripName }) => ({
    subject: `⏰ Recordatorio de pago — ${paymentName} vence en ${daysLeft} días`,
    html: emailWrapper("Recordatorio de pago", `
      <p style="color:#18181b;font-size:16px">Hola, <strong>${participantName}</strong></p>
      <p style="color:#52525b;font-size:14px;line-height:1.6">Te recordamos que tienes un pago pendiente para <strong>${tripName}</strong>.</p>
      <div style="margin:24px 0;padding:20px;background:#fff7ed;border-radius:12px;border-left:4px solid #f97316">
        <div style="color:#c2410c;font-weight:700;font-size:16px">${paymentName}</div>
        <div style="color:#52525b;font-size:14px;margin-top:8px">Importe: <strong>${amount}</strong></div>
        <div style="color:#52525b;font-size:14px">Fecha límite: <strong>${dueDate}</strong></div>
        <div style="margin-top:12px;padding:8px 12px;background:#fed7aa;border-radius:8px;color:#c2410c;font-weight:600;font-size:14px;display:inline-block">
          Quedan ${daysLeft} días
        </div>
      </div>
      <p style="color:#52525b;font-size:14px">Accede al portal para subir tu justificante de pago.</p>
    `),
  }),

  // Participante: documento pendiente de subir
  doc_reminder: ({ participantName, docName, tripName }) => ({
    subject: `📄 Documentación pendiente — ${docName}`,
    html: emailWrapper("Documentación pendiente", `
      <p style="color:#18181b;font-size:16px">Hola, <strong>${participantName}</strong></p>
      <p style="color:#52525b;font-size:14px;line-height:1.6">Aún tienes documentación pendiente de subir para <strong>${tripName}</strong>.</p>
      <div style="margin:24px 0;padding:16px;background:#fef2f2;border-radius:12px;border-left:4px solid #C0392B">
        <div style="color:#C0392B;font-weight:600;font-size:14px">📄 ${docName}</div>
        <div style="color:#52525b;font-size:13px;margin-top:4px">Estado: Pendiente de envío</div>
      </div>
      <p style="color:#52525b;font-size:14px">Accede al portal para subir tu documentación lo antes posible.</p>
    `),
  }),

  // Admin: nuevo documento subido
  admin_doc_uploaded: ({ participantName, docName, tripName }) => ({
    subject: `📥 Nuevo documento — ${participantName}`,
    html: emailWrapper("Nuevo documento pendiente de revisión", `
      <p style="color:#18181b;font-size:16px">Nuevo documento para revisar</p>
      <div style="margin:16px 0;padding:16px;background:#f4f4f5;border-radius:12px">
        <div style="font-size:14px;color:#3f3f46"><strong>Participante:</strong> ${participantName}</div>
        <div style="font-size:14px;color:#3f3f46;margin-top:4px"><strong>Documento:</strong> ${docName}</div>
        <div style="font-size:14px;color:#3f3f46;margin-top:4px"><strong>Viaje:</strong> ${tripName}</div>
      </div>
      <p style="color:#52525b;font-size:14px">Accede al panel de administración para revisar y confirmar el documento.</p>
    `),
  }),

  // Admin: nuevo justificante subido
  admin_payment_uploaded: ({ participantName, paymentName, tripName }) => ({
    subject: `💳 Nuevo justificante — ${participantName}`,
    html: emailWrapper("Nuevo justificante pendiente de revisión", `
      <p style="color:#18181b;font-size:16px">Nuevo justificante de pago para revisar</p>
      <div style="margin:16px 0;padding:16px;background:#f4f4f5;border-radius:12px">
        <div style="font-size:14px;color:#3f3f46"><strong>Participante:</strong> ${participantName}</div>
        <div style="font-size:14px;color:#3f3f46;margin-top:4px"><strong>Pago:</strong> ${paymentName}</div>
        <div style="font-size:14px;color:#3f3f46;margin-top:4px"><strong>Viaje:</strong> ${tripName}</div>
      </div>
      <p style="color:#52525b;font-size:14px">Accede al panel de administración para revisar y confirmar el pago.</p>
    `),
  }),

  // Admin: nueva pregunta
  admin_new_question: ({ participantName, question, tripName }) => ({
    subject: `❓ Nueva pregunta — ${participantName}`,
    html: emailWrapper("Nueva consulta de un participante", `
      <p style="color:#18181b;font-size:16px">Nueva pregunta recibida</p>
      <div style="margin:16px 0;padding:16px;background:#f4f4f5;border-radius:12px">
        <div style="font-size:14px;color:#3f3f46"><strong>Participante:</strong> ${participantName}</div>
        <div style="font-size:14px;color:#3f3f46;margin-top:4px"><strong>Viaje:</strong> ${tripName}</div>
      </div>
      <div style="margin:16px 0;padding:16px;background:#eff6ff;border-radius:12px;border-left:4px solid #3b82f6">
        <p style="color:#1e3a5f;font-size:14px;margin:0">${question}</p>
      </div>
      <p style="color:#52525b;font-size:14px">Accede al panel → Preguntas para responder.</p>
    `),
  }),
};

// ─── Handler ──────────────────────────────────────────────────────────────────

export async function POST(request) {
  const { error: authError } = await requireAuth(request);
  if (authError) return authError;

  try {
    const body = await request.json();
    const { type, to, participantId, data } = body;

    if (!type || !templates[type]) {
      return NextResponse.json({ error: "Tipo de notificación inválido" }, { status: 400 });
    }

    const { subject, html } = templates[type](data || {});
    const recipient = to || ADMIN_EMAIL;

    // Enviar email
    const { error: emailError } = await resend.emails.send({
      from: FROM,
      to: recipient,
      subject,
      html,
    });

    if (emailError) {
      console.error("Resend error:", emailError);
      return NextResponse.json({ error: emailError.message }, { status: 500 });
    }

    // Guardar notificación in-app si es para un participante
    if (participantId) {
      const inAppMessages = {
        doc_confirmed: { title: "Documento confirmado ✅", body: `Tu documento "${data.docName}" ha sido confirmado.` },
        doc_rejected: { title: "Documento rechazado ❌", body: `Tu documento "${data.docName}" necesita correcciones.` },
        payment_confirmed: { title: "Pago confirmado ✅", body: `Tu pago "${data.paymentName}" ha sido confirmado.` },
        question_replied: { title: "Nueva respuesta 💬", body: "El equipo de GIMELOOS ha respondido a tu consulta." },
        payment_reminder: { title: "Pago próximo ⏰", body: `"${data.paymentName}" vence en ${data.daysLeft} días.` },
        doc_reminder: { title: "Documentación pendiente 📄", body: `Tienes que subir: "${data.docName}".` },
      };

      const msg = inAppMessages[type];
      if (msg) {
        await supabaseAdmin.from("notifications").insert({
          participant_id: participantId,
          type,
          title: msg.title,
          body: msg.body,
        });
      }
    }

    return NextResponse.json({ ok: true });
  } catch (err) {
    console.error("notify error:", err);
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}

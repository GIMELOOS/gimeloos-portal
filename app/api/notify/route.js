import { NextResponse } from "next/server";
import { Resend } from "resend";
import { createClient } from "@supabase/supabase-js";
import { requireAdmin } from "@/lib/api-auth";

export const runtime = "nodejs";

const escapeHtml = (str) => String(str)
  .replace(/&/g, "&amp;")
  .replace(/</g, "&lt;")
  .replace(/>/g, "&gt;")
  .replace(/"/g, "&quot;")
  .replace(/'/g, "&#39;");

const resend = new Resend(process.env.RESEND_API_KEY);
const FROM = process.env.NOTIFY_FROM || "onboarding@resend.dev";
const ADMIN_EMAIL = process.env.ADMIN_EMAIL || "info@gimeloos.com";

const supabaseAdmin = createClient(
  process.env.NEXT_PUBLIC_SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY
);

// ─── Plantillas de email ──────────────────────────────────────────────────────

const baseStyle = `font-family:Arial,sans-serif;max-width:560px;margin:0 auto;background:#fff;border-radius:16px;overflow:hidden;border:1px solid #e4e4e7`;
const headerStyle = `background:#FF3131;padding:32px 32px 24px;text-align:center`;
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
      <div style="margin:24px 0;padding:16px;background:#fef2f2;border-radius:12px;border-left:4px solid #FF3131">
        <div style="color:#FF3131;font-weight:600;font-size:14px">📄 ${docName}</div>
        <div style="color:#52525b;font-size:13px;margin-top:4px">Estado: Pendiente de envío</div>
      </div>
      <p style="color:#52525b;font-size:14px">Accede al portal para subir tu documentación lo antes posible.</p>
    `),
  }),

  // Coordinador de colegio: recordatorio listado de alumnos
  school_reminder_listado: ({ schoolName, contactName, tripName }) => ({
    subject: `📋 Listado de alumnos pendiente — ${tripName}`,
    html: emailWrapper("Listado de alumnos pendiente", `
      <p style="color:#18181b;font-size:16px">Hola${contactName ? `, <strong>${contactName}</strong>` : ""}</p>
      <p style="color:#52525b;font-size:14px;line-height:1.6">Te recordamos que necesitamos el <strong>listado definitivo de alumnos</strong> del colegio <strong>${schoolName}</strong> para el viaje <strong>${tripName}</strong>.</p>
      <div style="margin:24px 0;padding:20px;background:#eff6ff;border-radius:12px;border-left:4px solid #3b82f6">
        <div style="color:#1d4ed8;font-weight:700;font-size:15px">📋 Acción requerida: listado de alumnos</div>
        <div style="color:#52525b;font-size:14px;margin-top:8px">Por favor, accede al portal escolar y sube el listado lo antes posible.</div>
      </div>
    `),
  }),

  // Coordinador de colegio: recordatorio alergias e intolerancias
  school_reminder_alergias: ({ schoolName, contactName, tripName }) => ({
    subject: `🍽️ Alergias e intolerancias pendientes — ${tripName}`,
    html: emailWrapper("Alergias e intolerancias pendientes", `
      <p style="color:#18181b;font-size:16px">Hola${contactName ? `, <strong>${contactName}</strong>` : ""}</p>
      <p style="color:#52525b;font-size:14px;line-height:1.6">Es imprescindible que nos comuniquéis las <strong>alergias e intolerancias alimentarias</strong> de los alumnos del colegio <strong>${schoolName}</strong> para el viaje <strong>${tripName}</strong>.</p>
      <div style="margin:24px 0;padding:20px;background:#fef9c3;border-radius:12px;border-left:4px solid #eab308">
        <div style="color:#854d0e;font-weight:700;font-size:15px">🍽️ Alergias e intolerancias</div>
        <div style="color:#52525b;font-size:14px;margin-top:8px">Accede al portal y rellena la información médica de cada alumno. Es obligatorio para garantizar su seguridad.</div>
      </div>
    `),
  }),

  // Coordinador de colegio: recordatorio documentación
  school_doc_reminder: ({ schoolName, contactName, tripName, pendingCount }) => ({
    subject: `📄 Documentación pendiente — ${tripName}`,
    html: emailWrapper("Documentación pendiente del viaje", `
      <p style="color:#18181b;font-size:16px">Hola${contactName ? `, <strong>${contactName}</strong>` : ""}</p>
      <p style="color:#52525b;font-size:14px;line-height:1.6">El colegio <strong>${schoolName}</strong> tiene documentación pendiente de enviar para el viaje <strong>${tripName}</strong>.</p>
      <div style="margin:24px 0;padding:20px;background:#fef2f2;border-radius:12px;border-left:4px solid #FF3131">
        <div style="color:#FF3131;font-weight:700;font-size:16px">📄 ${pendingCount > 0 ? `${pendingCount} documento${pendingCount !== 1 ? "s" : ""} pendiente${pendingCount !== 1 ? "s" : ""}` : "Documentación pendiente"}</div>
        <div style="color:#52525b;font-size:14px;margin-top:8px">Accede al portal escolar y sube los documentos requeridos lo antes posible.</div>
      </div>
    `),
  }),

  // Coordinador de colegio: recordatorio rooming
  school_reminder_rooming: ({ schoolName, contactName, tripName }) => ({
    subject: `🛏️ Asignación de habitaciones pendiente — ${tripName}`,
    html: emailWrapper("Rooming pendiente", `
      <p style="color:#18181b;font-size:16px">Hola${contactName ? `, <strong>${contactName}</strong>` : ""}</p>
      <p style="color:#52525b;font-size:14px;line-height:1.6">Necesitamos la <strong>asignación de habitaciones (rooming)</strong> del colegio <strong>${schoolName}</strong> para el viaje <strong>${tripName}</strong>.</p>
      <div style="margin:24px 0;padding:20px;background:#f0fdf4;border-radius:12px;border-left:4px solid #16a34a">
        <div style="color:#15803d;font-weight:700;font-size:15px">🛏️ Rooming pendiente</div>
        <div style="color:#52525b;font-size:14px;margin-top:8px">Accede al portal escolar → sección Rooming y completa la distribución de habitaciones.</div>
      </div>
    `),
  }),

  // Coordinador de colegio: recordatorio grupos
  school_reminder_grupos: ({ schoolName, contactName, tripName }) => ({
    subject: `👥 Grupos de actividad pendientes — ${tripName}`,
    html: emailWrapper("Grupos de actividad pendientes", `
      <p style="color:#18181b;font-size:16px">Hola${contactName ? `, <strong>${contactName}</strong>` : ""}</p>
      <p style="color:#52525b;font-size:14px;line-height:1.6">Los <strong>grupos de actividad</strong> del colegio <strong>${schoolName}</strong> para el viaje <strong>${tripName}</strong> están pendientes de definir.</p>
      <div style="margin:24px 0;padding:20px;background:#faf5ff;border-radius:12px;border-left:4px solid #9333ea">
        <div style="color:#7e22ce;font-weight:700;font-size:15px">👥 Grupos de actividad</div>
        <div style="color:#52525b;font-size:14px;margin-top:8px">Accede al portal escolar → sección Grupos y define los grupos de actividad.</div>
      </div>
    `),
  }),

  // Coordinador de colegio: recordatorio completo (todo)
  school_reminder_todo: ({ schoolName, contactName, tripName, pendingItems }) => ({
    subject: `⏰ Resumen de pendientes — ${tripName}`,
    html: emailWrapper("Resumen de pendientes del viaje", `
      <p style="color:#18181b;font-size:16px">Hola${contactName ? `, <strong>${contactName}</strong>` : ""}</p>
      <p style="color:#52525b;font-size:14px;line-height:1.6">Aquí tienes un resumen de todo lo que está pendiente del colegio <strong>${schoolName}</strong> para el viaje <strong>${tripName}</strong>:</p>
      <div style="margin:24px 0;border-radius:12px;overflow:hidden;border:1px solid #e4e4e7">
        ${(pendingItems || []).map((item) => `
          <div style="padding:14px 16px;border-bottom:1px solid #f4f4f5;display:flex;align-items:center;gap:10px">
            <span style="font-size:18px">${escapeHtml(item.icon)}</span>
            <div>
              <div style="font-weight:600;font-size:14px;color:#18181b">${escapeHtml(item.label)}</div>
              <div style="font-size:13px;color:#71717a;margin-top:2px">${escapeHtml(item.detail)}</div>
            </div>
          </div>`).join("")}
      </div>
      <p style="color:#52525b;font-size:14px">Accede al portal escolar para completar toda la información.</p>
    `),
  }),

  // Coordinador de colegio: recordatorio genérico
  school_reminder: ({ schoolName, contactName, tripName, message }) => ({
    subject: `⏰ Recordatorio GIMELOOS — ${tripName}`,
    html: emailWrapper("Recordatorio del viaje escolar", `
      <p style="color:#18181b;font-size:16px">Hola${contactName ? `, <strong>${contactName}</strong>` : ""}</p>
      <p style="color:#52525b;font-size:14px;line-height:1.6">El equipo de GIMELOOS quiere recordarte algo sobre el viaje <strong>${tripName}</strong> del colegio <strong>${schoolName}</strong>.</p>
      <div style="margin:24px 0;padding:20px;background:#fff7ed;border-radius:12px;border-left:4px solid #f97316">
        <p style="color:#3f3f46;font-size:14px;margin:0">${escapeHtml(message || "Por favor, revisa el estado de tu portal escolar.")}</p>
      </div>
    `),
  }),

  // Admin: colegio envió una pregunta
  admin_school_question: ({ schoolName, question }) => ({
    subject: `❓ Nueva pregunta de colegio — ${schoolName}`,
    html: emailWrapper("Nueva pregunta de colegio", `
      <p style="color:#18181b;font-size:16px">Nueva pregunta de <strong>${schoolName}</strong></p>
      <div style="margin:16px 0;padding:16px;background:#f4f4f5;border-radius:12px">
        <p style="color:#3f3f46;font-size:14px;margin:0">${escapeHtml(question)}</p>
      </div>
      <p style="color:#52525b;font-size:14px">Accede al panel de administración para responder.</p>
    `),
  }),

  // Coordinador de colegio: admin respondió a su pregunta
  school_question_replied: ({ schoolName, contactName, question, reply }) => ({
    subject: `💬 El equipo GIMELOOS ha respondido tu pregunta`,
    html: emailWrapper("Respuesta a tu pregunta", `
      <p style="color:#18181b;font-size:16px">Hola${contactName ? `, <strong>${contactName}</strong>` : ""}</p>
      <p style="color:#52525b;font-size:14px;line-height:1.6">El equipo de GIMELOOS ha respondido a una pregunta del colegio <strong>${schoolName}</strong>.</p>
      <div style="margin:16px 0;padding:16px;background:#f4f4f5;border-radius:12px">
        <div style="font-size:12px;color:#71717a;margin-bottom:4px">Tu pregunta:</div>
        <p style="color:#52525b;font-size:14px;margin:0">${escapeHtml(question)}</p>
      </div>
      <div style="margin:16px 0;padding:16px;background:#f0fdf4;border-radius:12px;border-left:4px solid #16a34a">
        <div style="font-size:12px;color:#71717a;margin-bottom:4px">Respuesta:</div>
        <p style="color:#15803d;font-size:14px;margin:0">${escapeHtml(reply)}</p>
      </div>
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
  const { error: authError } = await requireAdmin(request);
  if (authError) return authError;

  try {
    const body = await request.json();
    const { type, to, participantId, data } = body;

    if (!type || !templates[type]) {
      return NextResponse.json({ error: "Tipo de notificación inválido" }, { status: 400 });
    }

    // Intentar leer plantilla personalizada de Supabase
    let subject, html;
    const { data: dbTpl } = await supabaseAdmin.from("email_templates").select("subject, body").eq("id", type).maybeSingle();
    if (dbTpl) {
      // Interpolar variables {nombre}, {viaje}, etc.
      const d = data || {};
      const varMap = {
        nombre: d.participantName || d.schoolName || "",
        viaje: d.tripName || "",
        documento: d.docName || "",
        pago: d.paymentName || "",
        importe: d.amount || "",
        fecha: d.dueDate || "",
        dias: String(d.daysLeft || ""),
        pregunta: d.question || "",
        respuesta: d.reply || "",
        colegio: d.schoolName || "",
        // Sin coma: el admin escribe "Hola {coordinador}" y sale "Hola Juan" o "Hola " si no hay
        coordinador: d.contactName || "",
        // pendientes: soporta tanto pendingCount (número) como pendingItems (array)
        pendientes: String(d.pendingCount ?? d.pendingItems?.length ?? ""),
      };
      const interpolate = (str) => str.replace(/\{(\w+)\}/g, (_, k) => varMap[k] ?? `{${k}}`);
      subject = interpolate(dbTpl.subject);
      const bodyLines = interpolate(dbTpl.body).split("\n").map((l) => l.trim() ? `<p style="color:#52525b;font-size:14px;line-height:1.6;margin:0 0 12px">${escapeHtml(l)}</p>` : "").join("");
      html = emailWrapper(subject, bodyLines);
    } else {
      ({ subject, html } = templates[type](data || {}));
    }
    const recipient = to || ADMIN_EMAIL;

    // Enviar email
    const { error: emailError } = await resend.emails.send({
      from: FROM,
      to: recipient,
      subject,
      html,
    });

    if (emailError) {
      console.error("Resend error:", JSON.stringify(emailError), emailError?.message, emailError?.name);
      return NextResponse.json({ error: emailError.message || emailError.name || "Error al enviar email" }, { status: 500 });
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
        // No duplicar: si ya hay una notificación sin leer del mismo tipo y contenido, la actualizamos en vez de insertar
        const { data: existing } = await supabaseAdmin
          .from("notifications")
          .select("id")
          .eq("participant_id", participantId)
          .eq("type", type)
          .eq("body", msg.body)
          .eq("read", false)
          .limit(1);

        if (existing?.length) {
          await supabaseAdmin.from("notifications").update({ created_at: new Date().toISOString() }).eq("id", existing[0].id);
        } else {
          await supabaseAdmin.from("notifications").insert({
            participant_id: participantId,
            type,
            title: msg.title,
            body: msg.body,
          });
        }
      }
    }

    return NextResponse.json({ ok: true });
  } catch (err) {
    console.error("notify error:", err);
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}

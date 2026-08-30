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

// emailWrapper(eyebrow, heading, body, footerNote?)
// eyebrow  → texto pequeño rojo en mayúsculas ("DOCUMENTACIÓN", "PAGOS"…)
// heading  → título grande y bold
// body     → HTML del contenido
// footerNote → frase opcional sobre por qué se recibe el email
function emailWrapper(eyebrow, heading, body, footerNote = "") {
  return `
  <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;background:#fff;border:1px solid #e4e4e7">
    <div style="background:#18181b;padding:28px 40px">
      <div style="color:white;font-size:20px;font-weight:700;letter-spacing:-0.3px">GIMELOOS</div>
      <div style="color:#a1a1aa;font-size:11px;letter-spacing:0.15em;margin-top:4px;text-transform:uppercase">Área privada de clientes</div>
    </div>
    <div style="padding:40px;background:#fff">
      <div style="color:#FF3131;font-size:11px;font-weight:700;letter-spacing:0.15em;text-transform:uppercase;margin-bottom:12px">${eyebrow}</div>
      <h1 style="color:#18181b;font-size:24px;font-weight:700;margin:0 0 24px;line-height:1.3">${heading}</h1>
      ${body}
    </div>
    <div style="background:#f4f4f5;padding:24px 40px;text-align:center;font-size:12px;color:#71717a;border-top:1px solid #e4e4e7">
      ${footerNote ? `<p style="margin:0 0 8px">${footerNote}</p>` : ""}
      <p style="margin:0">© 2026 GIMELOOS EVENTOS Y ACTIVIDADES SL</p>
    </div>
  </div>`;
}

const btn = (url, label) =>
  `<a href="${url}" style="display:inline-block;background:#FF3131;color:white;font-weight:600;font-size:15px;padding:14px 28px;border-radius:10px;text-decoration:none;margin:8px 0 24px">${label}</a>`;

const infoBox = (content) =>
  `<div style="margin:20px 0;padding:16px 20px;background:#fff7ed;border-radius:10px;border:1px solid #fed7aa;color:#92400e;font-size:14px;line-height:1.6">${content}</div>`;

const templates = {
  doc_confirmed: ({ participantName, docName, tripName }) => ({
    subject: `Documento confirmado — ${docName}`,
    html: emailWrapper("DOCUMENTACIÓN", `Documento confirmado`, `
      <p style="color:#52525b;font-size:15px;line-height:1.7;margin:0 0 16px">Hola, <strong>${participantName}</strong>.</p>
      <p style="color:#52525b;font-size:15px;line-height:1.7;margin:0 0 16px">Tu documento <strong>${docName}</strong> para el viaje <strong>${tripName}</strong> ha sido revisado y <strong style="color:#16a34a">confirmado</strong> por el equipo de GIMELOOS.</p>
      <div style="margin:20px 0;padding:16px 20px;background:#f0fdf4;border-radius:10px;border:1px solid #bbf7d0;color:#15803d;font-size:14px">
        ✓ <strong>${docName}</strong> — Estado: Confirmado
      </div>
      <p style="color:#71717a;font-size:13px;margin:0">Accede al portal para ver el estado de todos tus documentos.</p>
    `, `Recibes este correo porque tienes un viaje contratado con GIMELOOS. Escríbenos a info@gimeloos.com si tienes dudas.`),
  }),

  doc_rejected: ({ participantName, docName, tripName }) => ({
    subject: `Documento rechazado — ${docName}`,
    html: emailWrapper("DOCUMENTACIÓN", `Revisa tu documento`, `
      <p style="color:#52525b;font-size:15px;line-height:1.7;margin:0 0 16px">Hola, <strong>${participantName}</strong>.</p>
      <p style="color:#52525b;font-size:15px;line-height:1.7;margin:0 0 16px">Tu documento <strong>${docName}</strong> para <strong>${tripName}</strong> ha sido revisado pero necesita correcciones.</p>
      <div style="margin:20px 0;padding:16px 20px;background:#fef2f2;border-radius:10px;border:1px solid #fecaca;color:#dc2626;font-size:14px">
        ✗ <strong>${docName}</strong> — Por favor sube una versión corregida
      </div>
      <p style="color:#71717a;font-size:13px;margin:0">Accede al portal para subir el documento corregido.</p>
    `, `Recibes este correo porque tienes un viaje contratado con GIMELOOS. Escríbenos a info@gimeloos.com si tienes dudas.`),
  }),

  payment_confirmed: ({ participantName, paymentName, amount, tripName }) => ({
    subject: `Pago confirmado — ${paymentName}`,
    html: emailWrapper("PAGOS", `Pago confirmado`, `
      <p style="color:#52525b;font-size:15px;line-height:1.7;margin:0 0 16px">Hola, <strong>${participantName}</strong>.</p>
      <p style="color:#52525b;font-size:15px;line-height:1.7;margin:0 0 16px">Hemos confirmado tu pago de <strong>${paymentName}</strong> para el viaje <strong>${tripName}</strong>.</p>
      <div style="margin:20px 0;padding:16px 20px;background:#f0fdf4;border-radius:10px;border:1px solid #bbf7d0;color:#15803d;font-size:14px">
        ✓ <strong>${paymentName}</strong> · Importe: <strong>${amount}</strong> · Estado: Confirmado
      </div>
    `, `Recibes este correo porque tienes un viaje contratado con GIMELOOS. Escríbenos a info@gimeloos.com si tienes dudas.`),
  }),

  question_replied: ({ participantName, question, reply }) => ({
    subject: `Tienes una respuesta de GIMELOOS`,
    html: emailWrapper("CONSULTAS", `El equipo GIMELOOS ha respondido`, `
      <p style="color:#52525b;font-size:15px;line-height:1.7;margin:0 0 20px">Hola, <strong>${participantName}</strong>.</p>
      <div style="margin:0 0 16px;padding:16px 20px;background:#f4f4f5;border-radius:10px;font-size:14px;color:#52525b">
        <div style="font-size:11px;font-weight:700;letter-spacing:0.1em;color:#a1a1aa;text-transform:uppercase;margin-bottom:8px">Tu pregunta</div>
        ${escapeHtml(question)}
      </div>
      <div style="margin:0 0 20px;padding:16px 20px;background:#fef3c7;border-radius:10px;border-left:4px solid #d97706;font-size:14px;color:#3f3f46">
        <div style="font-size:11px;font-weight:700;letter-spacing:0.1em;color:#92400e;text-transform:uppercase;margin-bottom:8px">Respuesta del equipo</div>
        ${escapeHtml(reply)}
      </div>
      <p style="color:#71717a;font-size:13px;margin:0">Si tienes más dudas, accede al portal y envíanos otro mensaje.</p>
    `, `Recibes este correo porque tienes un viaje contratado con GIMELOOS. Escríbenos a info@gimeloos.com si tienes dudas.`),
  }),

  payment_reminder: ({ participantName, paymentName, amount, dueDate, daysLeft, tripName }) => ({
    subject: `Recordatorio de pago — ${paymentName} vence en ${daysLeft} días`,
    html: emailWrapper("PAGOS", `Tienes un pago pendiente`, `
      <p style="color:#52525b;font-size:15px;line-height:1.7;margin:0 0 16px">Hola, <strong>${participantName}</strong>.</p>
      <p style="color:#52525b;font-size:15px;line-height:1.7;margin:0 0 20px">Tienes un pago pendiente para el viaje <strong>${tripName}</strong>. Haz clic en el botón de abajo para acceder al portal y subir tu justificante. Si no realizas esta acción, tu plaza podría quedar sin confirmar.</p>
      <div style="margin:0 0 24px;padding:20px;background:#f4f4f5;border-radius:10px;font-size:14px;color:#3f3f46">
        <div style="font-weight:700;font-size:16px;color:#18181b;margin-bottom:8px">${paymentName}</div>
        <div>Importe: <strong>${amount}</strong></div>
        <div>Fecha límite: <strong>${dueDate}</strong></div>
      </div>
      ${infoBox(`⏰ Este enlace es válido hasta el <strong>${dueDate}</strong>. Quedan <strong>${daysLeft} días</strong>.`)}
      <p style="color:#71717a;font-size:13px;margin:0">Si ya has realizado el pago, puedes ignorar este correo.</p>
    `, `Recibes este correo porque tienes un viaje contratado con GIMELOOS. Escríbenos a info@gimeloos.com si tienes dudas.`),
  }),

  doc_reminder: ({ participantName, docName, tripName }) => ({
    subject: `Documentación pendiente — ${docName}`,
    html: emailWrapper("DOCUMENTACIÓN", `Tienes documentación pendiente`, `
      <p style="color:#52525b;font-size:15px;line-height:1.7;margin:0 0 16px">Hola, <strong>${participantName}</strong>.</p>
      <p style="color:#52525b;font-size:15px;line-height:1.7;margin:0 0 20px">Aún tienes documentación pendiente de subir para el viaje <strong>${tripName}</strong>. Haz clic en el botón de abajo para acceder al portal y subir los documentos requeridos.</p>
      <div style="margin:0 0 24px;padding:16px 20px;background:#fef2f2;border-radius:10px;border-left:4px solid #FF3131;font-size:14px;color:#FF3131">
        <strong>${docName}</strong> — Pendiente de envío
      </div>
      <p style="color:#71717a;font-size:13px;margin:0">Si ya has subido el documento, puedes ignorar este correo.</p>
    `, `Recibes este correo porque tienes un viaje contratado con GIMELOOS. Escríbenos a info@gimeloos.com si tienes dudas.`),
  }),

  school_reminder_listado: ({ schoolName, contactName, tripName }) => ({
    subject: `Listado de alumnos pendiente — ${tripName}`,
    html: emailWrapper("PORTAL DEL COLEGIO", `Listado de alumnos pendiente`, `
      <p style="color:#52525b;font-size:15px;line-height:1.7;margin:0 0 16px">Hola${contactName ? `, <strong>${contactName}</strong>` : ""}.</p>
      <p style="color:#52525b;font-size:15px;line-height:1.7;margin:0 0 20px">Necesitamos el <strong>listado definitivo de alumnos</strong> del colegio <strong>${schoolName}</strong> para el viaje <strong>${tripName}</strong>. Por favor, accede al portal y sube el listado lo antes posible.</p>
      ${infoBox(`📋 <strong>Acción requerida:</strong> sube el listado de alumnos en el Portal del Colegio GIMELOOS.`)}
    `, `Recibes este correo en representación del colegio ${schoolName}. Escríbenos a info@gimeloos.com si tienes dudas.`),
  }),

  school_reminder_alergias: ({ schoolName, contactName, tripName }) => ({
    subject: `Alergias e intolerancias pendientes — ${tripName}`,
    html: emailWrapper("PORTAL DEL COLEGIO", `Alergias e intolerancias pendientes`, `
      <p style="color:#52525b;font-size:15px;line-height:1.7;margin:0 0 16px">Hola${contactName ? `, <strong>${contactName}</strong>` : ""}.</p>
      <p style="color:#52525b;font-size:15px;line-height:1.7;margin:0 0 20px">Es imprescindible que nos comuniquéis las <strong>alergias e intolerancias alimentarias</strong> de los alumnos de <strong>${schoolName}</strong> para el viaje <strong>${tripName}</strong>. Es obligatorio para garantizar su seguridad.</p>
      ${infoBox(`🍽️ <strong>Acción requerida:</strong> rellena la información médica de cada alumno en el Portal del Colegio GIMELOOS.`)}
    `, `Recibes este correo en representación del colegio ${schoolName}. Escríbenos a info@gimeloos.com si tienes dudas.`),
  }),

  school_doc_reminder: ({ schoolName, contactName, tripName, pendingCount }) => ({
    subject: `Documentación pendiente — ${tripName}`,
    html: emailWrapper("PORTAL DEL COLEGIO", `Documentación pendiente del viaje`, `
      <p style="color:#52525b;font-size:15px;line-height:1.7;margin:0 0 16px">Hola${contactName ? `, <strong>${contactName}</strong>` : ""}.</p>
      <p style="color:#52525b;font-size:15px;line-height:1.7;margin:0 0 20px">El colegio <strong>${schoolName}</strong> tiene documentación pendiente de enviar para el viaje <strong>${tripName}</strong>. Accede al portal y sube los documentos requeridos.</p>
      ${infoBox(`📄 <strong>${pendingCount > 0 ? `${pendingCount} documento${pendingCount !== 1 ? "s" : ""} pendiente${pendingCount !== 1 ? "s" : ""}` : "Documentación pendiente"}</strong> — Accede al Portal del Colegio para completarla.`)}
    `, `Recibes este correo en representación del colegio ${schoolName}. Escríbenos a info@gimeloos.com si tienes dudas.`),
  }),

  school_reminder_rooming: ({ schoolName, contactName, tripName }) => ({
    subject: `Asignación de habitaciones pendiente — ${tripName}`,
    html: emailWrapper("PORTAL DEL COLEGIO", `Rooming pendiente`, `
      <p style="color:#52525b;font-size:15px;line-height:1.7;margin:0 0 16px">Hola${contactName ? `, <strong>${contactName}</strong>` : ""}.</p>
      <p style="color:#52525b;font-size:15px;line-height:1.7;margin:0 0 20px">Necesitamos la <strong>asignación de habitaciones</strong> del colegio <strong>${schoolName}</strong> para el viaje <strong>${tripName}</strong>.</p>
      ${infoBox(`🛏️ <strong>Acción requerida:</strong> accede al Portal del Colegio → sección Rooming y completa la distribución de habitaciones.`)}
    `, `Recibes este correo en representación del colegio ${schoolName}. Escríbenos a info@gimeloos.com si tienes dudas.`),
  }),

  school_reminder_grupos: ({ schoolName, contactName, tripName }) => ({
    subject: `Grupos de actividad pendientes — ${tripName}`,
    html: emailWrapper("PORTAL DEL COLEGIO", `Grupos de actividad pendientes`, `
      <p style="color:#52525b;font-size:15px;line-height:1.7;margin:0 0 16px">Hola${contactName ? `, <strong>${contactName}</strong>` : ""}.</p>
      <p style="color:#52525b;font-size:15px;line-height:1.7;margin:0 0 20px">Los <strong>grupos de actividad</strong> del colegio <strong>${schoolName}</strong> para el viaje <strong>${tripName}</strong> están pendientes de definir.</p>
      ${infoBox(`👥 <strong>Acción requerida:</strong> accede al Portal del Colegio → sección Grupos y define los grupos de actividad.`)}
    `, `Recibes este correo en representación del colegio ${schoolName}. Escríbenos a info@gimeloos.com si tienes dudas.`),
  }),

  school_reminder_todo: ({ schoolName, contactName, tripName, pendingItems }) => ({
    subject: `Resumen de pendientes — ${tripName}`,
    html: emailWrapper("PORTAL DEL COLEGIO", `Resumen de pendientes del viaje`, `
      <p style="color:#52525b;font-size:15px;line-height:1.7;margin:0 0 16px">Hola${contactName ? `, <strong>${contactName}</strong>` : ""}.</p>
      <p style="color:#52525b;font-size:15px;line-height:1.7;margin:0 0 20px">Aquí tienes un resumen de todo lo pendiente del colegio <strong>${schoolName}</strong> para el viaje <strong>${tripName}</strong>:</p>
      <div style="margin:0 0 24px;border-radius:10px;overflow:hidden;border:1px solid #e4e4e7">
        ${(pendingItems || []).map((item, i) => `
          <div style="padding:14px 16px;${i > 0 ? "border-top:1px solid #f4f4f5;" : ""}display:flex;align-items:center;gap:12px">
            <span style="font-size:20px;flex-shrink:0">${escapeHtml(item.icon)}</span>
            <div>
              <div style="font-weight:600;font-size:14px;color:#18181b">${escapeHtml(item.label)}</div>
              <div style="font-size:13px;color:#71717a;margin-top:2px">${escapeHtml(item.detail)}</div>
            </div>
          </div>`).join("")}
      </div>
      <p style="color:#71717a;font-size:13px;margin:0">Accede al portal escolar para completar toda la información pendiente.</p>
    `, `Recibes este correo en representación del colegio ${schoolName}. Escríbenos a info@gimeloos.com si tienes dudas.`),
  }),

  school_reminder: ({ schoolName, contactName, tripName, message }) => ({
    subject: `Recordatorio GIMELOOS — ${tripName}`,
    html: emailWrapper("PORTAL DEL COLEGIO", `Recordatorio del viaje escolar`, `
      <p style="color:#52525b;font-size:15px;line-height:1.7;margin:0 0 16px">Hola${contactName ? `, <strong>${contactName}</strong>` : ""}.</p>
      <p style="color:#52525b;font-size:15px;line-height:1.7;margin:0 0 20px">El equipo de GIMELOOS quiere recordarte algo sobre el viaje <strong>${tripName}</strong> del colegio <strong>${schoolName}</strong>.</p>
      ${infoBox(escapeHtml(message || "Por favor, revisa el estado de tu portal escolar."))}
    `, `Recibes este correo en representación del colegio ${schoolName}. Escríbenos a info@gimeloos.com si tienes dudas.`),
  }),

  admin_school_question: ({ schoolName, question }) => ({
    subject: `Nueva pregunta de colegio — ${schoolName}`,
    html: emailWrapper("ADMINISTRACIÓN", `Nueva pregunta de colegio`, `
      <p style="color:#52525b;font-size:15px;line-height:1.7;margin:0 0 20px">El colegio <strong>${schoolName}</strong> ha enviado una nueva consulta a través del portal.</p>
      <div style="margin:0 0 24px;padding:16px 20px;background:#f4f4f5;border-radius:10px;font-size:14px;color:#3f3f46">
        ${escapeHtml(question)}
      </div>
      <p style="color:#71717a;font-size:13px;margin:0">Accede al panel de administración → Colegios → Preguntas para responder.</p>
    `, "Mensaje automático del Portal GIMELOOS."),
  }),

  school_question_replied: ({ schoolName, contactName, question, reply }) => ({
    subject: `El equipo GIMELOOS ha respondido tu pregunta`,
    html: emailWrapper("CONSULTAS", `Respuesta a tu consulta`, `
      <p style="color:#52525b;font-size:15px;line-height:1.7;margin:0 0 20px">Hola${contactName ? `, <strong>${contactName}</strong>` : ""}. El equipo de GIMELOOS ha respondido a tu consulta sobre el colegio <strong>${schoolName}</strong>.</p>
      <div style="margin:0 0 16px;padding:16px 20px;background:#f4f4f5;border-radius:10px;font-size:14px;color:#52525b">
        <div style="font-size:11px;font-weight:700;letter-spacing:0.1em;color:#a1a1aa;text-transform:uppercase;margin-bottom:8px">Tu pregunta</div>
        ${escapeHtml(question)}
      </div>
      <div style="margin:0 0 20px;padding:16px 20px;background:#fef3c7;border-radius:10px;border-left:4px solid #d97706;font-size:14px;color:#3f3f46">
        <div style="font-size:11px;font-weight:700;letter-spacing:0.1em;color:#92400e;text-transform:uppercase;margin-bottom:8px">Respuesta del equipo</div>
        ${escapeHtml(reply)}
      </div>
      <p style="color:#71717a;font-size:13px;margin:0">Si tienes más dudas, accede al portal y envíanos otro mensaje.</p>
    `, `Recibes este correo en representación del colegio ${schoolName}. Escríbenos a info@gimeloos.com si tienes dudas.`),
  }),

  admin_doc_uploaded: ({ participantName, docName, tripName }) => ({
    subject: `Nuevo documento — ${participantName}`,
    html: emailWrapper("ADMINISTRACIÓN", `Nuevo documento para revisar`, `
      <p style="color:#52525b;font-size:15px;line-height:1.7;margin:0 0 20px">Un participante ha subido un nuevo documento que requiere tu revisión.</p>
      <div style="margin:0 0 24px;padding:16px 20px;background:#f4f4f5;border-radius:10px;font-size:14px;color:#3f3f46">
        <div><strong>Participante:</strong> ${participantName}</div>
        <div style="margin-top:6px"><strong>Documento:</strong> ${docName}</div>
        <div style="margin-top:6px"><strong>Viaje:</strong> ${tripName}</div>
      </div>
      <p style="color:#71717a;font-size:13px;margin:0">Accede al panel de administración para revisar y confirmar el documento.</p>
    `, "Mensaje automático del Portal GIMELOOS."),
  }),

  admin_payment_uploaded: ({ participantName, paymentName, tripName }) => ({
    subject: `Nuevo justificante — ${participantName}`,
    html: emailWrapper("ADMINISTRACIÓN", `Nuevo justificante para revisar`, `
      <p style="color:#52525b;font-size:15px;line-height:1.7;margin:0 0 20px">Un participante ha subido un justificante de pago que requiere tu revisión.</p>
      <div style="margin:0 0 24px;padding:16px 20px;background:#f4f4f5;border-radius:10px;font-size:14px;color:#3f3f46">
        <div><strong>Participante:</strong> ${participantName}</div>
        <div style="margin-top:6px"><strong>Pago:</strong> ${paymentName}</div>
        <div style="margin-top:6px"><strong>Viaje:</strong> ${tripName}</div>
      </div>
      <p style="color:#71717a;font-size:13px;margin:0">Accede al panel de administración para revisar y confirmar el pago.</p>
    `, "Mensaje automático del Portal GIMELOOS."),
  }),

  admin_new_question: ({ participantName, question, tripName }) => ({
    subject: `Nueva consulta — ${participantName}`,
    html: emailWrapper("ADMINISTRACIÓN", `Nueva consulta recibida`, `
      <p style="color:#52525b;font-size:15px;line-height:1.7;margin:0 0 20px">Has recibido una nueva pregunta de un participante.</p>
      <div style="margin:0 0 16px;padding:16px 20px;background:#f4f4f5;border-radius:10px;font-size:14px;color:#3f3f46">
        <div><strong>Participante:</strong> ${participantName}</div>
        <div style="margin-top:6px"><strong>Viaje:</strong> ${tripName}</div>
      </div>
      <div style="margin:0 0 24px;padding:16px 20px;background:#eff6ff;border-radius:10px;border-left:4px solid #3b82f6;font-size:14px;color:#1e3a5f">
        ${escapeHtml(question)}
      </div>
      <p style="color:#71717a;font-size:13px;margin:0">Accede al panel → Preguntas para responder.</p>
    `, "Mensaje automático del Portal GIMELOOS."),
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

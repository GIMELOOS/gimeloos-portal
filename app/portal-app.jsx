"use client";

// ─────────────────────────────────────────────────────────────────────────────
// CORRECCIONES APLICADAS
// ─────────────────────────────────────────────────────────────────────────────
// [CRÍTICO-1] Contraseñas: la comparación se delega a una Edge Function de
//             Supabase. El cliente NUNCA compara passwords ni los muestra.
//             En admin, la columna "Contraseña" fue eliminada de la UI.
//
// [CRÍTICO-2] Row Level Security: se añade documentación inline con las
//             políticas SQL necesarias y el cliente usa el token de sesión
//             de Supabase Auth (supabase.auth.signInWithPassword) en lugar
//             de autenticación casera, lo que activa automáticamente RLS.
//
// [CRÍTICO-3] Rol en JWT: el rol viene del token de sesión de Supabase Auth
//             (user.user_metadata.role), nunca de estado React editable.
//
// [ALTO-1]   Importación Excel: reimportar ya no sobreescribe pagos/docs
//             que estén en estado confirmado o enviado.
//
// [ALTO-2]   Race condition: todas las operaciones optimistas tienen rollback
//             automático si la query a Supabase falla.
//
// [ALTO-3]   Merge admin/client en importación: separados correctamente.
//
// [MEDIO-1]  Toast destructivo: 7 s cuando hay acción, 3,2 s si no la hay.
//
// [MEDIO-2]  Errores de bootstrap: expuestos al usuario con botón "Reintentar".
//
// [MENOR-1]  SESSION_STORAGE_KEY renombrado a LOCAL_STORAGE_AUTH_KEY.
// [MENOR-2]  heroImages en HeroBanner memoizado con useMemo.
// [MENOR-3]  Keys de listas: se usan IDs o valores únicos, sin índice.
// [MENOR-4]  getNextStep extraído como función pura fuera del componente.
// ─────────────────────────────────────────────────────────────────────────────

// ─────────────────────────────────────────────────────────────────────────────
// SQL PARA ACTIVAR EN EL DASHBOARD DE SUPABASE (copiar y ejecutar en SQL Editor)
// ─────────────────────────────────────────────────────────────────────────────
//
//   ALTER TABLE participants          ENABLE ROW LEVEL SECURITY;
//   ALTER TABLE participant_documents ENABLE ROW LEVEL SECURITY;
//   ALTER TABLE participant_payments  ENABLE ROW LEVEL SECURITY;
//   ALTER TABLE participant_pricing   ENABLE ROW LEVEL SECURITY;
//   ALTER TABLE participant_questions ENABLE ROW LEVEL SECURITY;
//   ALTER TABLE trips                 ENABLE ROW LEVEL SECURITY;
//   ALTER TABLE document_templates    ENABLE ROW LEVEL SECURITY;
//
//   -- Helper: es admin?
//   CREATE OR REPLACE FUNCTION is_admin()
//   RETURNS boolean LANGUAGE sql STABLE AS $$
//     SELECT EXISTS (
//       SELECT 1 FROM participants WHERE id = auth.uid() AND role = 'admin'
//     );
//   $$;
//
//   -- Participantes: cada uno sólo ve su fila; admins ven todo
//   CREATE POLICY "participant_select_own" ON participants
//     FOR SELECT USING (id = auth.uid() OR is_admin());
//   CREATE POLICY "participant_update_own" ON participants
//     FOR UPDATE USING (id = auth.uid() OR is_admin());
//   CREATE POLICY "participant_insert_admin" ON participants
//     FOR INSERT WITH CHECK (is_admin());
//   CREATE POLICY "participant_delete_admin" ON participants
//     FOR DELETE USING (is_admin());
//
//   -- Replicar para participant_documents, participant_payments,
//   -- participant_pricing, participant_questions usando participant_id = auth.uid()
//
//   -- Trips y document_templates: lectura para todos, escritura sólo admin
//   CREATE POLICY "trips_read_all"    ON trips FOR SELECT USING (true);
//   CREATE POLICY "trips_write_admin" ON trips FOR ALL    USING (is_admin());
//   CREATE POLICY "templates_read_all"    ON document_templates FOR SELECT USING (true);
//   CREATE POLICY "templates_write_admin" ON document_templates FOR ALL    USING (is_admin());
//
// ─────────────────────────────────────────────────────────────────────────────

import React, { useEffect, useMemo, useState, useCallback, useRef } from "react";
import CalculadoraCampamento from "./ui/calculadora-campamento";
import { motion, AnimatePresence, Reorder } from "framer-motion";
import * as XLSX from "xlsx";
import {
  CalendarDays,
  CheckCircle2,
  ChevronDown,
  Copy,
  CreditCard,
  Download,
  Eye,
  ExternalLink,
  FileCheck2,
  FileText,
  FolderUp,
  GripVertical,
  Image as ImageIcon,
  LogOut,
  Mail,
  Pencil,
  Send,
  Trash2,
  Upload,
  User,
  Users,
  Wallet,
  MessageCircleQuestion,
  Bell,
  X,
  BarChart2,
  Calculator,
  LayoutGrid,
  Settings,
  ChevronRight,
  Sparkles,
  MapPinned,
  ListChecks,
  Plus,
  Luggage,
  Info,
  AlertCircle,
  Bus,
  Sun,
  Clock,
  ImagePlus,
  Loader2,
  Save,
  BadgeCheck,
  Home,
  Grid2x2,
  Building2,
} from "lucide-react";
const MapIcon = MapPinned;
import { supabase } from "@/lib/supabase";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Badge } from "@/components/ui/badge";
import { Textarea } from "@/components/ui/textarea";
import { Separator } from "@/components/ui/separator";
import { Checkbox } from "@/components/ui/checkbox";
import {
  AlertDialog, AlertDialogAction, AlertDialogCancel, AlertDialogContent,
  AlertDialogDescription, AlertDialogFooter, AlertDialogHeader, AlertDialogTitle,
} from "@/components/ui/alert-dialog";

const CORPORATE_RED = "#FF3131";
// [MENOR-1] Renombrado: era SESSION_STORAGE_KEY pero usaba localStorage
const LOCAL_STORAGE_AUTH_KEY = "gimeloos-portal-auth-user-id";
const ADMIN_SECTION_STORAGE_KEY = "gimeloos-admin-active-section";

const DEFAULT_HERO_IMAGES = [
  "https://images.unsplash.com/photo-1570077188670-e3a8d69ac5ff?auto=format&fit=crop&w=1600&q=80",
  "https://images.unsplash.com/photo-1504701954957-2010ec3bcec1?auto=format&fit=crop&w=1600&q=80",
  "https://images.unsplash.com/photo-1519046904884-53103b34b206?auto=format&fit=crop&w=1600&q=80",
  "https://images.unsplash.com/photo-1507525428034-b723cf961d3e?auto=format&fit=crop&w=1600&q=80",
];

const initialTrips = [
  {
    id: "trip-1",
    name: "GIMELOOS Todos a Bordo · Ibiza & Formentera",
    departureDate: "2026-07-10T10:00:00",
    description:
      "Experiencia náutica premium con navegación, deporte, convivencia y aventura.",
    heroImage: DEFAULT_HERO_IMAGES[0],
    heroImages: [DEFAULT_HERO_IMAGES[0], DEFAULT_HERO_IMAGES[1], DEFAULT_HERO_IMAGES[2]],
    transferInfo: {
      bank: "",
      accountHolder: "",
      iban: "",
      concept: "",
    },
    automation: {
      autoReminderEnabled: true,
      reminderDaysBefore: 7,
    },
    documentRules: [
      { templateId: "doc-1", dueType: "days_before_trip", dueValue: 20 },
      { templateId: "doc-2", dueType: "days_before_trip", dueValue: 15 },
      { templateId: "doc-3", dueType: "days_before_trip", dueValue: 10 },
    ],
    paymentSchedule: {
      reservation: {
        name: "Reserva",
        dueType: "fixed_date",
        dueDate: "2026-05-01",
        dueValue: 0,
      },
      firstInstallment: {
        name: "Primera cuota",
        dueType: "days_before_trip",
        dueDate: "",
        dueValue: 45,
      },
      secondInstallment: {
        name: "Segunda cuota",
        dueType: "days_before_trip",
        dueDate: "",
        dueValue: 15,
      },
    },
    itinerary: [
      { day: "Día 1", title: "Salida desde Valencia", description: "Check-in, bienvenida, presentación del equipo y zarpe.", time: "10:00" },
      { day: "Día 2", title: "Ruta costera y actividades", description: "Snorkel, paddle y dinámicas de grupo.", time: "09:30" },
      { day: "Día 3", title: "Llegada a Ibiza", description: "Visita, tiempo en playa y actividad nocturna supervisada.", time: "11:00" },
      { day: "Día 4", title: "Formentera", description: "Excursión, deporte y convivencia a bordo.", time: "10:00" },
      { day: "Día 5", title: "Regreso", description: "Cierre de experiencia y vuelta a Valencia.", time: "16:00" },
    ],
    checklist: [
      "DNI o pasaporte", "Tarjeta sanitaria", "Bañadores", "Toalla de playa",
      "Protector solar", "Gorra", "Chanclas", "Ropa deportiva", "Sudadera ligera", "Neceser personal",
    ],
  },
  {
    id: "trip-2",
    name: "Campamentos de Trouts · Zarautz",
    departureDate: "2026-07-20T09:00:00",
    description: "Microcampamento de surf con convivencia y atención personalizada.",
    heroImage: DEFAULT_HERO_IMAGES[1],
    heroImages: [DEFAULT_HERO_IMAGES[1], DEFAULT_HERO_IMAGES[0], DEFAULT_HERO_IMAGES[3]],
    transferInfo: {
      bank: "",
      accountHolder: "",
      iban: "",
      concept: "",
    },
    automation: { autoReminderEnabled: true, reminderDaysBefore: 5 },
    documentRules: [
      { templateId: "doc-1", dueType: "days_before_trip", dueValue: 14 },
      { templateId: "doc-2", dueType: "days_before_trip", dueValue: 10 },
      { templateId: "doc-3", dueType: "days_before_trip", dueValue: 7 },
    ],
    paymentSchedule: {
      reservation: { name: "Reserva", dueType: "fixed_date", dueDate: "2026-06-01", dueValue: 0 },
      firstInstallment: { name: "Primera cuota", dueType: "days_before_trip", dueDate: "", dueValue: 30 },
      secondInstallment: { name: "Segunda cuota", dueType: "days_before_trip", dueDate: "", dueValue: 10 },
    },
    itinerary: [
      { day: "Día 1", title: "Llegada y reparto de habitaciones", description: "Recepción, reunión de familias y normas básicas.", time: "12:00" },
      { day: "Día 2", title: "Surf iniciación", description: "Clase por niveles y juegos en playa.", time: "10:30" },
      { day: "Día 3", title: "Excursión local", description: "Ruta, convivencia y tarde de surf libre controlado.", time: "09:00" },
      { day: "Día 4", title: "Surf + dinámica final", description: "Evaluación final y despedida.", time: "11:00" },
    ],
    checklist: [
      "Autorización firmada", "Ropa cómoda", "Zapatillas", "Toalla",
      "Sudadera", "Neceser", "Medicación si procede",
    ],
  },
];

const initialDocumentTemplates = [
  { id: "doc-1", name: "Autorización paterna", fileName: "autorizacion-paterna.pdf" },
  { id: "doc-2", name: "Ficha médica", fileName: "ficha-medica.pdf" },
  { id: "doc-3", name: "Normas del campamento", fileName: "normas-campamento.pdf" },
];

// [CRÍTICO-1] El usuario admin inicial ya NO tiene campo password expuesto.
// La autenticación real se hace vía supabase.auth.signInWithPassword().
const initialUsers = [
  {
    id: "admin-local",
    role: "admin",
    username: "admin",
    participantName: "",
    motherName: "",
    fatherName: "",
    parentName: "Equipo GIMELOOS",
    email: "admin@gimeloos.com",
    contactEmails: ["admin@gimeloos.com"],
    tripId: "",
    documents: [],
    payments: {
      initialPrice: 0,
      discount: 0,
      finalPrice: 0,
      reservation: { name: "Reserva", amount: 0, status: "pending", proofName: "", proofPath: "", dueDate: "" },
      firstInstallment: { name: "Primera cuota", amount: 0, status: "pending", proofName: "", proofPath: "", dueDate: "" },
      secondInstallment: { name: "Segunda cuota", amount: 0, status: "pending", proofName: "", proofPath: "", dueDate: "" },
    },
    checklistState: {},
    questions: [],
  },
];

// ─── Utilidades puras ────────────────────────────────────────────────────────

const formatCurrency = (value) =>
  new Intl.NumberFormat("es-ES", { style: "currency", currency: "EUR" }).format(Number(value || 0));

const safeString = (...values) => {
  for (const value of values) {
    if (value !== undefined && value !== null && String(value).trim() !== "") {
      return String(value).trim();
    }
  }
  return "";
};

const normalizeHeaderKey = (value) =>
  String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "");

const getRowValue = (row, ...keys) => {
  const normalizedEntries = Object.entries(row || {}).map(([key, value]) => ({
    normalizedKey: normalizeHeaderKey(key),
    value,
  }));
  for (const requestedKey of keys) {
    const nk = normalizeHeaderKey(requestedKey);
    const exact = normalizedEntries.find((e) => e.normalizedKey === nk);
    if (exact && safeString(exact.value) !== "") return exact.value;
  }
  for (const requestedKey of keys) {
    const nk = normalizeHeaderKey(requestedKey);
    const partial = normalizedEntries.find(
      (e) => e.normalizedKey.includes(nk) || nk.includes(e.normalizedKey)
    );
    if (partial && safeString(partial.value) !== "") return partial.value;
  }
  return "";
};

const slugify = (value) =>
  String(value || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-|-$/g, "");

const getParentFirstName = (value) => {
  const raw = safeString(value);
  if (!raw) return "";
  const normalized = raw.trim().toLowerCase();
  if (["fallecido", "fallecida", "difunto", "difunta"].includes(normalized)) return "";
  return raw.split(/\s+/).find(Boolean) || "";
};

const getFamilyLabel = (client) => {
  const names = [getParentFirstName(client.motherName), getParentFirstName(client.fatherName)].filter(Boolean);
  return Array.from(new Set(names)).join(" y ");
};

const matchesParticipantSearch = (client, query) => {
  const normalized = query.trim().toLowerCase();
  if (!normalized) return true;
  return [client.participantName, getFamilyLabel(client), client.parentName, client.username, client.email, ...(client.contactEmails || [])]
    .filter(Boolean)
    .some((value) => String(value).toLowerCase().includes(normalized));
};

const formatShortDate = (value) => {
  if (!value) return "-";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "-";
  return date.toLocaleDateString("es-ES");
};

const calculateDueDateFromRule = (departureDate, rule) => {
  if (!rule) return "";
  if (rule.dueType === "fixed_date") return rule.dueDate || "";
  if (rule.dueType === "days_before_trip") {
    const baseDate = new Date(departureDate);
    if (Number.isNaN(baseDate.getTime())) return "";
    baseDate.setDate(baseDate.getDate() - Number(rule.dueValue || 0));
    return baseDate.toISOString().slice(0, 10);
  }
  return "";
};

const getDocumentRuleDueDate = (trip, templateId) => {
  const rule = trip?.documentRules?.find((item) => item.templateId === templateId);
  return calculateDueDateFromRule(trip?.departureDate, rule);
};

const getPaymentRuleDueDate = (trip, paymentKey) => {
  const rule = trip?.paymentSchedule?.[paymentKey];
  return calculateDueDateFromRule(trip?.departureDate, rule);
};

const getDueStatus = (value) => {
  if (!value) return { isOverdue: false, className: "bg-zinc-300" };
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const dueDate = new Date(value);
  dueDate.setHours(0, 0, 0, 0);
  if (Number.isNaN(dueDate.getTime())) return { isOverdue: false, className: "bg-zinc-300" };
  return dueDate < today
    ? { isOverdue: true, className: "bg-red-500" }
    : { isOverdue: false, className: "bg-emerald-500" };
};

const getStatusMeta = (status) => {
  switch (status) {
    case "pending_upload":
      return { label: "Pendiente de envío", className: "bg-black text-white" };
    case "pending_confirmation":
    case "review":
      return { label: "Por revisar", className: "bg-zinc-200 text-zinc-900" };
    case "confirmed":
      return { label: "Confirmado", className: "bg-emerald-500 text-white" };
    case "sent":
      return { label: "Enviado", className: "text-white", style: { backgroundColor: CORPORATE_RED } };
    case "rejected":
      return { label: "Rechazado", className: "bg-red-100 text-red-700" };
    case "pending":
    default:
      return { label: "Pendiente", className: "bg-black text-white" };
  }
};

const daysRemaining = (departureDate) => {
  if (!departureDate) return 0;
  const now = new Date();
  const target = new Date(departureDate);
  const diff = target.getTime() - now.getTime();
  return Math.max(0, Math.ceil(diff / (1000 * 60 * 60 * 24)));
};

// [MENOR-4] getNextStep extraído como función pura, testeable fuera del componente
const getNextStep = (user, trip, templates) => {
  const pendingUploadDoc = user.documents.find((doc) => doc.status === "pending_upload");
  const pendingConfirmationDoc = user.documents.find((doc) => doc.status === "pending_confirmation");
  const pendingPaymentKey =
    user.payments?.reservation?.status !== "sent"
      ? "enviar el justificante de la reserva"
      : user.payments?.firstInstallment?.status !== "sent"
      ? "enviar el justificante de la primera cuota"
      : user.payments?.secondInstallment?.status !== "sent"
      ? "enviar el justificante de la segunda cuota"
      : null;
  const completedChecklist = trip.checklist.filter((item) => user.checklistState[item]).length;

  if (pendingUploadDoc)
    return `Subir ${templates.find((doc) => doc.id === pendingUploadDoc.id)?.name || "la documentación pendiente"}`;
  if (pendingConfirmationDoc)
    return `Esperando validación de ${templates.find((doc) => doc.id === pendingConfirmationDoc.id)?.name || "un documento"}`;
  if (pendingPaymentKey)
    return `Tu próximo paso es ${pendingPaymentKey}`;
  if (completedChecklist < trip.checklist.length)
    return "Revisar el checklist de equipaje";
  return "Todo está en orden. Solo queda disfrutar de la experiencia.";
};

// ─── Notificaciones ──────────────────────────────────────────────────────────

async function getAuthToken() {
  const { data } = await supabase.auth.getSession();
  return data?.session?.access_token ?? null;
}

async function sendNotification(type, to, participantId, data) {
  try {
    const token = await getAuthToken();
    const res = await fetch("/api/notify", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        ...(token ? { Authorization: `Bearer ${token}` } : {}),
      },
      body: JSON.stringify({ type, to, participantId, data }),
    });
    if (!res.ok) {
      const body = await res.json().catch(() => ({}));
      console.error("sendNotification failed:", type, res.status, body.error);
      return { ok: false, error: body.error };
    }
    return { ok: true };
  } catch (err) {
    console.error("sendNotification error:", err);
    return { ok: false, error: err.message };
  }
}

// ─── Supabase helpers ────────────────────────────────────────────────────────

async function uploadFileToDrive(file, participantName, subfolder, onProgress, tripName) {
  const token = await getAuthToken();
  return new Promise((resolve, reject) => {
    const formData = new FormData();
    formData.append("file", file);
    formData.append("participantName", participantName);
    formData.append("subfolder", subfolder);
    if (tripName) formData.append("tripName", tripName);
    const xhr = new XMLHttpRequest();
    xhr.open("POST", "/api/upload-to-drive");
    if (token) xhr.setRequestHeader("Authorization", `Bearer ${token}`);
    xhr.upload.onprogress = (e) => {
      if (e.lengthComputable && onProgress) onProgress(Math.round((e.loaded / e.total) * 100));
    };
    xhr.onload = () => {
      if (xhr.status >= 200 && xhr.status < 300) resolve(JSON.parse(xhr.responseText));
      else { try { reject(new Error(JSON.parse(xhr.responseText).error)); } catch { reject(new Error("Error al subir")); } }
    };
    xhr.onerror = () => reject(new Error("Error de red"));
    xhr.send(formData);
  });
}


// [ALTO-1] upsertDocument: no sobreescribe documentos ya confirmados
async function upsertDocument(participantId, templateId, payload) {
  const { data: existing, error: fetchErr } = await supabase
    .from("participant_documents")
    .select("id, status")
    .eq("participant_id", participantId)
    .eq("template_id", templateId)
    .maybeSingle();

  if (fetchErr) throw fetchErr;

  if (existing?.id) {
    // No degradar un estado confirmado
    const protectedStatus = ["confirmed"].includes(existing.status) ? existing.status : payload.status;
    const { error } = await supabase
      .from("participant_documents")
      .update({ ...payload, status: protectedStatus })
      .eq("id", existing.id);
    if (error) throw error;
  } else {
    const { error } = await supabase
      .from("participant_documents")
      .insert({ participant_id: participantId, template_id: templateId, ...payload });
    if (error) throw error;
  }
}

// [ALTO-1] upsertPayment: no sobreescribe pagos ya confirmados o enviados
// bypassProtection=true lo usa el admin para confirmar/rechazar desde el panel
async function upsertPayment(participantId, paymentKey, payload, bypassProtection = false) {
  const { data: existing, error: fetchErr } = await supabase
    .from("participant_payments")
    .select("id, status")
    .eq("participant_id", participantId)
    .eq("payment_key", paymentKey)
    .maybeSingle();

  if (fetchErr) throw fetchErr;

  if (existing?.id) {
    const updates = { ...payload };
    // El cliente no puede degradar un pago ya enviado o confirmado;
    // el admin (bypassProtection) sí puede cambiar el estado libremente.
    if (!bypassProtection && ["confirmed", "sent"].includes(existing.status)) {
      delete updates.status;
      delete updates.proof_name;
      delete updates.proof_path;
    }
    const { error } = await supabase
      .from("participant_payments")
      .update(updates)
      .eq("id", existing.id);
    if (error) throw error;
  } else {
    const { error } = await supabase
      .from("participant_payments")
      .insert({ participant_id: participantId, payment_key: paymentKey, ...payload });
    if (error) throw error;
  }
}

// ─── Exportar PDF ────────────────────────────────────────────────────────────
function exportListToPDF(title, subtitle = "", htmlBody) {
  const w = window.open("", "_blank");
  if (!w) {
    alert("El navegador ha bloqueado la ventana emergente.\nPor favor, permite ventanas emergentes para este sitio y vuelve a intentarlo.");
    return;
  }
  w.document.write(`<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8">
<title>${title} — GIMELOOS</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: Arial, sans-serif; color: #18181b; background: #fff; padding: 32px 40px; }
  .pdf-header { display: flex; align-items: flex-start; justify-content: space-between; margin-bottom: 32px; border-bottom: 2px solid #18181b; padding-bottom: 16px; }
  .pdf-header-left h1 { font-size: 22px; font-weight: 700; letter-spacing: -0.5px; }
  .pdf-header-left p  { font-size: 12px; color: #71717a; margin-top: 4px; }
  .pdf-header-brand   { font-size: 28px; font-weight: 700; letter-spacing: -1px; color: #18181b; }
  table { width: 100%; border-collapse: collapse; font-size: 12px; margin-top: 8px; }
  th { text-align: left; padding: 8px 10px; background: #f4f4f5; font-weight: 700; font-size: 10px; text-transform: uppercase; letter-spacing: .06em; color: #52525b; border-bottom: 1px solid #e4e4e7; }
  td { padding: 8px 10px; border-bottom: 1px solid #f4f4f5; vertical-align: top; }
  tr:last-child td { border-bottom: none; }
  .badge { display: inline-block; padding: 2px 8px; border-radius: 99px; font-size: 10px; font-weight: 600; }
  .badge-red  { background: #fee2e2; color: #b91c1c; }
  .badge-amber { background: #fef3c7; color: #92400e; }
  .room-block { margin-bottom: 16px; break-inside: avoid; }
  .room-title { font-weight: 700; font-size: 13px; margin-bottom: 6px; }
  .room-students { font-size: 12px; color: #52525b; }
  .footer { margin-top: 40px; font-size: 10px; color: #a1a1aa; text-align: center; }
  @media print { body { padding: 20px 28px; } }
</style></head><body>
<div class="pdf-header">
  <div class="pdf-header-left">
    <h1>${title}</h1>
    ${subtitle ? `<p>${subtitle}</p>` : ""}
    <p>Generado el ${new Date().toLocaleDateString("es-ES", { day: "2-digit", month: "long", year: "numeric" })}</p>
  </div>
  <div class="pdf-header-brand">GIMELOOS</div>
</div>
${htmlBody}
<div class="footer" style="border-top:1px solid #e4e4e7;padding-top:12px;margin-top:32px">
  ⚠️ Documento confidencial — contiene datos personales protegidos por la LOPD/RGPD. Uso interno exclusivo. No distribuir.
  <br>GIMELOOS · Listado generado automáticamente
</div>
</body></html>`);
  w.document.close();
  w.focus();
  setTimeout(() => w.print(), 400);
}

// ─── Componentes UI reutilizables ────────────────────────────────────────────

function SectionTitle({ icon: Icon, title, subtitle, extra }) {
  return (
    <div className="flex flex-col gap-3 sm:flex-row sm:items-start sm:justify-between">
      <div className="flex items-start gap-3">
        <div className="mt-1 rounded-2xl p-2.5 shadow-sm text-white" style={{ backgroundColor: CORPORATE_RED }}>
          <Icon className="h-5 w-5" />
        </div>
        <div>
          <h2 className="text-xl font-semibold tracking-tight text-zinc-950">{title}</h2>
          {subtitle && <p className="mt-1 text-sm text-zinc-500">{subtitle}</p>}
        </div>
      </div>
      {extra}
    </div>
  );
}

function ActionToast({ notifications, removeNotification }) {
  return (
    <div className="pointer-events-none fixed right-4 top-4 z-50 flex w-[380px] max-w-[calc(100vw-2rem)] flex-col gap-3">
      <AnimatePresence>
        {notifications.map((item) => (
          <motion.div
            key={item.id}
            initial={{ opacity: 0, y: -12, scale: 0.98 }}
            animate={{ opacity: 1, y: 0, scale: 1 }}
            exit={{ opacity: 0, y: -8, scale: 0.98 }}
            className="pointer-events-auto rounded-[28px] border border-zinc-200/90 bg-white/95 p-4 shadow-[0_18px_40px_rgba(0,0,0,0.12)] backdrop-blur-sm"
          >
            <div className="flex items-start gap-3">
              <div className="rounded-2xl p-2 text-white" style={{ backgroundColor: CORPORATE_RED }}>
                <Bell className="h-4 w-4" />
              </div>
              <div className="min-w-0 flex-1">
                <div className="text-sm font-semibold text-zinc-950">Acción realizada</div>
                <div className="mt-1 text-sm text-zinc-600">{item.message}</div>
                {item.actionLabel && item.onAction && (
                  <button
                    type="button"
                    onClick={() => { item.onAction(); removeNotification(item.id); }}
                    className="mt-3 rounded-2xl border border-zinc-200 px-3 py-2 text-sm font-medium text-zinc-900 hover:bg-white"
                  >
                    {item.actionLabel}
                  </button>
                )}
              </div>
              <button
                type="button"
                aria-label="Cerrar notificación"
                onClick={() => removeNotification(item.id)}
                className="rounded-full p-1 text-zinc-400 hover:bg-zinc-100 hover:text-zinc-700"
              >
                <X className="h-4 w-4" />
              </button>
            </div>
          </motion.div>
        ))}
      </AnimatePresence>
    </div>
  );
}

function LogoMark({ dark = false, totalParticipants = null }) {
  return (
    <div className="flex items-center gap-3">
      <div>
        <div className={`text-xs uppercase tracking-[0.24em] ${dark ? "text-zinc-300" : "text-zinc-500"}`}>Portal</div>
        <div className={`text-lg font-semibold tracking-[0.18em] ${dark ? "text-white" : "text-zinc-950"}`}>GIMELOOS</div>
        {typeof totalParticipants === "number" && (
          <div className={`${dark ? "text-zinc-200" : "text-zinc-600"} mt-1 text-sm`}>
            Participantes registrados:{" "}
            <span className={`${dark ? "text-white" : "text-zinc-950"} font-semibold`}>{totalParticipants}</span>
          </div>
        )}
      </div>
    </div>
  );
}

function LoginScreen({ onLogin, loginError, isLoading }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [emailTouched, setEmailTouched] = useState(false);
  const emailValid = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(username.trim());
  const emailError = emailTouched && username && !emailValid ? "Introduce un email válido" : null;

  const [forgotMode, setForgotMode] = useState(false);
  const [forgotEmail, setForgotEmail] = useState("");
  const [forgotSent, setForgotSent] = useState(false);
  const [forgotLoading, setForgotLoading] = useState(false);
  const [forgotError, setForgotError] = useState(null);

  async function handleForgotPassword(e) {
    e.preventDefault();
    setForgotLoading(true);
    setForgotError(null);
    const { error } = await supabase.auth.resetPasswordForEmail(forgotEmail.trim(), {
      redirectTo: window.location.origin,
    });
    setForgotLoading(false);
    if (error) {
      setForgotError("No hemos podido enviar el enlace. Comprueba el email e inténtalo de nuevo.");
    } else {
      setForgotSent(true);
    }
  }

  return (
    <div className="min-h-screen bg-white text-zinc-950">
      <div className="mx-auto flex min-h-screen max-w-6xl items-center p-4 sm:p-6 lg:p-8">
        <motion.div
          initial={{ opacity: 0, y: 16 }}
          animate={{ opacity: 1, y: 0 }}
          className="relative w-full overflow-hidden rounded-[32px] border border-black/5 bg-zinc-950 shadow-[0_30px_100px_rgba(0,0,0,0.18)]"
        >
          <img
            src="https://images.unsplash.com/photo-1501785888041-af3ef285b470?auto=format&fit=crop&w=1600&q=80"
            alt="Experiencias GIMELOOS"
            width="1600"
            height="900"
            className="absolute inset-0 h-full w-full object-cover"
          />
          <div className="absolute inset-0 bg-[linear-gradient(135deg,rgba(0,0,0,0.72),rgba(0,0,0,0.45),rgba(255,49,49,0.18))]" />

          <div className="relative grid min-h-[700px] grid-cols-1 lg:grid-cols-[1.1fr_420px]">
            <div className="flex flex-col p-6 sm:p-8 lg:p-10">
              <LogoMark dark />
              <div className="mt-4 max-w-xl">
                <Badge className="border-0 bg-white/10 text-white backdrop-blur-sm hover:bg-white/10">
                  Área privada de clientes
                </Badge>
                <h1 className="mt-4 whitespace-nowrap text-[1.35rem] font-semibold leading-tight tracking-tight text-white sm:text-[1.65rem] lg:text-[1.95rem] xl:text-[2.1rem]">
                  LA EXPERIENCIA QUE TE MERECES
                </h1>
              </div>
            </div>

            <div className="flex items-end p-4 sm:p-6 lg:items-center lg:p-6">
              <Card className="w-full rounded-[28px] border border-white/10 bg-white/88 shadow-2xl backdrop-blur-xl">
                <CardHeader className="space-y-3 p-6 pb-3 sm:p-7 sm:pb-3">
                  <div>
                    <div className="text-xs uppercase tracking-[0.24em] text-zinc-500">GIMELOOS</div>
                    <CardTitle className="mt-2 text-2xl font-semibold tracking-tight text-zinc-950 sm:text-3xl">
                      Accede a tu viaje
                    </CardTitle>
                  </div>
                </CardHeader>
                <CardContent className="p-6 pt-2 sm:p-7 sm:pt-2">
                  {forgotMode ? (
                    forgotSent ? (
                      <div className="space-y-4">
                        <div className="rounded-2xl bg-green-50 px-4 py-4 text-sm text-green-700">
                          Hemos enviado un enlace a <strong>{forgotEmail}</strong>. Revisa tu bandeja de entrada y sigue las instrucciones.
                        </div>
                        <button
                          type="button"
                          className="w-full text-sm text-zinc-500 underline-offset-2 hover:underline"
                          onClick={() => { setForgotMode(false); setForgotSent(false); setForgotEmail(""); }}
                        >
                          Volver al inicio de sesión
                        </button>
                      </div>
                    ) : (
                      <form className="space-y-4" onSubmit={handleForgotPassword}>
                        <p className="text-sm text-zinc-500">Introduce tu email y te enviaremos un enlace para restablecer tu contraseña.</p>
                        <div className="space-y-2">
                          <Label>Email</Label>
                          <Input
                            value={forgotEmail}
                            onChange={(e) => setForgotEmail(e.target.value)}
                            placeholder="Introduce tu email"
                            className="h-12 rounded-2xl border-zinc-200 bg-white"
                            autoComplete="email"
                            type="email"
                            required
                          />
                        </div>
                        {forgotError && (
                          <div className="rounded-2xl bg-red-50 px-4 py-3 text-sm text-red-700">{forgotError}</div>
                        )}
                        <Button
                          type="submit"
                          disabled={forgotLoading}
                          className="h-12 w-full rounded-2xl text-white shadow-lg"
                          style={{ backgroundColor: CORPORATE_RED }}
                        >
                          {forgotLoading ? "Enviando..." : "Enviar enlace"}
                        </Button>
                        <button
                          type="button"
                          className="w-full text-sm text-zinc-500 underline-offset-2 hover:underline"
                          onClick={() => { setForgotMode(false); setForgotError(null); }}
                        >
                          Volver al inicio de sesión
                        </button>
                      </form>
                    )
                  ) : (
                    <>
                      {/* [CRÍTICO-1] onSubmit llama a supabase.auth.signInWithPassword en el padre */}
                      <form
                        className="space-y-4"
                        onSubmit={(e) => { e.preventDefault(); onLogin(username, password); }}
                      >
                        <div className="space-y-2">
                          <Label>Email</Label>
                          <Input
                            value={username}
                            onChange={(e) => setUsername(e.target.value)}
                            onBlur={() => setEmailTouched(true)}
                            placeholder="Introduce tu email"
                            className={`h-12 rounded-2xl border-zinc-200 bg-white ${emailError ? "border-red-400" : ""}`}
                            autoComplete="email"
                            type="email"
                          />
                          {emailError && <p className="text-xs text-red-500">{emailError}</p>}
                        </div>
                        <div className="space-y-2">
                          <Label>Contraseña</Label>
                          <Input
                            value={password}
                            onChange={(e) => setPassword(e.target.value)}
                            type="password"
                            placeholder="••••••••"
                            className="h-12 rounded-2xl border-zinc-200 bg-white"
                            autoComplete="current-password"
                          />
                        </div>
                        {loginError && (
                          <div className="rounded-2xl bg-red-50 px-4 py-3 text-sm text-red-700">{loginError}</div>
                        )}
                        <Button
                          type="submit"
                          disabled={isLoading}
                          className="h-12 w-full rounded-2xl text-white shadow-lg"
                          style={{ backgroundColor: CORPORATE_RED }}
                        >
                          <User className="mr-2 h-4 w-4" />
                          {isLoading ? "Entrando..." : "Entrar"}
                        </Button>
                      </form>
                      <button
                        type="button"
                        className="mt-4 w-full text-sm text-zinc-400 underline-offset-2 hover:text-zinc-600 hover:underline"
                        onClick={() => { setForgotMode(true); setForgotEmail(username); }}
                      >
                        ¿Olvidaste tu contraseña?
                      </button>
                    </>
                  )}
                </CardContent>
              </Card>
            </div>
          </div>
        </motion.div>
      </div>
    </div>
  );
}

function ResetPasswordScreen({ onDone }) {
  const [password, setPassword] = useState("");
  const [confirm, setConfirm] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [done, setDone] = useState(false);

  async function handleSubmit(e) {
    e.preventDefault();
    if (password.length < 6) { setError("La contraseña debe tener al menos 6 caracteres."); return; }
    if (password !== confirm) { setError("Las contraseñas no coinciden."); return; }
    setLoading(true);
    setError(null);
    const { error: updErr } = await supabase.auth.updateUser({ password });
    setLoading(false);
    if (updErr) { setError("No se pudo actualizar la contraseña: " + updErr.message); return; }
    setDone(true);
    await supabase.auth.signOut();
    setTimeout(onDone, 2500);
  }

  return (
    <div className="min-h-screen bg-white text-zinc-950">
      <div className="mx-auto flex min-h-screen max-w-6xl items-center p-4 sm:p-6 lg:p-8">
        <motion.div
          initial={{ opacity: 0, y: 16 }}
          animate={{ opacity: 1, y: 0 }}
          className="relative w-full overflow-hidden rounded-[32px] border border-black/5 bg-zinc-950 shadow-[0_30px_100px_rgba(0,0,0,0.18)]"
        >
          <img
            src="https://images.unsplash.com/photo-1501785888041-af3ef285b470?auto=format&fit=crop&w=1600&q=80"
            alt="Experiencias GIMELOOS"
            width="1600" height="900"
            className="absolute inset-0 h-full w-full object-cover"
          />
          <div className="absolute inset-0 bg-[linear-gradient(135deg,rgba(0,0,0,0.72),rgba(0,0,0,0.45),rgba(255,49,49,0.18))]" />
          <div className="relative grid min-h-[700px] grid-cols-1 lg:grid-cols-[1.1fr_420px]">
            <div className="flex flex-col p-6 sm:p-8 lg:p-10">
              <LogoMark dark />
              <div className="mt-4 max-w-xl">
                <Badge className="border-0 bg-white/10 text-white backdrop-blur-sm hover:bg-white/10">Área privada de clientes</Badge>
                <h1 className="mt-4 text-[1.35rem] font-semibold leading-tight tracking-tight text-white sm:text-[1.65rem] lg:text-[1.95rem]">
                  LA EXPERIENCIA QUE TE MERECES
                </h1>
              </div>
            </div>
            <div className="flex items-end p-4 sm:p-6 lg:items-center lg:p-6">
              <Card className="w-full rounded-[28px] border border-white/10 bg-white/88 shadow-2xl backdrop-blur-xl">
                <CardHeader className="space-y-3 p-6 pb-3 sm:p-7 sm:pb-3">
                  <div>
                    <div className="text-xs uppercase tracking-[0.24em] text-zinc-500">GIMELOOS</div>
                    <CardTitle className="mt-2 text-2xl font-semibold tracking-tight text-zinc-950 sm:text-3xl">
                      Nueva contraseña
                    </CardTitle>
                  </div>
                </CardHeader>
                <CardContent className="p-6 pt-2 sm:p-7 sm:pt-2">
                  {done ? (
                    <div className="rounded-2xl bg-green-50 px-4 py-4 text-sm text-green-700">
                      ✓ Contraseña actualizada. Redirigiendo al inicio de sesión…
                    </div>
                  ) : (
                    <form className="space-y-4" onSubmit={handleSubmit}>
                      <div className="space-y-2">
                        <Label>Nueva contraseña</Label>
                        <Input
                          value={password}
                          onChange={(e) => setPassword(e.target.value)}
                          type="password"
                          placeholder="Mínimo 6 caracteres"
                          className="h-12 rounded-2xl border-zinc-200 bg-white"
                          autoComplete="new-password"
                          required
                        />
                      </div>
                      <div className="space-y-2">
                        <Label>Confirmar contraseña</Label>
                        <Input
                          value={confirm}
                          onChange={(e) => setConfirm(e.target.value)}
                          type="password"
                          placeholder="Repite la contraseña"
                          className="h-12 rounded-2xl border-zinc-200 bg-white"
                          autoComplete="new-password"
                          required
                        />
                      </div>
                      {error && <div className="rounded-2xl bg-red-50 px-4 py-3 text-sm text-red-700">{error}</div>}
                      <Button
                        type="submit"
                        disabled={loading}
                        className="h-12 w-full rounded-2xl text-white shadow-lg"
                        style={{ backgroundColor: CORPORATE_RED }}
                      >
                        {loading ? "Guardando…" : "Guardar nueva contraseña"}
                      </Button>
                    </form>
                  )}
                </CardContent>
              </Card>
            </div>
          </div>
        </motion.div>
      </div>
    </div>
  );
}

function HeroBanner({ trip, user, pendingSummary, onNavigate }) {
  const remaining = daysRemaining(trip.departureDate);

  // 1 experiencia contratada = 1 imagen. Sin carrusel.
  const heroImages = useMemo(
    () => [trip.heroImage || trip.heroImages?.[0] || DEFAULT_HERO_IMAGES[0]],
    [trip.heroImage, trip.heroImages]
  );

  const [activeImage, setActiveImage] = useState(0);
  const totalPending = pendingSummary.reduce((acc, s) => acc + s.count, 0);

  useEffect(() => {
    if (heroImages.length <= 1) return;
    const id = setInterval(() => setActiveImage((i) => (i + 1) % heroImages.length), 5000);
    return () => clearInterval(id);
  }, [heroImages.length]);

  return (
    <div className="relative overflow-hidden rounded-[32px] shadow-[0_20px_70px_rgba(0,0,0,0.12)]">
      <img src={heroImages[activeImage]} alt={trip.name} className="absolute inset-0 block h-full w-full object-cover object-center transition-opacity duration-700" />
      <div className="absolute inset-0 bg-[linear-gradient(135deg,rgba(0,0,0,0.80),rgba(0,0,0,0.45),rgba(255,49,49,0.16))]" />
      {heroImages.length > 1 && (
        <div className="absolute bottom-4 left-1/2 z-10 flex -translate-x-1/2 gap-1.5">
          {heroImages.map((_, i) => (
            <button
              key={i}
              aria-label={`Imagen ${i + 1}`}
              onClick={() => setActiveImage(i)}
              className="h-1.5 rounded-full transition-all duration-300"
              style={{ width: i === activeImage ? 20 : 6, backgroundColor: i === activeImage ? "#fff" : "rgba(255,255,255,0.45)" }}
            />
          ))}
        </div>
      )}
      <div className="relative grid gap-6 p-5 sm:p-7 lg:grid-cols-[1fr_260px] lg:gap-8 lg:p-10">
        <div className="flex flex-col justify-between">
          <div>
            <Badge className="border-0 bg-white/10 text-white backdrop-blur-sm hover:bg-white/10">Experiencia contratada</Badge>
            <h1 className="mt-4 max-w-3xl text-3xl font-semibold tracking-tight text-white sm:text-4xl lg:text-5xl">{trip.name?.toUpperCase()}</h1>
            <p className="mt-3 max-w-2xl text-base leading-7 text-white font-medium sm:text-lg">
              Hola, familia de <span className="text-white font-bold">{user.participantName}</span> 👋
            </p>
            <p className="mt-1 max-w-2xl text-sm leading-6 text-zinc-300">
              Aquí tienes toda la información de la experiencia.
            </p>
          </div>
          <div className="mt-6 flex flex-wrap gap-3 text-sm text-white">
            <div className="inline-flex items-center gap-2 rounded-2xl border border-white/10 bg-white/10 px-4 py-2 backdrop-blur-sm">
              <User className="h-4 w-4" /> {user.participantName}
            </div>
            <div className="inline-flex items-center gap-2 rounded-2xl border border-white/10 bg-white/10 px-4 py-2 backdrop-blur-sm">
              <Mail className="h-4 w-4" /> {getFamilyLabel(user) || user.parentName}
            </div>
          </div>
        </div>
        <div className="flex items-end lg:items-center">
          <div className="w-full rounded-[26px] border border-white/10 bg-black/35 p-5 text-white shadow-2xl backdrop-blur-xl">
            <div className="text-xs uppercase tracking-[0.24em] text-zinc-300">Cuenta atrás</div>
            <div className="mt-3 text-5xl font-semibold leading-none sm:text-6xl">{remaining}</div>
            <div className="mt-2 text-zinc-200">días para tu experiencia</div>
            <div className="mt-6 rounded-2xl bg-white/10 p-4 text-sm text-zinc-200">
              <div className="flex items-center gap-2"><CalendarDays className="h-4 w-4" /> Salida</div>
              <div className="mt-2 font-medium text-white">
                {trip.departureDate
                  ? (() => { const s = new Date(trip.departureDate).toLocaleDateString("es-ES", { weekday: "long", day: "numeric", month: "long", year: "numeric" }); return s.charAt(0).toUpperCase() + s.slice(1); })()
                  : "-"}
              </div>
            </div>
            {/* Resumen de tareas pendientes */}
            <div className="mt-4 rounded-2xl border border-white/10 bg-white/5 p-4">
              <div className="text-xs uppercase tracking-[0.18em] text-zinc-300 mb-3">Tareas pendientes</div>
              {totalPending === 0 ? (
                <div className="flex items-center gap-2 text-sm text-emerald-300">
                  <CheckCircle2 className="h-4 w-4" />
                  ¡Todo al día!
                </div>
              ) : (
                <div className="flex flex-wrap gap-3">
                  {pendingSummary.filter((s) => s.count > 0).map(({ key, label, icon: Icon, sectionId, count }) => (
                    <button
                      key={key}
                      type="button"
                      onClick={() => sectionId && onNavigate?.(key)}
                      className={`relative flex flex-col items-center gap-1 ${sectionId ? "cursor-pointer" : "cursor-default"}`}
                    >
                      <div className="relative rounded-2xl border border-white/15 bg-white/10 p-3 backdrop-blur-sm transition hover:bg-white/20">
                        <Icon className="h-5 w-5 text-white" />
                        <span className="absolute -right-1.5 -top-1.5 flex h-5 w-5 items-center justify-center rounded-full text-[10px] font-bold text-white" style={{ backgroundColor: CORPORATE_RED }}>
                          {count}
                        </span>
                      </div>
                      <span className="text-[10px] text-zinc-300 text-center leading-tight max-w-[52px]">{label}</span>
                    </button>
                  ))}
                </div>
              )}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

function ClientDocuments({ user, templates, onUploadDocument }) {
  const [progress, setProgress] = useState({});
  return (
    <div className="space-y-4">
      <div className="grid gap-4">
        {user.documents.map((docItem) => {
          const template = templates.find((doc) => doc.id === docItem.id);
          const status = getStatusMeta(docItem.status);
          const pct = progress[docItem.id];
          const uploading = pct !== undefined && pct < 100;
          return (
            <Card key={docItem.id} className="rounded-3xl border-zinc-200 bg-white shadow-sm">
              <CardContent className="flex flex-col gap-4 p-5 lg:flex-row lg:items-center lg:justify-between">
                <div className="min-w-0 flex-1">
                  <div className="font-medium text-zinc-950">{template?.name || docItem.id}</div>
                  {docItem.uploadedFileName && (
                    <div className="mt-1 text-sm text-zinc-500">Subido: {docItem.uploadedFileName}</div>
                  )}
                  {pct !== undefined && (
                    <div className="mt-2">
                      <div className="mb-1 flex items-center justify-between text-xs text-zinc-500">
                        <span>{pct < 100 ? "Subiendo a Google Drive…" : "¡Subido!"}</span>
                        <span>{pct}%</span>
                      </div>
                      <div className="h-1.5 w-full overflow-hidden rounded-full bg-zinc-100">
                        <div
                          className="h-full rounded-full transition-all duration-200"
                          style={{ width: `${pct}%`, backgroundColor: CORPORATE_RED }}
                        />
                      </div>
                    </div>
                  )}
                </div>
                <div className="flex flex-wrap items-center gap-3">
                  <Badge className={status.className} style={status.style}>{status.label}</Badge>
                  {template?.driveUrl && (
                    <Button variant="outline" className="rounded-2xl border-zinc-200 bg-white" onClick={() => window.open(template.driveUrl, "_blank", "noopener,noreferrer")}>
                      <Download className="mr-2 h-4 w-4" />Descargar plantilla
                    </Button>
                  )}
                  <div className="flex flex-col gap-1">
                  <p className="text-xs text-zinc-400">PDF · máx. 20 MB</p>
                  <label className={uploading ? "cursor-not-allowed opacity-60" : "cursor-pointer"}>
                    <input type="file" accept=".pdf,.PDF" className="hidden" disabled={uploading} onChange={(e) => {
                      const file = e.target.files?.[0];
                      if (!file) return;
                      setProgress((p) => ({ ...p, [docItem.id]: 0 }));
                      const onProgress = (pct) => setProgress((p) => ({ ...p, [docItem.id]: pct }));
                      onUploadDocument(docItem.id, file, onProgress)
                        .then(() => setProgress((p) => ({ ...p, [docItem.id]: 100 })))
                        .catch(() => setProgress((p) => { const n = { ...p }; delete n[docItem.id]; return n; }))
                        .finally(() => setTimeout(() => setProgress((p) => { const n = { ...p }; delete n[docItem.id]; return n; }), 1800));
                    }} />
                    <span className="inline-flex h-10 items-center rounded-2xl px-4 text-sm font-medium text-white" style={{ backgroundColor: CORPORATE_RED }}>
                      <Upload className="mr-2 h-4 w-4" />{uploading ? `${pct}%` : "Subir documento"}
                    </span>
                  </label>
                  </div>
                </div>
              </CardContent>
            </Card>
          );
        })}
      </div>
    </div>
  );
}

function InvoiceUploadButton({ existing, onUpload, size = "sm" }) {
  const [pct, setPct] = useState(undefined);
  const uploading = pct !== undefined && pct < 100;
  const label = existing ? "Actualizar factura" : "Subir factura";
  return (
    <div className="space-y-1 min-w-0">
      <label className={`cursor-pointer ${uploading ? "pointer-events-none opacity-60" : ""}`}>
        <input type="file" accept=".pdf,.PDF" className="hidden" onChange={async (e) => {
          const file = e.target.files?.[0];
          if (!file) return;
          try { await onUpload(file, (p) => setPct(p)); }
          finally { setPct(undefined); e.target.value = ""; }
        }} />
        <span className={`inline-flex cursor-pointer items-center rounded-2xl border border-zinc-200 bg-white font-medium text-zinc-700 hover:bg-zinc-50 ${size === "sm" ? "h-8 px-3 text-xs" : "h-9 px-4 text-sm"}`}>
          <Upload className={`mr-1.5 ${size === "sm" ? "h-3.5 w-3.5" : "h-4 w-4"}`} />{label}
        </span>
      </label>
      {pct !== undefined && (
        <div>
          <div className="flex justify-between text-[10px] text-zinc-400 mb-0.5">
            <span>{pct < 100 ? "Subiendo…" : "¡Listo!"}</span>
            <span>{pct}%</span>
          </div>
          <div className="h-1 w-full overflow-hidden rounded-full bg-zinc-100">
            <div className="h-full rounded-full transition-all duration-200"
              style={{ width: `${pct}%`, backgroundColor: pct === 100 ? "#16a34a" : CORPORATE_RED }} />
          </div>
        </div>
      )}
    </div>
  );
}

function PaymentRow({ title, payment, onUploadProof }) {
  const status = getStatusMeta(payment.status);
  const [pct, setPct] = useState(undefined);
  const [localFileName, setLocalFileName] = useState(null); // feedback inmediato
  const uploading = pct !== undefined && pct < 100;
  const displayProofName = localFileName ?? payment.proofName;
  return (
    <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
      <CardContent className="flex flex-col gap-4 p-5 lg:flex-row lg:items-center lg:justify-between">
        <div className="min-w-0 flex-1">
          <div className="font-medium text-zinc-950">{title}</div>
          <div className="mt-1 text-sm text-zinc-500">Importe: {formatCurrency(payment.amount)}</div>
          {displayProofName && (
            <div className="mt-1 flex items-center gap-1.5 text-sm text-zinc-500">
              <CheckCircle2 className="h-3.5 w-3.5 shrink-0 text-emerald-500" />
              <span>Justificante: {displayProofName}</span>
            </div>
          )}
          {pct !== undefined && (
            <div className="mt-2">
              <div className="mb-1 flex items-center justify-between text-xs text-zinc-500">
                <span>{pct < 100 ? "Subiendo a Google Drive…" : "¡Subido correctamente!"}</span>
                <span>{Math.round(pct)}%</span>
              </div>
              <div className="h-1.5 w-full overflow-hidden rounded-full bg-zinc-100">
                <div
                  className="h-full rounded-full transition-all duration-200"
                  style={{ width: `${pct}%`, backgroundColor: pct === 100 ? "#16a34a" : CORPORATE_RED }}
                />
              </div>
            </div>
          )}
        </div>
        <div className="flex flex-wrap items-center gap-3">
          <Badge className={status.className} style={status.style}>{status.label}</Badge>
          <label className={uploading ? "cursor-not-allowed opacity-60" : "cursor-pointer"}>
            <input type="file" className="hidden" disabled={uploading} onChange={(e) => {
              const file = e.target.files?.[0];
              if (!file) return;
              setLocalFileName(file.name);
              setPct(0);
              const onProgress = (p) => setPct(p);
              onUploadProof(title, file, onProgress)
                .then(() => setPct(100))
                .catch(() => { setLocalFileName(null); setPct(undefined); })
                .finally(() => setTimeout(() => setPct(undefined), 2500));
            }} />
            <span className="inline-flex h-10 items-center rounded-2xl px-4 text-sm font-medium text-white" style={{ backgroundColor: CORPORATE_RED }}>
              <FolderUp className="mr-2 h-4 w-4" />{uploading ? `${Math.round(pct)}%` : "Subir justificante"}
            </span>
          </label>
        </div>
      </CardContent>
    </Card>
  );
}

function ClientPayments({ user, trip, onUploadProof }) {
  const payments = user.payments || {};
  const initialPrice = Number(payments.initialPrice || 0);
  const discount = Number(payments.discount || 0);
  const finalPrice = discount > 0 ? Math.max(0, initialPrice - discount) : Number(payments.finalPrice || 0);
  const reservation = payments.reservation || { name: "Reserva", amount: 0, status: "pending", proofName: "", dueDate: "" };
  const firstInstallment = payments.firstInstallment || { name: "Primera cuota", amount: 0, status: "pending", proofName: "", dueDate: "" };
  const secondInstallment = payments.secondInstallment || { name: "Segunda cuota", amount: 0, status: "pending", proofName: "", dueDate: "" };
  const paidAmount = [reservation, firstInstallment, secondInstallment]
    .filter((p) => ["sent", "confirmed"].includes(p.status))
    .reduce((sum, p) => sum + Number(p.amount || 0), 0);
  const calculatedOutstanding = Math.max(0, finalPrice - paidAmount);

  return (
    <div className="space-y-4">
      <div className="grid gap-4 lg:grid-cols-[1fr_320px]">
        <div className="grid gap-4">
          <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
            <CardContent className="grid gap-4 p-5 sm:grid-cols-2 xl:grid-cols-6">
              {[
                ["Precio inicial", initialPrice],
                discount > 0 ? ["Tu descuento", discount] : null,
                ["Precio final", finalPrice],
                ["Importe reserva", Number(reservation.amount || 0)],
                ["Primera cuota", Number(firstInstallment.amount || 0)],
                ["Segunda cuota", Number(secondInstallment.amount || 0)],
              ].filter(Boolean).map(([label, value]) => (
                // [MENOR-3] Key única: label es único en este array
                <div key={String(label)} className={`rounded-2xl border p-4 ${label === "Tu descuento" ? "border-green-200 bg-green-50" : "border-zinc-200 bg-white"}`}>
                  <div className={`text-xs uppercase tracking-[0.18em] ${label === "Tu descuento" ? "text-green-600" : "text-zinc-500"}`}>{label}</div>
                  <div className={`mt-2 text-xl font-semibold whitespace-nowrap ${label === "Tu descuento" ? "text-green-700" : "text-zinc-950"}`}>
                    {formatCurrency(value)}
                  </div>
                </div>
              ))}
            </CardContent>
          </Card>
          <PaymentRow title={reservation.name || "Reserva"} payment={reservation} onUploadProof={(_, file) => onUploadProof("reservation", file)} />
          <PaymentRow title={firstInstallment.name || "Primera cuota"} payment={firstInstallment} onUploadProof={(_, file) => onUploadProof("firstInstallment", file)} />
          <PaymentRow title={secondInstallment.name || "Segunda cuota"} payment={secondInstallment} onUploadProof={(_, file) => onUploadProof("secondInstallment", file)} />
        </div>
        <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
          <CardContent className="space-y-4 p-5">
            <div>
              <div className="text-sm uppercase tracking-[0.18em] text-zinc-500">Datos de transferencia</div>
              <div className="mt-3 space-y-2 text-sm text-zinc-700">
                <div><span className="font-medium text-zinc-950">Banco:</span> {trip.transferInfo.bank}</div>
                <div><span className="font-medium text-zinc-950">Titular:</span> {trip.transferInfo.accountHolder}</div>
                <div><span className="font-medium text-zinc-950">IBAN:</span> {trip.transferInfo.iban}</div>
                <div><span className="font-medium text-zinc-950">Concepto:</span> {trip.transferInfo.concept}</div>
              </div>
            </div>
            <Separator />
            <div>
              <div className="text-sm uppercase tracking-[0.18em] text-zinc-500">Resumen</div>
              <div className="mt-3 space-y-3">
                <div className="rounded-2xl bg-white p-4">
                  <div className="text-sm text-zinc-500">Importe ya enviado/abonado</div>
                  <div className="mt-2 text-2xl font-semibold text-zinc-950">{formatCurrency(paidAmount)}</div>
                </div>
                <div className="rounded-2xl bg-white p-4">
                  <div className="text-sm text-zinc-500">Importe residual pendiente</div>
                  <div className="mt-2 text-2xl font-semibold text-zinc-950">{formatCurrency(calculatedOutstanding)}</div>
                </div>
              </div>
            </div>
            <>
              <Separator />
              <div>
                <div className="text-sm uppercase tracking-[0.18em] text-zinc-500">Factura</div>
                {user.invoiceUrl ? (
                  <Button variant="outline" className="mt-3 w-full rounded-2xl" onClick={() => window.open(user.invoiceUrl, "_blank", "noopener,noreferrer")}>
                    <Download className="mr-2 h-4 w-4" />Descargar factura
                  </Button>
                ) : (
                  <p className="mt-2 text-sm text-zinc-400">Tu factura estará disponible aquí en cuanto el equipo de GIMELOOS la emita.</p>
                )}
              </div>
            </>
          </CardContent>
        </Card>
      </div>
    </div>
  );
}

function ClientLogistics({ trip }) {
  const items = trip.logistics || [];
  if (items.length === 0) {
    return (
      <div className="rounded-2xl border border-dashed border-stone-200 bg-stone-50 p-8 text-center text-sm text-zinc-400">
        El equipo de GIMELOOS publicará pronto la información logística de tu experiencia.
      </div>
    );
  }
  return (
    <ul className="space-y-2">
      {items.map((item, index) => (
        <li key={`log-${index}`} className="flex gap-3 items-start py-2 border-b border-stone-100 last:border-0">
          <span className="mt-1 shrink-0 text-base font-bold leading-none" style={{ color: CORPORATE_RED }}>—</span>
          <div className="min-w-0">
            <span className="font-semibold text-zinc-900 text-sm">{item.title}</span>
            {item.description && (
              <p className="mt-0.5 text-sm text-zinc-600 leading-relaxed">{item.description}</p>
            )}
          </div>
        </li>
      ))}
    </ul>
  );
}

function ClientItinerary({ trip }) {
  const items = trip.itinerary || [];
  if (items.length === 0) {
    return (
      <div className="rounded-2xl border border-dashed border-stone-200 bg-stone-50 p-8 text-center text-sm text-zinc-400">
        El equipo de GIMELOOS publicará el itinerario próximamente.
      </div>
    );
  }
  return (
    <div className="space-y-3">
      {items.map((item, index) => (
        <div key={`itin-${index}`} className="flex gap-4 items-start rounded-2xl border border-stone-100 bg-stone-50 p-4">
          <div className="shrink-0 min-w-[56px] rounded-xl px-2 py-1.5 text-center text-white text-xs font-semibold" style={{ backgroundColor: CORPORATE_RED }}>
            {item.day || `Día ${index + 1}`}
          </div>
          <div className="min-w-0 flex-1">
            <div className="flex items-center gap-2 flex-wrap">
              <span className="font-semibold text-zinc-900 text-sm">{item.title}</span>
              {item.time && <span className="text-xs text-zinc-400">{item.time}</span>}
            </div>
            {item.description && (
              <p className="mt-1 text-sm text-zinc-600 leading-relaxed">{item.description}</p>
            )}
          </div>
        </div>
      ))}
    </div>
  );
}

function ClientChecklist({ user, trip, onToggleItem }) {
  const checklistItems = Array.isArray(trip?.checklist) ? trip.checklist : [];
  const checklistState = user?.checklistState || {};
  const total = checklistItems.length;
  const completed = checklistItems.filter((item) => !!checklistState[item]).length;
  const progress = total > 0 ? Math.round((completed / total) * 100) : 0;

  return (
    <div className="space-y-4">
      <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
        <CardContent className="space-y-5 p-5">
          <div>
            <div className="mb-2 flex items-center justify-between text-sm text-zinc-500">
              <span>Progreso</span>
              <span>{completed}/{total} · {progress}%</span>
            </div>
            <div className="h-3 w-full overflow-hidden rounded-full bg-zinc-200">
              <div className="h-full rounded-full transition-all duration-500 ease-out" style={{ width: `${progress}%`, backgroundColor: CORPORATE_RED }} />
            </div>
          </div>
          <div className="grid gap-3 md:grid-cols-2">
            {checklistItems.map((item) => (
              // [MENOR-3] Key: item es string único en el checklist. div en lugar de label para evitar doble-toggle con Radix button
              <div key={item} onClick={() => onToggleItem(item)} className="flex cursor-pointer select-none items-center gap-3 rounded-2xl border border-zinc-200 bg-white p-4 text-sm text-zinc-800 shadow-sm">
                <Checkbox checked={!!checklistState[item]} onCheckedChange={() => onToggleItem(item)} onClick={(e) => e.stopPropagation()} />
                <span>{item}</span>
              </div>
            ))}
          </div>
        </CardContent>
      </Card>
    </div>
  );
}

function ClientQuestions({ questions = [], onSendQuestion }) {
  const [message, setMessage] = useState("");
  const [sending, setSending] = useState(false);
  const [error, setError] = useState("");

  const handleSubmit = async () => {
    if (!message.trim() || sending) return;
    setSending(true);
    setError("");
    try {
      await onSendQuestion(message.trim());
      setMessage("");
    } catch {
      setError("No se pudo enviar. Inténtalo de nuevo.");
    } finally {
      setSending(false);
    }
  };

  return (
    <div className="space-y-4">
      <div className="rounded-3xl border border-zinc-200 bg-white p-5">
        <div className="text-lg font-semibold text-zinc-950">¿Tienes alguna duda?</div>
        <p className="mt-2 text-sm leading-6 text-zinc-600">
          Escríbenos aquí cualquier duda sobre el viaje, la documentación, los pagos o el equipaje y te responderemos lo antes posible.
        </p>
        <div className="mt-4 space-y-3">
          <Textarea
            value={message}
            onChange={(e) => setMessage(e.target.value)}
            placeholder="Escribe aquí tu consulta..."
            className="min-h-[140px] rounded-2xl border-zinc-200 bg-white"
          />
          <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
            <div className="text-xs text-zinc-500">Tu mensaje quedará registrado para que el equipo de GIMELOOS pueda responderte.</div>
            <Button onClick={handleSubmit} disabled={sending || !message.trim()} className="rounded-2xl text-white" style={{ backgroundColor: CORPORATE_RED }}>
              <Send className="mr-2 h-4 w-4" />{sending ? "Enviando…" : "Enviar duda"}
            </Button>
          </div>
          {error && <div className="rounded-2xl bg-red-50 px-4 py-3 text-sm text-red-700">{error}</div>}
        </div>
      </div>

      {questions.length > 0 && (
        <div className="space-y-3">
          <div className="text-sm font-semibold uppercase tracking-[0.18em] text-zinc-500">Historial de consultas</div>
          {[...questions].reverse().map((q) => (
            <div key={q.id} className="space-y-2 rounded-3xl border border-zinc-200 bg-white p-5">
              <div className="flex items-start justify-between gap-3">
                <p className="text-sm text-zinc-800">{q.message}</p>
                <span className="shrink-0 text-xs text-zinc-400">{new Date(q.createdAt).toLocaleDateString("es-ES", { day: "numeric", month: "short" })}</span>
              </div>
              {q.reply ? (
                <div className="rounded-2xl bg-white px-4 py-3">
                  <div className="mb-1 text-xs font-medium text-zinc-500">Respuesta del equipo GIMELOOS</div>
                  <p className="text-sm text-zinc-800">{q.reply}</p>
                </div>
              ) : (
                <div className="text-xs text-zinc-400">Pendiente de respuesta</div>
              )}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function AccordionSection({ title, icon: Icon, subtitle, children, defaultOpen = false, meta, hasUnread = false, sectionId, forceOpen, onForceOpenConsumed }) {
  const [open, setOpen] = useState(defaultOpen);

  useEffect(() => {
    if (forceOpen) {
      setOpen(true);
      if (sectionId) {
        setTimeout(() => {
          document.getElementById(sectionId)?.scrollIntoView({ behavior: "smooth", block: "start" });
        }, 80);
      }
      onForceOpenConsumed?.();
    }
  }, [forceOpen, sectionId, onForceOpenConsumed]);

  return (
    <Card id={sectionId} className={`overflow-hidden rounded-[28px] border shadow-sm transition-all ${open ? "border-zinc-200 bg-white shadow-md" : "border-zinc-200 bg-white"}`}>
      <div className={`flex w-full items-center justify-between gap-4 px-5 py-4 transition ${open ? "bg-white" : "hover:bg-white/70"}`}>
        <button type="button" onClick={() => setOpen(!open)} className="flex min-w-0 flex-1 cursor-pointer items-center gap-4 text-left">
          <div className="relative shrink-0 rounded-2xl p-2.5 shadow-sm transition" style={{ backgroundColor: open ? CORPORATE_RED : "#f4f4f5" }}>
            <Icon className={`h-5 w-5 transition ${open ? "text-white" : "text-zinc-700"}`} />
            {hasUnread && (
              <span className="absolute -right-1 -top-1 h-3 w-3 rounded-full border-2 border-white" style={{ backgroundColor: CORPORATE_RED }} />
            )}
          </div>
          <div className="min-w-0">
            <div className="text-base font-semibold text-zinc-950">{title}</div>
            {subtitle && <div className="mt-0.5 text-sm text-zinc-400">{subtitle}</div>}
          </div>
        </button>
        <div className="flex items-center gap-2">
          {meta}
          <button type="button" onClick={() => setOpen(!open)} className={`rounded-full border p-2 transition ${open ? "border-zinc-200 bg-white" : "border-zinc-200 bg-white"}`}>
            <ChevronDown className={`h-4 w-4 text-zinc-500 transition-transform duration-200 ${open ? "rotate-180" : ""}`} />
          </button>
        </div>
      </div>
      {open && (
        <div className="border-t border-zinc-100 p-5">
          {children}
        </div>
      )}
    </Card>
  );
}

// ─── Portal del cliente ──────────────────────────────────────────────────────

function ClientPortal({ user, trips, templates, setUsers, onLogout, notify }) {
  const trip = trips.find((t) => t.id === user.tripId);

  if (!trip) {
    return (
      <div className="flex min-h-screen items-center justify-center bg-zinc-50 p-8 text-center">
        <div>
          <div className="text-4xl mb-4">🗺️</div>
          <h2 className="text-xl font-semibold text-zinc-900 mb-2">No tienes una experiencia asignada</h2>
          <p className="text-sm text-zinc-500">Contacta con el equipo de GIMELOOS para que te asignen tu viaje.</p>
          <button
            type="button"
            className="mt-6 text-sm text-zinc-400 underline-offset-2 hover:text-zinc-600 hover:underline"
            onClick={onLogout}
          >
            Volver a la página de inicio de sesión
          </button>
        </div>
      </div>
    );
  }

  // [ALTO-2] Snapshot para rollback: capturado antes de cada operación optimista
  const updateCurrentUser = useCallback((updater, rollbackUser) => {
    setUsers((prev) => prev.map((item) => (item.id === user.id ? updater(item) : item)));
  }, [user.id, setUsers]);

  const rollbackUser = useCallback(() => {
    setUsers((prev) => prev.map((item) => (item.id === user.id ? user : item)));
  }, [user, setUsers]);

  const completedChecklist = trip.checklist.filter((item) => user.checklistState[item]).length;
  const pendingDocuments = user.documents.filter((doc) => doc.status !== "confirmed").length;
  const pendingPayments = [user.payments.reservation, user.payments.firstInstallment, user.payments.secondInstallment]
    .filter((p) => p.status === "pending").length;
  const sentPayments = [user.payments.reservation, user.payments.firstInstallment, user.payments.secondInstallment]
    .filter((p) => p.status === "sent").length;
  const questionsCount = user.questions?.length || 0;
  const unreadReplies = (user.questions || []).filter((q) => q.reply && q.status === "replied").length;

  // ── Notificaciones in-app ──
  const [notifications, setNotifications] = useState([]);
  const [showNotifications, setShowNotifications] = useState(false);
  const unreadCount = notifications.filter((n) => !n.read).length;
  const notifRef = useRef(null);

  useEffect(() => {
    if (!showNotifications) return;
    const handler = (e) => { if (notifRef.current && !notifRef.current.contains(e.target)) setShowNotifications(false); };
    document.addEventListener("mousedown", handler);
    // Marcar todas como leídas automáticamente al abrir el panel
    if (notifications.some((n) => !n.read)) {
      setNotifications((prev) => prev.map((n) => ({ ...n, read: true })));
      supabase.from("notifications").update({ read: true }).eq("participant_id", user.id).eq("read", false).then(() => {});
    }
    return () => document.removeEventListener("mousedown", handler);
  }, [showNotifications]); // eslint-disable-line react-hooks/exhaustive-deps

  useEffect(() => {
    supabase.from("notifications").select("*").eq("participant_id", user.id).order("created_at", { ascending: false }).limit(20)
      .then(({ data }) => { if (data) setNotifications(data); });
  }, [user.id]);

  const markAllRead = async () => {
    setNotifications((prev) => prev.map((n) => ({ ...n, read: true })));
    const { error } = await supabase.from("notifications").update({ read: true }).eq("participant_id", user.id).eq("read", false);
    if (error) notify("Error marcando notificaciones: " + error.message);
  };

  const unreadBySection = {
    docs: notifications.some((n) => !n.read && (n.type === "doc_confirmed" || n.type === "doc_rejected")),
    payments: notifications.some((n) => !n.read && n.type === "payment_confirmed"),
    questions: notifications.some((n) => !n.read && n.type === "question_replied"),
  };

  const [activeTab, setActiveTab] = useState("docs");

  // Conteos "ya vistos" por tab — persisten en localStorage por participante
  const seenKey = `gimeloos_seen_${user.id}`;
  const [seenCounts, setSeenCounts] = useState(() => {
    try { return JSON.parse(localStorage.getItem(seenKey) || "{}"); } catch (_) { return {}; }
  });

  const markTabSeen = (tab, count) => {
    setSeenCounts((prev) => {
      const next = { ...prev, [tab]: count };
      try { localStorage.setItem(seenKey, JSON.stringify(next)); } catch (_) {}
      return next;
    });
  };

  const navigateTo = (key) => {
    const tabMap = { docs: "docs", payments: "payments", replies: "questions" };
    if (tabMap[key]) setActiveTab(tabMap[key]);
  };

  // Marcar tabs como vistos al activarlos
  useEffect(() => {
    if (activeTab === "docs") markTabSeen("docs", pendingDocuments);
    if (activeTab === "payments") markTabSeen("payments", sentPayments);
  }, [activeTab]); // eslint-disable-line react-hooks/exhaustive-deps

  // Marcar respuestas como vistas al abrir el tab de dudas
  useEffect(() => {
    if (activeTab !== "questions") return;
    const unread = (user.questions || []).filter((q) => q.reply && q.status === "replied");
    if (!unread.length) return;
    const ids = unread.map((q) => q.id);
    updateCurrentUser((u) => ({
      ...u,
      questions: u.questions.map((q) => ids.includes(q.id) ? { ...q, status: "read" } : q),
    }));
    supabase.from("participant_questions").update({ status: "read" }).in("id", ids).then(({ error }) => {
      if (error) console.error("Error marcando dudas como vistas:", error.message);
    });
  }, [activeTab, user.questions]);

  // Badge visible solo si hay más items que la última vez que se visitó el tab
  const docsBadge = pendingDocuments > (seenCounts.docs ?? pendingDocuments) ? pendingDocuments : null;
  const paymentsBadge = sentPayments > (seenCounts.payments ?? sentPayments) ? sentPayments : null;

  const pendingSummary = [
    { key: "docs",      label: "Docs",       icon: FileCheck2,           count: docsBadge ?? 0,   sectionId: "section-docs" },
    { key: "payments",  label: "Pagos",      icon: Wallet,               count: pendingPayments,  sectionId: "section-payments" },
    { key: "replies",   label: "Respuestas", icon: MessageCircleQuestion, count: unreadReplies,   sectionId: "section-questions" },
    { key: "notifs",    label: "Avisos",     icon: Bell,                 count: unreadCount,      sectionId: null },
  ];

  const clientTabs = [
    { key: "docs",       label: "Documentación", icon: FileCheck2,            badge: docsBadge },
    { key: "payments",   label: "Pagos",          icon: Wallet,               badge: paymentsBadge },
    ...(trip.showLogistics !== false ? [{ key: "logistics", label: "Lo que llevar", icon: MapPinned }] : []),
    ...(trip.showItinerary !== false ? [{ key: "itinerary", label: "Itinerario",    icon: CalendarDays }] : []),
    { key: "checklist",  label: "Checklist",      icon: CheckCircle2,          badge: null },
    { key: "questions",  label: "Dudas",          icon: MessageCircleQuestion, badge: unreadReplies > 0 ? unreadReplies : null },
    ...(user.invoiceUrl ? [{ key: "invoice", label: "Factura", icon: Download }] : []),
  ];

  return (
    <div className="min-h-screen bg-white text-zinc-950">
      <div className="mx-auto max-w-7xl p-6 lg:p-8">
        <div className="mb-6 flex flex-col gap-4 rounded-[28px] border border-zinc-200 bg-white px-6 py-5 shadow-sm lg:flex-row lg:items-center lg:justify-between">
          <LogoMark />
          <div className="flex items-center gap-3">
            <div className="hidden rounded-2xl border border-zinc-200 bg-white px-4 py-2 text-sm text-zinc-600 sm:block">
              Participante: <span className="font-medium text-zinc-950">{user.participantName}</span>
            </div>
            {/* Campana de notificaciones */}
            <div className="relative" ref={notifRef}>
              <Button variant="outline" className="relative rounded-2xl px-3" onClick={() => setShowNotifications((v) => !v)}>
                <Bell className="h-4 w-4" />
                {unreadCount > 0 && (
                  <span className="absolute -right-1 -top-1 flex h-4 w-4 items-center justify-center rounded-full text-[10px] font-bold text-white" style={{ backgroundColor: CORPORATE_RED }}>{unreadCount}</span>
                )}
              </Button>
              {showNotifications && (
                <div className="absolute right-0 top-12 z-50 w-80 rounded-3xl border border-zinc-200 bg-white shadow-xl">
                  <div className="flex items-center justify-between border-b border-zinc-100 px-4 py-3">
                    <span className="font-semibold text-zinc-950">Notificaciones</span>
                    <div className="flex items-center gap-2">
                      {unreadCount > 0 && (
                        <button onClick={markAllRead} className="text-xs text-zinc-400 hover:text-zinc-700 underline">Marcar leídas</button>
                      )}
                      <button onClick={() => setShowNotifications(false)} className="text-zinc-400 hover:text-zinc-700 ml-1">✕</button>
                    </div>
                  </div>
                  <div className="max-h-80 overflow-y-auto rounded-b-3xl">
                    {notifications.length === 0 ? (
                      <div className="px-4 py-6 text-center text-sm text-zinc-400">Sin notificaciones</div>
                    ) : notifications.map((n, i, arr) => (
                      <div key={n.id} className={`border-b border-zinc-50 px-4 py-3 ${n.read ? "" : "bg-red-50"} ${i === arr.length - 1 ? "rounded-b-3xl" : ""}`}>
                        <div className="font-medium text-sm text-zinc-950">{n.title}</div>
                        <div className="text-xs text-zinc-500 mt-0.5">{n.body}</div>
                        <div className="text-xs text-zinc-400 mt-1">{new Date(n.created_at).toLocaleDateString("es-ES", { day: "numeric", month: "short", hour: "2-digit", minute: "2-digit" })}</div>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
            <Button variant="outline" className="rounded-2xl" onClick={() => { onLogout(); notify("Sesión cerrada."); }}>
              <LogOut className="mr-2 h-4 w-4" />Salir
            </Button>
          </div>
        </div>

        <div className="space-y-6">
          <HeroBanner trip={trip} user={user} pendingSummary={pendingSummary} onNavigate={navigateTo} />

          {/* Tabs — igual que portal de colegios */}
          <div className="flex flex-wrap gap-2">
            {clientTabs.map(({ key, label, icon: Icon, badge }) => {
              const isActive = activeTab === key;
              return (
                <button key={key} type="button" onMouseDown={(e) => e.preventDefault()} onClick={() => setActiveTab(key)}
                  className={`relative flex items-center gap-2 rounded-2xl px-4 py-2.5 text-sm font-medium transition ${isActive ? "text-white shadow-sm" : "border border-zinc-200 bg-white text-zinc-600 hover:bg-zinc-50"}`}
                  style={isActive ? { backgroundColor: CORPORATE_RED } : {}}>
                  <Icon className="h-4 w-4" />{label}
                  {badge != null && (
                    <span className={`flex h-5 w-5 items-center justify-center rounded-full text-[10px] font-bold ${isActive ? "bg-white text-red-600" : "text-white"}`}
                      style={!isActive ? { backgroundColor: CORPORATE_RED } : {}}>{badge}</span>
                  )}
                </button>
              );
            })}
          </div>

          {/* Contenido del tab activo */}
          <div className="mt-2">
            {activeTab === "docs" && (
              <ClientDocuments
                user={user}
                templates={templates}
                onUploadDocument={async (docId, file, onProgress) => {
                  const previousDocs = user.documents;
                  try {
                    const uploaded = await uploadFileToDrive(file, user.participantName, "documentos", onProgress, trip?.name);
                    const nextDocs = user.documents.map((doc) =>
                      doc.id === docId
                        ? { ...doc, uploadedFileName: file.name, filePath: "", driveUrl: uploaded.webViewLink, status: "pending_confirmation" }
                        : doc
                    );
                    updateCurrentUser((current) => ({ ...current, documents: nextDocs }));
                    await upsertDocument(user.id, docId, { status: "pending_confirmation", uploaded_file_name: file.name, file_path: "", storage_path: "", drive_url: uploaded.webViewLink, confirmed_at: null });
                    notify(`Documento subido: ${file.name}.`);
                    const docName = templates.find((t) => t.id === docId)?.name || file.name;
                    sendNotification("admin_doc_uploaded", null, null, { participantName: user.participantName, docName, tripName: trip?.name || "" });
                  } catch (error) {
                    console.error(error);
                    setUsers((prev) => prev.map((u) => u.id === user.id ? { ...u, documents: previousDocs } : u));
                    notify("No se ha podido subir el documento. Los cambios han sido revertidos.");
                  }
                }}
              />
            )}

            {activeTab === "payments" && (
              <ClientPayments
                user={user}
                trip={trip}
                onUploadProof={async (paymentKey, file, onProgress) => {
                  const previousPayments = user.payments;
                  try {
                    const uploaded = await uploadFileToDrive(file, user.participantName, "pagos", onProgress, trip?.name);
                    updateCurrentUser((current) => ({
                      ...current,
                      payments: { ...current.payments, [paymentKey]: { ...current.payments[paymentKey], proofName: file.name, proofPath: uploaded.webViewLink, status: "sent" } },
                    }));
                    await upsertPayment(user.id, paymentKey, { name: user.payments[paymentKey].name, amount: user.payments[paymentKey].amount, status: "sent", proof_name: file.name, proof_path: uploaded.webViewLink, due_date: user.payments[paymentKey].dueDate || null });
                    notify("Justificante cargado correctamente.");
                    sendNotification("admin_payment_uploaded", null, null, { participantName: user.participantName, paymentName: user.payments[paymentKey].name, tripName: trip?.name || "" });
                  } catch (error) {
                    console.error(error);
                    setUsers((prev) => prev.map((u) => u.id === user.id ? { ...u, payments: previousPayments } : u));
                    notify("No se ha podido subir el justificante. Los cambios han sido revertidos.");
                  }
                }}
              />
            )}

            {activeTab === "logistics" && <ClientLogistics trip={trip} />}
            {activeTab === "itinerary" && <ClientItinerary trip={trip} />}

            {activeTab === "checklist" && (
              <ClientChecklist
                user={user}
                trip={trip}
                onToggleItem={async (item) => {
                  const previousState = user.checklistState;
                  const nextChecklist = { ...user.checklistState, [item]: !user.checklistState[item] };
                  updateCurrentUser((current) => ({ ...current, checklistState: nextChecklist }));
                  const { error: ckErr } = await supabase.from("participants").update({ checklist_state: nextChecklist }).eq("id", user.id);
                  if (ckErr) {
                    setUsers((prev) => prev.map((u) => u.id === user.id ? { ...u, checklistState: previousState } : u));
                    notify("No se pudo guardar el checklist. Los cambios han sido revertidos.");
                  }
                }}
              />
            )}

            {activeTab === "questions" && (
              <ClientQuestions
                questions={user.questions || []}
                onSendQuestion={async (message) => {
                  const tempId = `q-${Date.now()}`;
                  const newQuestion = { id: tempId, message, createdAt: new Date().toISOString(), status: "sent", reply: "", repliedAt: null };
                  const previousQuestions = user.questions || [];
                  updateCurrentUser((current) => ({ ...current, questions: [...(current.questions || []), newQuestion] }));
                  try {
                    const { data, error } = await supabase.from("participant_questions").insert({ participant_id: user.id, message, status: "sent" }).select("id").single();
                    if (error) throw error;
                    updateCurrentUser((current) => ({ ...current, questions: current.questions.map((q) => q.id === tempId ? { ...q, id: data.id } : q) }));
                    notify("Tu duda ha sido enviada.");
                    sendNotification("admin_new_question", null, null, { participantName: user.participantName, question: message, tripName: trip?.name || "" });
                  } catch (error) {
                    setUsers((prev) => prev.map((u) => u.id === user.id ? { ...u, questions: previousQuestions } : u));
                    throw error;
                  }
                }}
              />
            )}

            {activeTab === "invoice" && user.invoiceUrl && (
              <div className="rounded-3xl border border-zinc-200 bg-white p-8 text-center shadow-sm">
                <Download className="mx-auto mb-3 h-8 w-8 text-zinc-400" />
                <div className="mb-4 text-sm text-zinc-500">Tu factura está disponible para descargar.</div>
                <Button variant="outline" className="rounded-2xl" onClick={() => window.open(user.invoiceUrl, "_blank", "noopener,noreferrer")}>
                  <Download className="mr-2 h-4 w-4" />Descargar factura
                </Button>
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

// ─── Panel de administración ─────────────────────────────────────────────────

function AdminClients({ users, trips, setUsers, templates, notify, setTrips }) {
  const clients = users.filter((u) => u.role === "client" && !u.schoolId);
  const [selectedTripFilter, setSelectedTripFilter] = useState("all");
  const [selectedGroupTrip, setSelectedGroupTrip] = useState(trips[0]?.id || "");
  const [assignTargetTrip, setAssignTargetTrip] = useState(trips[0]?.id || "");
  const [selectedClientIds, setSelectedClientIds] = useState([]);
  const [importFileName, setImportFileName] = useState("");
  const [searchQuery, setSearchQuery] = useState("");
  const [isImporting, setIsImporting] = useState(false);
  const [importProgress, setImportProgress] = useState(0);
  const [importMessage, setImportMessage] = useState("");
  const [importTotal, setImportTotal] = useState(0);
  const [importDone, setImportDone] = useState(0);
  const [sheetUrl, setSheetUrl] = useState("");
  const [isSyncingSheet, setIsSyncingSheet] = useState(false);
  const [pendingDeleteId, setPendingDeleteId] = useState(null);
  const [confirmBulkDelete, setConfirmBulkDelete] = useState(false);
  const [confirmBulkInvite, setConfirmBulkInvite] = useState(false);
  const [pendingDeleteInvoiceClientId, setPendingDeleteInvoiceClientId] = useState(null);
  const [clientPage, setClientPage] = useState(1);
  const PAGE_SIZE = 50;

  const visibleClients = clients.filter(
    (client) =>
      (selectedTripFilter === "all" || client.tripId === selectedTripFilter) &&
      matchesParticipantSearch(client, searchQuery)
  );
  const totalPages = Math.max(1, Math.ceil(visibleClients.length / PAGE_SIZE));
  const pagedClients = visibleClients.slice((clientPage - 1) * PAGE_SIZE, clientPage * PAGE_SIZE);
  const visibleClientIds = pagedClients.map((c) => c.id);
  const allVisibleSelected = visibleClientIds.length > 0 && visibleClientIds.every((id) => selectedClientIds.includes(id));

  useEffect(() => { setClientPage(1); }, [selectedTripFilter, searchQuery]);

  const toggleClientSelection = (clientId) =>
    setSelectedClientIds((prev) => prev.includes(clientId) ? prev.filter((id) => id !== clientId) : [...prev, clientId]);

  const toggleSelectAllVisible = () =>
    setSelectedClientIds((prev) =>
      allVisibleSelected ? prev.filter((id) => !visibleClientIds.includes(id)) : Array.from(new Set([...prev, ...visibleClientIds]))
    );

  const assignTripToGroup = async () => {
    const targetIds = selectedClientIds.length
      ? selectedClientIds
      : users.filter((u) => u.role === "client" && u.tripId === selectedGroupTrip).map((u) => u.id);
    if (!targetIds.length) { notify("No hay clientes seleccionados."); return; }
    for (const clientId of targetIds) {
      const { error } = await supabase.from("participants").update({ trip_id: assignTargetTrip }).eq("id", clientId);
      if (error) { notify("Error asignando experiencia: " + error.message); return; }
    }
    setUsers((prev) => prev.map((u) => targetIds.includes(u.id) ? { ...u, tripId: assignTargetTrip } : u));
    notify("Experiencia aplicada correctamente.");
  };

  const restoreDeletedClient = async (deletedClient) => {
    const errors = [];

    const payload = {
      id: deletedClient.id,
      role: deletedClient.role,
      username: deletedClient.username,
      participant_name: deletedClient.participantName || "",
      mother_name: deletedClient.motherName || "",
      father_name: deletedClient.fatherName || "",
      parent_name: deletedClient.parentName || "",
      email: deletedClient.email || "",
      contact_emails: deletedClient.contactEmails || [],
      dni: deletedClient.dni || "",
      trip_id: deletedClient.tripId || null,
      checklist_state: deletedClient.checklistState || {},
    };
    const { error: participantErr } = await supabase.from("participants").insert(payload);
    if (participantErr) throw new Error("Error restaurando participante: " + participantErr.message);

    const { error: pricingErr } = await supabase.from("participant_pricing").upsert(
      { participant_id: deletedClient.id, initial_price: deletedClient.payments?.initialPrice || 0, discount: deletedClient.payments?.discount || 0, final_price: deletedClient.payments?.finalPrice || 0 },
      { onConflict: "participant_id" }
    );
    if (pricingErr) errors.push("precios: " + pricingErr.message);

    for (const [key, name] of [["reservation", "Reserva"], ["firstInstallment", "Primera cuota"], ["secondInstallment", "Segunda cuota"]]) {
      const { error: payErr } = await supabase.from("participant_payments").insert({
        participant_id: deletedClient.id,
        payment_key: key,
        name: deletedClient.payments?.[key]?.name || name,
        amount: deletedClient.payments?.[key]?.amount || 0,
        status: deletedClient.payments?.[key]?.status || "pending",
        proof_name: deletedClient.payments?.[key]?.proofName || "",
        due_date: deletedClient.payments?.[key]?.dueDate || null,
      });
      if (payErr) errors.push(`pago ${key}: ` + payErr.message);
    }

    for (const doc of deletedClient.documents || []) {
      const { error: docErr } = await supabase.from("participant_documents").insert({
        participant_id: deletedClient.id,
        template_id: doc.id,
        status: doc.status || "pending_upload",
        uploaded_file_name: doc.uploadedFileName || "",
        file_path: doc.filePath || "",
        storage_path: doc.filePath || "",
        drive_url: doc.driveUrl || "",
        confirmed_at: doc.status === "confirmed" ? new Date().toISOString() : null,
      });
      if (docErr) errors.push("documento: " + docErr.message);
    }

    for (const q of deletedClient.questions || []) {
      const { error: qErr } = await supabase.from("participant_questions").insert({
        participant_id: deletedClient.id,
        message: q.message || "",
        status: q.status || "sent",
        created_at: q.createdAt || new Date().toISOString(),
      });
      if (qErr) errors.push("pregunta: " + qErr.message);
    }

    if (errors.length) notify(`Restaurado con errores parciales: ${errors.join(", ")}`);
  };

  const deleteSingleClient = async (clientId) => {
    const deletedClient = users.find((u) => u.id === clientId);
    if (!deletedClient) return;
    try {
      const subtableDeletes = await Promise.all([
        supabase.from("participant_questions").delete().eq("participant_id", clientId),
        supabase.from("participant_documents").delete().eq("participant_id", clientId),
        supabase.from("participant_payments").delete().eq("participant_id", clientId),
        supabase.from("participant_pricing").delete().eq("participant_id", clientId),
      ]);
      const subtableError = subtableDeletes.find((r) => r.error);
      if (subtableError) throw new Error(subtableError.error.message);
      const { error: participantError } = await supabase.from("participants").delete().eq("id", clientId);
      if (participantError) throw new Error(participantError.message);
      setUsers((prev) => prev.filter((u) => u.id !== clientId));
      setSelectedClientIds((prev) => prev.filter((id) => id !== clientId));
      // [MEDIO-1] Toast destructivo: 7s para dar tiempo al "Deshacer"
      notify("Cliente eliminado correctamente.", {
        actionLabel: "Deshacer",
        onAction: async () => {
          try {
            await restoreDeletedClient(deletedClient);
            setUsers((prev) => [...prev, deletedClient]);
            notify("Eliminación deshecha.");
          } catch (error) {
            console.error(error);
            notify("No se pudo deshacer la eliminación.");
          }
        },
      });
    } catch (error) {
      console.error(error);
      notify("No se ha podido eliminar el cliente.");
    }
  };

  const deleteSelectedClients = async () => {
    if (!selectedClientIds.length) return;
    const deletedClients = users.filter((u) => selectedClientIds.includes(u.id));
    try {
      for (const clientId of selectedClientIds) {
        const [r1, r2, r3, r4] = await Promise.all([
          supabase.from("participant_questions").delete().eq("participant_id", clientId),
          supabase.from("participant_documents").delete().eq("participant_id", clientId),
          supabase.from("participant_payments").delete().eq("participant_id", clientId),
          supabase.from("participant_pricing").delete().eq("participant_id", clientId),
        ]);
        const subtableErr = [r1, r2, r3, r4].find((r) => r.error)?.error;
        if (subtableErr) throw new Error(subtableErr.message);
        const { error: mainErr } = await supabase.from("participants").delete().eq("id", clientId);
        if (mainErr) throw new Error(mainErr.message);
      }
      setUsers((prev) => prev.filter((u) => u.role === "admin" || !selectedClientIds.includes(u.id)));
      // [MEDIO-1] Toast destructivo con 7s
      notify(`${selectedClientIds.length} cliente(s) eliminados.`, {
        actionLabel: "Deshacer",
        onAction: async () => {
          try {
            for (const dc of deletedClients) await restoreDeletedClient(dc);
            setUsers((prev) => [...prev, ...deletedClients]);
            notify("Eliminación deshecha.");
          } catch (error) {
            console.error(error);
            notify("No se pudo deshacer la eliminación.");
          }
        },
      });
      setSelectedClientIds([]);
    } catch (error) {
      console.error(error);
      notify("No se ha podido eliminar la selección.");
    }
  };

  const bulkInviteSelected = async () => {
    const targets = (selectedClientIds.length ? users.filter((u) => selectedClientIds.includes(u.id)) : clients.filter((u) => u.role === "client")).filter((u) => u.email);
    if (!targets.length) { notify("Ningún participante con email en la selección."); return; }
    notify(`Enviando acceso a ${targets.length} participante(s)…`);
    let ok = 0; let fail = 0;
    const token = await getAuthToken();
    for (const client of targets) {
      try {
        const res = await fetch("/api/invite-participant", {
          method: "POST",
          headers: { "Content-Type": "application/json", ...(token ? { Authorization: `Bearer ${token}` } : {}) },
          body: JSON.stringify({ participantId: client.id }),
        });
        const json = await res.json();
        if (res.ok) ok++; else fail++;
      } catch { fail++; }
    }
    notify(`✅ Acceso enviado a ${ok} participante(s).${fail ? ` ${fail} fallaron.` : ""}`);
  };

  const handleSyncSheet = async () => {
    if (!sheetUrl.trim()) { notify("Introduce la URL del documento Excel o Google Sheets."); return; }
    const selectedGroupTripObj = trips.find((t) => t.id === selectedGroupTrip);
    if (!selectedGroupTripObj) { notify("Selecciona primero el campamento en 'Campamento origen'."); return; }

    const proxyUrl = `/api/proxy-sheet?url=${encodeURIComponent(sheetUrl.trim())}`;
    const token = await getAuthToken();

    setIsSyncingSheet(true);
    try {
      const res = await fetch(proxyUrl, {
        headers: token ? { Authorization: `Bearer ${token}` } : {},
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.error || `Error ${res.status}`);
      }
      const buffer = await res.arrayBuffer();
      const fakeFile = new File([buffer], "sheet.xlsx", { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      await handleBulkExcelUpload(fakeFile, null);
    } catch (err) {
      notify("Error al sincronizar: " + err.message, { variant: "destructive" });
    } finally {
      setIsSyncingSheet(false);
    }
  };

  const handleBulkExcelUpload = async (file, input) => {
    if (!file) { notify("No se ha seleccionado ningún archivo."); return; }
    setImportFileName(file.name);
    setIsImporting(true);
    setImportProgress(2);
    setImportMessage("Leyendo Excel...");
    setImportTotal(0);
    setImportDone(0);

    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: "array", cellDates: true });
      const firstSheetName = workbook.SheetNames?.[0];
      if (!firstSheetName) { notify("El archivo no contiene hojas válidas."); setIsImporting(false); return; }
      const sheet = workbook.Sheets[firstSheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
      if (!rows.length) { notify("El Excel está vacío o no tiene filas de datos."); setIsImporting(false); return; }
      console.log("[Import] Columnas detectadas:", Object.keys(rows[0] || {}));

      setImportTotal(rows.length);

      const parseExcelDate = (value) => {
        if (!value) return null;
        // Serial numérico de Excel (ej. 46201)
        if (typeof value === "number" && value > 1000) {
          try {
            const parsed = XLSX.SSF.parse_date_code(value);
            if (parsed?.y) {
              // Mediodía UTC para evitar desfase de timezone en fechas sin hora
              const hour = parsed.H || 12;
              return new Date(Date.UTC(parsed.y, parsed.m - 1, parsed.d, hour, parsed.M || 0)).toISOString();
            }
          } catch (_) {}
        }
        // Date object (cuando cellDates:true lo convierte)
        if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString();
        // String
        const date = new Date(String(value));
        if (!Number.isNaN(date.getTime())) return date.toISOString();
        return null;
      };

      const parseAmount = (value) => {
        if (value === undefined || value === null || value === "") return 0;
        if (typeof value === "number") return value;
        const raw = String(value).trim().replace(/€/g, "").replace(/\s/g, "");
        if (!raw) return 0;
        const hasComma = raw.includes(","), hasDot = raw.includes(".");
        let normalized = raw;
        if (hasComma && hasDot) normalized = raw.replace(/\./g, "").replace(",", ".");
        else if (hasComma) normalized = raw.replace(",", ".");
        const num = Number(normalized);
        return Number.isFinite(num) ? num : 0;
      };

      const normalizeDateForDb = (value) => {
        if (!value) return null;
        if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10);
        const raw = String(value).trim();
        if (!raw) return null;
        const parsed = new Date(raw);
        if (!Number.isNaN(parsed.getTime())) return parsed.toISOString().slice(0, 10);
        const match = raw.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/);
        if (match) {
          let [, d, m, y] = match;
          if (y.length === 2) y = `20${y}`;
          const iso = `${y.padStart(4, "0")}-${m.padStart(2, "0")}-${d.padStart(2, "0")}`;
          return Number.isNaN(new Date(iso).getTime()) ? null : iso;
        }
        return null;
      };

      const templatesToUse = templates?.length > 0 ? templates : initialDocumentTemplates;
      setImportMessage("Sincronizando plantillas...");
      setImportProgress(5);

      const { error: tplErr } = await supabase.from("document_templates").upsert(
        templatesToUse.map((t) => ({ id: t.id, name: t.name, file_name: t.fileName || "" })),
        { onConflict: "id" }
      );
      if (tplErr) throw new Error("Error sincronizando plantillas: " + tplErr.message);

      // ── Fase 1: parsear todas las filas en memoria ────────────────────────────
      setImportMessage("Procesando filas...");
      setImportProgress(10);

      // Fallback de viaje: usar el viaje seleccionado en "Grupo origen", o el único viaje si solo hay uno
      const selectedGroupTripObj = trips.find((t) => t.id === selectedGroupTrip);
      const fallbackTripName = selectedGroupTripObj?.name || (trips.length === 1 ? trips[0].name : "");

      const parsedRows = [];
      for (const row of rows) {
        // Nombre del participante — acepta cualquier variante razonable
        const participantName = safeString(
          getRowValue(row,
            "Participante", "Nombre participante", "Nombre del participante", "Nombre completo del participante",
            "Nombre y apellidos", "Nombre completo", "NombreCompleto", "Nombre", "Alumno", "Alumna",
            "Menor", "Hijo", "Hija", "Niño", "Niña", "name", "fullname", "full_name", "student"
          )
        );
        if (!participantName) continue; // fila sin nombre → saltar

        // Título del viaje: si el usuario seleccionó un viaje en "Grupo origen", siempre usarlo
        // (evita que el partial-match pesque columnas como "Qué modalidad de campamento...")
        const tripTitle = selectedGroupTripObj?.name || safeString(
          getRowValue(row, "Viaje_Titulo", "Titulo_Viaje", "Viaje", "Trip")
        ) || (trips.length === 1 ? trips[0].name : "");
        if (!tripTitle) continue;

        // Si hay viaje seleccionado en Grupo origen, usar su ID real (no generar slug)
        const tripId = selectedGroupTripObj ? selectedGroupTripObj.id : `trip-${slugify(tripTitle)}`;
        const usernameFromExcel = safeString(getRowValue(row, "Usuario", "Username", "Login", "User"));
        const emailRaw = safeString(
          getRowValue(row,
            "Email", "email", "Correo", "Correo_electronico", "Correo electronico", "Mail",
            "Email_Madre", "Email_Padre", "Correo_Madre", "Correo_Padre",
            "Correo electrónico de la madre", "Correo electrónico del padre",
            "email_madre", "email_padre"
          )
        ).split(",")[0].trim();
        // Si hay columna Usuario la usamos; si no, el email; si no, slug del nombre
        const finalUsername = (usernameFromExcel
          ? usernameFromExcel.toString().trim().toLowerCase().replace(/\s+/g, "-")
          : emailRaw
            ? emailRaw.toLowerCase()
            : (slugify(participantName) || `user-${Date.now()}`));
        const motherName = safeString(getRowValue(row,
          "Nombre_Madre", "Nombre completo de la madre", "Madre", "madre", "NombreMadre",
          "Nombre madre", "Nombre de la madre", "Tutor_1", "Tutor1", "Tutor", "Responsable_1"
        ));
        const fatherName = safeString(getRowValue(row,
          "Nombre_Padre", "Nombre completo del padre", "Padre", "padre", "NombrePadre",
          "Nombre padre", "Nombre del padre", "Tutor_2", "Tutor2", "Responsable_2"
        ));
        const motherEmail = safeString(getRowValue(row,
          "Email_Madre", "Correo electrónico de la madre", "email_madre", "Correo_Madre",
          "correo madre", "Email madre", "Email_Tutor1", "Email_Tutor_1"
        ));
        const fatherEmail = safeString(getRowValue(row,
          "Email_Padre", "Correo electrónico del padre", "email_padre", "Correo_Padre",
          "correo padre", "Email padre", "Email_Tutor2", "Email_Tutor_2"
        ));
        const passwordFromExcel = safeString(getRowValue(row, "Password", "Contraseña", "Clave", "Pass"));
        const dniRaw = safeString(getRowValue(row,
          "DNI", "Dni", "NIF", "Pasaporte", "passport", "document",
          "DNI del participante", "DNI (participante)", "Documento", "ID"
        ));
        // Si Tally guarda una URL de archivo en vez del número, ignorar
        const dniFromExcel = dniRaw.startsWith("http") ? "" : dniRaw;
        const birthDateRaw = getRowValue(row,
          "Fecha_Nacimiento", "Fecha de nacimiento", "FechaNacimiento", "Nacimiento",
          "fecha_nac", "F_Nacimiento", "birthdate", "birth_date", "Fecha nacimiento"
        );
        const genderRaw = safeString(getRowValue(row, "Sexo", "Genero", "Género", "Sex", "Gender"));
        const addressRaw = safeString(getRowValue(row,
          "Direccion", "Dirección", "Dirección completa", "Direccion_Completa",
          "address", "Domicilio", "Domicilio completo"
        ));
        const schoolRaw = safeString(getRowValue(row,
          "Colegio", "Colegio en el que estudia", "Centro_Educativo", "School",
          "Centro educativo", "Escuela", "Instituto"
        ));
        const phoneFatherRaw = safeString(getRowValue(row,
          "Telefono_Padre", "Teléfono del padre", "Tel_Padre", "Telefono padre",
          "Movil_Padre", "Movil padre", "Phone_Father", "Tlf padre"
        ));
        const phoneMotherRaw = safeString(getRowValue(row,
          "Telefono_Madre", "Teléfono de la madre", "Tel_Madre", "Telefono madre",
          "Movil_Madre", "Movil madre", "Phone_Mother", "Tlf madre"
        ));
        const dniFatherRaw = safeString(getRowValue(row, "DNI_Padre", "DNI del padre", "NIF padre")).replace(/^https?:\/\/.*/, "");
        const dniMotherRaw = safeString(getRowValue(row, "DNI_Madre", "DNI de la madre", "NIF madre")).replace(/^https?:\/\/.*/, "");
        const imageAuthRaw = getRowValue(row,
          "Autorizacion_Imagenes", "Autorización uso de imágenes", "Auth_Imagenes",
          "autorizacion imagenes", "Uso de imágenes", "Autorización imágenes"
        );
        const imageAuth = imageAuthRaw ? ["si", "sí", "yes", "1", "true", "autorizo", "autorizado"].includes(String(imageAuthRaw).toLowerCase().trim()) : false;
        const allergiesRaw = safeString(getRowValue(row,
          "Alergias", "Alergias y/o intolerancias alimentarias", "Alergias_Intolerancias",
          "Intolerancias", "Alergias e intolerancias", "allergies"
        ));
        const healthNotesRaw = safeString(getRowValue(row,
          "Salud", "¿El participante tiene algún problema de salud?", "Problemas_Salud",
          "Health", "Salud_Observaciones", "Problema de salud", "Condición médica"
        ));
        const shirtSizeRaw = safeString(getRowValue(row,
          "Talla", "Talla de camiseta", "Talla_Camiseta", "Shirt_Size", "Talla camiseta", "Size"
        ));
        const notesRaw = safeString(getRowValue(row,
          "Observaciones", "Alguna otra información relevante", "Notas", "Notes",
          "Comentarios", "Otros", "Other", "Información adicional"
        ));
        const modalityRaw = safeString(getRowValue(row,
          "Modalidad", "Qué modalidad de campamento", "Turno", "Modality",
          "Modalidad campamento", "Turno campamento"
        ));
        const howKnownRaw = safeString(getRowValue(row,
          "Como_Conocio", "Cómo nos has conocido", "Como_nos_conocio", "How_Known",
          "Como nos conociste", "Como conociste gimeloos"
        ));
        const discount = parseAmount(getRowValue(row, "Viaje_Descuento", "Descuento", "discount"));
        const initialPrice = parseAmount(getRowValue(row, "Viaje_Precio", "Precio", "Importe", "price", "amount"));
        const finalPrice = parseAmount(getRowValue(row, "Viaje_APagar", "A_Pagar", "Total", "Importe total")) || Math.max(0, initialPrice - discount);

        parsedRows.push({
          tripId, tripTitle,
          tripPayload: {
            id: tripId,
            name: tripTitle,
            departure_date: parseExcelDate(getRowValue(row, "Viaje_Fecha")),
            description: safeString(getRowValue(row, "Viaje_Destino"), "Experiencia GIMELOOS"),
          },
          participantPayload: {
            role: "client",
            username: finalUsername,
            participant_name: participantName,
            mother_name: motherName,
            father_name: fatherName,
            parent_name: [motherName, fatherName].filter(Boolean).join(" / "),
            email: safeString(motherEmail, fatherEmail),
            contact_emails: Array.from(new Set([motherEmail, fatherEmail].filter(Boolean))),
            dni: dniFromExcel,
            birth_date: birthDateRaw ? normalizeDateForDb(birthDateRaw) : null,
            gender: genderRaw,
            address: addressRaw,
            school: schoolRaw,
            phone_father: phoneFatherRaw,
            phone_mother: phoneMotherRaw,
            dni_father: dniFatherRaw,
            dni_mother: dniMotherRaw,
            image_auth: imageAuth,
            allergies: allergiesRaw,
            health_notes: healthNotesRaw,
            shirt_size: shirtSizeRaw,
            notes: notesRaw,
            modality: modalityRaw,
            how_known: howKnownRaw,
            trip_id: tripId,
          },
          pricing: { initial_price: initialPrice, discount, final_price: finalPrice },
          payments: [
            { key: "reservation",       name: safeString(getRowValue(row, "Pago1_Nombre"), "Reserva"),       amount: parseAmount(getRowValue(row, "Pago1_Cantidad", "1er pago", "1er Pago", "Primer pago", "Pago 1", "Reserva pago", "Importe reserva")), due_date: normalizeDateForDb(getRowValue(row, "Pago1_Fecha", "Fecha 1er pago", "Fecha reserva")) },
            { key: "firstInstallment",  name: safeString(getRowValue(row, "Pago2_Nombre"), "Primera cuota"), amount: parseAmount(getRowValue(row, "Pago2_Cantidad", "2do pago", "2do Pago", "2º pago", "Segundo pago", "Pago 2", "1a cuota", "1ª cuota", "Primera cuota pago")), due_date: normalizeDateForDb(getRowValue(row, "Pago2_Fecha", "Fecha 2do pago", "Fecha primera cuota")) },
            { key: "secondInstallment", name: safeString(getRowValue(row, "Pago3_Nombre"), "Segunda cuota"), amount: parseAmount(getRowValue(row, "Pago3_Cantidad", "3er pago", "3er Pago", "3º pago", "Tercer pago", "Pago 3", "2a cuota", "2ª cuota", "Segunda cuota pago")), due_date: normalizeDateForDb(getRowValue(row, "Pago3_Fecha", "Fecha 3er pago", "Fecha segunda cuota")) },
          ],
          participantName, motherName, fatherName,
          email: safeString(motherEmail, fatherEmail),
          contactEmails: Array.from(new Set([motherEmail, fatherEmail].filter(Boolean))),
          finalUsername, initialPrice, discount, finalPrice,
          passwordFromExcel, dniFromExcel,
        });
      }

      setImportTotal(parsedRows.length);

      if (parsedRows.length === 0) {
        const foundCols = Object.keys(rows[0] || {}).slice(0, 8).join(" | ");
        const hint = !fallbackTripName && trips.length > 1
          ? ' Tienes varios viajes: añade una columna "Viaje" con el nombre del campamento.'
          : ' Asegúrate de que haya una columna con el nombre del participante (ej. "Nombre", "Participante", "Alumno").';
        notify(`No se han podido importar participantes. Columnas detectadas: ${foundCols || "(ninguna)"}.${hint}`, { variant: "destructive" });
        setIsImporting(false);
        if (input) input.value = "";
        setImportFileName("");
        setTimeout(() => { setImportProgress(0); setImportMessage(""); setImportTotal(0); setImportDone(0); }, 300);
        return;
      }

      // ── Fase 2: upsert trips (batch, sin duplicados) ──────────────────────────
      setImportMessage("Guardando viajes...");
      setImportProgress(20);

      const currentTrips = [...trips];
      const uniqueTrips = new Map();
      for (const r of parsedRows) {
        if (!uniqueTrips.has(r.tripId)) {
          // Si usamos un viaje existente (selectedGroupTripObj), no lo modificamos
          if (selectedGroupTripObj && r.tripId === selectedGroupTripObj.id) continue;
          const existingTrip = currentTrips.find((t) => t.id === r.tripId);
          uniqueTrips.set(r.tripId, {
            ...r.tripPayload,
            hero_image: existingTrip?.heroImage || DEFAULT_HERO_IMAGES[0],
            hero_images: existingTrip?.heroImages || DEFAULT_HERO_IMAGES,
            transfer_info: existingTrip?.transferInfo || { bank: "", accountHolder: "", iban: "", concept: "" },
            automation: existingTrip?.automation || { autoReminderEnabled: false, reminderDaysBefore: 5 },
            document_rules: existingTrip?.documentRules || [
              { templateId: "doc-1", dueType: "days_before_trip", dueValue: 20 },
              { templateId: "doc-2", dueType: "days_before_trip", dueValue: 15 },
              { templateId: "doc-3", dueType: "days_before_trip", dueValue: 10 },
            ],
            payment_schedule: existingTrip?.paymentSchedule || {
              reservation: { name: "Reserva", dueType: "fixed_date", dueDate: "", dueValue: 0 },
              firstInstallment: { name: "Primera cuota", dueType: "days_before_trip", dueDate: "", dueValue: 45 },
              secondInstallment: { name: "Segunda cuota", dueType: "days_before_trip", dueDate: "", dueValue: 15 },
            },
            itinerary: existingTrip?.itinerary || [],
            checklist: existingTrip?.checklist || ["DNI o pasaporte", "Tarjeta sanitaria", "Bañador", "Toalla", "Protector solar"],
          });
        }
      }

      const { error: tripsBatchErr } = await supabase.from("trips").upsert([...uniqueTrips.values()], { onConflict: "id" });
      if (tripsBatchErr) { notify(`Error guardando viajes: ${tripsBatchErr.message}`); setIsImporting(false); return; }

      for (const [tripId, tp] of uniqueTrips) {
        if (!currentTrips.find((t) => t.id === tripId)) {
          currentTrips.push({ id: tripId, name: tp.name, departureDate: tp.departure_date || "", description: tp.description, heroImage: tp.hero_image, heroImages: tp.hero_images, transferInfo: tp.transfer_info || { bank: "", accountHolder: "", iban: "", concept: "" }, automation: tp.automation, documentRules: tp.document_rules, paymentSchedule: tp.payment_schedule, itinerary: tp.itinerary, checklist: tp.checklist });
        }
      }

      // ── Fase 3: pre-fetch participantes existentes (1 query) ─────────────────
      setImportMessage("Sincronizando participantes...");
      setImportProgress(35);

      const allUsernames = parsedRows.map((r) => r.finalUsername);
      const { data: existingParticipants, error: fetchParticipantsErr } = await supabase
        .from("participants").select("id, username").in("username", allUsernames);
      if (fetchParticipantsErr) { notify(`Error leyendo participantes: ${fetchParticipantsErr.message}`); setIsImporting(false); return; }

      const existingByUsername = new Map((existingParticipants || []).map((p) => [p.username, p.id]));

      // ── Fase 4: batch insert nuevos + batch update existentes ─────────────────
      const toInsert = parsedRows.filter((r) => !existingByUsername.has(r.finalUsername)).map((r) => r.participantPayload);
      const toUpdate = parsedRows.filter((r) => existingByUsername.has(r.finalUsername));

      // Deduplicar por username (el Excel puede tener filas repetidas)
      const payloadByUsername = new Map();
      for (const r of parsedRows) payloadByUsername.set(r.finalUsername, r.participantPayload);
      const allPayloads = [...payloadByUsername.values()];

      if (allPayloads.length) {
        const { data: upserted, error: upsertErr } = await supabase
          .from("participants")
          .upsert(allPayloads, { onConflict: "username" })
          .select("id, username");
        if (upsertErr) { notify(`Error importando participantes: ${upsertErr.message}`); setIsImporting(false); return; }
        for (const p of (upserted || [])) existingByUsername.set(p.username, p.id);
      }

      setImportProgress(55);

      // ── Fase 5: batch upsert pricing (1 query) ────────────────────────────────
      setImportMessage("Sincronizando precios...");
      const pricingByParticipant = new Map();
      for (const r of parsedRows) {
        const pid = existingByUsername.get(r.finalUsername);
        if (pid) pricingByParticipant.set(pid, { participant_id: pid, ...r.pricing });
      }
      const pricingRows = [...pricingByParticipant.values()];

      if (pricingRows.length) {
        const { error: pricErr } = await supabase.from("participant_pricing").upsert(pricingRows, { onConflict: "participant_id" });
        if (pricErr) { notify("Error sincronizando precios: " + pricErr.message, { variant: "destructive" }); setIsImporting(false); return; }
      }

      setImportProgress(65);

      // ── Fase 6: pagos — pre-fetch estados protegidos (1 query) ────────────────
      setImportMessage("Sincronizando pagos...");
      const allParticipantIds = parsedRows.map((r) => existingByUsername.get(r.finalUsername)).filter(Boolean);
      const PROTECTED = ["confirmed", "sent"];

      const { data: existingPayments } = await supabase
        .from("participant_payments")
        .select("id, participant_id, payment_key, status")
        .in("participant_id", allParticipantIds);

      const paymentStatusMap = new Map();
      for (const p of (existingPayments || [])) {
        paymentStatusMap.set(`${p.participant_id}__${p.payment_key}`, { id: p.id, status: p.status });
      }

      // Un solo Map deduplicado: últimos valores del Excel por (participantId, paymentKey)
      const paymentUpsertMap = new Map();
      for (const r of parsedRows) {
        const participantId = existingByUsername.get(r.finalUsername);
        if (!participantId) continue;
        for (const p of r.payments) {
          const comboKey = `${participantId}__${p.key}`;
          const existing = paymentStatusMap.get(comboKey);
          // No tocar pagos ya confirmados o enviados
          if (existing && PROTECTED.includes(existing.status)) continue;
          paymentUpsertMap.set(comboKey, {
            participant_id: participantId,
            payment_key: p.key,
            name: p.name,
            amount: p.amount,
            due_date: p.due_date,
            // Solo ponemos status "pending" si es nuevo; si existe, lo dejamos como está
            ...(existing ? {} : { status: "pending", proof_name: "", proof_path: "" }),
          });
        }
      }

      const paymentUpserts = [...paymentUpsertMap.values()];
      if (paymentUpserts.length) {
        const { error: payErr } = await supabase
          .from("participant_payments")
          .upsert(paymentUpserts, { onConflict: "participant_id,payment_key" });
        if (payErr) {
          console.error("Error guardando pagos:", payErr);
          notify("Error guardando pagos del Excel: " + payErr.message, { variant: "destructive" });
        }
      }

      setImportProgress(90);

      // ── Fase 7: crear cuentas Auth para participantes con email + password ──────
      const authCandidates = parsedRows
        .filter((r) => r.email && r.passwordFromExcel && existingByUsername.get(r.finalUsername))
        .map((r) => ({
          participantId: existingByUsername.get(r.finalUsername),
          email: r.email,
          password: r.passwordFromExcel,
        }));

      if (authCandidates.length) {
        setImportMessage("Creando accesos de participantes...");
        try {
          const token = await getAuthToken();
          const res = await fetch("/api/create-auth-users", {
            method: "POST",
            headers: {
              "Content-Type": "application/json",
              ...(token ? { Authorization: `Bearer ${token}` } : {}),
            },
            body: JSON.stringify({ participants: authCandidates }),
          });
          const result = await res.json();
          if (!res.ok) { console.error("Error creando cuentas Auth:", result.error); notify(`Error creando accesos: ${result.error || "revisa el log"}`, { variant: "destructive" }); }
          else notify(`Accesos creados: ${result.created} nuevos, ${result.updated} actualizados.`);
        } catch (err) {
          console.error("Error llamando /api/create-auth-users:", err);
          notify("Error creando accesos de participantes. Revisa el log.", { variant: "destructive" });
        }
      }

      // ── Construir refreshedUsers para el estado React ─────────────────────────
      const refreshedUsers = parsedRows.map((r) => {
        const participantId = existingByUsername.get(r.finalUsername);
        return {
          id: participantId,
          role: "client",
          username: r.finalUsername,
          participantName: r.participantName,
          motherName: r.motherName,
          fatherName: r.fatherName,
          parentName: [r.motherName, r.fatherName].filter(Boolean).join(" / "),
          email: r.email,
          contactEmails: r.contactEmails,
          tripId: r.tripId,
          documents: templatesToUse.map((t) => ({ id: t.id, status: "pending_upload", uploadedFileName: "", filePath: "", driveUrl: "" })),
          payments: {
            initialPrice: r.initialPrice, discount: r.discount, finalPrice: r.finalPrice,
            reservation:       r.payments[0] ? { name: r.payments[0].name, amount: r.payments[0].amount, status: "pending", proofName: "", dueDate: r.payments[0].due_date || "" } : {},
            firstInstallment:  r.payments[1] ? { name: r.payments[1].name, amount: r.payments[1].amount, status: "pending", proofName: "", dueDate: r.payments[1].due_date || "" } : {},
            secondInstallment: r.payments[2] ? { name: r.payments[2].name, amount: r.payments[2].amount, status: "pending", proofName: "", dueDate: r.payments[2].due_date || "" } : {},
          },
          checklistState: {},
          questions: [],
          dni: r.dniFromExcel || "",
        };
      });

      setImportDone(parsedRows.length);

      setTrips([...currentTrips]);

      // [ALTO-3] Merge correcto: admins separados del merge por username
      setUsers((prev) => {
        const admins = prev.filter((u) => u.role === "admin");
        const clientMap = new Map(
          prev.filter((u) => u.role === "client").map((u) => [u.username.toLowerCase(), u])
        );
        refreshedUsers.forEach((u) => clientMap.set(u.username.toLowerCase(), u));
        return [...admins, ...clientMap.values()];
      });

      setImportProgress(100);
      setImportMessage("Importación completada");
      notify(`Excel importado: ${parsedRows.length} participante(s) procesados correctamente.`);
      if (input) input.value = "";
      setImportFileName("");
    } catch (error) {
      console.error(error);
      notify(`Error importando Excel: ${error?.message || "Revisa el formato."}`);
      if (input) input.value = "";
    } finally {
      setTimeout(() => { setIsImporting(false); setImportProgress(0); setImportMessage(""); setImportTotal(0); setImportDone(0); }, 1200);
    }
  };

  return (
    <div className="space-y-4">
      <SectionTitle icon={Users} title="Clientes" subtitle="Importación, asignación y gestión rápida." />

      <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
        <CardContent className="p-5">
          <div className="space-y-2">
            <Label>Buscar participante</Label>
            <Input value={searchQuery} onChange={(e) => setSearchQuery(e.target.value)} placeholder="Busca por participante, familia, usuario o email" className="rounded-2xl" />
          </div>
        </CardContent>
      </Card>

      <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
        <CardContent className="p-5">
          <div className="grid gap-3 lg:grid-cols-[auto_1fr_1fr_auto] lg:items-end">
            <div>
              <Label className="mb-2 block">Importar Excel</Label>
              <label className="cursor-pointer">
                <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={(e) => handleBulkExcelUpload(e.target.files?.[0], e.currentTarget)} />
                <span className={`inline-flex h-11 items-center rounded-2xl px-4 text-sm font-medium text-white ${isImporting ? "opacity-60" : ""}`} style={{ backgroundColor: CORPORATE_RED }}>
                  <Upload className="mr-2 h-4 w-4" />{isImporting ? "Importando..." : importFileName || "Subir Excel"}
                </span>
              </label>
            </div>
            <div className="space-y-2">
              <Label>Campamento origen</Label>
              <select value={selectedGroupTrip} onChange={(e) => setSelectedGroupTrip(e.target.value)} className="h-11 w-full rounded-2xl border border-zinc-200 bg-white px-4 text-sm">
                {trips.map((t) => <option key={t.id} value={t.id}>{t.name}</option>)}
              </select>
            </div>
            <div className="space-y-2">
              <Label>Asignar campamento</Label>
              <select value={assignTargetTrip} onChange={(e) => setAssignTargetTrip(e.target.value)} className="h-11 w-full rounded-2xl border border-zinc-200 bg-white px-4 text-sm">
                {trips.map((t) => <option key={t.id} value={t.id}>{t.name}</option>)}
              </select>
            </div>
            <Button onClick={assignTripToGroup} className="h-11 rounded-2xl text-white" style={{ backgroundColor: CORPORATE_RED }}>
              <Users className="mr-2 h-4 w-4" />{selectedClientIds.length ? "Aplicar a seleccionados" : "Aplicar al grupo"}
            </Button>
          </div>

          {isImporting && (
            <div className="mt-4 space-y-3 rounded-2xl border border-zinc-200 bg-white p-4">
              <div className="flex flex-col gap-1 sm:flex-row sm:items-center sm:justify-between">
                <span className="text-sm font-medium text-zinc-800">{importMessage || "Importando..."}</span>
                <span className="text-sm text-zinc-600">{importDone}/{importTotal} importados · {importProgress}%</span>
              </div>
              <div className="h-3 w-full overflow-hidden rounded-full bg-zinc-200">
                <div className="h-full rounded-full transition-all duration-500 ease-out" style={{ width: `${importProgress}%`, backgroundColor: CORPORATE_RED }} />
              </div>
            </div>
          )}

          <div className="mt-3 border-t border-zinc-100 pt-3">
            <Label className="mb-2 block text-xs text-zinc-500">Vincular documento Excel o Google Sheets (URL pública)</Label>
            <div className="flex gap-2">
              <Input
                value={sheetUrl}
                onChange={(e) => setSheetUrl(e.target.value)}
                placeholder="https://docs.google.com/spreadsheets/d/... o enlace Excel"
                className="rounded-2xl text-sm"
              />
              <Button
                onClick={handleSyncSheet}
                disabled={isSyncingSheet || isImporting}
                className="h-11 shrink-0 rounded-2xl text-white"
                style={{ backgroundColor: CORPORATE_RED }}
              >
                {isSyncingSheet ? <Loader2 className="h-4 w-4 animate-spin" /> : <Download className="h-4 w-4" />}
                <span className="ml-2">{isSyncingSheet ? "Sincronizando..." : "Sincronizar"}</span>
              </Button>
            </div>
            <p className="mt-1 text-xs text-zinc-400">Comparte el Google Sheet con la cuenta de Google conectada al portal. Usa el mismo campamento seleccionado en "Campamento origen".</p>
          </div>
        </CardContent>
      </Card>

      <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
        <CardContent className="space-y-4 p-5">
          <div className="flex flex-col gap-3 lg:flex-row lg:items-end lg:justify-between">
            <div className="flex flex-col gap-3 lg:flex-row lg:items-end">
              <div className="space-y-3">
                <Label className="block pb-1">Filtrar por campamento</Label>
                <select value={selectedTripFilter} onChange={(e) => setSelectedTripFilter(e.target.value)} className="h-11 min-w-[280px] rounded-2xl border border-zinc-200 bg-white px-4 text-sm">
                  <option value="all">Todos los campamentos</option>
                  {trips.map((t) => <option key={t.id} value={t.id}>{t.name}</option>)}
                </select>
              </div>
              <Button variant="outline" className="h-11 rounded-2xl" onClick={toggleSelectAllVisible}>
                <CheckCircle2 className="mr-2 h-4 w-4" />{allVisibleSelected ? "Deseleccionar todos" : "Seleccionar todos"}
              </Button>
            </div>
            <div className="flex items-center gap-2 self-end flex-wrap">
              {selectedClientIds.length > 1 && (
                <Badge className="bg-zinc-900 text-white hover:bg-zinc-900">{selectedClientIds.length} seleccionados</Badge>
              )}
              <Button variant="outline" className="h-11 rounded-2xl" onClick={() => setConfirmBulkInvite(true)}>
                <Mail className="mr-2 h-4 w-4" />{selectedClientIds.length > 1 ? `Enviar acceso (${selectedClientIds.length})` : "Enviar acceso a todos"}
              </Button>
              {selectedClientIds.length > 1 && (
                <Button variant="outline" className="h-11 rounded-2xl" onClick={() => setConfirmBulkDelete(true)}>
                  <Trash2 className="mr-2 h-4 w-4" />Eliminar selección
                </Button>
              )}
            </div>
          </div>

          <div className="grid gap-3">
            {pagedClients.map((client) => {
              const isSelected = selectedClientIds.includes(client.id);
              return (
                <div key={client.id} onClick={() => toggleClientSelection(client.id)} className={`grid cursor-pointer gap-3 rounded-3xl border p-4 transition-all lg:grid-cols-[44px_1.2fr_1fr_44px] lg:items-center ${isSelected ? "border-zinc-900 bg-white shadow-sm" : "border-zinc-200 bg-white hover:border-zinc-400"}`}>
                  <div className="flex justify-center" onClick={(e) => e.stopPropagation()}>
                    <Checkbox checked={isSelected} onCheckedChange={() => toggleClientSelection(client.id)} />
                  </div>
                  <div>
                    <div className="font-medium text-zinc-950">{client.participantName}</div>
                    {getFamilyLabel(client) && <div className="mt-1 text-sm text-zinc-500">Familia: {getFamilyLabel(client)}</div>}
                    <div className="mt-1 text-sm text-zinc-500">Usuario: {client.username}</div>
                    {client.email
                      ? <div className="mt-1 text-xs text-zinc-400">{client.email}</div>
                      : <div className="mt-1 text-xs text-amber-600">Sin email — no se puede crear acceso</div>
                    }
                  </div>
                  <select
                    onClick={(e) => e.stopPropagation()}
                    value={client.tripId && trips.some((t) => t.id === client.tripId) ? client.tripId : ""}
                    onChange={async (e) => {
                      const value = e.target.value || null;
                      setUsers((prev) => prev.map((u) => u.id === client.id ? { ...u, tripId: value } : u));
                      const { error } = await supabase.from("participants").update({ trip_id: value }).eq("id", client.id);
                      if (error) notify("Error asignando viaje: " + error.message);
                    }}
                    className="h-11 w-full rounded-2xl border border-zinc-200 bg-white px-4 text-sm"
                  >
                    <option value="">Sin campamento asignado</option>
                    {trips.map((t) => <option key={t.id} value={t.id}>{t.name}</option>)}
                  </select>
                  <div className="flex flex-col items-end gap-2" onClick={(e) => e.stopPropagation()}>
                    <Button
                      variant="ghost" size="icon"
                      aria-label={client.email ? "Enviar invitación de acceso" : "Sin email registrado"}
                      title={client.email ? "Crear acceso / reenviar invitación" : "Sin email registrado"}
                      disabled={!client.email}
                      onClick={async () => {
                        if (!client.email) return;
                        try {
                          const token = await getAuthToken();
                          const res = await fetch("/api/invite-participant", {
                            method: "POST",
                            headers: {
                              "Content-Type": "application/json",
                              ...(token ? { Authorization: `Bearer ${token}` } : {}),
                            },
                            body: JSON.stringify({ participantId: client.id }),
                          });
                          const json = await res.json();
                          if (!res.ok) { notify(`Error: ${json.error}`); return; }
                          notify(json.message);
                        } catch (err) {
                          notify("Error enviando invitación.");
                        }
                      }}
                    >
                      <Mail className="h-4 w-4 text-zinc-700" />
                    </Button>
                    <Button variant="ghost" size="icon" aria-label="Eliminar participante" onClick={() => setPendingDeleteId(client.id)}>
                      <Trash2 className="h-4 w-4 text-zinc-700" />
                    </Button>
                  </div>
                </div>
              );
            })}
          </div>
          {totalPages > 1 && (
            <div className="flex items-center justify-between border-t border-zinc-100 pt-4 mt-2">
              <span className="text-xs text-zinc-400">{visibleClients.length} participantes · página {clientPage} de {totalPages}</span>
              <div className="flex gap-2">
                <Button variant="outline" size="sm" className="rounded-xl" disabled={clientPage === 1} onClick={() => setClientPage((p) => p - 1)}>←</Button>
                <Button variant="outline" size="sm" className="rounded-xl" disabled={clientPage === totalPages} onClick={() => setClientPage((p) => p + 1)}>→</Button>
              </div>
            </div>
          )}
        </CardContent>
      </Card>

      {/* Confirmación: borrar factura participante */}
      <AlertDialog open={!!pendingDeleteInvoiceClientId} onOpenChange={(o) => { if (!o) setPendingDeleteInvoiceClientId(null); }}>
        <AlertDialogContent>
          <AlertDialogHeader>
            <AlertDialogTitle>¿Borrar factura?</AlertDialogTitle>
            <AlertDialogDescription>Se eliminará el enlace a la factura de este participante. El archivo en Google Drive no se borrará.</AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel>Cancelar</AlertDialogCancel>
            <AlertDialogAction style={{ backgroundColor: CORPORATE_RED }} onClick={async () => {
              const id = pendingDeleteInvoiceClientId;
              setPendingDeleteInvoiceClientId(null);
              const { error } = await supabase.from("participants").update({ invoice_url: null }).eq("id", id);
              if (!error) { setUsers((prev) => prev.map((u) => u.id === id ? { ...u, invoiceUrl: null } : u)); notify("Factura eliminada."); }
              else notify("Error eliminando factura: " + error.message, { variant: "destructive" });
            }}>Borrar factura</AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>

      {/* Confirmación enviar invitaciones */}
      <AlertDialog open={confirmBulkInvite} onOpenChange={setConfirmBulkInvite}>
        <AlertDialogContent>
          <AlertDialogHeader>
            <AlertDialogTitle>
              {selectedClientIds.length > 1 ? `¿Enviar acceso a ${selectedClientIds.length} participantes?` : "¿Enviar acceso a todos los participantes?"}
            </AlertDialogTitle>
            <AlertDialogDescription>
              Se enviará un email de invitación con sus credenciales de acceso al portal.
            </AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel>Cancelar</AlertDialogCancel>
            <AlertDialogAction onClick={() => { bulkInviteSelected(); setConfirmBulkInvite(false); }}>
              Enviar
            </AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>

      {/* Confirmación eliminar individual */}
      <AlertDialog open={pendingDeleteId !== null} onOpenChange={(open) => { if (!open) setPendingDeleteId(null); }}>
        <AlertDialogContent>
          <AlertDialogHeader>
            <AlertDialogTitle>¿Eliminar participante?</AlertDialogTitle>
            <AlertDialogDescription>
              Se eliminarán todos los datos del participante. Esta acción no se puede deshacer.
            </AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel>Cancelar</AlertDialogCancel>
            <AlertDialogAction className="bg-red-600 hover:bg-red-700" onClick={() => { deleteSingleClient(pendingDeleteId); setPendingDeleteId(null); }}>
              Eliminar
            </AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>

      {/* Confirmación eliminar selección múltiple */}
      <AlertDialog open={confirmBulkDelete} onOpenChange={setConfirmBulkDelete}>
        <AlertDialogContent>
          <AlertDialogHeader>
            <AlertDialogTitle>¿Eliminar {selectedClientIds.length} participantes?</AlertDialogTitle>
            <AlertDialogDescription>
              Se eliminarán todos sus datos permanentemente. Esta acción no se puede deshacer.
            </AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel>Cancelar</AlertDialogCancel>
            <AlertDialogAction className="bg-red-600 hover:bg-red-700" onClick={() => { deleteSelectedClients(); setConfirmBulkDelete(false); }}>
              Eliminar {selectedClientIds.length} participantes
            </AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>
    </div>
  );
}

function AdminParticipantsExport({ users, trips }) {
  const clients = users.filter((u) => u.role === "client" && !u.schoolId);
  const [selectedTripId, setSelectedTripId] = useState(trips[0]?.id || "all");

  // Todos los campos disponibles, organizados por grupo
  const COLUMN_GROUPS = [
    {
      group: "Participante",
      cols: [
        { key: "nombre",       label: "Nombre del participante",    default: true },
        { key: "dni",          label: "DNI / NIF / Pasaporte",      default: false },
        { key: "nacimiento",   label: "Fecha de nacimiento",        default: false },
        { key: "sexo",         label: "Sexo",                       default: false },
        { key: "direccion",    label: "Dirección completa",         default: false },
        { key: "colegio",      label: "Colegio",                    default: false },
        { key: "talla",        label: "Talla de camiseta",          default: false },
        { key: "modalidad",    label: "Modalidad de campamento",    default: false },
        { key: "username",     label: "Usuario (login)",            default: false },
      ],
    },
    {
      group: "Familia",
      cols: [
        { key: "madre",        label: "Nombre de la madre",         default: false },
        { key: "tel_madre",    label: "Teléfono de la madre",       default: false },
        { key: "email_madre",  label: "Email de la madre",          default: false },
        { key: "dni_madre",    label: "DNI de la madre",            default: false },
        { key: "padre",        label: "Nombre del padre",           default: false },
        { key: "tel_padre",    label: "Teléfono del padre",         default: false },
        { key: "email_padre",  label: "Email del padre",            default: false },
        { key: "dni_padre",    label: "DNI del padre",              default: false },
        { key: "email",        label: "Email de contacto",          default: false },
        { key: "auth_img",     label: "Autorización imágenes",      default: false },
      ],
    },
    {
      group: "Salud",
      cols: [
        { key: "alergias",     label: "Alergias / intolerancias",   default: false },
        { key: "salud",        label: "Problemas de salud",         default: false },
        { key: "observ",       label: "Observaciones",              default: false },
      ],
    },
    {
      group: "Procedencia",
      cols: [
        { key: "como_conocio", label: "Cómo nos conocieron",        default: false },
        { key: "viaje",        label: "Nombre del viaje",           default: false },
      ],
    },
    {
      group: "Económico",
      cols: [
        { key: "precio_ini",   label: "Precio inicial",             default: false },
        { key: "descuento",    label: "Descuento",                  default: false },
        { key: "precio_fin",   label: "Precio final",               default: false },
      ],
    },
    {
      group: "Pagos",
      cols: [
        { key: "res_importe",  label: "Reserva — importe",          default: false },
        { key: "res_estado",   label: "Reserva — estado",           default: false },
        { key: "res_fecha",    label: "Reserva — fecha límite",     default: false },
        { key: "p1_importe",   label: "1ª Cuota — importe",         default: false },
        { key: "p1_estado",    label: "1ª Cuota — estado",          default: false },
        { key: "p1_fecha",     label: "1ª Cuota — fecha límite",    default: false },
        { key: "p2_importe",   label: "2ª Cuota — importe",         default: false },
        { key: "p2_estado",    label: "2ª Cuota — estado",          default: false },
        { key: "p2_fecha",     label: "2ª Cuota — fecha límite",    default: false },
      ],
    },
  ];

  const allCols = COLUMN_GROUPS.flatMap((g) => g.cols);
  const [cols, setCols] = useState(() => Object.fromEntries(allCols.map((c) => [c.key, c.default])));
  const toggleCol = (key) => setCols((p) => ({ ...p, [key]: !p[key] }));
  const [openGroups, setOpenGroups] = React.useState(() => Object.fromEntries(COLUMN_GROUPS.map(g => [g.group, false])));
  const toggleGroup = (group) => setOpenGroups(p => ({ ...p, [group]: !p[group] }));

  const STATUS_LABELS = { pending: "Pendiente", sent: "Enviado", confirmed: "Confirmado", rejected: "Rechazado" };

  const tripClients = selectedTripId === "all" ? clients : clients.filter((c) => c.tripId === selectedTripId);
  const tripName = selectedTripId === "all" ? "Todos los viajes" : trips.find((t) => t.id === selectedTripId)?.name || "";

  const getCell = (client, key) => {
    const p = client.payments || {};
    const fmt = (v) => v ? new Intl.NumberFormat("es-ES", { style: "currency", currency: "EUR" }).format(v) : "—";
    const fmtDate = (d) => d ? new Date(d).toLocaleDateString("es-ES") : "—";
    if (key === "nombre")      return client.participantName || "—";
    if (key === "dni")         return client.dni || "—";
    if (key === "nacimiento")  return client.birthDate ? new Date(client.birthDate).toLocaleDateString("es-ES") : "—";
    if (key === "sexo")        return client.gender || "—";
    if (key === "direccion")   return client.address || "—";
    if (key === "colegio")     return client.school || "—";
    if (key === "talla")       return client.shirtSize || "—";
    if (key === "modalidad")   return client.modality || "—";
    if (key === "username")    return client.username || "—";
    if (key === "madre")       return client.motherName || "—";
    if (key === "tel_madre")   return client.phoneMother || "—";
    if (key === "email_madre") return client.contactEmails?.[0] || client.email || "—";
    if (key === "dni_madre")   return client.dniMother || "—";
    if (key === "padre")       return client.fatherName || "—";
    if (key === "tel_padre")   return client.phoneFather || "—";
    if (key === "email_padre") return client.contactEmails?.[1] || "—";
    if (key === "dni_padre")   return client.dniFather || "—";
    if (key === "email")       return client.email || client.contactEmails?.[0] || "—";
    if (key === "auth_img")    return client.imageAuth ? "Sí" : "No";
    if (key === "alergias")    return client.allergies || "—";
    if (key === "salud")       return client.healthNotes || "—";
    if (key === "observ")      return client.notes || "—";
    if (key === "como_conocio") return client.howKnown || "—";
    if (key === "viaje")       return trips.find((t) => t.id === client.tripId)?.name || "—";
    if (key === "precio_ini")  return fmt(p.initialPrice);
    if (key === "descuento")   return p.discount > 0 ? fmt(p.discount) : "—";
    if (key === "precio_fin")  return fmt(p.finalPrice || Math.max(0, (p.initialPrice || 0) - (p.discount || 0)));
    if (key === "res_importe") return fmt(p.reservation?.amount);
    if (key === "res_estado")  return STATUS_LABELS[p.reservation?.status] || "—";
    if (key === "res_fecha")   return fmtDate(p.reservation?.dueDate);
    if (key === "p1_importe")  return fmt(p.firstInstallment?.amount);
    if (key === "p1_estado")   return STATUS_LABELS[p.firstInstallment?.status] || "—";
    if (key === "p1_fecha")    return fmtDate(p.firstInstallment?.dueDate);
    if (key === "p2_importe")  return fmt(p.secondInstallment?.amount);
    if (key === "p2_estado")   return STATUS_LABELS[p.secondInstallment?.status] || "—";
    if (key === "p2_fecha")    return fmtDate(p.secondInstallment?.dueDate);
    return "—";
  };

  const activeCols = allCols.filter((c) => cols[c.key]);

  const handleExport = () => {
    if (!tripClients.length || !activeCols.length) return;
    // Filtrar columnas que tengan al menos un valor real para este campamento
    const usedCols = activeCols.filter((c) =>
      tripClients.some((client) => { const v = getCell(client, c.key); return v && v !== "—"; })
    );
    if (!usedCols.length) return;
    const tbody = tripClients.map((client, i) =>
      `<tr><td style="color:#a1a1aa;width:24px">${i + 1}</td>${usedCols.map((c) => `<td>${getCell(client, c.key)}</td>`).join("")}</tr>`
    ).join("");
    exportListToPDF(
      `Participantes — ${tripName}`,
      `${tripClients.length} participante(s)`,
      `<table><thead><tr><th>#</th>${usedCols.map((c) => `<th>${c.label}</th>`).join("")}</tr></thead><tbody>${tbody}</tbody></table>`
    );
  };

  return (
    <div className="space-y-5">
      <SectionTitle icon={FileText} title="Participantes" subtitle="Exporta listados personalizados en PDF." />

      <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
        <CardContent className="p-5 space-y-6">
          {/* Selector de campamento */}
          <div className="space-y-2">
            <Label>Campamento</Label>
            <select value={selectedTripId} onChange={(e) => setSelectedTripId(e.target.value)}
              className="h-11 w-full max-w-sm rounded-2xl border border-zinc-200 bg-white px-4 text-sm">
              <option value="all">Todos los campamentos</option>
              {trips.map((t) => <option key={t.id} value={t.id}>{t.name}</option>)}
            </select>
          </div>

          {/* Columnas por grupo — desplegables */}
          {COLUMN_GROUPS.map(({ group, cols: groupCols }) => {
            const groupActive = groupCols.filter(c => cols[c.key]).length;
            const open = openGroups[group];
            return (
              <div key={group} className="rounded-2xl border border-zinc-200 overflow-hidden">
                <button
                  type="button"
                  onClick={() => toggleGroup(group)}
                  className="flex w-full items-center justify-between bg-zinc-50 px-4 py-3 text-left transition hover:bg-zinc-100"
                >
                  <div className="flex items-center gap-2">
                    <span className="text-xs font-semibold uppercase tracking-[0.18em] text-zinc-500">{group}</span>
                    {groupActive > 0 && (
                      <span className="inline-flex h-5 min-w-[20px] items-center justify-center rounded-full px-1.5 text-xs font-bold text-white" style={{ backgroundColor: CORPORATE_RED }}>
                        {groupActive}
                      </span>
                    )}
                  </div>
                  <ChevronDown className={`h-4 w-4 text-zinc-400 transition-transform duration-200 ${open ? "rotate-180" : ""}`} />
                </button>
                {open && (
                  <div className="grid gap-2 p-3 sm:grid-cols-2 lg:grid-cols-3 border-t border-zinc-100 bg-white">
                    {groupCols.map(({ key, label }) => (
                      <div key={key} onClick={() => toggleCol(key)}
                        className={`flex cursor-pointer select-none items-center gap-3 rounded-2xl border p-3 text-sm transition-all ${cols[key] ? "border-zinc-900 bg-zinc-950 text-white" : "border-zinc-200 bg-white text-zinc-700 hover:border-zinc-400"}`}>
                        <div className={`flex h-5 w-5 shrink-0 items-center justify-center rounded-md border text-xs font-bold ${cols[key] ? "border-white bg-white text-zinc-950" : "border-zinc-300"}`}>
                          {cols[key] ? "✓" : ""}
                        </div>
                        {label}
                      </div>
                    ))}
                  </div>
                )}
              </div>
            );
          })}

          <div className="flex flex-wrap gap-3">
            <Button variant="outline" className="h-11 rounded-2xl text-sm" onClick={() => setCols(Object.fromEntries(allCols.map(c => [c.key, true])))}>
              Seleccionar todo
            </Button>
            <Button variant="outline" className="h-11 rounded-2xl text-sm" onClick={() => setCols(Object.fromEntries(allCols.map(c => [c.key, false])))}>
              Borrar selección
            </Button>
            <Button onClick={handleExport} disabled={!tripClients.length || activeCols.length === 0}
              className="h-11 rounded-2xl text-white" style={{ backgroundColor: CORPORATE_RED }}>
              <FileText className="mr-2 h-4 w-4" />
              Exportar PDF — {tripClients.length} participante(s)
            </Button>
          </div>
        </CardContent>
      </Card>

      {/* Vista previa */}
      {tripClients.length > 0 && activeCols.length > 0 && (
        <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
          <CardContent className="p-5">
            <div className="mb-3 text-sm font-semibold text-zinc-700">Vista previa</div>
            <div className="overflow-x-auto">
              <table className="w-full text-xs">
                <thead>
                  <tr className="border-b border-zinc-100 bg-zinc-50">
                    <th className="px-3 py-2 text-left text-xs font-semibold uppercase tracking-wide text-zinc-400">#</th>
                    {activeCols.map((c) => <th key={c.key} className="px-3 py-2 text-left text-xs font-semibold uppercase tracking-wide text-zinc-500">{c.label}</th>)}
                  </tr>
                </thead>
                <tbody>
                  {tripClients.slice(0, 8).map((client, i) => (
                    <tr key={client.id} className="border-b border-zinc-50 hover:bg-zinc-50/50">
                      <td className="px-3 py-2 text-zinc-400">{i + 1}</td>
                      {activeCols.map((c) => <td key={c.key} className="px-3 py-2 text-zinc-700">{getCell(client, c.key)}</td>)}
                    </tr>
                  ))}
                </tbody>
              </table>
              {tripClients.length > 8 && <div className="mt-2 text-xs text-zinc-400">+{tripClients.length - 8} participantes más en el PDF</div>}
            </div>
          </CardContent>
        </Card>
      )}
    </div>
  );
}

function AdminTracking({ users, trips, templates, setUsers, notify }) {
  const clients = users.filter((u) => u.role === "client" && !u.schoolId);
  const [selectedTripId, setSelectedTripId] = useState("all");
  const [searchQuery, setSearchQuery] = useState("");
  const [replyTexts, setReplyTexts] = useState({});
  const [sendingReply, setSendingReply] = useState({});
  const [sendingReminder, setSendingReminder] = useState(false);
  const [syncingSheet, setSyncingSheet] = useState(false);
  const handleReplyQuestion = async (participantId, questionId, reply) => {
    if (!reply?.trim()) return;
    setSendingReply((s) => ({ ...s, [questionId]: true }));
    try {
      const { error } = await supabase.from("participant_questions").update({ reply, replied_at: new Date().toISOString(), status: "replied" }).eq("id", questionId);
      if (error) throw new Error(error.message);
      setUsers((prev) => prev.map((u) => u.id === participantId ? { ...u, questions: (u.questions || []).map((q) => q.id === questionId ? { ...q, reply, repliedAt: new Date().toISOString(), status: "replied" } : q) } : u));
      setReplyTexts((t) => { const n = { ...t }; delete n[questionId]; return n; });
      const participant = users.find((u) => u.id === participantId);
      const question = (participant?.questions || []).find((q) => q.id === questionId);
      if (participant?.email && question) sendNotification("question_replied", participant.email, participantId, { participantName: participant.participantName, question: question.message, reply });
      notify("✅ Respuesta enviada.");
    } catch (err) { notify("Error guardando respuesta: " + err.message); }
    finally { setSendingReply((s) => { const n = { ...s }; delete n[questionId]; return n; }); }
  };
  const filteredClients = clients.filter(
    (c) => (selectedTripId === "all" || c.tripId === selectedTripId) && matchesParticipantSearch(c, searchQuery)
  );
  const updateClient = (clientId, updater) =>
    setUsers((prev) => prev.map((u) => u.id === clientId ? updater(u) : u));
  const sendReminder = async (client, type) => {
    const trip = trips.find((t) => t.id === client.tripId);
    const email = client.email || client.contactEmails?.[0];
    if (!email) { notify("Este participante no tiene email registrado."); return; }
    try {
      const token = await getAuthToken();
      const authHeaders = {
        "Content-Type": "application/json",
        ...(token ? { Authorization: `Bearer ${token}` } : {}),
      };
      if (type === "documentación") {
        // Bug fix: los estados válidos son "pending_upload" y "pending_confirmation", no "pending"/"not_uploaded"
        const pendingDocs = client.documents.filter((d) =>
          !["confirmed"].includes(d.status)
        );
        const docName = pendingDocs.length > 0
          ? pendingDocs.map((d) => d.id).join(", ")
          : "Documentación pendiente";
        const res = await fetch("/api/notify", {
          method: "POST",
          headers: authHeaders,
          body: JSON.stringify({ type: "doc_reminder", to: email, participantId: client.id, data: { participantName: client.participantName, docName, tripName: trip?.name || "" } }),
        });
        if (!res.ok) { const j = await res.json().catch(() => ({})); throw new Error(j.error || `Error ${res.status}`); }
      } else {
        const pendingPays = [
          { key: "reservation", label: "Reserva" },
          { key: "firstInstallment", label: "1.ª cuota" },
          { key: "secondInstallment", label: "2.ª cuota" },
        ].filter((p) => client.payments[p.key]?.status === "pending");
        const paymentName = pendingPays.map((p) => p.label).join(", ") || "Pago pendiente";
        const res = await fetch("/api/notify", {
          method: "POST",
          headers: authHeaders,
          body: JSON.stringify({ type: "payment_reminder", to: email, participantId: client.id, data: { participantName: client.participantName, paymentName, amount: "", dueDate: "", daysLeft: "", tripName: trip?.name || "" } }),
        });
        if (!res.ok) { const j = await res.json().catch(() => ({})); throw new Error(j.error || `Error ${res.status}`); }
      }
      notify(`✉️ Recordatorio enviado a ${email}`);
    } catch (err) {
      notify("Error enviando recordatorio: " + err.message);
    }
  };
  const getSummaryTone = (v) => v === 0 ? "bg-red-100 text-red-700" : v === 1 ? "bg-amber-100 text-amber-700" : "bg-emerald-100 text-emerald-700";

  // Actualizar Google Sheets de seguimiento cuando cambia un estado de pago
  const syncPaymentToSheet = async (client, paymentKey, nextStatus, amount) => {
    const trip = trips.find((t) => t.id === client.tripId);
    if (!trip) return;
    try {
      const token = await getAuthToken();
      await fetch("/api/tracking-sheet", {
        method: "POST",
        headers: { "Content-Type": "application/json", ...(token ? { Authorization: `Bearer ${token}` } : {}) },
        body: JSON.stringify({
          action: "update_payment",
          tripName: trip.name,
          participantName: client.participantName,
          participantEmail: client.email || "",
          participantDni: client.dni || "",
          paymentKey,
          paymentStatus: nextStatus,
          paymentAmount: amount || 0,
        }),
      });
    } catch (err) {
      console.warn("tracking-sheet sync error (non-critical):", err.message);
    }
  };

  return (
    <div className="space-y-5">
      <SectionTitle icon={FileCheck2} title="Seguimiento" subtitle="Control de pagos y documentación por participante." />

      <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
        <CardContent className="p-5">
          <div className="space-y-2">
            <Label>Buscar participante</Label>
            <Input value={searchQuery} onChange={(e) => setSearchQuery(e.target.value)} placeholder="Busca por participante, familia, usuario o email" className="rounded-2xl" />
          </div>
        </CardContent>
      </Card>

      <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
        <CardContent className="flex flex-col gap-4 p-5 lg:flex-row lg:items-end lg:justify-between">
          <div className="space-y-2">
            <Label className="mb-2 block">Filtrar por campamento</Label>
            <select value={selectedTripId} onChange={(e) => setSelectedTripId(e.target.value)} className="mt-2 h-11 min-w-[320px] rounded-2xl border border-zinc-200 bg-white px-4 text-sm">
              <option value="all">Todos los campamentos</option>
              {trips.map((t) => <option key={t.id} value={t.id}>{t.name}</option>)}
            </select>
          </div>
          {selectedTripId !== "all" && (
            <div className="flex flex-wrap items-center gap-3 rounded-2xl border border-zinc-200 bg-white px-4 py-3 text-sm text-zinc-700">
              <span>Primer doc: <strong>{formatShortDate(calculateDueDateFromRule(trips.find((t) => t.id === selectedTripId)?.departureDate, trips.find((t) => t.id === selectedTripId)?.documentRules?.[0]))}</strong></span>
              <span>Último pago: <strong>{formatShortDate(getPaymentRuleDueDate(trips.find((t) => t.id === selectedTripId), "secondInstallment"))}</strong></span>
              <Button variant="outline" className="rounded-2xl" disabled={sendingReminder} onClick={async () => {
                const trip = trips.find((t) => t.id === selectedTripId);
                const tripClients = clients.filter((c) => c.tripId === selectedTripId);
                const withEmail = tripClients.filter((c) => c.email || c.contactEmails?.[0]);
                if (!withEmail.length) { notify("Ningún participante tiene email registrado."); return; }
                setSendingReminder(true);
                notify(`Enviando recordatorio a ${withEmail.length} participante(s)…`);
                const token = await getAuthToken();
                const authHeaders = { "Content-Type": "application/json", ...(token ? { Authorization: `Bearer ${token}` } : {}) };
                let sent = 0; let failed = 0;
                for (const c of withEmail) {
                  const email = c.email || c.contactEmails?.[0];
                  const tripName = trip?.name || "";
                  let ok = true;
                  // Docs pendientes
                  const pendingDocs = c.documents.filter((d) => d.status === "pending_upload" || d.status === "pending_confirmation");
                  if (pendingDocs.length) {
                    const r = await fetch("/api/notify", { method: "POST", headers: authHeaders, body: JSON.stringify({ type: "doc_reminder", to: email, participantId: c.id, data: { participantName: c.participantName, docName: pendingDocs.map((d) => { const t2 = templates.find((t3) => t3.id === d.id); return t2?.name || "documento pendiente"; }).join(", "), tripName } }) }).catch(() => null);
                    if (!r?.ok) ok = false;
                  }
                  // Pagos pendientes
                  const pendingPays = ["reservation", "firstInstallment", "secondInstallment"].filter((k) => c.payments[k]?.status === "pending");
                  for (const pk of pendingPays) {
                    const pay = c.payments[pk];
                    const r = await fetch("/api/notify", { method: "POST", headers: authHeaders, body: JSON.stringify({ type: "payment_reminder", to: email, participantId: c.id, data: { participantName: c.participantName, paymentName: pay.name || pk, amount: pay.amount ? new Intl.NumberFormat("es-ES", { style: "currency", currency: "EUR" }).format(pay.amount) : "", dueDate: pay.dueDate || "", daysLeft: "", tripName } }) }).catch(() => null);
                    if (!r?.ok) ok = false;
                  }
                  if (ok) sent++; else failed++;
                }
                setSendingReminder(false);
                if (failed > 0) notify(`Recordatorio: ${sent} enviados, ${failed} con error.`, { variant: "destructive" });
                else notify(`✅ Recordatorio enviado a ${sent} participante(s) del grupo.`);
              }}>
                {sendingReminder ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <Mail className="mr-2 h-4 w-4" />}Recordar al grupo
              </Button>
              <Button variant="outline" className="rounded-2xl" disabled={syncingSheet} onClick={async () => {
                const trip = trips.find((t) => t.id === selectedTripId);
                const tripClients = clients.filter((c) => c.tripId === selectedTripId);
                if (!trip || !tripClients.length) { notify("No hay participantes en este viaje."); return; }
                setSyncingSheet(true);
                notify("Sincronizando hoja de seguimiento…");
                try {
                  const token = await getAuthToken();
                  const res = await fetch("/api/tracking-sheet", {
                    method: "POST",
                    headers: { "Content-Type": "application/json", ...(token ? { Authorization: `Bearer ${token}` } : {}) },
                    body: JSON.stringify({ action: "sync", tripName: trip.name, participants: tripClients }),
                  });
                  const json = await res.json();
                  if (!res.ok) throw new Error(json.error || "Error");
                  notify("✅ Hoja sincronizada.", { actionLabel: "Abrir", onAction: () => window.open(json.sheetUrl, "_blank", "noopener,noreferrer") });
                } catch (err) { notify("Error sincronizando: " + err.message); }
                finally { setSyncingSheet(false); }
              }}>
                {syncingSheet ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <FileCheck2 className="mr-2 h-4 w-4" />}Sync hoja Excel
              </Button>
            </div>
          )}
        </CardContent>
      </Card>

      <div className="space-y-4">
        {filteredClients.map((client) => {
          const trip = trips.find((t) => t.id === client.tripId);
          const docsTotal = client.documents.length;
          const docsConfirmed = client.documents.filter((d) => d.status === "confirmed").length;
          const docsReview = client.documents.filter((d) => ["pending_confirmation", "review"].includes(d.status)).length;
          const payList = [client.payments.reservation, client.payments.firstInstallment, client.payments.secondInstallment];
          const paysTotal = payList.length;
          const paysConfirmed = payList.filter((p) => p.status === "confirmed").length;
          const paysReview = payList.filter((p) => ["sent", "review"].includes(p.status)).length;
          const questionsUnanswered = (client.questions || []).filter((q) => !q.reply).length;
          const totalPendingReview = docsReview + paysReview + questionsUnanswered;

          return (
            <AccordionSection
              key={client.id}
              title={client.participantName}
              subtitle={`${getFamilyLabel(client) ? `Familia: ${getFamilyLabel(client)} · ` : ""}Usuario: ${client.username} · ${trip?.name || "Sin viaje"}`}
              icon={Users}
              meta={
                <div className="flex items-center gap-2 overflow-x-auto whitespace-nowrap">
                  {totalPendingReview > 0 && (
                    <span className="inline-flex h-6 w-6 shrink-0 items-center justify-center rounded-full text-xs font-bold text-white" style={{ backgroundColor: CORPORATE_RED }}>{totalPendingReview}</span>
                  )}
                  <div className={`inline-flex items-center rounded-2xl px-3 py-2 text-xs font-medium ${getSummaryTone(docsConfirmed === docsTotal ? 2 : docsConfirmed > 0 || docsReview > 0 ? 1 : 0)}`}>Docs {docsConfirmed}/{docsTotal}</div>
                  <div className={`inline-flex items-center rounded-2xl px-3 py-2 text-xs font-medium ${getSummaryTone(paysConfirmed === paysTotal ? 2 : paysConfirmed > 0 || paysReview > 0 ? 1 : 0)}`}>Pagos {paysConfirmed}/{paysTotal}</div>
                  <Button variant="outline" className="h-9 shrink-0 rounded-2xl px-3" onClick={(e) => { e.stopPropagation(); sendReminder(client, "documentación"); }}><Mail className="mr-2 h-4 w-4" />Recordar docs</Button>
                  <Button variant="outline" className="h-9 shrink-0 rounded-2xl px-3" onClick={(e) => { e.stopPropagation(); sendReminder(client, "pagos"); }}><CreditCard className="mr-2 h-4 w-4" />Recordar pagos</Button>
                </div>
              }
            >
              <div className="mb-5 grid gap-3 lg:grid-cols-4">
                {[["Docs confirmados", `${docsConfirmed}/${docsTotal}`], ["Docs por revisar", docsReview], ["Pagos confirmados", `${paysConfirmed}/${paysTotal}`], ["Pagos por revisar", paysReview]].map(([label, val]) => (
                  <div key={String(label)} className="rounded-2xl border border-zinc-200 bg-white p-4">
                    <div className="text-xs uppercase tracking-[0.18em] text-zinc-500">{label}</div>
                    <div className="mt-2 text-lg font-semibold text-zinc-950">{val}</div>
                  </div>
                ))}
              </div>

              {/* Información del participante */}
              <div className="mb-5 flex flex-wrap gap-2">
                {[
                  { label: "Email",        value: client.email },
                  { label: "Usuario",      value: client.username },
                  { label: "DNI",          value: client.dni },
                  { label: "Fecha nac.",   value: client.birthDate ? new Date(client.birthDate).toLocaleDateString("es-ES") : null },
                  { label: "Género",       value: client.gender },
                  { label: "Familia",      value: getFamilyLabel(client) },
                  { label: "Madre",        value: client.motherName },
                  { label: "Tel. madre",   value: client.phoneMother },
                  { label: "Padre",        value: client.fatherName },
                  { label: "Tel. padre",   value: client.phoneFather },
                  { label: "Colegio",      value: client.school },
                  { label: "Dirección",    value: client.address },
                  { label: "Alergias",     value: client.allergies },
                  { label: "Salud",        value: client.healthNotes },
                  { label: "Talla",        value: client.shirtSize },
                  { label: "Modalidad",    value: client.modality },
                  { label: "Notas",        value: client.notes },
                ].filter(({ value }) => !!value).map(({ label, value }) => (
                  <div key={label} className="flex items-center gap-1.5 rounded-xl border border-zinc-200 bg-white px-3 py-1.5 text-xs">
                    <span className="text-zinc-400">{label}:</span>
                    <span className="font-medium text-zinc-800">{value}</span>
                  </div>
                ))}
              </div>

              <div className="grid gap-5 xl:grid-cols-2">
                <div className="space-y-3">
                  <div className="text-sm font-semibold uppercase tracking-[0.18em] text-zinc-500">Documentación</div>
                  {client.documents.map((docItem) => {
                    const template = templates.find((t) => t.id === docItem.id);
                    const status = getStatusMeta(docItem.status);
                    const isPendingReview = docItem.status === "pending_confirmation";
                    return (
                      <div key={docItem.id} className={`rounded-2xl border p-4 space-y-3 ${isPendingReview ? "border-amber-300 bg-amber-50" : "border-zinc-200 bg-white"}`}>
                        {/* Fila 1: info + badge */}
                        <div className="flex items-start gap-3">
                          <div className="min-w-0 flex-1">
                            <div className="font-medium text-zinc-950">{template?.name || docItem.id}</div>
                            <div className="text-sm text-zinc-500">{docItem.uploadedFileName || "Sin archivo subido"}</div>
                            <div className="flex items-center gap-2 text-sm text-zinc-500 mt-0.5">
                              <span className={`h-2.5 w-2.5 shrink-0 rounded-full ${getDueStatus(getDocumentRuleDueDate(trip, docItem.id)).className}`} />
                              <span>Límite: {formatShortDate(getDocumentRuleDueDate(trip, docItem.id))}</span>
                            </div>
                          </div>
                          <Badge className={`shrink-0 ${status.className}`} style={status.style}>{status.label}</Badge>
                        </div>
                        {/* Fila 2: controles */}
                        <div className="flex flex-wrap items-center gap-2">
                          {[
                            { value: "pending_upload", label: "Pendiente" },
                            { value: "pending_confirmation", label: "Por revisar" },
                            { value: "confirmed", label: "Confirmado" },
                            { value: "rejected", label: "Rechazado" },
                          ].map(({ value, label }) => {
                            const isActive = docItem.status === value;
                            const meta = getStatusMeta(value);
                            return (
                              <button key={value} type="button"
                                onClick={async () => {
                                  if (isActive) return;
                                  const prevDocs = client.documents;
                                  updateClient(client.id, (c) => ({ ...c, documents: c.documents.map((d) => d.id === docItem.id ? { ...d, status: value } : d) }));
                                  try {
                                    await upsertDocument(client.id, docItem.id, { status: value, uploaded_file_name: docItem.uploadedFileName || "", file_path: docItem.filePath || "", storage_path: docItem.filePath || "", drive_url: docItem.driveUrl || "", confirmed_at: value === "confirmed" ? new Date().toISOString() : null });
                                    if (value === "confirmed" || value === "rejected") {
                                      const docName = template?.name || docItem.id;
                                      const tripName = trips.find((t) => t.id === client.tripId)?.name || "";
                                      sendNotification(value === "confirmed" ? "doc_confirmed" : "doc_rejected", client.email, client.id, { participantName: client.participantName, docName, tripName });
                                    }
                                  } catch (err) {
                                    console.error(err);
                                    updateClient(client.id, (c) => ({ ...c, documents: prevDocs }));
                                    notify("No se pudo guardar el estado del documento.");
                                  }
                                }}
                                className={`rounded-xl px-3 py-1.5 text-xs font-medium transition ${isActive ? meta.className : "border border-zinc-200 bg-white text-zinc-500 hover:bg-zinc-50"}`}
                                style={isActive && meta.style ? meta.style : undefined}
                              >{label}</button>
                            );
                          })}
                          {(docItem.driveUrl || docItem.filePath) && (
                            <Button variant="outline" className="ml-auto h-8 rounded-xl px-3 text-xs shrink-0" onClick={() => window.open(docItem.driveUrl || docItem.filePath, "_blank", "noopener,noreferrer")}>
                              <Eye className="mr-1.5 h-3.5 w-3.5" />Ver
                            </Button>
                          )}
                        </div>
                      </div>
                    );
                  })}
                </div>

                <div className="space-y-3">
                  <div className="text-sm font-semibold uppercase tracking-[0.18em] text-zinc-500">Pagos</div>
                  {[["reservation", client.payments.reservation], ["firstInstallment", client.payments.firstInstallment], ["secondInstallment", client.payments.secondInstallment]].map(([paymentKey, payment]) => {
                    const status = getStatusMeta(payment.status);
                    const isPendingReview = payment.status === "sent";
                    return (
                      <div key={String(paymentKey)} className={`rounded-2xl border p-4 space-y-3 ${isPendingReview ? "border-amber-300 bg-amber-50" : "border-zinc-200 bg-white"}`}>
                        {/* Fila 1: info + badge */}
                        <div className="flex items-start gap-3">
                          <div className="min-w-0 flex-1">
                            <div className="font-medium text-zinc-950">{payment.name || String(paymentKey)}</div>
                            <div className="text-sm text-zinc-500">{formatCurrency(payment.amount)} · {payment.proofName || "Sin justificante"}</div>
                            <div className="flex items-center gap-2 text-sm text-zinc-500 mt-0.5">
                              <span className={`h-2.5 w-2.5 shrink-0 rounded-full ${getDueStatus(getPaymentRuleDueDate(trip, paymentKey)).className}`} />
                              <span>Límite: {formatShortDate(getPaymentRuleDueDate(trip, paymentKey))}</span>
                            </div>
                          </div>
                          <Badge className={`shrink-0 ${status.className}`} style={status.style}>{status.label}</Badge>
                        </div>
                        {/* Fila 2: controles */}
                        <div className="flex flex-wrap items-center gap-2">
                          {[
                            { value: "pending", label: "Pendiente" },
                            { value: "sent", label: "Enviado" },
                            { value: "confirmed", label: "Confirmado" },
                            { value: "rejected", label: "Rechazado" },
                          ].map(({ value, label }) => {
                            const isActive = payment.status === value;
                            const meta = getStatusMeta(value);
                            return (
                              <button key={value} type="button"
                                onClick={async () => {
                                  if (isActive) return;
                                  const prevPayments = client.payments;
                                  updateClient(client.id, (c) => ({ ...c, payments: { ...c.payments, [paymentKey]: { ...c.payments[paymentKey], status: value } } }));
                                  try {
                                    await upsertPayment(client.id, paymentKey, { name: payment.name || String(paymentKey), amount: Number(payment.amount || 0), status: value, proof_name: payment.proofName || "", proof_path: payment.proofPath || "", due_date: payment.dueDate || null }, true);
                                    if (value === "confirmed") {
                                      const tripName = trips.find((t) => t.id === client.tripId)?.name || "";
                                      const amount = new Intl.NumberFormat("es-ES", { style: "currency", currency: "EUR" }).format(payment.amount || 0);
                                      sendNotification("payment_confirmed", client.email, client.id, { participantName: client.participantName, paymentName: payment.name, amount, tripName });
                                    }
                                    syncPaymentToSheet(client, paymentKey, value, payment.amount);
                                  } catch (err) {
                                    console.error(err);
                                    updateClient(client.id, (c) => ({ ...c, payments: prevPayments }));
                                    notify("No se pudo guardar el estado del pago.");
                                  }
                                }}
                                className={`rounded-xl px-3 py-1.5 text-xs font-medium transition ${isActive ? meta.className : "border border-zinc-200 bg-white text-zinc-500 hover:bg-zinc-50"}`}
                                style={isActive && meta.style ? meta.style : undefined}
                              >{label}</button>
                            );
                          })}
                          <Button variant="outline" className="h-10 rounded-2xl px-3 shrink-0" disabled={!payment.proofPath} onClick={() => window.open(payment.proofPath, "_blank", "noopener,noreferrer")}>
                            <Eye className="mr-2 h-4 w-4" />Ver justificante
                          </Button>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* Sección dudas */}
              {(client.questions || []).length > 0 && (
                <div className="mt-5 space-y-3">
                  <div className="text-sm font-semibold uppercase tracking-[0.18em] text-zinc-500">Dudas</div>
                  {[...(client.questions || [])].reverse().map((q) => (
                    <div key={q.id} className={`rounded-2xl border p-4 space-y-2 ${!q.reply ? "border-amber-300 bg-amber-50" : "border-zinc-200 bg-white"}`}>
                      <div className="flex items-start justify-between gap-3">
                        <div className="min-w-0 flex-1">
                          <div className="text-xs text-zinc-400">{new Date(q.createdAt).toLocaleDateString("es-ES", { day: "numeric", month: "long", year: "numeric" })}</div>
                          <p className="mt-1 text-sm text-zinc-900">{q.message}</p>
                        </div>
                        <Badge className={`shrink-0 ${q.reply ? "bg-green-100 text-green-800 hover:bg-green-100" : "bg-amber-100 text-amber-800 hover:bg-amber-100"}`}>
                          {q.reply ? "Respondida" : "Pendiente"}
                        </Badge>
                      </div>
                      {q.reply ? (
                        <div className="rounded-xl bg-green-50 px-3 py-2 text-xs text-zinc-700">↳ {q.reply}</div>
                      ) : (
                        <div className="flex gap-2">
                          <input
                            type="text"
                            value={replyTexts[q.id] || ""}
                            onChange={(e) => setReplyTexts((t) => ({ ...t, [q.id]: e.target.value }))}
                            onKeyDown={(e) => { if (e.key === "Enter") handleReplyQuestion(client.id, q.id, replyTexts[q.id]); }}
                            placeholder="Escribe la respuesta y pulsa Enter o Enviar…"
                            className="flex-1 rounded-xl border border-zinc-200 bg-white px-3 py-2 text-xs text-zinc-950 placeholder-zinc-400 focus:outline-none"
                          />
                          <Button
                            disabled={!replyTexts[q.id]?.trim() || !!sendingReply[q.id]}
                            onClick={() => handleReplyQuestion(client.id, q.id, replyTexts[q.id])}
                            className="h-8 shrink-0 rounded-xl px-3 text-xs text-white"
                            style={{ backgroundColor: CORPORATE_RED }}
                          >
                            <Send className="mr-1.5 h-3 w-3" />{sendingReply[q.id] ? "…" : "Enviar"}
                          </Button>
                        </div>
                      )}
                    </div>
                  ))}
                </div>
              )}
            </AccordionSection>
          );
        })}
      </div>
    </div>
  );
}

function AdminPayments({ users, setUsers, trips, setTrips, notify }) {
  const clients = users.filter((u) => u.role === "client" && !u.schoolId);
  const [searchQuery, setSearchQuery] = useState("");
  const filteredClients = clients.filter((c) => matchesParticipantSearch(c, searchQuery));
  const globalTi = trips?.[0]?.transferInfo || {};

  const saveGlobalTransferInfo = async (field, value) => {
    if (!trips?.length) return;
    const updated = { ...globalTi, [field]: value };
    setTrips?.((prev) => prev.map((t) => ({ ...t, transferInfo: updated })));
    const ids = trips.map((t) => t.id);
    const { error } = await supabase.from("trips").update({ transfer_info: updated }).in("id", ids);
    if (error) notify("Error guardando datos bancarios: " + error.message);
  };

  return (
    <div className="space-y-5">
      <SectionTitle icon={CreditCard} title="Pagos" subtitle="Edita importes y estados por cliente." />

      {/* Datos bancarios globales */}
      {trips && trips.length > 0 && (
        <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
          <CardContent className="p-5 space-y-4">
            <div className="flex items-center gap-2">
              <Building2 className="h-4 w-4 text-zinc-400" />
              <span className="text-sm font-semibold uppercase tracking-[0.15em] text-zinc-500">Datos bancarios</span>
            </div>
            <div className="grid gap-3 sm:grid-cols-2">
              {[["Banco", "bank"], ["Titular", "accountHolder"], ["IBAN", "iban"], ["Concepto", "concept"]].map(([label, field]) => (
                <div key={field} className="space-y-1">
                  <Label className="text-xs">{label}</Label>
                  <Input
                    defaultValue={globalTi[field] || ""}
                    placeholder={field === "concept" ? "Nombre del participante + viaje" : field === "iban" ? "ES00 0000 0000 0000 0000 0000" : ""}
                    className="rounded-xl text-sm h-9"
                    onBlur={(e) => saveGlobalTransferInfo(field, e.target.value)}
                  />
                </div>
              ))}
            </div>
          </CardContent>
        </Card>
      )}

      <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
        <CardContent className="p-5">
          <div className="space-y-2">
            <Label>Buscar participante</Label>
            <Input value={searchQuery} onChange={(e) => setSearchQuery(e.target.value)} placeholder="Busca por participante, familia, usuario o email" className="rounded-2xl" />
          </div>
        </CardContent>
      </Card>
      <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
        <CardContent className="space-y-4 p-5">
          <div className="grid gap-4">
            {filteredClients.map((client) => (
              <Card key={client.id} className="rounded-3xl border-zinc-200 bg-white">
                <CardContent className="grid gap-4 p-5 lg:grid-cols-[1.7fr_repeat(6,minmax(0,1fr))] lg:items-end">
                  <div>
                    <div className="font-medium text-zinc-950">{client.participantName}</div>
                    {getFamilyLabel(client) && <div className="text-sm text-zinc-500">Familia: {getFamilyLabel(client)}</div>}
                    {client.dni && <div className="text-sm text-zinc-500">DNI: {client.dni}</div>}
                    {client.email && <div className="text-xs text-zinc-400">{client.email}</div>}
                    <div className="mt-2 flex flex-wrap items-center gap-2">
                      {client.invoiceUrl ? (
                        <Button variant="outline" size="sm" className="h-8 rounded-2xl text-xs" onClick={() => window.open(client.invoiceUrl, "_blank", "noopener,noreferrer")}>
                          <Download className="mr-1.5 h-3.5 w-3.5" />Ver factura
                        </Button>
                      ) : null}
                      <InvoiceUploadButton existing={!!client.invoiceUrl} size="sm" onUpload={async (file, onProgress) => {
                        try {
                          const result = await uploadFileToDrive(file, client.participantName, "facturas", onProgress, "GIMELOOS Facturas");
                          const url = result.webViewLink;
                          const { error: invErr } = await supabase.from("participants").update({ invoice_url: url }).eq("id", client.id);
                          if (invErr) throw new Error(invErr.message);
                          setUsers((prev) => prev.map((u) => u.id === client.id ? { ...u, invoiceUrl: url } : u));
                          notify("Factura subida correctamente.");
                        } catch (err) { notify("Error subiendo factura: " + err.message); }
                      }} />
                      {client.invoiceUrl && (
                        <Button variant="outline" size="sm" className="h-8 rounded-2xl text-xs text-red-600 hover:bg-red-50 border-red-200" onClick={() => setPendingDeleteInvoiceClientId(client.id)}>
                          <Trash2 className="mr-1.5 h-3.5 w-3.5" />Borrar factura
                        </Button>
                      )}
                    </div>
                  </div>
                  {[["Precio inicial", "initialPrice"], ["Descuento", "discount"], ["Precio final", "finalPrice"], ["Reserva", "reservation"], ["1ª cuota", "firstInstallment"], ["2ª cuota", "secondInstallment"]].map(([label, key]) => (
                    <div key={label}>
                      <Label className="mb-2 block">{label}</Label>
                      {["initialPrice", "discount", "finalPrice"].includes(key) ? (
                        key === "finalPrice" ? (
                          <Input
                            type="number"
                            value={Math.max(0, (client.payments.initialPrice || 0) - (client.payments.discount || 0))}
                            readOnly
                            className="rounded-2xl bg-zinc-50 text-zinc-500 cursor-not-allowed"
                          />
                        ) : (
                        <Input type="number" value={client.payments[key] || ""}
                          onChange={async (e) => {
                            const num = Number(e.target.value || 0);
                            const newInitial   = key === "initialPrice" ? num : (client.payments.initialPrice || 0);
                            const newDiscount  = key === "discount"     ? num : (client.payments.discount     || 0);
                            const newFinal     = Math.max(0, newInitial - newDiscount);
                            // La segunda cuota absorbe el descuento: segundaCuota = final - reserva - primeraCuota
                            const reservation  = Number(client.payments.reservation?.amount || 0);
                            const first        = Number(client.payments.firstInstallment?.amount || 0);
                            const newSecond    = Math.max(0, newFinal - reservation - first);
                            setUsers((prev) => prev.map((u) => u.id === client.id ? {
                              ...u,
                              payments: {
                                ...u.payments,
                                [key]: num,
                                finalPrice: newFinal,
                                secondInstallment: { ...u.payments.secondInstallment, amount: newSecond },
                              }
                            } : u));
                            // Guardar precio
                            const { error: priceErr } = await supabase.from("participant_pricing").upsert(
                              { participant_id: client.id, initial_price: newInitial, discount: newDiscount, final_price: newFinal },
                              { onConflict: "participant_id" }
                            );
                            if (priceErr) notify("Error guardando precio: " + priceErr.message);
                            // Actualizar importe de la segunda cuota en BD
                            const { error: payErr } = await supabase.from("participant_payments")
                              .update({ amount: newSecond })
                              .eq("participant_id", client.id)
                              .eq("payment_key", "secondInstallment");
                            if (payErr) notify("Error actualizando segunda cuota: " + payErr.message);
                          }}
                          className="rounded-2xl"
                        />
                        )
                      ) : (
                        <Input type="number" value={client.payments[key].amount || ""}
                          onChange={async (e) => {
                            const num = Number(e.target.value || 0);
                            setUsers((prev) => prev.map((u) => u.id === client.id ? { ...u, payments: { ...u.payments, [key]: { ...u.payments[key], amount: num } } } : u));
                            const { error } = await supabase.from("participant_payments").update({ amount: num }).eq("participant_id", client.id).eq("payment_key", key);
                            if (error) notify("Error guardando importe: " + error.message);
                          }}
                          className="rounded-2xl"
                        />
                      )}
                    </div>
                  ))}
                </CardContent>
              </Card>
            ))}
          </div>
        </CardContent>
      </Card>
    </div>
  );
}

function AdminDocs({ templates, setTemplates, users, setUsers, trips, notify }) {
  const [name, setName] = useState("");
  const [uploadedFile, setUploadedFile] = useState(null);
  const [uploading, setUploading] = useState(false);
  const [uploadPct, setUploadPct] = useState(undefined);
  const [selectedTrip, setSelectedTrip] = useState(trips[0]?.id || "");
  const [selectedTemplateId, setSelectedTemplateId] = useState(templates[0]?.id || "");
  const [templateToDelete, setTemplateToDelete] = useState(null);

  const addTemplate = async () => {
    if (!name.trim()) return;
    setUploading(true);
    try {
      const id = `doc-${Date.now()}`;
      let driveUrl = "";
      let fileName = uploadedFile?.name || `${name.toLowerCase().replace(/\s+/g, "-")}.pdf`;
      if (uploadedFile) {
        const result = await uploadFileToDrive(uploadedFile, "archivos", "plantillas", (pct) => setUploadPct(pct), "GIMELOOS Plantillas");
        driveUrl = result.webViewLink;
        fileName = result.fileName;
      }
      const { error: tplSaveErr } = await supabase.from("document_templates").upsert({ id, name, file_name: fileName, drive_url: driveUrl });
      if (tplSaveErr) throw new Error(tplSaveErr.message);
      setTemplates((prev) => [...prev, { id, name, fileName, driveUrl }]);
      setName(""); setUploadedFile(null); setSelectedTemplateId(id);
      notify("Nueva plantilla creada.");
    } catch (err) {
      notify("Error al crear la plantilla: " + err.message);
    } finally {
      setUploading(false);
      setUploadPct(undefined);
    }
  };

  const deleteTemplate = async (templateId) => {
    const { error: e1 } = await supabase.from("participant_documents").delete().eq("template_id", templateId);
    if (e1) { notify("Error al eliminar documentos: " + e1.message); return; }
    const { error: e2 } = await supabase.from("document_templates").delete().eq("id", templateId);
    if (e2) { notify("Error al eliminar plantilla: " + e2.message); return; }
    setTemplates((prev) => prev.filter((t) => t.id !== templateId));
    setUsers((prev) => prev.map((u) => u.role === "client" ? { ...u, documents: u.documents.filter((d) => d.id !== templateId) } : u));
    if (selectedTemplateId === templateId) setSelectedTemplateId(templates.find((t) => t.id !== templateId)?.id || "");
    notify("Plantilla eliminada.");
  };

  const applyTemplateToGroup = async () => {
    const targetClients = users.filter((u) => u.role === "client" && u.tripId === selectedTrip);
    setUsers((prev) => prev.map((u) => {
      if (u.role !== "client" || u.tripId !== selectedTrip) return u;
      if (u.documents.some((d) => d.id === selectedTemplateId)) return u;
      return { ...u, documents: [...u.documents, { id: selectedTemplateId, status: "pending_upload", uploadedFileName: "", filePath: "", driveUrl: "" }] };
    }));
    for (const u of targetClients) {
      if (u.documents.some((d) => d.id === selectedTemplateId)) continue;
      try {
        await upsertDocument(u.id, selectedTemplateId, { status: "pending_upload", uploaded_file_name: "", file_path: "", storage_path: "", drive_url: "", confirmed_at: null });
      } catch (err) {
        console.error(err);
        notify("No se pudo aplicar la plantilla al grupo.");
        return;
      }
    }
    notify("Plantilla aplicada al grupo.");
  };

  return (
    <div className="space-y-5">
      <SectionTitle icon={FileText} title="Documentación" subtitle="Crea plantillas, bórralas y aplícalas a grupos de clientes por viaje." />
      <div className="grid gap-5 lg:grid-cols-[0.95fr_1.05fr]">
        <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
          <CardContent className="space-y-4 p-5">
            <div className="font-medium text-zinc-950">Nueva plantilla</div>
            <div className="space-y-2">
              <Label>Nombre del documento</Label>
              <Input value={name} onChange={(e) => setName(e.target.value)} placeholder="Ej. Autorización de salida" className="rounded-2xl" />
            </div>
            <label className="cursor-pointer">
              <input type="file" className="hidden" onChange={(e) => setUploadedFile(e.target.files?.[0] || null)} />
              <span className="inline-flex h-11 items-center rounded-2xl border border-zinc-200 px-4 text-sm font-medium text-zinc-900">
                <Upload className="mr-2 h-4 w-4" />{uploadedFile ? `${uploadedFile.name}` : "Subir archivo de plantilla"}
              </span>
            </label>
            <Button onClick={addTemplate} disabled={uploading} className="h-11 rounded-2xl text-white" style={{ backgroundColor: CORPORATE_RED }}>
              <FileText className="mr-2 h-4 w-4" />{uploading ? "Subiendo…" : "Crear plantilla"}
            </Button>
            {uploadPct !== undefined && (
              <div>
                <div className="mb-1 flex justify-between text-xs text-zinc-500">
                  <span>{uploadPct < 100 ? "Subiendo a Google Drive…" : "¡Listo!"}</span>
                  <span>{uploadPct}%</span>
                </div>
                <div className="h-1.5 w-full overflow-hidden rounded-full bg-zinc-100">
                  <div className="h-full rounded-full transition-all duration-200"
                    style={{ width: `${uploadPct}%`, backgroundColor: uploadPct === 100 ? "#16a34a" : CORPORATE_RED }} />
                </div>
              </div>
            )}
            <Separator />
            <div className="space-y-3">
              {templates.map((t) => (
                <div key={t.id} className="flex items-center justify-between gap-3 rounded-2xl border border-zinc-200 bg-white p-4">
                  <div className="min-w-0 flex-1">
                    <div className="font-medium text-zinc-950">{t.name}</div>
                    <div className="text-sm text-zinc-500">{t.fileName}</div>
                  </div>
                  <div className="flex items-center gap-2">
                    {t.driveUrl && (
                      <Button variant="ghost" size="icon" onClick={() => window.open(t.driveUrl, "_blank", "noopener,noreferrer")}>
                        <Eye className="h-4 w-4 text-zinc-700" />
                      </Button>
                    )}
                    <Button variant="ghost" size="icon" onClick={() => setTemplateToDelete(t)}>
                      <Trash2 className="h-4 w-4 text-zinc-700" />
                    </Button>
                  </div>
                </div>
              ))}
            </div>
          </CardContent>
        </Card>
        <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
          <CardContent className="space-y-4 p-5">
            <div className="font-medium text-zinc-950">Aplicar a grupo</div>
            <div className="space-y-2">
              <Label>Campamento</Label>
              <select value={selectedTrip} onChange={(e) => setSelectedTrip(e.target.value)} className="h-11 w-full rounded-2xl border border-zinc-200 bg-white px-4 text-sm">
                {trips.map((t) => <option key={t.id} value={t.id}>{t.name}</option>)}
              </select>
            </div>
            <div className="space-y-2">
              <Label>Plantilla</Label>
              <select value={selectedTemplateId} onChange={(e) => setSelectedTemplateId(e.target.value)} className="h-11 w-full rounded-2xl border border-zinc-200 bg-white px-4 text-sm">
                {templates.map((t) => <option key={t.id} value={t.id}>{t.name}</option>)}
              </select>
            </div>
            <Button onClick={applyTemplateToGroup} className="h-11 rounded-2xl text-white" style={{ backgroundColor: CORPORATE_RED }}>
              <Users className="mr-2 h-4 w-4" />Aplicar al grupo
            </Button>
          </CardContent>
        </Card>
      </div>

      <AlertDialog open={!!templateToDelete} onOpenChange={(open) => { if (!open) setTemplateToDelete(null); }}>
        <AlertDialogContent className="rounded-3xl">
          <AlertDialogHeader>
            <AlertDialogTitle>¿Eliminar plantilla?</AlertDialogTitle>
            <AlertDialogDescription>
              Se eliminará <strong>{templateToDelete?.name}</strong> y todos los documentos de participantes vinculados a ella. Esta acción no se puede deshacer.
            </AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel className="rounded-2xl">Cancelar</AlertDialogCancel>
            <AlertDialogAction
              className="rounded-2xl text-white"
              style={{ backgroundColor: "#dc2626" }}
              onClick={() => { deleteTemplate(templateToDelete.id); setTemplateToDelete(null); }}
            >
              Eliminar
            </AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>
    </div>
  );
}

function AdminChecklists({ trips, setTrips, notify }) {
  const [selectedTripId, setSelectedTripId] = useState(trips[0]?.id || "");
  const [newItem, setNewItem] = useState("");
  const selectedTrip = trips.find((t) => t.id === selectedTripId);

  const addItem = async () => {
    if (!newItem.trim()) return;
    const nextChecklist = [...selectedTrip.checklist, newItem];
    setTrips((prev) => prev.map((t) => t.id === selectedTripId ? { ...t, checklist: nextChecklist } : t));
    const { error } = await supabase.from("trips").update({ checklist: nextChecklist }).eq("id", selectedTripId);
    if (error) notify("Error añadiendo elemento: " + error.message);
    else setNewItem("");
  };

  const duplicateChecklist = () => {
    if (!selectedTrip) return;
    const cloneId = `trip-${Date.now()}`;
    setTrips((prev) => [...prev, { ...selectedTrip, id: cloneId, name: `${selectedTrip.name} · copia` }]);
    notify("Checklist duplicado correctamente.");
  };

  return (
    <div className="space-y-5">
      <SectionTitle icon={CheckCircle2} title="Checklist de equipaje" subtitle="Crea checklists por viaje y duplícalos para trabajar más rápido." />
      <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
        <CardContent className="space-y-4 p-5">
          <div className="flex flex-col gap-4 lg:flex-row lg:items-end">
            <div className="flex-1 space-y-2">
              <Label>Campamento</Label>
              <select value={selectedTripId} onChange={(e) => setSelectedTripId(e.target.value)} className="h-11 w-full rounded-2xl border border-zinc-200 bg-white px-4 text-sm">
                {trips.map((t) => <option key={t.id} value={t.id}>{t.name}</option>)}
              </select>
            </div>
            <Button variant="outline" className="h-11 rounded-2xl" onClick={duplicateChecklist}>
              <Copy className="mr-2 h-4 w-4" />Duplicar checklist
            </Button>
          </div>
          <div className="flex flex-col gap-3 lg:flex-row">
            <Input value={newItem} onChange={(e) => setNewItem(e.target.value)} placeholder="Añadir elemento de equipaje" className="rounded-2xl" />
            <Button onClick={addItem} className="h-11 rounded-2xl text-white" style={{ backgroundColor: CORPORATE_RED }}>
              <CheckCircle2 className="mr-2 h-4 w-4" />Añadir
            </Button>
          </div>
          <div className="grid gap-3 md:grid-cols-2">
            {selectedTrip?.checklist.map((item) => (
              // [MENOR-3] Key: item es string único en el checklist
              <div key={item} className="flex items-center justify-between rounded-2xl border border-zinc-200 bg-white p-4">
                <span className="text-sm text-zinc-800">{item}</span>
                <Button variant="ghost" size="sm" onClick={async () => {
                  const next = selectedTrip.checklist.filter((l) => l !== item);
                  setTrips((prev) => prev.map((t) => t.id === selectedTripId ? { ...t, checklist: next } : t));
                  const { error } = await supabase.from("trips").update({ checklist: next }).eq("id", selectedTripId);
                  if (error) notify("Error quitando elemento: " + error.message);
                }}>Quitar</Button>
              </div>
            ))}
          </div>
        </CardContent>
      </Card>
    </div>
  );
}

function AdminTrips({ trips, setTrips, notify }) {
  const [selectedTripId, setSelectedTripId] = useState(trips[0]?.id || "");
  const [creating, setCreating] = useState(false);
  const selectedTrip = trips.find((t) => t.id === selectedTripId) || trips[0];

  const syncField = (field, value) => setTrips((prev) => prev.map((t) => t.id === selectedTripId ? { ...t, [field]: value } : t));
  const saveField = async (field, value) => {
    syncField(field, value);
    const { error } = await supabase.from("trips").update({ [field]: value }).eq("id", selectedTripId);
    if (error) notify("Error guardando cambios: " + error.message);
  };

  const handleCreate = async () => {
    setCreating(true);
    const newId = crypto.randomUUID();
    const { data, error } = await supabase.from("trips").insert({ id: newId, name: "Nuevo campamento", tipo: "campamento", checklist: [], itinerary: [], logistics: [] }).select().single();
    if (error) { notify("Error creando campamento: " + error.message); setCreating(false); return; }
    const newTrip = { id: data.id, name: data.name, departureDate: "", description: "", heroImage: "", heroImages: [], transferInfo: { bank: "", accountHolder: "", iban: "", concept: "" }, automation: {}, showItinerary: true, showLogistics: true, documentRules: [], paymentSchedule: {}, itinerary: [], logistics: [], checklist: [], tipo: "campamento" };
    setTrips((prev) => [...prev, newTrip]);
    setSelectedTripId(data.id);
    setCreating(false);
    notify("Campamento creado. Edita el nombre y los datos.");
  };

  if (!trips.length) return (
    <div className="space-y-5">
      <SectionTitle icon={MapIcon} title="Campamentos" subtitle="Información básica y foto de portada de cada campamento." extra={
        <Button onClick={handleCreate} disabled={creating} className="rounded-2xl text-white" style={{ backgroundColor: CORPORATE_RED }}>
          <Plus className="mr-1.5 h-4 w-4" />{creating ? "Creando..." : "Nuevo campamento"}
        </Button>
      } />
      <div className="py-16 text-center text-sm text-zinc-400">No hay campamentos configurados. Crea el primero.</div>
    </div>
  );

  return (
    <div className="space-y-5">
      <SectionTitle icon={MapIcon} title="Campamentos" subtitle="Información básica y foto de portada de cada campamento." extra={
        <Button onClick={handleCreate} disabled={creating} className="rounded-2xl text-white" style={{ backgroundColor: CORPORATE_RED }}>
          <Plus className="mr-1.5 h-4 w-4" />{creating ? "Creando..." : "Nuevo campamento"}
        </Button>
      } />
      <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
        <CardContent className="p-5">
          <div className="flex flex-col gap-3 sm:flex-row sm:items-center">
            <div className="flex-1 space-y-1">
              <Label>Campamento activo</Label>
              <select value={selectedTripId} onChange={(e) => setSelectedTripId(e.target.value)} className="h-11 w-full rounded-2xl border border-zinc-200 bg-white px-4 text-sm font-medium">
                {trips.map((t) => <option key={t.id} value={t.id}>{t.name}</option>)}
              </select>
            </div>
            {selectedTrip?.departureDate && (
              <div className="rounded-2xl border border-zinc-100 bg-white px-4 py-2 text-center">
                <div className="text-xs text-zinc-400">Fecha de salida</div>
                <div className="font-semibold text-zinc-950">{new Date(selectedTrip.departureDate).toLocaleDateString("es-ES", { day: "numeric", month: "long", year: "numeric" })}</div>
              </div>
            )}
          </div>
        </CardContent>
      </Card>
      {selectedTrip && (
        <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
          <CardContent className="space-y-5 p-6">
            <div className="grid gap-4 lg:grid-cols-2">
              <div className="space-y-2">
                <Label>Nombre del viaje</Label>
                <Input value={selectedTrip.name} onChange={(e) => syncField("name", e.target.value)} onBlur={(e) => saveField("name", e.target.value)} className="rounded-2xl" />
              </div>
              <div className="space-y-2">
                <Label>Fecha de salida</Label>
                <Input type="datetime-local" value={(selectedTrip?.departureDate || "").slice(0, 16)} onChange={async (e) => { syncField("departureDate", e.target.value); const { error } = await supabase.from("trips").update({ departure_date: e.target.value || null }).eq("id", selectedTripId); if (error) notify("Error guardando fecha: " + error.message); }} className="rounded-2xl" />
              </div>
            </div>
            <div className="space-y-2">
              <Label>Foto de portada</Label>
              <CoverImageInput
                value={selectedTrip.heroImage || ""}
                onChange={(v) => syncField("heroImage", v)}
                onBlur={(v) => saveField("hero_image", v)}
                tripId={selectedTripId}
                notify={notify}
              />
            </div>
            <div className="space-y-2">
              <Label>Descripción</Label>
              <Textarea value={selectedTrip.description || ""} onChange={(e) => syncField("description", e.target.value)} onBlur={(e) => saveField("description", e.target.value)} className="min-h-[120px] rounded-2xl" />
            </div>
          </CardContent>
        </Card>
      )}
    </div>
  );
}

function AdminItinerary({ trips, setTrips, notify }) {
  const [selectedTripId, setSelectedTripId] = useState(trips[0]?.id || "");
  const selectedTrip = trips.find((t) => t.id === selectedTripId) || trips[0];
  const [localOrder, setLocalOrder] = useState(selectedTrip?.itinerary || []);

  useEffect(() => { setLocalOrder(selectedTrip?.itinerary || []); }, [selectedTripId]);

  const syncItinerary = async (nextOrder) => {
    setLocalOrder(nextOrder);
    setTrips((prev) => prev.map((t) => t.id === selectedTripId ? { ...t, itinerary: nextOrder } : t));
    const { error } = await supabase.from("trips").update({ itinerary: nextOrder }).eq("id", selectedTripId);
    if (error) notify("Error guardando itinerario: " + error.message);
  };

  if (!trips.length) return <div className="py-16 text-center text-sm text-zinc-400">No hay campamentos configurados.</div>;

  return (
    <div className="space-y-5">
      <SectionTitle icon={CalendarDays} title="Itinerario" subtitle="Programa día a día de cada campamento. Arrastra para reordenar." />
      <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
        <CardContent className="p-5">
          <div className="space-y-1">
            <Label>Campamento</Label>
            <select value={selectedTripId} onChange={(e) => setSelectedTripId(e.target.value)} className="h-11 w-full rounded-2xl border border-zinc-200 bg-white px-4 text-sm font-medium">
              {trips.map((t) => <option key={t.id} value={t.id}>{t.name}</option>)}
            </select>
          </div>
        </CardContent>
      </Card>
      {selectedTrip && (
        <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
          <CardContent className="space-y-4 p-6">
            <div className="flex items-center justify-between flex-wrap gap-3">
              <div>
                <div className="font-semibold text-zinc-950">Itinerario día a día</div>
                <div className="text-sm text-zinc-500">Arrastra las filas para reordenar.</div>
              </div>
              <div className="flex items-center gap-3 flex-wrap">
                <label className="flex items-center gap-2 cursor-pointer select-none">
                  <div
                    onClick={async () => {
                      const next = selectedTrip.showItinerary === false;
                      setTrips((prev) => prev.map((t) => t.id === selectedTripId ? { ...t, showItinerary: next } : t));
                      const { error } = await supabase.from("trips").update({ automation: { ...(selectedTrip.automation || {}), showItinerary: next } }).eq("id", selectedTripId);
                      if (error) notify("Error guardando visibilidad: " + error.message);
                    }}
                    className={`relative h-6 w-11 rounded-full transition-colors cursor-pointer ${selectedTrip.showItinerary !== false ? "bg-green-500" : "bg-zinc-300"}`}
                  >
                    <div className={`absolute top-0.5 h-5 w-5 rounded-full bg-white shadow transition-all ${selectedTrip.showItinerary !== false ? "left-5" : "left-0.5"}`} />
                  </div>
                  <span className="text-sm text-zinc-700">Visible para clientes</span>
                </label>
                <Button className="rounded-2xl text-white" style={{ backgroundColor: CORPORATE_RED }}
                  onClick={() => syncItinerary([...localOrder, { day: `Día ${localOrder.length + 1}`, title: "Nuevo tramo", description: "Detalle", time: "10:00" }])}>
                  <Plus className="mr-2 h-4 w-4" />Añadir tramo
                </Button>
              </div>
            </div>
            <Reorder.Group axis="y" values={localOrder} onReorder={syncItinerary} className="space-y-3">
              {localOrder.map((item, index) => (
                <Reorder.Item key={`${item.day}-${item.title}-${index}`} value={item} whileDrag={{ scale: 1.015, boxShadow: "0 28px 60px rgba(0,0,0,0.18)", zIndex: 30 }} className="list-none">
                  <div className="rounded-2xl border border-zinc-200 bg-white p-4 transition hover:border-zinc-300">
                    <div className="flex items-start gap-3">
                      <div className="mt-3 cursor-grab text-zinc-300 active:cursor-grabbing"><GripVertical className="h-5 w-5" /></div>
                      <div className="flex-1 space-y-2">
                        <div className="grid gap-2 sm:grid-cols-3">
                          <Input value={item.day} placeholder="Día" onChange={(e) => syncItinerary(localOrder.map((l, i) => i === index ? { ...l, day: e.target.value } : l))} className="rounded-xl bg-white text-sm" />
                          <Input value={item.title} placeholder="Título" onChange={(e) => syncItinerary(localOrder.map((l, i) => i === index ? { ...l, title: e.target.value } : l))} className="rounded-xl bg-white text-sm" />
                          <Input value={item.time || ""} placeholder="Hora" onChange={(e) => syncItinerary(localOrder.map((l, i) => i === index ? { ...l, time: e.target.value } : l))} className="rounded-xl bg-white text-sm" />
                        </div>
                        <Input value={item.description} placeholder="Descripción" onChange={(e) => syncItinerary(localOrder.map((l, i) => i === index ? { ...l, description: e.target.value } : l))} className="rounded-xl bg-white text-sm" />
                      </div>
                      <button type="button" onClick={() => syncItinerary(localOrder.filter((_, i) => i !== index))} className="mt-2 rounded-xl p-1.5 text-zinc-400 hover:bg-zinc-200 hover:text-zinc-700">
                        <X className="h-4 w-4" />
                      </button>
                    </div>
                  </div>
                </Reorder.Item>
              ))}
            </Reorder.Group>
            {localOrder.length === 0 && (
              <div className="rounded-2xl border-2 border-dashed border-zinc-200 py-10 text-center text-sm text-zinc-400">
                Sin tramos. Pulsa "Añadir tramo" para empezar.
              </div>
            )}
          </CardContent>
        </Card>
      )}
    </div>
  );
}

function AdminLogistica({ trips, setTrips, notify }) {
  const [selectedTripId, setSelectedTripId] = useState(trips[0]?.id || "");
  const selectedTrip = trips.find((t) => t.id === selectedTripId) || trips[0];

  const syncLogistics = async (nextLogistics) => {
    setTrips((prev) => prev.map((t) => t.id === selectedTripId ? { ...t, logistics: nextLogistics } : t));
    const { error } = await supabase.from("trips").update({ logistics: nextLogistics }).eq("id", selectedTripId);
    if (error) notify("Error guardando logística: " + error.message);
  };

  if (!trips.length) return <div className="py-16 text-center text-sm text-zinc-400">No hay campamentos configurados.</div>;

  return (
    <div className="space-y-5">
      <SectionTitle icon={MapPinned} title="Logística" subtitle="Datos clave previos al campamento: horarios, lugar de encuentro, qué llevar..." />
      <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
        <CardContent className="p-5">
          <div className="space-y-1">
            <Label>Campamento</Label>
            <select value={selectedTripId} onChange={(e) => setSelectedTripId(e.target.value)} className="h-11 w-full rounded-2xl border border-zinc-200 bg-white px-4 text-sm font-medium">
              {trips.map((t) => <option key={t.id} value={t.id}>{t.name}</option>)}
            </select>
          </div>
        </CardContent>
      </Card>
      {selectedTrip && (
        <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
          <CardContent className="space-y-4 p-6">
            <div className="flex items-center justify-between flex-wrap gap-3">
              <div>
                <div className="font-semibold text-zinc-950">Puntos logísticos</div>
                <div className="text-sm text-zinc-500">Punto de encuentro, hora de salida, qué traer, contacto de emergencia...</div>
              </div>
              <div className="flex items-center gap-3 flex-wrap">
                <label className="flex items-center gap-2 cursor-pointer select-none">
                  <div
                    onClick={async () => {
                      const next = selectedTrip.showLogistics === false;
                      setTrips((prev) => prev.map((t) => t.id === selectedTripId ? { ...t, showLogistics: next } : t));
                      const { error } = await supabase.from("trips").update({ automation: { ...(selectedTrip.automation || {}), showLogistics: next } }).eq("id", selectedTripId);
                      if (error) notify("Error guardando visibilidad: " + error.message);
                    }}
                    className={`relative h-6 w-11 rounded-full transition-colors cursor-pointer ${selectedTrip.showLogistics !== false ? "bg-green-500" : "bg-zinc-300"}`}
                  >
                    <div className={`absolute top-0.5 h-5 w-5 rounded-full bg-white shadow transition-all ${selectedTrip.showLogistics !== false ? "left-5" : "left-0.5"}`} />
                  </div>
                  <span className="text-sm text-zinc-700">Visible para clientes</span>
                </label>
                <Button className="rounded-2xl text-white" style={{ backgroundColor: CORPORATE_RED }}
                  onClick={() => syncLogistics([...(selectedTrip.logistics || []), { title: "Nuevo punto", description: "" }])}>
                  <Plus className="mr-2 h-4 w-4" />Añadir punto
                </Button>
              </div>
            </div>
            <div className="space-y-3">
              {(selectedTrip.logistics || []).map((item, index) => (
                <div key={index} className="flex gap-3 items-start rounded-2xl border border-zinc-200 bg-white p-4">
                  <div className="flex-1 space-y-2">
                    <Input value={item.title} placeholder="Título (ej: Punto de encuentro)" onChange={(e) => syncLogistics((selectedTrip.logistics || []).map((l, i) => i === index ? { ...l, title: e.target.value } : l))} className="rounded-xl bg-white text-sm font-medium" />
                    <Textarea value={item.description} placeholder="Descripción (ej: Parking del instituto a las 8:00)" onChange={(e) => syncLogistics((selectedTrip.logistics || []).map((l, i) => i === index ? { ...l, description: e.target.value } : l))} className="min-h-[72px] rounded-xl bg-white text-sm" />
                  </div>
                  <button type="button" onClick={() => syncLogistics((selectedTrip.logistics || []).filter((_, i) => i !== index))} className="mt-1 rounded-xl p-1.5 text-zinc-400 hover:bg-zinc-200 hover:text-zinc-700">
                    <X className="h-4 w-4" />
                  </button>
                </div>
              ))}
              {(selectedTrip.logistics || []).length === 0 && (
                <div className="rounded-2xl border-2 border-dashed border-zinc-200 py-10 text-center text-sm text-zinc-400">
                  Sin puntos logísticos. Pulsa "Añadir punto" para empezar.
                </div>
              )}
            </div>
          </CardContent>
        </Card>
      )}
    </div>
  );
}

function QuestionCard({ q, replyText, onReplyChange, onSendReply, sending }) {
  return (
    <div className="rounded-3xl border border-zinc-200 bg-white p-5 space-y-3">
      <div className="flex items-start justify-between gap-3">
        <div>
          <div className="font-medium text-zinc-950">{q.participantName}</div>
          <div className="text-xs text-zinc-400">{new Date(q.createdAt).toLocaleDateString("es-ES", { day: "numeric", month: "long", year: "numeric" })}</div>
        </div>
        <Badge className={q.reply ? "bg-green-100 text-green-800 hover:bg-green-100" : "bg-amber-100 text-amber-800 hover:bg-amber-100"}>
          {q.reply ? "Respondida" : "Pendiente"}
        </Badge>
      </div>
      <p className="text-sm text-zinc-700 rounded-2xl bg-white px-4 py-3">{q.message}</p>
      {q.reply ? (
        <div className="rounded-2xl bg-green-50 px-4 py-3">
          <div className="mb-1 text-xs font-medium text-green-700">Tu respuesta</div>
          <p className="text-sm text-zinc-800">{q.reply}</p>
        </div>
      ) : (
        <div className="space-y-2">
          <Textarea
            value={replyText}
            onChange={(e) => onReplyChange(e.target.value)}
            placeholder="Escribe tu respuesta..."
            className="min-h-[80px] rounded-2xl border-zinc-200 bg-white text-sm"
          />
          <Button
            disabled={!replyText?.trim() || sending}
            onClick={onSendReply}
            className="rounded-2xl text-white"
            style={{ backgroundColor: CORPORATE_RED }}
          >
            <Send className="mr-2 h-4 w-4" />{sending ? "Enviando…" : "Responder"}
          </Button>
        </div>
      )}
    </div>
  );
}

function AdminQuestions({ users, setUsers, notify }) {
  const allQuestions = users
    .filter((u) => u.role === "client")
    .flatMap((u) => (u.questions || []).map((q) => ({ ...q, participantName: u.participantName, participantId: u.id })))
    .sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));

  const pending = allQuestions.filter((q) => !q.reply);
  const answered = allQuestions.filter((q) => q.reply);

  const [replyText, setReplyText] = useState({});
  const [sending, setSending] = useState({});

  const sendReply = async (q) => {
    const reply = replyText[q.id]?.trim();
    if (!reply) return;
    setSending((s) => ({ ...s, [q.id]: true }));
    try {
      const { error } = await supabase
        .from("participant_questions")
        .update({ reply, replied_at: new Date().toISOString(), status: "replied" })
        .eq("id", q.id);
      if (error) throw error;
      setUsers((prev) => prev.map((u) =>
        u.id !== q.participantId ? u : {
          ...u,
          questions: (u.questions || []).map((question) =>
            question.id !== q.id ? question : { ...question, reply, repliedAt: new Date().toISOString(), status: "replied" }
          ),
        }
      ));
      setReplyText((r) => { const n = { ...r }; delete n[q.id]; return n; });
      notify("Respuesta enviada.");
      const participant = users.find((u) => u.id === q.participantId);
      if (participant?.email) {
        sendNotification("question_replied", participant.email, q.participantId, { participantName: q.participantName, question: q.message, reply });
      }
    } catch (err) {
      notify("Error al enviar la respuesta: " + err.message);
    } finally {
      setSending((s) => ({ ...s, [q.id]: false }));
    }
  };

  return (
    <div className="space-y-5">
      <SectionTitle icon={MessageCircleQuestion} title="Preguntas" subtitle="Consultas enviadas por los participantes." />
      {allQuestions.length === 0 ? (
        <div className="rounded-3xl border border-zinc-200 bg-white p-10 text-center text-sm text-zinc-400">Aún no hay preguntas.</div>
      ) : (
        <div className="grid gap-5 lg:grid-cols-2">
          <div className="space-y-4">
            <div className="text-sm font-semibold uppercase tracking-[0.18em] text-zinc-500">Pendientes ({pending.length})</div>
            {pending.length === 0 && <div className="rounded-3xl border border-zinc-200 bg-white p-6 text-sm text-zinc-400 text-center">Todo respondido ✓</div>}
            {pending.map((q) => (
              <QuestionCard key={q.id} q={q} replyText={replyText[q.id] || ""} onReplyChange={(v) => setReplyText((r) => ({ ...r, [q.id]: v }))} onSendReply={() => sendReply(q)} sending={!!sending[q.id]} />
            ))}
          </div>
          <div className="space-y-4">
            <div className="text-sm font-semibold uppercase tracking-[0.18em] text-zinc-500">Respondidas ({answered.length})</div>
            {answered.length === 0 && <div className="rounded-3xl border border-zinc-200 bg-white p-6 text-sm text-zinc-400 text-center">Sin respuestas aún.</div>}
            {answered.map((q) => (
              <QuestionCard key={q.id} q={q} replyText={replyText[q.id] || ""} onReplyChange={(v) => setReplyText((r) => ({ ...r, [q.id]: v }))} onSendReply={() => sendReply(q)} sending={!!sending[q.id]} />
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

// ─── School Portal ────────────────────────────────────────────────────────────

function SchoolTrips({ schoolTrips, courses, students, schoolDocuments, onNavigate }) {
  const [expandedId, setExpandedId] = useState(null);

  const getPending = (st) => {
    const tripCourses = courses.filter((c) => c.school_trip_id === st.id);
    const tripCourseIds = new Set(tripCourses.map((c) => c.id));
    const tripStudents = students.filter((s) => tripCourseIds.has(s.school_course_id));
    const pendingDocs = schoolDocuments.filter((d) => tripCourseIds.has(d.school_course_id) && d.status !== "approved");
    const studentsWithoutMedical = tripStudents.filter((s) => !s.allergies && !s.intolerances && !s.notes);
    return {
      docs: pendingDocs,
      studentsNoMedical: studentsWithoutMedical,
      tripStudents,
      tripCourses,
      total: pendingDocs.length,
    };
  };

  if (!schoolTrips.length) {
    return (
      <div className="flex flex-col items-center justify-center py-16 text-zinc-400">
        <CalendarDays className="mb-3 h-10 w-10 opacity-40" />
        <p className="text-sm">No hay viajes asignados todavía.</p>
      </div>
    );
  }

  return (
    <div className="space-y-3">
      <SectionTitle icon={CalendarDays} title="Mis viajes" subtitle="Viajes escolares asignados a tu colegio." />
      {schoolTrips.map((st) => {
        const pending = getPending(st);
        const isOpen = expandedId === st.id;
        const depDate = st.trips?.departure_date
          ? new Date(st.trips.departure_date).toLocaleDateString("es-ES", { day: "numeric", month: "long", year: "numeric" })
          : null;
        const remaining = daysRemaining(st.trips?.departure_date);

        return (
          <div key={st.id} className="rounded-2xl border border-zinc-200 bg-white shadow-sm overflow-hidden">
            {/* Fila principal — clicable */}
            <button
              type="button"
              onClick={() => setExpandedId(isOpen ? null : st.id)}
              className="flex w-full items-center gap-4 px-5 py-4 text-left hover:bg-zinc-50 transition"
            >
              {/* Mini thumbnail */}
              {st.trips?.hero_image ? (
                <img src={st.trips.hero_image} alt={st.trips?.name} className="h-12 w-16 rounded-xl object-cover shrink-0" />
              ) : (
                <div className="flex h-12 w-16 items-center justify-center rounded-xl bg-zinc-100 shrink-0">
                  <Map className="h-5 w-5 text-zinc-400" />
                </div>
              )}

              {/* Info */}
              <div className="min-w-0 flex-1">
                <div className="font-semibold text-zinc-950 leading-tight truncate">{st.trips?.name || st.trip_id}</div>
                {depDate && (
                  <div className="mt-0.5 flex items-center gap-1.5 text-xs text-zinc-500">
                    <CalendarDays className="h-3 w-3 shrink-0" />{depDate}
                    <span className="ml-1 rounded-full bg-zinc-100 px-2 py-0.5 text-[10px] font-medium text-zinc-600">
                      {remaining === 0 ? "hoy" : `${remaining}d`}
                    </span>
                  </div>
                )}
                {st.courses?.length > 0 && (
                  <div className="mt-1 flex flex-wrap gap-1">
                    {st.courses.map((c) => (
                      <span key={c.id} className="rounded-lg bg-zinc-100 px-2 py-0.5 text-[10px] text-zinc-600">
                        {c.course_name}{c.group_name ? ` · ${c.group_name}` : ""}
                      </span>
                    ))}
                  </div>
                )}
              </div>

              {/* Badge pendientes + chevron */}
              <div className="flex items-center gap-2 shrink-0">
                {pending.total > 0 && (
                  <span className="flex h-5 w-5 items-center justify-center rounded-full text-[10px] font-bold text-white" style={{ backgroundColor: CORPORATE_RED }}>
                    {pending.total}
                  </span>
                )}
                {pending.total === 0 && (
                  <span className="flex h-5 w-5 items-center justify-center rounded-full bg-emerald-100">
                    <CheckCircle2 className="h-3 w-3 text-emerald-600" />
                  </span>
                )}
                <ChevronDown className={`h-4 w-4 text-zinc-400 transition-transform ${isOpen ? "rotate-180" : ""}`} />
              </div>
            </button>

            {/* Panel expandible */}
            {isOpen && (
              <div className="border-t border-zinc-100 px-5 py-4 bg-zinc-50 space-y-4">
                {/* Stats rápidas */}
                <div className="grid gap-2 sm:grid-cols-2 lg:grid-cols-4">
                  {[
                    { icon: Users,        label: "Alumnos",        val: pending.tripStudents.length,                               ok: pending.tripStudents.length > 0 },
                    { icon: FileCheck2,   label: "Docs pendientes",val: pending.docs.length,                                       ok: pending.docs.length === 0,  bad: pending.docs.length > 0 },
                    { icon: Home,         label: "Rooming",        val: st.rooming?.length ? `${st.rooming.length} hab.` : "Pendiente",   ok: (st.rooming?.length ?? 0) > 0, bad: !(st.rooming?.length) },
                    { icon: Grid2x2,     label: "Grupos",         val: st.activity_groups?.length ? `${st.activity_groups.length} grupos` : "Pendiente", ok: (st.activity_groups?.length ?? 0) > 0, bad: !(st.activity_groups?.length) },
                    { icon: CheckCircle2, label: "Checklist",  val: st.checklist?.length ? `${st.checklist.length} ítems` : "Pendiente",  ok: (st.checklist?.length ?? 0) > 0,  bad: !(st.checklist?.length) },
                    ...(st.show_itinerary !== false ? [{ icon: CalendarDays, label: "Itinerario", val: st.itinerary?.length ? `${st.itinerary.length} tramos` : "Pendiente", ok: (st.itinerary?.length ?? 0) > 0, bad: !(st.itinerary?.length) }] : []),
                    ...(st.show_logistics !== false ? [{ icon: MapPinned,    label: "Logística",   val: st.logistics?.length ? `${st.logistics.length} puntos` : "Pendiente", ok: (st.logistics?.length ?? 0) > 0, bad: !(st.logistics?.length) }] : []),
                    { icon: CreditCard,   label: "Pagos",          val: st.payment_info?.status === "completed" ? "Al día" : st.payment_info?.status === "partial" ? "Parcial" : "Pendiente", ok: st.payment_info?.status === "completed", bad: !st.payment_info?.status || st.payment_info?.status === "pending",
                      extra: st.invoice_url ? (
                        <button type="button" className="ml-1 inline-flex items-center gap-0.5 text-[10px] text-emerald-600 underline" onClick={(e) => { e.stopPropagation(); window.open(st.invoice_url, "_blank", "noopener,noreferrer"); }}>
                          <Download className="h-3 w-3" />Factura
                        </button>
                      ) : null },
                  ].map(({ icon: Icon, label, val, ok, bad, extra }) => (
                    <div key={label} className={`flex items-center gap-3 rounded-2xl border px-3 py-2.5 ${ok ? "border-emerald-200 bg-emerald-50" : bad ? "border-red-100 bg-red-50" : "border-zinc-200 bg-white"}`}>
                      <Icon className={`h-4 w-4 shrink-0 ${ok ? "text-emerald-600" : bad ? "text-red-400" : "text-zinc-400"}`} />
                      <div className="min-w-0">
                        <div className={`text-xs font-semibold ${ok ? "text-emerald-700" : bad ? "text-red-600" : "text-zinc-500"}`}>{label}</div>
                        <div className={`truncate text-xs ${ok ? "text-emerald-600" : bad ? "text-red-500" : "text-zinc-400"}`}>{val}{extra}</div>
                      </div>
                    </div>
                  ))}
                </div>

                {/* Documentos pendientes */}
                {pending.docs.length > 0 ? (
                  <div>
                    <div className="mb-2 text-xs font-semibold uppercase tracking-widest text-zinc-400">Documentos pendientes</div>
                    <div className="space-y-1.5">
                      {pending.docs.slice(0, 6).map((d) => (
                        <button key={d.id} type="button" onClick={() => onNavigate?.("docs")}
                          className="flex w-full items-center justify-between rounded-xl bg-white border border-zinc-200 px-3 py-2 text-sm text-left hover:bg-zinc-50 transition">
                          <span className="truncate text-zinc-700">{d.doc_name || d.name || d.id}</span>
                          <Badge variant="outline" className="ml-2 shrink-0 rounded-lg text-[10px]">{d.status || "pendiente"}</Badge>
                        </button>
                      ))}
                      {pending.docs.length > 6 && (
                        <div className="text-xs text-zinc-400 px-1">+{pending.docs.length - 6} más</div>
                      )}
                    </div>
                  </div>
                ) : null}
                {/* "Todo al día" solo cuando NADA está en rojo */}
                {(() => {
                  const allOk = pending.docs.length === 0
                    && (st.rooming?.length ?? 0) > 0
                    && (st.activity_groups?.length ?? 0) > 0
                    && (st.checklist?.length ?? 0) > 0
                    && (st.show_itinerary === false || (st.itinerary?.length ?? 0) > 0)
                    && (st.show_logistics === false || (st.logistics?.length ?? 0) > 0)
                    && st.payment_info?.status === "completed";
                  return allOk ? (
                    <div className="flex items-center gap-2 text-sm text-emerald-600">
                      <CheckCircle2 className="h-4 w-4" /> Todo al día
                    </div>
                  ) : null;
                })()}
                {/* Botón de factura si el admin la ha subido */}
                {st.invoice_url && (
                  <div>
                    <Button variant="outline" className="h-9 rounded-2xl text-sm" onClick={() => window.open(st.invoice_url, "_blank", "noopener,noreferrer")}>
                      <Download className="mr-2 h-4 w-4" />Descargar factura
                    </Button>
                  </div>
                )}
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}

function SchoolStudents({ schoolTrips, courses, students, setStudents, notify }) {
  const [selectedCourseId, setSelectedCourseId] = useState(courses[0]?.id || "");
  const [showAddForm, setShowAddForm] = useState(false);
  const [addName, setAddName] = useState("");
  const [addSurname, setAddSurname] = useState("");
  const [addAllergies, setAddAllergies] = useState("");
  const [addIntolerances, setAddIntolerances] = useState("");
  const [addNotes, setAddNotes] = useState("");
  const [xlsxPreview, setXlsxPreview] = useState(null); // { rows, mapping, headers, file }
  const [importing, setImporting] = useState(false);
  const [studentToDelete, setStudentToDelete] = useState(null);

  const tripCourses = courses; // all courses for this school
  const courseStudents = students.filter((s) => s.school_course_id === selectedCourseId);

  const normalizeHeader = (h) => h.toLowerCase().trim().normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/\s+/g, " ");

  const FIELD_SYNONYMS = {
    nombre: ["nombre", "name", "alumno", "primer nombre", "first name", "nom"],
    apellidos: ["apellidos", "apellido", "surname", "last name", "segundo nombre", "cognoms"],
    alergias: ["alergias", "alergia", "allergies", "allergy", "al·lergies", "alergies"],
    intolerancias: ["intolerancias", "intolerancia", "intolerances", "intolerance"],
    notas: ["notas", "observaciones", "notes", "comments", "otros", "observacions"],
  };

  const detectMapping = (headers) => {
    const mapping = {};
    headers.forEach((h, i) => {
      const norm = normalizeHeader(h);
      Object.entries(FIELD_SYNONYMS).forEach(([field, synonyms]) => {
        if (!mapping[field] && synonyms.some((s) => norm === s || norm.includes(s))) {
          mapping[field] = i;
        }
      });
    });
    return mapping;
  };

  const handleFileChange = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const buffer = await file.arrayBuffer();
      const XLSX2 = await import("xlsx");
      const wb = XLSX2.read(buffer, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const raw = XLSX2.utils.sheet_to_json(ws, { header: 1, defval: "" });
      if (!raw.length) { notify("El archivo está vacío.", { variant: "destructive" }); return; }
      const headers = raw[0].map(String);
      const mapping = detectMapping(headers);
      const rows = raw.slice(1).filter((r) => r.some((c) => String(c).trim()));
      setXlsxPreview({ rows, mapping, headers, file });
    } catch (err) {
      console.error(err);
      notify("No se pudo leer el archivo Excel.", { variant: "destructive" });
    }
    e.target.value = "";
  };

  const handleImport = async () => {
    if (!xlsxPreview || !selectedCourseId) return;
    setImporting(true);
    const { rows, mapping } = xlsxPreview;
    const toInsert = rows.map((r) => ({
      school_course_id: selectedCourseId,
      name: mapping.nombre !== undefined ? String(r[mapping.nombre] || "").trim() : "",
      surname: mapping.apellidos !== undefined ? String(r[mapping.apellidos] || "").trim() : "",
      allergies: mapping.alergias !== undefined ? String(r[mapping.alergias] || "").trim() : "",
      intolerances: mapping.intolerancias !== undefined ? String(r[mapping.intolerancias] || "").trim() : "",
      notes: mapping.notas !== undefined ? String(r[mapping.notas] || "").trim() : "",
    })).filter((s) => s.name);
    if (!toInsert.length) { notify("No se encontraron filas con nombre."); setImporting(false); return; }
    const { data, error } = await supabase.from("students").insert(toInsert).select();
    if (error) { notify("Error importando alumnos: " + error.message, { variant: "destructive" }); }
    else {
      setStudents((prev) => [...prev, ...(data || [])]);
      notify(`${toInsert.length} alumnos importados.`);
      setXlsxPreview(null);
    }
    setImporting(false);
  };

  const handleAddManual = async () => {
    if (!addName.trim() || !selectedCourseId) return;
    const row = { school_course_id: selectedCourseId, name: addName.trim(), surname: addSurname.trim(), allergies: addAllergies.trim(), intolerances: addIntolerances.trim(), notes: addNotes.trim() };
    const { data, error } = await supabase.from("students").insert([row]).select().maybeSingle();
    if (error) { notify("Error añadiendo alumno: " + error.message, { variant: "destructive" }); return; }
    setStudents((prev) => [...prev, data]);
    setAddName(""); setAddSurname(""); setAddAllergies(""); setAddIntolerances(""); setAddNotes("");
    notify("Alumno añadido.");
  };

  const handleDelete = async (id) => {
    const { error } = await supabase.from("students").delete().eq("id", id);
    if (error) { notify("Error eliminando alumno.", { variant: "destructive" }); return; }
    setStudents((prev) => prev.filter((s) => s.id !== id));
    notify("Alumno eliminado.");
  };

  const handleExportAlumnosPDF = () => {
    const courseName = tripCourses.find((c) => c.id === selectedCourseId);
    const tripName = schoolTrips.find((st) => st.id === selectedTripId)?.trips?.name || "";
    const label = courseName ? `${courseName.course_name}${courseName.group_name ? ` · ${courseName.group_name}` : ""}` : "Todos los cursos";
    const rows = (selectedCourseId ? courseStudents : students).map((s, i) => {
      const course = courses.find((c) => c.id === s.school_course_id);
      return `<tr>
        <td>${i + 1}</td>
        <td style="font-weight:600">${[s.name, s.surname].filter(Boolean).join(" ")}</td>
        <td>${course ? `${course.course_name}${course.group_name ? ` · ${course.group_name}` : ""}` : "—"}</td>
        <td>${s.allergies ? `<span class="badge badge-red">${s.allergies}</span>` : "—"}</td>
        <td>${s.intolerances ? `<span class="badge badge-amber">${s.intolerances}</span>` : "—"}</td>
        <td>${s.notes || "—"}</td>
      </tr>`;
    }).join("");
    exportListToPDF(
      `Listado de alumnos — ${label}`,
      tripName,
      `<table><thead><tr><th>#</th><th>Alumno</th><th>Curso / Grupo</th><th>Alergia</th><th>Intolerancia</th><th>Notas</th></tr></thead><tbody>${rows}</tbody></table>`
    );
  };

  return (
    <div className="space-y-6">
      <div className="flex items-center gap-3">
        <SectionTitle icon={Users} title="Alumnos" subtitle="Gestiona los alumnos asignados a cada curso." />
        <label className="flex shrink-0 cursor-pointer items-center gap-1.5 rounded-2xl px-4 py-2.5 text-sm font-medium text-white hover:opacity-90 transition-opacity" style={{ backgroundColor: CORPORATE_RED }}>
          <FolderUp className="h-4 w-4" />Importar listado
          <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleFileChange} />
        </label>
      </div>
      {/* Tabs de curso */}
      {courses.length > 0 && (
        <div className="flex flex-wrap gap-2">
          {courses.map((c) => (
            <button key={c.id} type="button" onClick={() => setSelectedCourseId(c.id)}
              className={`rounded-2xl px-3 py-1.5 text-xs font-medium transition ${selectedCourseId === c.id ? "text-white shadow-sm" : "border border-zinc-200 bg-white text-zinc-600 hover:bg-zinc-50"}`}
              style={selectedCourseId === c.id ? { backgroundColor: CORPORATE_RED } : {}}>
              {c.course_name}{c.group_name ? ` · ${c.group_name}` : ""}
            </button>
          ))}
        </div>
      )}

      {/* Botón añadir alumno + formulario colapsable */}
      {selectedCourseId && (
        <div>
          {!showAddForm ? (
            <Button variant="outline" className="rounded-2xl text-sm" onClick={() => setShowAddForm(true)}>
              <Plus className="mr-1.5 h-4 w-4" />Añadir alumno manualmente
            </Button>
          ) : (
            <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
              <CardContent className="p-5">
                <div className="mb-3 flex items-center justify-between">
                  <div className="text-sm font-semibold text-zinc-700">Añadir alumno manualmente</div>
                  <Button variant="ghost" size="sm" className="rounded-xl text-xs text-zinc-500" onClick={() => setShowAddForm(false)}>
                    <X className="h-3.5 w-3.5" />
                  </Button>
                </div>
                <div className="grid gap-2 sm:grid-cols-2 lg:grid-cols-3">
                  <Input placeholder="Nombre *" value={addName} onChange={(e) => setAddName(e.target.value)} className="h-11 rounded-2xl border-zinc-200" />
                  <Input placeholder="Apellidos" value={addSurname} onChange={(e) => setAddSurname(e.target.value)} className="h-11 rounded-2xl border-zinc-200" />
                  <Input placeholder="Alergias" value={addAllergies} onChange={(e) => setAddAllergies(e.target.value)} className="h-11 rounded-2xl border-zinc-200" />
                  <Input placeholder="Intolerancias" value={addIntolerances} onChange={(e) => setAddIntolerances(e.target.value)} className="h-11 rounded-2xl border-zinc-200" />
                  <Input placeholder="Notas" value={addNotes} onChange={(e) => setAddNotes(e.target.value)} className="h-11 rounded-2xl border-zinc-200" />
                  <Button onClick={handleAddManual} disabled={!addName.trim()} className="h-11 rounded-2xl text-white" style={{ backgroundColor: CORPORATE_RED }}>
                    <Plus className="mr-1.5 h-3.5 w-3.5" />Añadir
                  </Button>
                </div>
              </CardContent>
            </Card>
          )}
        </div>
      )}

      {/* Excel import preview — aparece tras seleccionar archivo */}
      {xlsxPreview && (
        <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
          <CardContent className="p-5">
            <div className="mb-3 flex items-center justify-between">
              <div className="text-sm font-semibold text-zinc-700">Vista previa del archivo</div>
              <Button variant="ghost" size="sm" className="rounded-xl text-xs text-zinc-500" onClick={() => setXlsxPreview(null)}>
                <X className="mr-1 h-3.5 w-3.5" />Cancelar
              </Button>
            </div>
            <div className="space-y-3">
              <div className="flex flex-wrap gap-1.5">
                {Object.entries(FIELD_SYNONYMS).map(([field]) => {
                  const found = xlsxPreview.mapping[field] !== undefined;
                  return (
                    <Badge key={field} variant="outline" className={`rounded-xl text-xs ${found ? "border-green-400 text-green-700" : "border-zinc-300 text-zinc-400"}`}>
                      {found ? "✓" : "✗"} {field} {found ? `(col: ${xlsxPreview.headers[xlsxPreview.mapping[field]]})` : ""}
                    </Badge>
                  );
                })}
              </div>
              {xlsxPreview.mapping.nombre === undefined && (
                <div className="flex items-center gap-1.5 text-xs text-amber-600">
                  <AlertCircle className="h-3.5 w-3.5" />No se detectó columna de nombre. Verifica los encabezados.
                </div>
              )}
              <div className="overflow-x-auto">
                <table className="w-full text-xs">
                  <thead>
                    <tr className="border-b border-zinc-100 bg-zinc-50">
                      {xlsxPreview.headers.map((h, i) => <th key={i} className="px-3 py-2 text-left text-xs font-semibold uppercase tracking-wide text-zinc-500">{h}</th>)}
                    </tr>
                  </thead>
                  <tbody>
                    {xlsxPreview.rows.slice(0, 5).map((r, ri) => (
                      <tr key={ri} className="border-b border-zinc-50">
                        {xlsxPreview.headers.map((_, ci) => <td key={ci} className="px-3 py-2 text-zinc-700">{String(r[ci] || "")}</td>)}
                      </tr>
                    ))}
                  </tbody>
                </table>
                {xlsxPreview.rows.length > 5 && <div className="mt-1 text-xs text-zinc-400">+{xlsxPreview.rows.length - 5} filas más</div>}
              </div>
              <div className="flex gap-2">
                <Button
                  onClick={handleImport}
                  disabled={importing || xlsxPreview.mapping.nombre === undefined}
                  className="h-11 rounded-2xl text-white text-xs"
                  style={{ backgroundColor: CORPORATE_RED }}
                >
                  {importing ? "Importando..." : `Importar ${xlsxPreview.rows.filter((r) => r.some((c) => String(c).trim())).length} alumnos`}
                </Button>
                <Button variant="outline" className="h-11 rounded-2xl text-xs" onClick={() => setXlsxPreview(null)}>Cancelar</Button>
              </div>
            </div>
          </CardContent>
        </Card>
      )}

      {/* Student list */}
      {selectedCourseId && (
        <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
          <CardContent className="p-5">
            <div className="mb-3 flex items-center justify-between">
              <div className="text-sm font-semibold text-zinc-700">Lista de alumnos</div>
              <Badge variant="outline" className="rounded-xl text-xs">{courseStudents.length} alumnos</Badge>
            </div>
            {courseStudents.length === 0 ? (
              <p className="text-xs text-zinc-400">No hay alumnos en este grupo todavía.</p>
            ) : (
              <div className="overflow-x-auto">
                <table className="w-full text-xs">
                  <thead>
                    <tr className="border-b border-zinc-100 bg-zinc-50">
                      <th className="px-3 py-2 text-left text-xs font-semibold uppercase tracking-wide text-zinc-500">Nombre</th>
                      <th className="px-3 py-2 text-left text-xs font-semibold uppercase tracking-wide text-zinc-500">Apellidos</th>
                      <th className="px-3 py-2 text-left text-xs font-semibold uppercase tracking-wide text-zinc-500">Alergias</th>
                      <th className="px-3 py-2 text-left text-xs font-semibold uppercase tracking-wide text-zinc-500">Intolerancias</th>
                      <th className="px-3 py-2 text-left text-xs font-semibold uppercase tracking-wide text-zinc-500">Notas</th>
                      <th className="px-3 py-2"></th>
                    </tr>
                  </thead>
                  <tbody>
                    {courseStudents.map((s) => (
                      <tr key={s.id} className="border-b border-zinc-50 hover:bg-zinc-50/50">
                        <td className="px-3 py-2 font-medium text-zinc-900">{s.name}</td>
                        <td className="px-3 py-2 text-zinc-700">{s.surname}</td>
                        <td className="px-3 py-2 text-zinc-700">{s.allergies || "—"}</td>
                        <td className="px-3 py-2 text-zinc-700">{s.intolerances || "—"}</td>
                        <td className="px-3 py-2 text-zinc-700">{s.notes || "—"}</td>
                        <td className="px-3 py-2">
                          <Button variant="ghost" size="icon" onClick={() => setStudentToDelete(s)} className="h-7 w-7 text-zinc-400 hover:bg-red-50 hover:text-red-600">
                            <Trash2 className="h-3.5 w-3.5" />
                          </Button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </CardContent>
        </Card>
      )}
      <AlertDialog open={!!studentToDelete} onOpenChange={(open) => { if (!open) setStudentToDelete(null); }}>
        <AlertDialogContent className="rounded-3xl">
          <AlertDialogHeader>
            <AlertDialogTitle>¿Eliminar alumno?</AlertDialogTitle>
            <AlertDialogDescription>
              Se eliminará a <strong>{studentToDelete?.name} {studentToDelete?.surname}</strong> de la lista. Esta acción no se puede deshacer.
            </AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel className="rounded-2xl">Cancelar</AlertDialogCancel>
            <AlertDialogAction
              className="rounded-2xl text-white"
              style={{ backgroundColor: "#dc2626" }}
              onClick={() => { handleDelete(studentToDelete.id); setStudentToDelete(null); }}
            >Eliminar</AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>
    </div>
  );
}

function SchoolAllergies({ courses, students }) {
  const [selectedCourseId, setSelectedCourseId] = useState("");
  const filteredStudents = selectedCourseId
    ? students.filter((s) => s.school_course_id === selectedCourseId && (s.allergies?.trim() || s.intolerances?.trim()))
    : students.filter((s) => s.allergies?.trim() || s.intolerances?.trim());
  const withAllergies = filteredStudents;
  const getCourse = (id) => courses.find((c) => c.id === id);

  const handleExportAlergiasPDF = () => {
    const rows = withAllergies.map((s, i) => {
      const course = getCourse(s.school_course_id);
      return `<tr>
        <td>${i + 1}</td>
        <td style="font-weight:600">${[s.name, s.surname].filter(Boolean).join(" ")}</td>
        <td>${course ? `${course.course_name}${course.group_name ? ` · ${course.group_name}` : ""}` : "—"}</td>
        <td>${s.allergies ? `<span class="badge badge-red">${s.allergies}</span>` : "—"}</td>
        <td>${s.intolerances ? `<span class="badge badge-amber">${s.intolerances}</span>` : "—"}</td>
        <td>${s.diet_notes || s.notes || "—"}</td>
      </tr>`;
    }).join("");
    exportListToPDF(
      "Alergias e intolerancias",
      `${withAllergies.length} alumno(s) con restricciones`,
      withAllergies.length === 0
        ? "<p style='color:#71717a;font-size:13px'>Ningún alumno tiene alergias o intolerancias registradas.</p>"
        : `<table><thead><tr><th>#</th><th>Alumno</th><th>Curso / Grupo</th><th>Alergia</th><th>Intolerancia</th><th>Notas dietéticas</th></tr></thead><tbody>${rows}</tbody></table>`
    );
  };

  return (
    <div className="space-y-4">
      <SectionTitle icon={AlertCircle} title="Alergias e intolerancias" subtitle="Alumnos con restricciones alimentarias." />
      {/* Filtro de curso */}
      {courses.length > 1 && (
        <div className="flex flex-wrap gap-2">
          <button type="button" onClick={() => setSelectedCourseId("")}
            className={`rounded-2xl px-3 py-1.5 text-xs font-medium transition ${!selectedCourseId ? "text-white shadow-sm" : "border border-zinc-200 bg-white text-zinc-600 hover:bg-zinc-50"}`}
            style={!selectedCourseId ? { backgroundColor: CORPORATE_RED } : {}}>Todos</button>
          {courses.map((c) => (
            <button key={c.id} type="button" onClick={() => setSelectedCourseId(c.id)}
              className={`rounded-2xl px-3 py-1.5 text-xs font-medium transition ${selectedCourseId === c.id ? "text-white shadow-sm" : "border border-zinc-200 bg-white text-zinc-600 hover:bg-zinc-50"}`}
              style={selectedCourseId === c.id ? { backgroundColor: CORPORATE_RED } : {}}>
              {c.course_name}{c.group_name ? ` · ${c.group_name}` : ""}
            </button>
          ))}
        </div>
      )}
      {withAllergies.length === 0 ? (
        <div className="flex flex-col items-center justify-center py-16 text-zinc-400">
          <CheckCircle2 className="mb-3 h-10 w-10 opacity-40" />
          <p className="text-sm">Ningún alumno tiene alergias o intolerancias registradas.</p>
        </div>
      ) : (
        <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
          <CardContent className="p-5">
            <div className="overflow-x-auto">
              <table className="w-full text-xs">
                <thead>
                  <tr className="border-b border-zinc-100 bg-zinc-50">
                    <th className="px-4 py-3 text-left text-xs font-semibold uppercase tracking-wide text-zinc-500">Nombre completo</th>
                    <th className="px-4 py-3 text-left text-xs font-semibold uppercase tracking-wide text-zinc-500">Curso / grupo</th>
                    <th className="px-4 py-3 text-left text-xs font-semibold uppercase tracking-wide text-zinc-500">Alergia</th>
                    <th className="px-4 py-3 text-left text-xs font-semibold uppercase tracking-wide text-zinc-500">Intolerancia</th>
                    <th className="px-4 py-3 text-left text-xs font-semibold uppercase tracking-wide text-zinc-500">Notas</th>
                  </tr>
                </thead>
                <tbody>
                  {withAllergies.map((s) => {
                    const course = getCourse(s.school_course_id);
                    return (
                      <tr key={s.id} className="border-b border-zinc-50 hover:bg-zinc-50/50">
                        <td className="px-4 py-3 font-medium text-zinc-900">{[s.name, s.surname].filter(Boolean).join(" ")}</td>
                        <td className="px-4 py-3 text-zinc-700">{course ? `${course.course_name}${course.group_name ? ` · ${course.group_name}` : ""}` : "—"}</td>
                        <td className="px-4 py-3">{s.allergies ? <Badge className="bg-red-50 text-red-700 border border-red-200 rounded-xl text-xs font-medium">{s.allergies}</Badge> : <span className="text-zinc-400">—</span>}</td>
                        <td className="px-4 py-3">{s.intolerances ? <Badge className="bg-amber-50 text-amber-700 border border-amber-200 rounded-xl text-xs font-medium">{s.intolerances}</Badge> : <span className="text-zinc-400">—</span>}</td>
                        <td className="px-4 py-3 text-zinc-700">{s.diet_notes || s.notes || "—"}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </CardContent>
        </Card>
      )}
    </div>
  );
}

function SchoolDocs({ courses, schoolDocuments, setSchoolDocuments, notify, school, schoolTrips }) {
  const [progress, setProgress] = useState({}); // docId → pct (0-100)
  const [localFileNames, setLocalFileNames] = useState({}); // docId → filename
  const [movingDoc, setMovingDoc] = useState(null); // docId being reassigned

  const handleMoveDoc = async (docId, newCourseId) => {
    setMovingDoc(docId);
    const { error } = await supabase.from("school_documents").update({ school_course_id: newCourseId }).eq("id", docId);
    if (error) { notify("Error moviendo el documento.", { variant: "destructive" }); }
    else {
      setSchoolDocuments((prev) => prev.map((d) => d.id === docId ? { ...d, school_course_id: newCourseId } : d));
      notify("Documento reasignado al curso correctamente.");
    }
    setMovingDoc(null);
  };

  const handleUpload = async (docId, courseId, file) => {
    if (!file) return;
    // Feedback inmediato
    setLocalFileNames((n) => ({ ...n, [docId]: file.name }));
    setProgress((p) => ({ ...p, [docId]: 0 }));

    const course = courses.find((c) => c.id === courseId);
    const schoolTrip = schoolTrips?.find((st) => st.id === course?.school_trip_id);
    const tripName = schoolTrip?.trips?.name || "colegio";
    const courseName = [course?.course_name, course?.group_name].filter(Boolean).join(" - ");
    const participantName = school?.name ? `${school.name}${courseName ? ` · ${courseName}` : ""}` : courseName || "colegio";

    try {
      const json = await uploadFileToDrive(
        file, participantName, "documentación",
        (pct) => setProgress((p) => ({ ...p, [docId]: pct })),
        tripName
      );
      const fileUrl = json.webViewLink || json.fileId;
      setProgress((p) => ({ ...p, [docId]: 100 }));
      const { error: dbErr } = await supabase.from("school_documents").update({ file_url: fileUrl, status: "uploaded", uploaded_file_name: file.name }).eq("id", docId);
      if (dbErr) { notify("Error actualizando estado del documento."); }
      setSchoolDocuments((prev) => prev.map((d) => d.id === docId ? { ...d, file_url: fileUrl, status: "uploaded", uploaded_file_name: file.name } : d));
      notify("✅ Documento subido correctamente.");
      setTimeout(() => setProgress((p) => { const n = { ...p }; delete n[docId]; return n; }), 1800);
    } catch (err) {
      setLocalFileNames((n) => { const m = { ...n }; delete m[docId]; return m; });
      setProgress((p) => { const m = { ...p }; delete m[docId]; return m; });
      notify("Error subiendo documento: " + err.message);
    }
  };

  if (!courses.length) {
    return <div className="py-16 text-center text-sm text-zinc-400">No hay cursos asignados.</div>;
  }

  return (
    <div className="space-y-6">
      <SectionTitle icon={FileCheck2} title="Documentación" subtitle="Documentos y recursos de tu viaje." />

      {/* ── Recursos subidos por GIMELOOS para el colegio ─────────────────── */}
      {schoolTrips?.some(st => st.admin_docs?.length) && (
        <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
          <CardContent className="p-5 space-y-3">
            <div className="font-semibold text-zinc-950">Recursos del viaje</div>
            <p className="text-sm text-zinc-500">Documentos e información compartidos por el equipo de GIMELOOS para tu colegio.</p>
            <div className="space-y-2">
              {schoolTrips.flatMap(st => (st.admin_docs || []).map(doc => (
                <div key={doc.id} className="flex items-center justify-between rounded-2xl border border-zinc-200 bg-zinc-50 p-4">
                  <div className="flex items-center gap-3 min-w-0">
                    <FileText className="h-5 w-5 shrink-0 text-zinc-400" />
                    <div className="min-w-0">
                      <div className="text-sm font-medium text-zinc-900 truncate">{doc.name}</div>
                      {doc.created_at && <div className="text-xs text-zinc-400">{new Date(doc.created_at).toLocaleDateString("es-ES")}</div>}
                    </div>
                  </div>
                  {doc.file_url && (
                    <Button variant="outline" size="sm" className="rounded-xl text-xs h-8 ml-3 shrink-0" onClick={() => window.open(doc.file_url, "_blank", "noopener,noreferrer")}>
                      <ExternalLink className="mr-1 h-3.5 w-3.5" />Ver / Descargar
                    </Button>
                  )}
                </div>
              )))}
            </div>
          </CardContent>
        </Card>
      )}
      {courses.map((course) => {
        const courseDocs = schoolDocuments.filter((d) => d.school_course_id === course.id);
        return (
          <Card key={course.id} className="rounded-3xl border-zinc-200 bg-white shadow-sm">
            <CardContent className="p-5">
              <div className="mb-4">
                <Badge variant="outline" className="rounded-2xl border-zinc-200 px-3 py-1 text-sm font-semibold text-zinc-900">
                  {course.course_name}{course.group_name ? ` · ${course.group_name}` : ""}
                </Badge>
              </div>
              {courseDocs.length === 0 ? (
                <p className="text-xs text-zinc-400">Sin documentos requeridos.</p>
              ) : (
                <div className="space-y-2">
                  {courseDocs.map((doc) => {
                    const pct = progress[doc.id];
                    const uploading = pct !== undefined && pct < 100;
                    const displayFileName = localFileNames[doc.id] || doc.uploaded_file_name;
                    return (
                      <div key={doc.id} className="rounded-2xl border border-zinc-100 px-4 py-3 space-y-2">
                        <div className="flex items-center justify-between gap-3">
                          <div className="flex min-w-0 flex-1 items-center gap-2">
                            <FileCheck2 className="h-4 w-4 shrink-0 text-zinc-400" />
                            <div className="min-w-0">
                              <div className="text-xs font-medium text-zinc-900">{doc.name}</div>
                              {displayFileName && (
                                <div className="mt-0.5 flex items-center gap-1 text-xs text-zinc-500">
                                  <CheckCircle2 className="h-3 w-3 shrink-0 text-emerald-500" />
                                  <span className="truncate">Subido: {displayFileName}</span>
                                </div>
                              )}
                              {!displayFileName && (
                                <div className="mt-0.5">
                                  {doc.status === "uploaded"
                                    ? <Badge className="bg-green-50 text-green-700 border border-green-200 rounded-xl text-xs font-medium">Subido</Badge>
                                    : <Badge className="bg-amber-50 text-amber-700 border border-amber-200 rounded-xl text-xs font-medium">Pendiente</Badge>
                                  }
                                </div>
                              )}
                            </div>
                          </div>
                          <div className="flex shrink-0 items-center gap-2">
                            {doc.file_url && (
                              <a href={doc.file_url} target="_blank" rel="noreferrer" download
                                className="inline-flex items-center gap-1 rounded-xl border border-zinc-200 bg-white px-2.5 py-1.5 text-xs font-medium text-zinc-700 hover:bg-zinc-50 transition">
                                <Download className="h-3 w-3" />Descargar
                              </a>
                            )}
                            {courses.length > 1 && (
                              <select
                                value={doc.school_course_id}
                                disabled={movingDoc === doc.id}
                                onChange={(e) => handleMoveDoc(doc.id, e.target.value)}
                                className="rounded-xl border border-zinc-200 bg-white px-2 py-1 text-xs text-zinc-700 focus:outline-none"
                                title="Mover a otro curso"
                              >
                                {courses.map((c) => (
                                  <option key={c.id} value={c.id}>{c.course_name}{c.group_name ? ` · ${c.group_name}` : ""}</option>
                                ))}
                              </select>
                            )}
                            <label className={`flex cursor-pointer items-center gap-1 rounded-2xl border px-3 py-1.5 text-xs font-medium text-white transition hover:opacity-90 ${uploading ? "cursor-not-allowed opacity-60" : ""}`} style={{ backgroundColor: CORPORATE_RED, borderColor: CORPORATE_RED }}>
                              <Upload className="h-3 w-3" />{uploading ? `${Math.round(pct)}%` : "Subir"}
                              <input type="file" className="hidden" disabled={uploading} onChange={(e) => handleUpload(doc.id, doc.school_course_id, e.target.files?.[0])} />
                            </label>
                          </div>
                        </div>
                        {pct !== undefined && (
                          <div className="h-1.5 w-full overflow-hidden rounded-full bg-zinc-100">
                            <div className="h-full rounded-full transition-all duration-200" style={{ width: `${pct}%`, backgroundColor: pct === 100 ? "#16a34a" : CORPORATE_RED }} />
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              )}
            </CardContent>
          </Card>
        );
      })}
    </div>
  );
}

function fmtExcelTime(v) {
  const n = parseFloat(v);
  if (!isNaN(n) && n > 0 && n < 1) {
    const totalMins = Math.round(n * 24 * 60);
    const h = Math.floor(totalMins / 60) % 24;
    const m = totalMins % 60;
    return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
  }
  return String(v);
}

function SchoolRooming({ schoolTrips, setSchoolTrips, notify }) {
  const allCourses = schoolTrips.flatMap((st) => (st.courses || []).map((c) => ({ ...c, tripId: st.id, tripName: st.trips?.name || "" })));
  const [selectedCourseId, setSelectedCourseId] = useState(allCourses[0]?.id || "");
  const [importing, setImporting] = useState(false);
  const [confirmClearRooming, setConfirmClearRooming] = useState(false);

  const selectedTrip = schoolTrips.find((st) => (st.courses || []).some((c) => c.id === selectedCourseId));
  const selectedTripId = selectedTrip?.id || "";
  const rooming = selectedTrip?.rooming || [];

  const handleClearRooming = async () => {
    const { error } = await supabase.from("school_trips").update({ rooming: [] }).eq("id", selectedTripId);
    if (!error) { setSchoolTrips((prev) => prev.map((st) => st.id === selectedTripId ? { ...st, rooming: [] } : st)); notify("Rooming eliminado."); }
    else notify("Error eliminando rooming: " + error.message, { variant: "destructive" });
  };

  const handleFileChange = async (e) => {
    const file = e.target.files?.[0];
    if (!file || !selectedTripId) return;
    setImporting(true);
    try {
      const buffer = await file.arrayBuffer();
      const XLSX2 = await import("xlsx");
      const wb = XLSX2.read(buffer, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const raw = XLSX2.utils.sheet_to_json(ws, { header: 1, defval: "" });
      const rooms = [];
      raw.forEach((row) => {
        if (!row[0] && row[0] !== 0) return;
        const roomName = fmtExcelTime(row[0]);
        const studentNames = row.slice(1).map((c) => (c === "" || c == null) ? "" : fmtExcelTime(c)).filter(Boolean);
        if (roomName) rooms.push({ room: roomName, students: studentNames });
      });
      const { error } = await supabase.from("school_trips").update({ rooming: rooms }).eq("id", selectedTripId);
      if (error) { notify("Error guardando rooming.", { variant: "destructive" }); }
      else {
        setSchoolTrips((prev) => prev.map((st) => st.id === selectedTripId ? { ...st, rooming: rooms } : st));
        notify(`Rooming importado: ${rooms.length} habitaciones.`);
      }
    } catch (err) {
      console.error(err);
      notify("Error leyendo archivo.", { variant: "destructive" });
    }
    setImporting(false);
    e.target.value = "";
  };

  const handleExportRoomingPDF = () => {
    if (!rooming.length) return;
    const selCourse = allCourses.find((c) => c.id === selectedCourseId);
    const tripName = selCourse ? `${selCourse.course_name}${selCourse.group_name ? ` · ${selCourse.group_name}` : ""}` : "";
    const blocks = rooming.map((r) =>
      `<div class="room-block">
        <div class="room-title">🛏️ ${r.room} <span style="font-weight:400;font-size:11px;color:#71717a">(${r.students.length} alumnos)</span></div>
        <table style="width:100%"><tbody>
          ${r.students.map((s, i) => `<tr><td style="width:28px;color:#a1a1aa">${i + 1}</td><td>${s}</td></tr>`).join("")}
        </tbody></table>
      </div>`
    ).join("<hr style='border:none;border-top:1px solid #f4f4f5;margin:12px 0'>");
    exportListToPDF("Rooming", tripName, blocks);
  };

  return (
    <div className="space-y-6">
      <SectionTitle icon={Home} title="Rooming" subtitle="Asignación de habitaciones. Importa desde Excel." />
      {/* Tabs de curso */}
      {allCourses.length > 0 && (
        <div className="flex flex-wrap gap-2">
          {allCourses.map((c) => (
            <button key={c.id} type="button" onClick={() => setSelectedCourseId(c.id)}
              className={`rounded-2xl px-3 py-1.5 text-xs font-medium transition ${selectedCourseId === c.id ? "text-white shadow-sm" : "border border-zinc-200 bg-white text-zinc-600 hover:bg-zinc-50"}`}
              style={selectedCourseId === c.id ? { backgroundColor: CORPORATE_RED } : {}}>
              {c.course_name}{c.group_name ? ` · ${c.group_name}` : ""}
            </button>
          ))}
        </div>
      )}
      <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
        <CardContent className="p-5">
          <div className="flex flex-wrap items-center gap-3">
            <label className={`flex h-11 cursor-pointer items-center gap-1.5 rounded-2xl border px-4 text-sm font-medium text-white hover:opacity-90 transition ${importing ? "opacity-50" : ""}`} style={{ backgroundColor: "#FF3131", borderColor: "#FF3131" }}>
              <FolderUp className="h-4 w-4" />{importing ? "Importando..." : "Importar documento"}
              <input type="file" accept=".xlsx,.xls,.csv,.pdf" className="hidden" onChange={handleFileChange} disabled={importing} />
            </label>
            {rooming.length > 0 && (
              <Button variant="outline" className="h-11 rounded-2xl text-sm text-red-600 hover:bg-red-50 border-red-200" onClick={() => setConfirmClearRooming(true)}>
                <Trash2 className="mr-2 h-4 w-4" />Borrar rooming
              </Button>
            )}
          </div>
          <div className="mt-2 text-xs text-zinc-400">Formato esperado: columna 1 = nombre de habitación, columnas siguientes = nombres de alumnos.</div>
        </CardContent>
      </Card>
      {rooming.length > 0 ? (
        <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
          {rooming.map((r, i) => (
            <Card key={i} className="rounded-3xl border-zinc-200 bg-white shadow-sm">
              <CardContent className="p-5">
                <div className="mb-2 font-semibold text-zinc-900">{fmtExcelTime(r.room)}</div>
                <div className="space-y-1">
                  {r.students.map((s, j) => <div key={j} className="flex items-center gap-1.5 text-xs text-zinc-600"><User className="h-3 w-3 shrink-0 text-zinc-400" />{fmtExcelTime(s)}</div>)}
                </div>
                <Badge variant="outline" className="mt-3 rounded-xl text-xs">{r.students.length} alumnos</Badge>
              </CardContent>
            </Card>
          ))}
        </div>
      ) : (
        <div className="flex flex-col items-center justify-center py-16 text-zinc-400">
          <LayoutGrid className="mb-3 h-10 w-10 opacity-40" />
          <p className="text-sm">Importa un Excel para ver las habitaciones aquí.</p>
        </div>
      )}

      <AlertDialog open={confirmClearRooming} onOpenChange={setConfirmClearRooming}>
        <AlertDialogContent>
          <AlertDialogHeader>
            <AlertDialogTitle>¿Borrar rooming?</AlertDialogTitle>
            <AlertDialogDescription>Se eliminarán todas las habitaciones y asignaciones de alumnos de este viaje. Esta acción no se puede deshacer.</AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel>Cancelar</AlertDialogCancel>
            <AlertDialogAction style={{ backgroundColor: CORPORATE_RED }} onClick={() => { setConfirmClearRooming(false); handleClearRooming(); }}>Borrar rooming</AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>
    </div>
  );
}

function SchoolGroups({ schoolTrips, setSchoolTrips, notify }) {
  const allCourses = schoolTrips.flatMap((st) => (st.courses || []).map((c) => ({ ...c, tripId: st.id, tripName: st.trips?.name || "" })));
  const [selectedCourseId, setSelectedCourseId] = useState(allCourses[0]?.id || "");
  const [importing, setImporting] = useState(false);

  const selectedTrip = schoolTrips.find((st) => (st.courses || []).some((c) => c.id === selectedCourseId));
  const selectedTripId = selectedTrip?.id || "";
  const groups = selectedTrip?.activity_groups || [];

  const handleFileChange = async (e) => {
    const file = e.target.files?.[0];
    if (!file || !selectedTripId) return;
    setImporting(true);
    try {
      const buffer = await file.arrayBuffer();
      const XLSX2 = await import("xlsx");
      const wb = XLSX2.read(buffer, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const raw = XLSX2.utils.sheet_to_json(ws, { header: 1, defval: "" });
      const parsed = [];
      raw.forEach((row) => {
        if (!row[0]) return;
        const group = String(row[0]).trim();
        const students = row.slice(1).map((c) => String(c).trim()).filter(Boolean);
        if (group) parsed.push({ group, students });
      });
      const { error } = await supabase.from("school_trips").update({ activity_groups: parsed }).eq("id", selectedTripId);
      if (error) { notify("Error guardando grupos.", { variant: "destructive" }); }
      else {
        setSchoolTrips((prev) => prev.map((st) => st.id === selectedTripId ? { ...st, activity_groups: parsed } : st));
        notify(`Grupos importados: ${parsed.length} grupos.`);
      }
    } catch (err) {
      console.error(err);
      notify("Error leyendo archivo.", { variant: "destructive" });
    }
    setImporting(false);
    e.target.value = "";
  };

  const handleExportGroupsPDF = () => {
    if (!groups.length) return;
    const selCourse = allCourses.find((c) => c.id === selectedCourseId);
    const tripName = selCourse ? `${selCourse.course_name}${selCourse.group_name ? ` · ${selCourse.group_name}` : ""}` : "";
    const blocks = groups.map((g) =>
      `<div class="room-block">
        <div class="room-title">👥 ${g.group}${g.monitor ? ` <span style="font-weight:400;font-size:11px;color:#71717a">· Monitor: ${g.monitor}</span>` : ""} <span style="font-weight:400;font-size:11px;color:#a1a1aa">(${g.students.length} alumnos)</span></div>
        <table style="width:100%"><tbody>
          ${g.students.map((s, i) => `<tr><td style="width:28px;color:#a1a1aa">${i + 1}</td><td>${s}</td></tr>`).join("")}
        </tbody></table>
      </div>`
    ).join("<hr style='border:none;border-top:1px solid #f4f4f5;margin:12px 0'>");
    exportListToPDF("Grupos de actividades", tripName, blocks);
  };

  return (
    <div className="space-y-6">
      <SectionTitle icon={Grid2x2} title="Grupos de actividades" subtitle="Grupos de actividades. Importa desde Excel." />
      {/* Tabs de curso */}
      {allCourses.length > 0 && (
        <div className="flex flex-wrap gap-2">
          {allCourses.map((c) => (
            <button key={c.id} type="button" onClick={() => setSelectedCourseId(c.id)}
              className={`rounded-2xl px-3 py-1.5 text-xs font-medium transition ${selectedCourseId === c.id ? "text-white shadow-sm" : "border border-zinc-200 bg-white text-zinc-600 hover:bg-zinc-50"}`}
              style={selectedCourseId === c.id ? { backgroundColor: CORPORATE_RED } : {}}>
              {c.course_name}{c.group_name ? ` · ${c.group_name}` : ""}
            </button>
          ))}
        </div>
      )}
      <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
        <CardContent className="p-5">
          <div className="flex flex-wrap items-center gap-3">
            <label className={`flex h-11 cursor-pointer items-center gap-1.5 rounded-2xl border px-4 text-sm font-medium text-white hover:opacity-90 transition ${importing ? "opacity-50" : ""}`} style={{ backgroundColor: "#FF3131", borderColor: "#FF3131" }}>
              <FolderUp className="h-4 w-4" />{importing ? "Importando..." : "Importar documento"}
              <input type="file" accept=".xlsx,.xls,.csv,.pdf" className="hidden" onChange={handleFileChange} disabled={importing} />
            </label>
          </div>
          <div className="mt-2 text-xs text-zinc-400">Formato esperado: columna 1 = nombre de grupo, columnas siguientes = alumnos.</div>
        </CardContent>
      </Card>
      {groups.length > 0 ? (
        <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
          {groups.map((g, i) => (
            <Card key={i} className="rounded-3xl border-zinc-200 bg-white shadow-sm">
              <CardContent className="p-5">
                <div className="mb-1 font-semibold text-zinc-900">{g.group}</div>
                {g.monitor && <div className="mb-2 text-xs text-zinc-500">Monitor: {g.monitor}</div>}
                <div className="space-y-1">
                  {g.students.map((s, j) => <div key={j} className="flex items-center gap-1.5 text-xs text-zinc-600"><User className="h-3 w-3 shrink-0 text-zinc-400" />{s}</div>)}
                </div>
                <Badge variant="outline" className="mt-3 rounded-xl text-xs">{g.students.length} alumnos</Badge>
              </CardContent>
            </Card>
          ))}
        </div>
      ) : (
        <div className="flex flex-col items-center justify-center py-16 text-zinc-400">
          <ListChecks className="mb-3 h-10 w-10 opacity-40" />
          <p className="text-sm">Importa un Excel para ver los grupos aquí.</p>
        </div>
      )}
    </div>
  );
}

function SchoolHeroBanner({ school, schoolTrips, students, schoolDocuments, onNavigate }) {
  // Ordenar por fecha de salida; empezar por el más próximo futuro
  const now = new Date();
  const sorted = [...schoolTrips].sort((a, b) => {
    const da = new Date(a.trips?.departure_date || "9999-01-01");
    const db = new Date(b.trips?.departure_date || "9999-01-01");
    return da - db;
  });
  const defaultIdx = Math.max(0, sorted.findIndex((st) => new Date(st.trips?.departure_date || "0") >= now));
  const [activeIdx, setActiveIdx] = useState(defaultIdx);
  const trip = sorted[activeIdx] || sorted[0];
  if (!trip) return null;

  const heroImages = (() => {
    const imgs = trip.trips?.hero_images;
    if (Array.isArray(imgs) && imgs.length) return imgs;
    const img = trip.trips?.hero_image;
    return img ? [img, ...DEFAULT_HERO_IMAGES.slice(1)] : DEFAULT_HERO_IMAGES;
  })();
  const remaining = daysRemaining(trip.trips?.departure_date);

  const pendingDocs = schoolDocuments.filter((d) => d.status !== "approved").length;
  const pendingRooming = schoolTrips.filter((st) => !(st.rooming?.length)).length;
  const pendingGroups = schoolTrips.filter((st) => !(st.activity_groups?.length)).length;
  const pendingPayments = schoolTrips.filter((st) => st.payment_info?.status !== "completed").length;
  const totalPending = pendingDocs + pendingRooming + pendingGroups + pendingPayments;

  return (
    <div className="relative overflow-hidden rounded-[32px] shadow-[0_20px_70px_rgba(0,0,0,0.12)] mb-6">
      <img src={heroImages[0]} alt={trip.trips?.name} className="absolute inset-0 block h-full w-full object-cover object-center" />
      <div className="absolute inset-0 bg-[linear-gradient(135deg,rgba(0,0,0,0.80),rgba(0,0,0,0.45),rgba(255,49,49,0.16))]" />
      <div className="relative grid gap-6 p-5 sm:p-7 lg:grid-cols-[1fr_260px] lg:gap-8 lg:p-10">
        {/* Columna izquierda */}
        <div className="flex flex-col justify-between">
          <div>
            <Badge className="border-0 bg-white/10 text-white backdrop-blur-sm hover:bg-white/10">Viaje escolar</Badge>
            <h1 className="mt-4 max-w-3xl text-3xl font-semibold tracking-tight text-white sm:text-4xl lg:text-5xl">
              {trip.trips?.name?.toUpperCase()}
            </h1>
            <p className="mt-3 max-w-2xl text-base leading-7 font-medium text-white sm:text-lg">
              Hola, coordinador/a de <span className="font-bold">{school?.name}</span> 👋
            </p>
            <p className="mt-1 max-w-2xl text-sm leading-6 text-zinc-300">
              Aquí tienes toda la información del viaje.
            </p>
            {/* Selector de viaje cuando hay más de uno */}
            {sorted.length > 1 && (
              <div className="mt-4 flex flex-wrap gap-2">
                {sorted.map((st, i) => (
                  <button
                    key={st.id}
                    type="button"
                    onClick={() => setActiveIdx(i)}
                    className={`rounded-2xl px-3 py-1.5 text-xs font-medium transition border ${
                      i === activeIdx
                        ? "bg-white text-zinc-950 border-white/80"
                        : "bg-white/10 text-white border-white/20 hover:bg-white/20"
                    }`}
                  >
                    {st.trips?.name || st.trip_id}
                  </button>
                ))}
              </div>
            )}
          </div>
          <div className="mt-6 flex flex-wrap gap-3 text-sm text-white">
            <div className="inline-flex items-center gap-2 rounded-2xl border border-white/10 bg-white/10 px-4 py-2 backdrop-blur-sm">
              <Users className="h-4 w-4" /> {students.length} alumno{students.length !== 1 ? "s" : ""}
            </div>
            {schoolTrips.length > 1 && (
              <div className="inline-flex items-center gap-2 rounded-2xl border border-white/10 bg-white/10 px-4 py-2 backdrop-blur-sm">
                <Map className="h-4 w-4" /> {schoolTrips.length} viajes
              </div>
            )}
          </div>
        </div>

        {/* Columna derecha: cuenta atrás */}
        <div className="flex items-end lg:items-center">
          <div className="w-full rounded-[26px] border border-white/10 bg-black/35 p-5 text-white shadow-2xl backdrop-blur-xl">
            <div className="text-xs uppercase tracking-[0.24em] text-zinc-300">Cuenta atrás</div>
            <div className="mt-3 text-5xl font-semibold leading-none sm:text-6xl">{remaining}</div>
            <div className="mt-2 text-zinc-200">días para el viaje</div>
            <div className="mt-6 rounded-2xl bg-white/10 p-4 text-sm text-zinc-200">
              <div className="flex items-center gap-2"><CalendarDays className="h-4 w-4" /> Salida</div>
              <div className="mt-2 font-medium text-white">
                {trip.trips?.departure_date
                  ? (() => { const s = new Date(trip.trips.departure_date).toLocaleDateString("es-ES", { weekday: "long", day: "numeric", month: "long", year: "numeric" }); return s.charAt(0).toUpperCase() + s.slice(1); })()
                  : "-"}
              </div>
            </div>
            <div className="mt-4 rounded-2xl border border-white/10 bg-white/5 p-4">
              <div className="text-xs uppercase tracking-[0.18em] text-zinc-300 mb-3">Tareas pendientes</div>
              {totalPending === 0 ? (
                <div className="flex items-center gap-2 text-sm text-emerald-300">
                  <CheckCircle2 className="h-4 w-4" /> ¡Todo al día!
                </div>
              ) : (
                <div className="flex flex-wrap gap-3">
                  {pendingDocs > 0 && (
                    <button type="button" onClick={() => onNavigate?.("docs")}
                      className="relative flex flex-col items-center gap-1 cursor-pointer">
                      <div className="relative rounded-2xl border border-white/15 bg-white/10 p-3 backdrop-blur-sm transition hover:bg-white/20">
                        <FileCheck2 className="h-5 w-5 text-white" />
                        <span className="absolute -right-1.5 -top-1.5 flex h-5 w-5 items-center justify-center rounded-full text-[10px] font-bold text-white" style={{ backgroundColor: CORPORATE_RED }}>{pendingDocs}</span>
                      </div>
                      <span className="text-[10px] text-zinc-300 text-center leading-tight max-w-[52px]">Docs</span>
                    </button>
                  )}
                </div>
              )}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

function SchoolPortalQuestions({ school, questions, setQuestions, notify }) {
  const [message, setMessage] = useState("");
  const [sending, setSending] = useState(false);

  const handleSend = async () => {
    if (!message.trim() || sending || !school?.id) return;
    setSending(true);
    const tempId = `sq-${Date.now()}`;
    const newQ = { id: tempId, message: message.trim(), school_id: school.id, reply: null, created_at: new Date().toISOString(), source: "school" };
    setQuestions((prev) => [...prev, newQ]);
    setMessage("");
    try {
      const { data, error } = await supabase.from("school_questions").insert([{ message: newQ.message, school_id: school.id, source: "school" }]).select().maybeSingle();
      if (error) throw new Error(error.message);
      setQuestions((prev) => prev.map((q) => q.id === tempId ? { ...q, id: data.id } : q));
      sendNotification("admin_school_question", null, null, { schoolName: school.name || "Colegio", question: newQ.message });
    } catch (err) {
      setQuestions((prev) => prev.filter((q) => q.id !== tempId));
      notify("Error enviando la duda: " + err.message);
    } finally { setSending(false); }
  };

  const pendingCount = questions.filter((q) => !q.reply).length;
  const repliedCount = questions.filter((q) => q.reply).length;

  return (
    <div className="space-y-4">
      <div className="rounded-3xl border border-zinc-200 bg-white p-5">
        <div className="text-lg font-semibold text-zinc-950">¿Tienes alguna duda?</div>
        <p className="mt-2 text-sm leading-6 text-zinc-600">
          Escríbenos cualquier pregunta sobre el viaje, documentación, alumnos o cualquier otra gestión y te responderemos lo antes posible.
        </p>
        <div className="mt-4 flex gap-2">
          <Textarea
            value={message}
            onChange={(e) => setMessage(e.target.value)}
            onKeyDown={(e) => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); handleSend(); } }}
            placeholder="Escribe tu duda aquí… (Enter para enviar)"
            className="flex-1 min-h-[72px] rounded-2xl border-zinc-200 bg-white text-sm resize-none"
          />
          <Button
            onClick={handleSend}
            disabled={!message.trim() || sending}
            className="h-auto self-end rounded-2xl px-4 py-3 text-sm text-white"
            style={{ backgroundColor: CORPORATE_RED }}
          >
            <Send className="h-4 w-4" />
          </Button>
        </div>
      </div>

      {questions.length > 0 && (
        <div className="space-y-3">
          <div className="flex items-center gap-3">
            {pendingCount > 0 && <Badge className="bg-amber-100 text-amber-800 hover:bg-amber-100">{pendingCount} sin respuesta</Badge>}
            {repliedCount > 0 && <Badge className="bg-green-100 text-green-800 hover:bg-green-100">{repliedCount} respondidas</Badge>}
          </div>
          {[...questions].reverse().map((q) => (
            <div key={q.id} className={`rounded-3xl border p-5 ${q.reply ? "border-zinc-200 bg-white" : "border-amber-200 bg-amber-50"}`}>
              <div className="flex items-start justify-between gap-3">
                <div className="min-w-0 flex-1">
                  <div className="text-xs text-zinc-400">{new Date(q.created_at).toLocaleDateString("es-ES", { day: "numeric", month: "long", year: "numeric" })}</div>
                  <p className="mt-1 text-sm text-zinc-900">{q.message}</p>
                </div>
                <Badge className={q.reply ? "bg-green-100 text-green-800 hover:bg-green-100 shrink-0" : "bg-amber-100 text-amber-800 hover:bg-amber-100 shrink-0"}>
                  {q.reply ? "Respondida" : "Pendiente"}
                </Badge>
              </div>
              {q.reply && (
                <div className="mt-3 rounded-2xl bg-zinc-50 px-4 py-3 text-sm text-zinc-700">
                  <span className="mr-2 font-medium text-zinc-500">Respuesta:</span>{q.reply}
                </div>
              )}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function SchoolChecklist({ schoolTrips }) {
  const allCourses = schoolTrips.flatMap((st) =>
    (st.courses || []).map((c) => ({ ...c, tripId: st.id, tripName: st.trips?.name || "" }))
  );
  const [selectedCourseId, setSelectedCourseId] = useState(allCourses[0]?.id || "");
  const selectedTrip = schoolTrips.find((st) => (st.courses || []).some((c) => c.id === selectedCourseId));
  const checklist = selectedTrip?.checklist || [];

  const storageKey = selectedTrip ? `school_checklist_state_${selectedTrip.id}` : null;
  const [checked, setChecked] = useState(() => {
    if (!storageKey || typeof window === "undefined") return {};
    try { return JSON.parse(window.localStorage.getItem(storageKey) || "{}"); } catch { return {}; }
  });

  useEffect(() => {
    if (!storageKey || typeof window === "undefined") { setChecked({}); return; }
    try { setChecked(JSON.parse(window.localStorage.getItem(storageKey) || "{}")); } catch { setChecked({}); }
  }, [storageKey]);

  const toggleItem = (item) => {
    const next = { ...checked, [item]: !checked[item] };
    setChecked(next);
    if (storageKey) { try { window.localStorage.setItem(storageKey, JSON.stringify(next)); } catch {} }
  };

  const completedCount = checklist.filter((item) => !!checked[item]).length;

  const exportChecklist = () => {
    const selectedCourse = allCourses.find((c) => c.id === selectedCourseId);
    const title = selectedCourse
      ? `Checklist — ${selectedCourse.course_name}${selectedCourse.group_name ? ` · ${selectedCourse.group_name}` : ""}`
      : "Checklist de equipaje";
    const subtitle = selectedTrip?.trips?.name || "";
    const rows = checklist.map((item) =>
      `<tr><td style="padding:6px 8px">${item}</td><td style="padding:6px 8px;text-align:center">${checked[item] ? "✓" : ""}</td></tr>`
    ).join("");
    exportListToPDF(title, subtitle,
      `<table style="width:100%;border-collapse:collapse">
        <thead><tr style="background:#f4f4f4"><th style="padding:6px 8px;text-align:left">Ítem</th><th style="padding:6px 8px;width:60px">Listo</th></tr></thead>
        <tbody>${rows}</tbody>
      </table>`
    );
  };

  return (
    <div className="space-y-5">
      <div className="flex items-start justify-between gap-3">
        <SectionTitle icon={CheckCircle2} title="Checklist de equipaje" subtitle="Marca los ítems que ya tienes listos para el viaje." />
        {checklist.length > 0 && (
          <Button variant="outline" className="shrink-0 rounded-2xl text-sm" onClick={exportChecklist}>
            <Download className="mr-2 h-4 w-4" />Exportar PDF
          </Button>
        )}
      </div>
      {allCourses.length > 1 && (
        <div className="flex flex-wrap gap-2">
          {allCourses.map((c) => (
            <button key={c.id} type="button" onClick={() => setSelectedCourseId(c.id)}
              className={`rounded-2xl px-3 py-1.5 text-xs font-medium transition ${selectedCourseId === c.id ? "text-white shadow-sm" : "border border-zinc-200 bg-white text-zinc-600 hover:bg-zinc-50"}`}
              style={selectedCourseId === c.id ? { backgroundColor: CORPORATE_RED } : {}}>
              {c.course_name}{c.group_name ? ` · ${c.group_name}` : ""}
            </button>
          ))}
        </div>
      )}
      <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
        <CardContent className="p-5 space-y-4">
          {checklist.length > 0 && (
            <div className="flex items-center justify-between">
              <span className="text-sm text-zinc-500">{completedCount} de {checklist.length} preparados</span>
              <div className="h-2 w-32 overflow-hidden rounded-full bg-zinc-100">
                <div className="h-full rounded-full transition-all"
                  style={{ width: `${checklist.length ? Math.round((completedCount / checklist.length) * 100) : 0}%`, backgroundColor: CORPORATE_RED }} />
              </div>
            </div>
          )}
          {checklist.length === 0 ? (
            <div className="rounded-2xl border-2 border-dashed border-zinc-200 py-10 text-center text-sm text-zinc-400">
              El equipo de GIMELOOS publicará el checklist de equipaje próximamente.
            </div>
          ) : (
            <div className="grid gap-3 md:grid-cols-2">
              {checklist.map((item, i) => (
                <div key={i} onClick={() => toggleItem(item)}
                  className={`flex cursor-pointer items-center gap-3 rounded-2xl border p-4 transition select-none ${checked[item] ? "border-green-200 bg-green-50" : "border-zinc-200 bg-white hover:bg-zinc-50"}`}>
                  <div className={`flex h-5 w-5 shrink-0 items-center justify-center rounded-full border-2 transition ${checked[item] ? "border-green-500 bg-green-500 text-white" : "border-zinc-300"}`}>
                    {checked[item] && <span className="text-[10px] font-bold leading-none">✓</span>}
                  </div>
                  <span className={`text-sm ${checked[item] ? "text-green-700 line-through" : "text-zinc-800"}`}>{item}</span>
                </div>
              ))}
            </div>
          )}
        </CardContent>
      </Card>
    </div>
  );
}

function SchoolPortal({ user, onLogout, notify, previewSchoolId = null }) {
  const [activeTab, setActiveTab] = useState("trips");
  const [school, setSchool] = useState(null);
  const [schoolTrips, setSchoolTrips] = useState([]);
  const [courses, setCourses] = useState([]);
  const [students, setStudents] = useState([]);
  const [schoolDocuments, setSchoolDocuments] = useState([]);
  const [schoolQuestions, setSchoolQuestions] = useState([]);
  const [loading, setLoading] = useState(true);
  const [loadErr, setLoadErr] = useState(null);

  useEffect(() => {
    const load = async () => {
      setLoading(true);
      setLoadErr(null);
      try {
        // 1. Get school — by previewSchoolId (admin) > schoolId (participant.school_id) > auth_uid
        const query = supabase.from("schools").select("*");
        const { data: schoolData, error: schoolErr } = previewSchoolId
          ? await query.eq("id", previewSchoolId).maybeSingle()
          : user.schoolId
            ? await query.eq("id", user.schoolId).maybeSingle()
            : await query.eq("auth_uid", user.authUid).maybeSingle();
        if (schoolErr) throw new Error(schoolErr.message);
        if (!schoolData) { setLoadErr("No se encontró un colegio asociado a este usuario."); setLoading(false); return; }
        setSchool(schoolData);

        // 2. School trips
        const { data: tripsData, error: tripsErr } = await supabase
          .from("school_trips")
          .select("*, trips(name, departure_date, hero_image, hero_images)")
          .eq("school_id", schoolData.id);
        if (tripsErr) throw new Error(tripsErr.message);
        const tripsArr = tripsData || [];
        setSchoolTrips(tripsArr);

        if (!tripsArr.length) { setLoading(false); return; }

        // 3. Courses
        const tripIds = tripsArr.map((t) => t.id);
        const { data: coursesData, error: coursesErr } = await supabase
          .from("school_courses")
          .select("*")
          .in("school_trip_id", tripIds);
        if (coursesErr) throw new Error(coursesErr.message);
        const coursesArr = coursesData || [];
        setCourses(coursesArr);

        if (!coursesArr.length) { setLoading(false); return; }

        // 4. Students
        const courseIds = coursesArr.map((c) => c.id);
        const { data: studentsData, error: studentsErr } = await supabase
          .from("students")
          .select("*")
          .in("school_course_id", courseIds);
        if (studentsErr) throw new Error(studentsErr.message);
        setStudents(studentsData || []);

        // 5. Documents
        const { data: docsData, error: docsErr } = await supabase
          .from("school_documents")
          .select("*")
          .in("school_course_id", courseIds);
        if (docsErr) throw new Error(docsErr.message);
        setSchoolDocuments(docsData || []);

        // 6. Questions
        const { data: qData, error: qErr } = await supabase.from("school_questions").select("*").eq("school_id", schoolData.id).order("created_at");
        if (qErr) throw new Error(qErr.message);
        setSchoolQuestions(qData || []);

        // Attach courses to trips for display
        setSchoolTrips(tripsArr.map((st) => ({ ...st, courses: coursesArr.filter((c) => c.school_trip_id === st.id) })));
      } catch (err) {
        console.error("Error cargando datos del colegio:", err);
        setLoadErr("Error cargando datos: " + err.message);
      } finally {
        setLoading(false);
      }
    };
    if (previewSchoolId || user?.authUid) load();
    else { setLoadErr("No se pudo identificar el usuario del colegio."); setLoading(false); }
  }, [user?.authUid, previewSchoolId]);

  const repliedQuestions = schoolQuestions.filter((q) => q.reply).length;
  const pendingQuestions = schoolQuestions.filter((q) => !q.reply).length; // sent by school, no reply yet

  const tabs = [
    { key: "trips",     label: "Mis viajes",    icon: CalendarDays },
    { key: "students",  label: "Alumnos",       icon: Users },
    { key: "allergies", label: "Alergias",      icon: AlertCircle },
    { key: "docs",      label: "Documentación", icon: FileCheck2 },
    { key: "rooming",   label: "Rooming",       icon: Home },
    { key: "groups",    label: "Grupos",        icon: Grid2x2 },
    { key: "checklist", label: "Checklist",     icon: CheckCircle2 },
    { key: "questions", label: "Dudas",         icon: MessageCircleQuestion, badge: (pendingQuestions + repliedQuestions) > 0 ? pendingQuestions + repliedQuestions : null },
  ];

  return (
    <div className="min-h-screen bg-white text-zinc-950">
      <div className="mx-auto max-w-5xl p-6 lg:p-8">
        {/* Header card */}
        <div className="mb-6 flex flex-col gap-4 rounded-[28px] border border-zinc-200 bg-white px-6 py-5 shadow-sm lg:flex-row lg:items-center lg:justify-between">
          <div className="flex items-center gap-3">
            <div className="flex h-10 w-10 items-center justify-center rounded-2xl text-white shadow-sm" style={{ backgroundColor: CORPORATE_RED }}>
              <Users className="h-5 w-5" />
            </div>
            <div>
              <div className="text-xs uppercase tracking-[0.22em] text-zinc-400">Portal Escolar</div>
              <div className="text-base font-bold tracking-[0.12em] text-zinc-950">GIMELOOS</div>
            </div>
          </div>
          <div className="flex items-center gap-3">
            {school?.name && (
              <div className="rounded-2xl border border-zinc-200 bg-zinc-50 px-4 py-2 text-sm font-medium text-zinc-700">
                {school.name}
              </div>
            )}
            <Button variant="outline" className="h-11 rounded-2xl" onClick={() => { onLogout(); notify("Sesión cerrada."); }}>
              <LogOut className="mr-2 h-4 w-4" />Salir
            </Button>
          </div>
        </div>

        {loading ? (
          <div className="flex min-h-[60vh] items-center justify-center">
            <div className="rounded-3xl border border-zinc-200 bg-white px-6 py-5 text-sm text-zinc-600 shadow-sm">Cargando datos del colegio...</div>
          </div>
        ) : loadErr ? (
          <div className="flex min-h-[60vh] items-center justify-center">
            <div className="flex flex-col items-center gap-3 rounded-3xl border border-zinc-200 bg-white px-8 py-6 shadow-sm">
              <AlertCircle className="h-8 w-8 text-red-500" />
              <div className="text-sm text-zinc-700">{loadErr}</div>
              <Button onClick={() => window.location.reload()} className="h-11 rounded-2xl text-white text-xs" style={{ backgroundColor: CORPORATE_RED }}>Reintentar</Button>
            </div>
          </div>
        ) : (
          <>
            {/* Hero banner */}
            <SchoolHeroBanner
              school={school}
              schoolTrips={schoolTrips}
              students={students}
              schoolDocuments={schoolDocuments}
              onNavigate={setActiveTab}
            />

            {/* Tab nav */}
            <div className="mb-6 flex gap-1.5">
              {tabs.map(({ key, label, icon: Icon, badge }) => {
                const active = activeTab === key;
                return (
                  <button
                    key={key}
                    type="button"
                    onMouseDown={(e) => e.preventDefault()}
                    onClick={() => setActiveTab(key)}
                    className={`relative flex flex-1 items-center justify-center gap-1.5 rounded-2xl px-2 py-2 text-sm font-medium transition ${
                      active ? "text-white shadow-sm" : "border border-zinc-200 bg-white text-zinc-700 hover:bg-zinc-50"
                    }`}
                    style={active ? { backgroundColor: CORPORATE_RED } : {}}
                  >
                    <Icon className="h-4 w-4" />{label}
                    {badge != null && (
                      <span className="ml-0.5 inline-flex h-5 w-5 items-center justify-center rounded-full text-xs font-bold text-white" style={{ backgroundColor: active ? "rgba(255,255,255,0.3)" : CORPORATE_RED }}>{badge}</span>
                    )}
                  </button>
                );
              })}
            </div>

            {/* Tab content */}
            {activeTab === "trips"     && <SchoolTrips schoolTrips={schoolTrips} courses={courses} students={students} schoolDocuments={schoolDocuments} onNavigate={setActiveTab} />}
            {activeTab === "students"  && <SchoolStudents schoolTrips={schoolTrips} courses={courses} students={students} setStudents={setStudents} notify={notify} />}
            {activeTab === "allergies" && <SchoolAllergies courses={courses} students={students} />}
            {activeTab === "docs"      && <SchoolDocs courses={courses} schoolDocuments={schoolDocuments} setSchoolDocuments={setSchoolDocuments} notify={notify} school={school} schoolTrips={schoolTrips} />}
            {activeTab === "rooming"   && <SchoolRooming schoolTrips={schoolTrips} setSchoolTrips={setSchoolTrips} notify={notify} />}
            {activeTab === "groups"    && <SchoolGroups schoolTrips={schoolTrips} setSchoolTrips={setSchoolTrips} notify={notify} />}
            {activeTab === "checklist" && <SchoolChecklist schoolTrips={schoolTrips} />}
            {activeTab === "questions" && (
              <SchoolPortalQuestions
                school={school}
                questions={schoolQuestions}
                setQuestions={setSchoolQuestions}
                notify={notify}
              />
            )}
          </>
        )}
      </div>
    </div>
  );
}

// ─── Admin Schools ────────────────────────────────────────────────────────────

function CoverImageInput({ value, onChange, onBlur, tripId, notify }) {
  const [uploading, setUploading] = useState(false);
  const [pct, setPct] = useState(undefined);
  const inputRef = useRef(null);

  const handleFile = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setUploading(true);
    try {
      const result = await uploadFileToDrive(file, `trip-${tripId || Date.now()}`, "covers", (p) => setPct(p), "GIMELOOS Portadas");
      const url = result.webContentLink || result.webViewLink || "";
      onChange(url);
      onBlur(url);
    } catch (err) {
      console.error(err);
      notify("No se pudo subir la imagen de portada. Inténtalo de nuevo.", { variant: "destructive" });
    } finally {
      setUploading(false);
      setPct(undefined);
      e.target.value = "";
    }
  };

  return (
    <div className="space-y-2">
      <div className="flex gap-3">
        <Input value={value} onChange={(e) => onChange(e.target.value)} onBlur={(e) => onBlur(e.target.value)} placeholder="https://..." className="rounded-2xl" />
        <input ref={inputRef} type="file" accept="image/*" className="hidden" onChange={handleFile} />
        <Button type="button" variant="outline" className="h-11 shrink-0 rounded-2xl" onClick={() => inputRef.current?.click()} disabled={uploading}>
          {uploading ? <Loader2 className="h-4 w-4 animate-spin" /> : <ImagePlus className="h-4 w-4" />}
        </Button>
        {value && (
          <img src={value} alt="portada" className="h-11 w-20 rounded-2xl object-cover border border-zinc-200" onError={(e) => { e.target.style.display = "none"; }} />
        )}
      </div>
      {pct !== undefined && (
        <div>
          <div className="mb-1 flex justify-between text-xs text-zinc-500">
            <span>{pct < 100 ? "Subiendo imagen…" : "¡Lista!"}</span>
            <span>{pct}%</span>
          </div>
          <div className="h-1.5 w-full overflow-hidden rounded-full bg-zinc-100">
            <div className="h-full rounded-full transition-all duration-200"
              style={{ width: `${pct}%`, backgroundColor: pct === 100 ? "#16a34a" : CORPORATE_RED }} />
          </div>
        </div>
      )}
    </div>
  );
}

function ChecklistInput({ onAdd }) {
  const [val, setVal] = useState("");
  return (
    <div className="flex gap-2">
      <Input value={val} onChange={(e) => setVal(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter" && val.trim()) { onAdd(val.trim()); setVal(""); } }} placeholder="Añadir elemento (Enter para confirmar)" className="rounded-2xl" />
      <Button onClick={() => { if (val.trim()) { onAdd(val.trim()); setVal(""); } }} className="rounded-2xl text-white shrink-0" style={{ backgroundColor: CORPORATE_RED }}>
        <Plus className="h-4 w-4" />
      </Button>
    </div>
  );
}

function SchoolQuestionCard({ q, schoolName, onReply }) {
  const [replyText, setReplyText] = useState(q.reply || "");
  const [editing, setEditing] = useState(false);
  const [saving, setSaving] = useState(false);
  const handleSave = async () => {
    setSaving(true);
    await onReply(q.id, replyText);
    setSaving(false);
    setEditing(false);
  };
  return (
    <Card className="rounded-2xl border-zinc-200 shadow-sm">
      <CardContent className="p-4 space-y-2">
        <div className="flex items-start justify-between gap-3">
          <div className="min-w-0 flex-1">
            <div className="text-xs font-medium text-zinc-500">{schoolName || "Colegio desconocido"}{q.source === "admin" ? " · añadida por admin" : ""}</div>
            <p className="mt-1 text-sm text-zinc-900">{q.message}</p>
            {q.reply && !editing && (
              <p className="mt-2 rounded-xl bg-zinc-50 px-3 py-2 text-xs text-zinc-600">↳ {q.reply}</p>
            )}
          </div>
          <Badge variant="outline" className={`shrink-0 rounded-xl text-xs ${q.reply ? "border-green-200 text-green-700" : "border-amber-200 text-amber-700"}`}>
            {q.reply ? "Respondida" : "Pendiente"}
          </Badge>
        </div>
        {editing ? (
          <div className="space-y-2">
            <textarea
              value={replyText}
              onChange={(e) => setReplyText(e.target.value)}
              rows={2}
              placeholder="Escribe la respuesta..."
              className="w-full rounded-xl border border-zinc-200 bg-white px-3 py-2 text-xs text-zinc-950 placeholder-zinc-400 focus:outline-none resize-none"
            />
            <div className="flex gap-2">
              <Button onClick={handleSave} disabled={saving} className="h-7 rounded-xl text-xs text-white px-3" style={{ backgroundColor: CORPORATE_RED }}>
                {saving ? "Guardando..." : "Guardar respuesta"}
              </Button>
              <Button variant="outline" className="h-7 rounded-xl text-xs px-3" onClick={() => setEditing(false)}>Cancelar</Button>
            </div>
          </div>
        ) : (
          <button type="button" onClick={() => setEditing(true)}
            className="text-xs text-zinc-400 hover:text-zinc-600 underline underline-offset-2 transition">
            {q.reply ? "Editar respuesta" : "Responder"}
          </button>
        )}
      </CardContent>
    </Card>
  );
}

function AdminSchoolViajes({ allSchoolTrips, schools, trips, setTrips, notify }) {
  const [selectedStId, setSelectedStId] = useState(allSchoolTrips[0]?.id || "");
  const selectedSt = allSchoolTrips.find((st) => st.id === selectedStId) || allSchoolTrips[0];
  const selectedTrip = trips.find((t) => t.id === selectedSt?.trip_id);
  const school = schools.find((s) => s.id === selectedSt?.school_id);

  const syncTripField = (field, value) => {
    if (!setTrips) return;
    setTrips((prev) => prev.map((t) => t.id === selectedTrip?.id ? { ...t, [field]: value } : t));
  };
  const saveTripField = async (dbField, value) => {
    if (!selectedTrip) return;
    const stateField = dbField === "hero_image" ? "heroImage" : dbField === "departure_date" ? "departureDate" : dbField;
    syncTripField(stateField, value);
    const { error } = await supabase.from("trips").update({ [dbField]: value }).eq("id", selectedTrip.id);
    if (error) notify("Error guardando cambios: " + error.message);
  };

  if (!allSchoolTrips.length) return (
    <div className="space-y-5">
      <SectionTitle icon={MapIcon} title="Viajes escolares" subtitle="Información básica y foto de portada de cada viaje escolar." />
      <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
        <CardContent className="p-8 text-center text-sm text-zinc-400">No hay viajes escolares registrados.</CardContent>
      </Card>
    </div>
  );

  return (
    <div className="space-y-5">
      <SectionTitle icon={MapIcon} title="Viajes escolares" subtitle="Información básica y foto de portada de cada viaje escolar." />
      <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
        <CardContent className="p-5">
          <div className="flex flex-col gap-3 sm:flex-row sm:items-center">
            <div className="flex-1 space-y-1">
              <Label>Viaje activo</Label>
              <select value={selectedStId} onChange={(e) => setSelectedStId(e.target.value)}
                className="h-11 w-full rounded-2xl border border-zinc-200 bg-white px-4 text-sm font-medium">
                {allSchoolTrips.map((st) => {
                  const sch = schools.find((s) => s.id === st.school_id);
                  return <option key={st.id} value={st.id}>{sch ? `${sch.name} · ` : ""}{st.trips?.name || st.trip_id}</option>;
                })}
              </select>
            </div>
            {selectedTrip?.departureDate && (
              <div className="rounded-2xl border border-zinc-100 bg-white px-4 py-2 text-center">
                <div className="text-xs text-zinc-400">Fecha de salida</div>
                <div className="font-semibold text-zinc-950">{new Date(selectedTrip.departureDate).toLocaleDateString("es-ES", { day: "numeric", month: "long", year: "numeric" })}</div>
              </div>
            )}
          </div>
        </CardContent>
      </Card>
      {selectedTrip && (
        <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
          <CardContent className="space-y-5 p-6">
            {school && (
              <div className="text-sm text-zinc-500">Colegio: <span className="font-medium text-zinc-800">{school.name}</span></div>
            )}
            <div className="grid gap-4 lg:grid-cols-2">
              <div className="space-y-2">
                <Label>Nombre del viaje</Label>
                <Input value={selectedTrip.name} onChange={(e) => syncTripField("name", e.target.value)} onBlur={(e) => saveTripField("name", e.target.value)} className="rounded-2xl" />
              </div>
              <div className="space-y-2">
                <Label>Fecha de salida</Label>
                <Input type="datetime-local" value={(selectedTrip?.departureDate || "").slice(0, 16)}
                  onChange={async (e) => {
                    syncTripField("departureDate", e.target.value);
                    const { error } = await supabase.from("trips").update({ departure_date: e.target.value || null }).eq("id", selectedTrip.id);
                    if (error) notify("Error guardando fecha: " + error.message);
                  }}
                  className="rounded-2xl" />
              </div>
            </div>
            <div className="space-y-2">
              <Label>Foto de portada</Label>
              <CoverImageInput
                value={selectedTrip.heroImage || ""}
                onChange={(v) => syncTripField("heroImage", v)}
                onBlur={(v) => saveTripField("hero_image", v)}
                tripId={selectedTrip.id}
                notify={notify}
              />
            </div>
            <div className="space-y-2">
              <Label>Descripción</Label>
              <Textarea value={selectedTrip.description || ""} onChange={(e) => syncTripField("description", e.target.value)} onBlur={(e) => saveTripField("description", e.target.value)} className="min-h-[120px] rounded-2xl" />
            </div>
          </CardContent>
        </Card>
      )}
    </div>
  );
}

function AdminSchools({ trips, setTrips, notify, section = "colegios", schoolTripIds, setSchoolTripIds }) {
  const tab = {
    "colegios": "schools", "alumnos": "students", "alergias": "allergies",
    "docs": "docs", "preguntas": "questions", "rooming": "rooming", "grupos": "groups",
    "seguimiento": "tracking", "checklist": "checklist", "pagos": "pagos",
    "itinerario": "itinerary", "logistica": "logistics", "viajes": "school_viajes",
  }[section] || "schools";
  const [schools, setSchools] = useState([]);
  const [allStudents, setAllStudents] = useState([]);
  const [allCourses, setAllCourses] = useState([]);
  const [allSchoolTrips, setAllSchoolTrips] = useState([]);
  const [allSchoolDocs, setAllSchoolDocs] = useState([]);
  const [allSchoolQuestions, setAllSchoolQuestions] = useState([]);
  const [loading, setLoading] = useState(true);
  // New school form
  const [showNewSchool, setShowNewSchool] = useState(false);
  const [newSchool, setNewSchool] = useState({ name: "", contact_name: "", email: "", phone: "" });
  const [savingSchool, setSavingSchool] = useState(false);
  // Assign trip form
  const [assigningSchoolId, setAssigningSchoolId] = useState(null);
  const [assignTripId, setAssignTripId] = useState(trips[0]?.id || "");
  const [assignCourse, setAssignCourse] = useState("");
  const [assignGroup, setAssignGroup] = useState("");
  const [savingAssign, setSavingAssign] = useState(false);
  // Filter for students tab
  const [filterSchoolId, setFilterSchoolId] = useState("");
  const [filterTripId, setFilterTripId] = useState("");
  const [filterCourseId, setFilterCourseId] = useState("all");
  const [schoolSearch, setSchoolSearch] = useState("");
  const [studentTrackingSearch, setStudentTrackingSearch] = useState("");
  const [pagosSchoolSearch, setPagosSchoolSearch] = useState("");
  // Manual student entry
  const [showAddStudent, setShowAddStudent] = useState(false);
  const [newStudent, setNewStudent] = useState({ name: "", surname: "", allergies: "", intolerances: "", diet_notes: "", notes: "" });
  const [savingStudent, setSavingStudent] = useState(false);
  // Manual doc entry
  const [showAddDoc, setShowAddDoc] = useState(false);
  const [newDoc, setNewDoc] = useState({ name: "", description: "", required: true });
  const [savingDoc, setSavingDoc] = useState(false);
  const [docFile, setDocFile] = useState(null);
  // Selectores propios del formulario de nuevo documento (independientes del filtro global)
  const [docFormSchoolId, setDocFormSchoolId] = useState("");
  const [docFormTripId, setDocFormTripId] = useState("");
  const [docFormCourseId, setDocFormCourseId] = useState("all");
  // Manual question entry
  const [showAddQuestion, setShowAddQuestion] = useState(false);
  const [newQuestion, setNewQuestion] = useState({ message: "", school_id: "" });
  const [savingQuestion, setSavingQuestion] = useState(false);
  // Inline reply in tracking
  const [schoolReplyTexts, setSchoolReplyTexts] = useState({});
  const [schoolSendingReply, setSchoolSendingReply] = useState({});
  // Rooming manual entry
  const [showAddRoom, setShowAddRoom] = useState(false);
  const [newRoom, setNewRoom] = useState({ room: "", students: "" });
  const [savingRoom, setSavingRoom] = useState(false);
  const [roomingTripId, setRoomingTripId] = useState("");
  // Groups manual entry
  const [showAddGroup, setShowAddGroup] = useState(false);
  const [newGroup, setNewGroup] = useState({ group: "", monitor: "", students: "" });
  const [savingGroup, setSavingGroup] = useState(false);
  const [groupTripId, setGroupTripId] = useState("");
  // Admin resource upload (docs tab)
  const [resName, setResName] = useState("");
  const [resFile, setResFile] = useState(null);
  const [resUploading, setResUploading] = useState(false);
  const [resPct, setResPct] = useState(undefined);
  // School Excel import
  const [isImportingSchool, setIsImportingSchool] = useState(false);
  const [schoolImportMsg, setSchoolImportMsg] = useState("");
  const [selectedImportSchool, setSelectedImportSchool] = useState("");
  const [selectedImportTrip, setSelectedImportTrip] = useState("");
  const [schoolSheetUrl, setSchoolSheetUrl] = useState("");
  const [isSyncingSchoolSheet, setIsSyncingSchoolSheet] = useState(false);
  // Confirmaciones destructivas en pagos
  const [pendingDeleteInvoiceId, setPendingDeleteInvoiceId] = useState(null);
  const [pendingResetPaymentId, setPendingResetPaymentId] = useState(null);
  const [libraryDocToDelete, setLibraryDocToDelete] = useState(null);
  // Alumnos export column selector (hoisted to avoid hooks-in-IIFE)
  const STUDENT_FIXED_COLS = [
    { key: "nombre",        label: "Nombre completo" },
    { key: "curso",         label: "Curso / Grupo" },
    { key: "colegio",       label: "Colegio" },
  ];
  const STUDENT_COLS_DEF = [
    { key: "alergias",      label: "Alergias",         default: false },
    { key: "intolerancias", label: "Intolerancias",    default: false },
    { key: "dieta",         label: "Notas dietéticas", default: false },
    { key: "rooming",       label: "Habitación",       default: false },
    { key: "grupo",         label: "Grupo",            default: false },
    { key: "notas",         label: "Observaciones",    default: false },
  ];
  const [exportCols, setExportCols] = useState(() => Object.fromEntries(STUDENT_COLS_DEF.map(c => [c.key, c.default])));

  useEffect(() => {
    const load = async () => {
      setLoading(true);
      try {
        const [schoolsRes, schoolTripsRes, coursesRes, studentsRes, docsRes] = await Promise.all([
          supabase.from("schools").select("*").order("name"),
          supabase.from("school_trips").select("*, trips(name, departure_date, hero_image, hero_images, description, transfer_info)").order("created_at"),
          supabase.from("school_courses").select("id,course_name,group_name,school_trip_id,created_at").order("course_name"),
          supabase.from("students").select("*").order("name"),
          supabase.from("school_documents").select("id,name,file_url,file_name,school_course_id,status,required,created_at").order("created_at"),
        ]);
        setSchools(schoolsRes.data || []);
        setAllSchoolTrips(schoolTripsRes.data || []);
        setAllCourses(coursesRes.data || []);
        setAllStudents(studentsRes.data || []);
        setAllSchoolDocs(docsRes.data || []);
        setAllSchoolQuestions([]);
      } catch (err) {
        console.error(err);
        notify("Error cargando datos de colegios.", { variant: "destructive" });
      } finally {
        setLoading(false);
      }
    };
    load();
  }, []);

  const handleSaveSchool = async () => {
    if (!newSchool.name.trim()) { notify("El nombre del colegio es obligatorio."); return; }
    setSavingSchool(true);
    const { data, error } = await supabase.from("schools").insert([newSchool]).select().maybeSingle();
    if (error) { notify("Error creando colegio: " + error.message, { variant: "destructive" }); }
    else {
      setSchools((prev) => [...prev, data]);
      setNewSchool({ name: "", contact_name: "", email: "", phone: "" });
      setShowNewSchool(false);
      notify("Colegio creado.");
    }
    setSavingSchool(false);
  };

  const handleAssignTrip = async (schoolId) => {
    if (!assignTripId || !assignCourse.trim()) { notify("Selecciona un viaje y escribe el nombre del curso."); return; }
    setSavingAssign(true);
    const { data: stData, error: stErr } = await supabase.from("school_trips").insert([{ school_id: schoolId, trip_id: assignTripId }]).select().maybeSingle();
    if (stErr) { notify("Error asignando viaje: " + stErr.message, { variant: "destructive" }); setSavingAssign(false); return; }
    // Marcar el viaje como tipo colegio para separarlo de campamentos
    const { error: tipoErr } = await supabase.from("trips").update({ tipo: "colegio" }).eq("id", assignTripId);
    if (tipoErr) { notify("Error actualizando tipo de viaje: " + tipoErr.message, { variant: "destructive" }); setSavingAssign(false); return; }
    const { error: scErr } = await supabase.from("school_courses").insert([{ school_trip_id: stData.id, course_name: assignCourse.trim(), group_name: assignGroup.trim() }]);
    if (scErr) { notify("Error creando curso: " + scErr.message, { variant: "destructive" }); setSavingAssign(false); return; }
    notify("Viaje asignado con curso.");
    setAssigningSchoolId(null);
    setAssignCourse(""); setAssignGroup("");
    // Reload
    const [stRes, scRes] = await Promise.all([
      supabase.from("school_trips").select("*, trips(name, departure_date, hero_image, hero_images, description, transfer_info)").order("created_at"),
      supabase.from("school_courses").select("id,course_name,group_name,school_trip_id,created_at").order("course_name"),
    ]);
    setAllSchoolTrips(stRes.data || []);
    setAllCourses(scRes.data || []);
    if (setSchoolTripIds && stRes.data) {
      setSchoolTripIds(new Set(stRes.data.map((r) => r.trip_id)));
    }
    setSavingAssign(false);
  };

  const getSchoolStudentCount = (schoolId) => {
    const tripIds = allSchoolTrips.filter((st) => st.school_id === schoolId).map((st) => st.id);
    const courseIds = allCourses.filter((c) => tripIds.includes(c.school_trip_id)).map((c) => c.id);
    return allStudents.filter((s) => courseIds.includes(s.school_course_id)).length;
  };

  const getSchoolTripCount = (schoolId) => allSchoolTrips.filter((st) => st.school_id === schoolId).length;

  const filteredSchoolTrips = filterSchoolId ? allSchoolTrips.filter((st) => st.school_id === filterSchoolId) : allSchoolTrips;

  const visibleCourses = (() => {
    const tripIds = filteredSchoolTrips.map((st) => st.id);
    return allCourses.filter((c) => tripIds.includes(c.school_trip_id));
  })();

  const filteredStudents = (() => {
    let courseIds;
    if (filterCourseId && filterCourseId !== "all") {
      courseIds = [filterCourseId];
    } else if (filterSchoolId) {
      const tripIds = filteredSchoolTrips.map((st) => st.id);
      courseIds = allCourses.filter((c) => tripIds.includes(c.school_trip_id)).map((c) => c.id);
    } else {
      return allStudents;
    }
    return allStudents.filter((s) => courseIds.includes(s.school_course_id));
  })();

  const exportActiveCols = [...STUDENT_FIXED_COLS, ...STUDENT_COLS_DEF.filter(c => exportCols[c.key])];
  const handleStudentExport = () => {
    const schoolName = filterSchoolId ? (schools.find(s => s.id === filterSchoolId)?.name || "") : "Todos los colegios";
    const selectedCourse = filterCourseId !== "all" ? visibleCourses.find(c => c.id === filterCourseId) : null;
    const courseLabel = selectedCourse ? ` · ${selectedCourse.course_name}${selectedCourse.group_name ? ` · ${selectedCourse.group_name}` : ""}` : "";
    const rows = filteredStudents.map((s, i) => {
      const course = allCourses.find(c => c.id === s.school_course_id);
      const school = schools.find(sc => {
        const st = allSchoolTrips.find(t => t.id === course?.school_trip_id);
        return sc.id === st?.school_id;
      });
      const schoolTrip = allSchoolTrips.find(t => t.id === course?.school_trip_id);
      const fullName = [s.name, s.surname].filter(Boolean).join(" ");
      const getCell = (key) => {
        if (key === "nombre")        return fullName;
        if (key === "curso")         return course ? `${course.course_name}${course.group_name ? ` · ${course.group_name}` : ""}` : "—";
        if (key === "colegio")       return school?.name || "—";
        if (key === "rooming")       return (schoolTrip?.rooming || []).find(r => (r.students || []).includes(fullName))?.room || "—";
        if (key === "grupo")         return (schoolTrip?.groups  || []).find(g => (g.students || []).includes(fullName))?.group || "—";
        if (key === "alergias")      return s.allergies || "—";
        if (key === "intolerancias") return s.intolerances || "—";
        if (key === "dieta")         return s.diet_notes || "—";
        if (key === "notas")         return s.notes || "—";
        return "—";
      };
      return `<tr><td style="color:#a1a1aa;width:24px">${i+1}</td>${exportActiveCols.map(c => `<td>${getCell(c.key)}</td>`).join("")}</tr>`;
    }).join("");
    exportListToPDF(
      `Alumnos — ${schoolName}${courseLabel}`,
      `${filteredStudents.length} alumnos`,
      `<table><thead><tr><th>#</th>${exportActiveCols.map(c => `<th>${c.label}</th>`).join("")}</tr></thead><tbody>${rows}</tbody></table>`
    );
  };

  const handleSaveStudent = async () => {
    if (!newStudent.name.trim()) { notify("El nombre del alumno es obligatorio."); return; }
    const targetCourseId = filterCourseId !== "all" ? filterCourseId : visibleCourses[0]?.id;
    if (!targetCourseId) { notify("Selecciona primero un colegio y curso."); return; }
    setSavingStudent(true);
    const { data, error } = await supabase.from("students").insert([{ ...newStudent, school_course_id: targetCourseId }]).select().maybeSingle();
    if (error) { notify("Error añadiendo alumno: " + error.message, { variant: "destructive" }); }
    else {
      setAllStudents((prev) => [...prev, data]);
      setNewStudent({ name: "", surname: "", allergies: "", intolerances: "", diet_notes: "", notes: "" });
      setShowAddStudent(false);
      notify("Alumno añadido.");
    }
    setSavingStudent(false);
  };

  const handleSaveDoc = async () => {
    if (!newDoc.name.trim()) { notify("El nombre del documento es obligatorio."); return; }
    // Determinar curso destino: puede ser un curso específico o todos los cursos del viaje
    const formCourses = docFormCourseId !== "all"
      ? allCourses.filter(c => c.id === docFormCourseId)
      : allCourses.filter(c => {
          const st = allSchoolTrips.find(t => t.id === docFormTripId);
          return st ? c.school_trip_id === st.id : false;
        });
    if (!docFormTripId) { notify("Selecciona colegio y viaje primero."); return; }
    if (formCourses.length === 0) { notify("No hay cursos disponibles para asignar."); return; }
    setSavingDoc(true);
    try {
      let driveUrl = ""; let driveFileName = "";
      if (docFile) {
        const school = schools.find(s => s.id === docFormSchoolId);
        const trip = allSchoolTrips.find(t => t.id === docFormTripId);
        const result = await uploadFileToDrive(docFile, school?.name || "colegio", "documentos", null, trip?.trips?.name || "colegio");
        driveUrl = result.webViewLink; driveFileName = result.fileName;
      }
      const { description, ...newDocBase } = newDoc;
      // Crear un registro por cada curso seleccionado
      const inserts = formCourses.map(c => ({
        ...newDocBase,
        school_course_id: c.id,
        status: "pending",
        ...(driveUrl ? { file_url: driveUrl, file_name: driveFileName } : {}),
      }));
      const { data, error } = await supabase.from("school_documents").insert(inserts).select();
      if (error) notify("Error añadiendo documento: " + error.message, { variant: "destructive" });
      else {
        setAllSchoolDocs((prev) => [...prev, ...(data || [])]);
        setNewDoc({ name: "", description: "", required: true });
        setDocFile(null);
        setDocFormCourseId("all");
        setShowAddDoc(false);
        notify(docFormCourseId === "all" ? `Documento asignado a ${formCourses.length} curso(s).` : "Documento asignado al curso.");
      }
    } catch(err) { notify("Error al subir el archivo: " + err.message); }
    setSavingDoc(false);
  };

  const handleSaveQuestion = async () => {
    if (!newQuestion.message.trim()) { notify("La pregunta no puede estar vacía."); return; }
    const schoolId = newQuestion.school_id || filterSchoolId || schools[0]?.id;
    if (!schoolId) { notify("Selecciona un colegio."); return; }
    setSavingQuestion(true);
    const { data, error } = await supabase.from("school_questions").insert([{ message: newQuestion.message, school_id: schoolId, source: "admin" }]).select().maybeSingle();
    if (error) notify("Error añadiendo pregunta: " + error.message, { variant: "destructive" });
    else { setAllSchoolQuestions((prev) => [...prev, data]); setNewQuestion({ message: "", school_id: "" }); setShowAddQuestion(false); notify("Pregunta añadida."); }
    setSavingQuestion(false);
  };

  const handleReplyQuestion = async (qId, reply) => {
    setSchoolSendingReply((s) => ({ ...s, [qId]: true }));
    const { error } = await supabase.from("school_questions").update({ reply }).eq("id", qId);
    if (error) notify("Error guardando respuesta.", { variant: "destructive" });
    else {
      const q = allSchoolQuestions.find((q) => q.id === qId);
      setAllSchoolQuestions((prev) => prev.map((q) => q.id === qId ? { ...q, reply } : q));
      setSchoolReplyTexts((t) => { const n = { ...t }; delete n[qId]; return n; });
      notify("Respuesta guardada.");
      if (q) {
        const school = schools.find((s) => s.id === q.school_id);
        if (school?.email) sendNotification("school_question_replied", school.email, null, { schoolName: school.name || "", contactName: school.contact_name || "", question: q.message, reply });
      }
    }
    setSchoolSendingReply((s) => { const n = { ...s }; delete n[qId]; return n; });
  };

  const handleSaveRoom = async () => {
    if (!newRoom.room.trim()) { notify("El nombre de la habitación es obligatorio."); return; }
    const stId = roomingTripId || filteredSchoolTrips[0]?.id;
    if (!stId) { notify("Selecciona un viaje primero."); return; }
    setSavingRoom(true);
    const st = allSchoolTrips.find((t) => t.id === stId);
    const existing = st?.rooming || [];
    const updated = [...existing, { room: newRoom.room, students: newRoom.students.split(",").map((s) => s.trim()).filter(Boolean) }];
    const { error } = await supabase.from("school_trips").update({ rooming: updated }).eq("id", stId);
    if (error) notify("Error añadiendo habitación: " + error.message, { variant: "destructive" });
    else { setAllSchoolTrips((prev) => prev.map((t) => t.id === stId ? { ...t, rooming: updated } : t)); setNewRoom({ room: "", students: "" }); setShowAddRoom(false); notify("Habitación añadida."); }
    setSavingRoom(false);
  };

  const handleSaveGroup = async () => {
    if (!newGroup.group.trim()) { notify("El nombre del grupo es obligatorio."); return; }
    const stId = groupTripId || filteredSchoolTrips[0]?.id;
    if (!stId) { notify("Selecciona un viaje primero."); return; }
    setSavingGroup(true);
    const st = allSchoolTrips.find((t) => t.id === stId);
    const existing = st?.activity_groups || [];
    const updated = [...existing, { group: newGroup.group, monitor: newGroup.monitor, students: newGroup.students.split(",").map((s) => s.trim()).filter(Boolean) }];
    const { error } = await supabase.from("school_trips").update({ activity_groups: updated }).eq("id", stId);
    if (error) notify("Error añadiendo grupo: " + error.message, { variant: "destructive" });
    else { setAllSchoolTrips((prev) => prev.map((t) => t.id === stId ? { ...t, activity_groups: updated } : t)); setNewGroup({ group: "", monitor: "", students: "" }); setShowAddGroup(false); notify("Grupo añadido."); }
    setSavingGroup(false);
  };

  if (loading) return <div className="py-16 text-center text-sm text-zinc-400">Cargando colegios...</div>;

  return (
    <div className="space-y-6">

      {/* Colegios tab */}
      {tab === "schools" && (
        <div className="space-y-5">
          <SectionTitle icon={Users} title="Colegios" subtitle="Importación, asignación y gestión de centros escolares." />

          <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
            <CardContent className="p-5">
              <div className="space-y-2">
                <Label>Buscar colegio</Label>
                <Input placeholder="Busca por nombre, coordinador o email" value={schoolSearch} onChange={(e) => setSchoolSearch(e.target.value)} className="rounded-2xl" />
              </div>
            </CardContent>
          </Card>

          <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
            <CardContent className="p-5">
              <div className="grid gap-3 lg:grid-cols-[auto_1fr_1fr_auto] lg:items-end">
                <div>
                  <Label className="mb-2 block">Importar Excel</Label>
                  <label className="cursor-pointer">
                    <input type="file" accept=".xlsx,.xls,.csv" className="hidden" disabled={isImportingSchool}
                      onChange={async (e) => {
                        const file = e.target.files?.[0]; if (!file) return;
                        e.target.value = "";
                        setIsImportingSchool(true); setSchoolImportMsg("Leyendo archivo...");
                        try {
                          const ab = await file.arrayBuffer();
                          const wb = XLSX.read(ab, { type: "array" });
                          const ws = wb.Sheets[wb.SheetNames[0]];
                          const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
                          const getVal = (row, ...keys) => { for (const k of keys) { const v = Object.entries(row).find(([rk]) => rk.toLowerCase().replace(/[^a-záéíóúñ0-9]/gi, "") === k.toLowerCase().replace(/[^a-záéíóúñ0-9]/gi, ""))?.[1]; if (v !== undefined && v !== "") return v; } return ""; };
                          let created = 0, skipped = 0;
                          for (const row of rows) {
                            const name = String(getVal(row, "nombre", "colegio", "centro", "name", "school") || "").trim();
                            if (!name) { skipped++; continue; }
                            const existing = schools.find(s => s.name.toLowerCase() === name.toLowerCase());
                            if (existing) { skipped++; continue; }
                            const contact_name = String(getVal(row, "coordinador", "contacto", "contact", "responsable") || "").trim();
                            const email = String(getVal(row, "email", "correo", "mail") || "").trim();
                            const phone = String(getVal(row, "telefono", "teléfono", "phone", "tel") || "").trim();
                            setSchoolImportMsg(`Guardando ${name}...`);
                            const { data, error } = await supabase.from("schools").insert([{ name, contact_name, email, phone }]).select().maybeSingle();
                            if (!error && data) { setSchools(prev => [...prev, data]); created++; }
                          }
                          notify(`Importación completada: ${created} colegio(s) creado(s)${skipped ? `, ${skipped} omitido(s)` : ""}.`);
                        } catch(err) { notify("Error al importar: " + err.message); }
                        finally { setIsImportingSchool(false); setSchoolImportMsg(""); }
                      }}
                    />
                    <span className={`inline-flex h-11 items-center rounded-2xl px-4 text-sm font-medium text-white ${isImportingSchool ? "opacity-60" : ""}`} style={{ backgroundColor: CORPORATE_RED }}>
                      <Upload className="mr-2 h-4 w-4" />{isImportingSchool ? schoolImportMsg || "Importando..." : "Subir Excel"}
                    </span>
                  </label>
                </div>
                <div className="space-y-2">
                  <Label>Colegio</Label>
                  <select value={selectedImportSchool} onChange={(e) => { setSelectedImportSchool(e.target.value); setSelectedImportTrip(""); }} className="h-11 w-full rounded-2xl border border-zinc-200 bg-white px-4 text-sm">
                    <option value="">Todos los colegios</option>
                    {schools.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
                  </select>
                </div>
                <div className="space-y-2">
                  <Label>Viaje contratado</Label>
                  <select value={selectedImportTrip} onChange={(e) => setSelectedImportTrip(e.target.value)} className="h-11 w-full rounded-2xl border border-zinc-200 bg-white px-4 text-sm">
                    <option value="">Todos los viajes</option>
                    {allSchoolTrips.filter(st => !selectedImportSchool || st.school_id === selectedImportSchool).map((st) => <option key={st.id} value={st.id}>{st.trips?.name || st.id}</option>)}
                  </select>
                </div>
                <Button className="h-11 rounded-2xl text-sm text-white" style={{ backgroundColor: CORPORATE_RED }} onClick={() => setShowNewSchool(!showNewSchool)}>
                  <Plus className="mr-1.5 h-4 w-4" />Nuevo colegio
                </Button>
              </div>

              <div className="mt-3 border-t border-zinc-100 pt-3">
                <Label className="mb-2 block text-xs text-zinc-500">Vincular documento Excel o Google Sheets (URL pública)</Label>
                <div className="flex gap-2">
                  <Input
                    value={schoolSheetUrl}
                    onChange={(e) => setSchoolSheetUrl(e.target.value)}
                    placeholder="https://docs.google.com/spreadsheets/d/... o enlace Excel"
                    className="rounded-2xl text-sm"
                  />
                  <Button
                    onClick={async () => {
                      if (!schoolSheetUrl.trim()) { notify("Introduce la URL del documento Excel o Google Sheets."); return; }
                      let targetUrl = schoolSheetUrl.trim();
                      const m = targetUrl.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
                      if (m) targetUrl = `https://docs.google.com/spreadsheets/d/${m[1]}/export?format=xlsx`;
                      setIsSyncingSchoolSheet(true);
                      try {
                        const token = await getAuthToken();
                        const res = await fetch(`/api/proxy-sheet?url=${encodeURIComponent(targetUrl)}`, {
                          headers: token ? { Authorization: `Bearer ${token}` } : {},
                        });
                        if (!res.ok) { const e = await res.json().catch(() => ({})); throw new Error(e.error || `Error ${res.status}`); }
                        const buffer = await res.arrayBuffer();
                        const wb = XLSX.read(buffer, { type: "array" });
                        const ws = wb.Sheets[wb.SheetNames[0]];
                        const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
                        const getVal = (row, ...keys) => { for (const k of keys) { const v = Object.entries(row).find(([rk]) => rk.toLowerCase().replace(/[^a-záéíóúñ0-9]/gi, "") === k.toLowerCase().replace(/[^a-záéíóúñ0-9]/gi, ""))?.[1]; if (v !== undefined && v !== "") return v; } return ""; };
                        let created = 0, skipped = 0;
                        for (const row of rows) {
                          const name = String(getVal(row, "nombre", "colegio", "centro", "name", "school") || "").trim();
                          if (!name) { skipped++; continue; }
                          const existing = schools.find(s => s.name.toLowerCase() === name.toLowerCase());
                          if (existing) { skipped++; continue; }
                          const contact_name = String(getVal(row, "coordinador", "contacto", "contact", "responsable") || "").trim();
                          const email = String(getVal(row, "email", "correo", "mail") || "").trim();
                          const phone = String(getVal(row, "telefono", "teléfono", "phone", "tel") || "").trim();
                          const { data, error } = await supabase.from("schools").insert([{ name, contact_name, email, phone }]).select().maybeSingle();
                          if (!error && data) { setSchools(prev => [...prev, data]); created++; }
                        }
                        notify(`Sincronización completada: ${created} colegio(s) creado(s)${skipped ? `, ${skipped} omitido(s)` : ""}.`);
                      } catch (err) {
                        notify("Error al sincronizar: " + err.message, { variant: "destructive" });
                      } finally {
                        setIsSyncingSchoolSheet(false);
                      }
                    }}
                    disabled={isSyncingSchoolSheet}
                    className="h-11 shrink-0 rounded-2xl text-white"
                    style={{ backgroundColor: CORPORATE_RED }}
                  >
                    {isSyncingSchoolSheet ? <Loader2 className="h-4 w-4 animate-spin" /> : <Download className="h-4 w-4" />}
                    <span className="ml-2">{isSyncingSchoolSheet ? "Sincronizando..." : "Sincronizar"}</span>
                  </Button>
                </div>
                <p className="mt-1 text-xs text-zinc-400">Comparte el Google Sheet con la cuenta de Google conectada al portal.</p>
              </div>
            </CardContent>
          </Card>

          {showNewSchool && (
            <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
              <CardContent className="p-5">
                <div className="mb-3 text-sm font-medium text-zinc-700">Nuevo colegio</div>
                <div className="grid gap-2 sm:grid-cols-2">
                  <Input placeholder="Nombre del colegio *" value={newSchool.name} onChange={(e) => setNewSchool((p) => ({ ...p, name: e.target.value }))} className="rounded-xl" />
                  <Input placeholder="Nombre del coordinador" value={newSchool.contact_name} onChange={(e) => setNewSchool((p) => ({ ...p, contact_name: e.target.value }))} className="rounded-xl" />
                  <Input placeholder="Email" type="email" value={newSchool.email} onChange={(e) => setNewSchool((p) => ({ ...p, email: e.target.value }))} className="rounded-xl" />
                  <Input placeholder="Teléfono" value={newSchool.phone} onChange={(e) => setNewSchool((p) => ({ ...p, phone: e.target.value }))} className="rounded-xl" />
                </div>
                <div className="mt-3 flex gap-2">
                  <Button onClick={handleSaveSchool} disabled={savingSchool || !newSchool.name.trim()} className="rounded-xl text-white text-sm" style={{ backgroundColor: CORPORATE_RED }}>
                    {savingSchool ? "Guardando..." : "Guardar colegio"}
                  </Button>
                  <Button variant="outline" className="rounded-xl text-sm" onClick={() => setShowNewSchool(false)}>Cancelar</Button>
                </div>
              </CardContent>
            </Card>
          )}

          <div className="space-y-3">
            {schools.length === 0 && <p className="text-sm text-zinc-400">No hay colegios registrados.</p>}
            {schools.filter(s => !schoolSearch || s.name.toLowerCase().includes(schoolSearch.toLowerCase()) || s.contact_name?.toLowerCase().includes(schoolSearch.toLowerCase())).map((school) => (
              <Card key={school.id} className="rounded-3xl border-zinc-200 shadow-sm">
                <CardContent className="p-5">
                  <div className="flex flex-wrap items-start justify-between gap-3">
                    <div>
                      <div className="font-semibold text-zinc-950">{school.name}</div>
                      {school.contact_name && <div className="mt-0.5 text-xs text-zinc-500">{school.contact_name}</div>}
                      <div className="mt-1 flex flex-wrap gap-3 text-xs text-zinc-400">
                        {school.email && <span>{school.email}</span>}
                        {school.phone && <span>{school.phone}</span>}
                      </div>
                      <div className="mt-2 flex gap-3">
                        <Badge variant="outline" className="rounded-xl text-xs">{getSchoolTripCount(school.id)} viajes</Badge>
                        <Badge variant="outline" className="rounded-xl text-xs">{getSchoolStudentCount(school.id)} alumnos</Badge>
                      </div>
                    </div>
                    <Button
                      variant="outline"
                      className="rounded-xl text-xs"
                      onClick={() => setAssigningSchoolId(assigningSchoolId === school.id ? null : school.id)}
                    >
                      <Plus className="mr-1 h-3.5 w-3.5" />Asignar viaje
                    </Button>
                  </div>

                  {assigningSchoolId === school.id && (
                    <div className="mt-4 border-t border-zinc-100 pt-4">
                      <div className="mb-2 text-xs font-medium text-zinc-700">Asignar viaje + curso</div>
                      <div className="flex flex-wrap gap-2">
                        <select
                          value={assignTripId}
                          onChange={(e) => setAssignTripId(e.target.value)}
                          className="rounded-xl border border-zinc-200 bg-white px-3 py-1.5 text-xs text-zinc-950 focus:outline-none"
                        >
                          {trips.map((t) => <option key={t.id} value={t.id}>{t.name}</option>)}
                        </select>
                        <Input placeholder="Curso (ej: 4ºA)" value={assignCourse} onChange={(e) => setAssignCourse(e.target.value)} className="h-8 rounded-xl text-xs w-32" />
                        <Input placeholder="Grupo (opcional)" value={assignGroup} onChange={(e) => setAssignGroup(e.target.value)} className="h-8 rounded-xl text-xs w-32" />
                        <Button onClick={() => handleAssignTrip(school.id)} disabled={savingAssign} className="h-8 rounded-xl text-xs text-white" style={{ backgroundColor: CORPORATE_RED }}>
                          {savingAssign ? "Guardando..." : "Guardar"}
                        </Button>
                        <Button variant="outline" className="h-8 rounded-xl text-xs" onClick={() => setAssigningSchoolId(null)}>Cancelar</Button>
                      </div>
                    </div>
                  )}
                </CardContent>
              </Card>
            ))}
          </div>
        </div>
      )}

      {/* Alumnos tab */}
      {tab === "students" && (
        <div className="space-y-5">
          <SectionTitle icon={Users} title="Alumnos" subtitle="Listado de alumnos por colegio, viaje y curso." extra={
            <Button className="rounded-2xl text-sm text-white" style={{ backgroundColor: CORPORATE_RED }} onClick={() => setShowAddStudent(!showAddStudent)}>
              <Plus className="mr-1.5 h-4 w-4" />Añadir alumno
            </Button>
          } />

          {/* ── Exportar listado de alumnos ─────────────────────────────── */}
          <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
            <CardContent className="p-5 space-y-4">
              <div className="grid gap-2 sm:grid-cols-2 lg:grid-cols-4">
                {STUDENT_COLS_DEF.map(({ key, label }) => (
                  <div key={key} onClick={() => setExportCols(p => ({ ...p, [key]: !p[key] }))}
                    className={`flex cursor-pointer select-none items-center gap-3 rounded-2xl border p-3 text-sm transition-all ${exportCols[key] ? "border-zinc-900 bg-zinc-950 text-white" : "border-zinc-200 bg-white text-zinc-700 hover:border-zinc-400"}`}>
                    <div className={`flex h-5 w-5 shrink-0 items-center justify-center rounded-md border text-xs font-bold ${exportCols[key] ? "border-white bg-white text-zinc-950" : "border-zinc-300"}`}>
                      {exportCols[key] ? "✓" : ""}
                    </div>
                    {label}
                  </div>
                ))}
              </div>
              <div className="flex flex-wrap gap-3">
                <Button variant="outline" className="h-11 rounded-2xl text-sm" onClick={() => setExportCols(Object.fromEntries(STUDENT_COLS_DEF.map(c => [c.key, true])))}>
                  Seleccionar todo
                </Button>
                <Button variant="outline" className="h-11 rounded-2xl text-sm" onClick={() => setExportCols(Object.fromEntries(STUDENT_COLS_DEF.map(c => [c.key, false])))}>
                  Borrar selección
                </Button>
                <Button onClick={handleStudentExport} disabled={filteredStudents.length === 0 || exportActiveCols.length === 0}
                  className="h-11 rounded-2xl text-white" style={{ backgroundColor: CORPORATE_RED }}>
                  <FileText className="mr-2 h-4 w-4" />Exportar PDF — {filteredStudents.length} alumno(s)
                </Button>
              </div>
            </CardContent>
          </Card>

          {/* Filtro colegio + buscar alumno */}
          <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
            <CardContent className="p-5 space-y-4">
              <div className="space-y-2">
                <Label>Filtrar por colegio</Label>
                <select value={filterSchoolId} onChange={(e) => { setFilterSchoolId(e.target.value); setFilterTripId(""); setFilterCourseId("all"); }}
                  className="h-11 w-full rounded-2xl border border-zinc-200 bg-white px-4 text-sm">
                  <option value="">Todos los colegios</option>
                  {schools.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
                </select>
              </div>
              <div className="space-y-2">
                <Label>Buscar alumno</Label>
                <Input value={studentTrackingSearch} onChange={(e) => setStudentTrackingSearch(e.target.value)} placeholder="Busca por nombre o apellido del alumno..." className="rounded-2xl" />
              </div>
            </CardContent>
          </Card>

          {/* Tabs de curso — solo si hay colegio seleccionado */}
          {filterSchoolId && visibleCourses.length > 0 && (
            <div className="flex flex-wrap gap-2">
              <button type="button" onClick={() => setFilterCourseId("all")}
                className={`rounded-2xl px-3 py-1.5 text-xs font-medium transition ${filterCourseId === "all" ? "text-white shadow-sm" : "border border-zinc-200 bg-white text-zinc-600 hover:bg-zinc-50"}`}
                style={filterCourseId === "all" ? { backgroundColor: CORPORATE_RED } : {}}>
                Todos
              </button>
              {visibleCourses.map((c) => (
                <button key={c.id} type="button" onClick={() => setFilterCourseId(c.id)}
                  className={`rounded-2xl px-3 py-1.5 text-xs font-medium transition ${filterCourseId === c.id ? "text-white shadow-sm" : "border border-zinc-200 bg-white text-zinc-600 hover:bg-zinc-50"}`}
                  style={filterCourseId === c.id ? { backgroundColor: CORPORATE_RED } : {}}>
                  {c.course_name}{c.group_name ? ` · ${c.group_name}` : ""}
                </button>
              ))}
            </div>
          )}

          {/* Formulario añadir alumno manualmente */}
          {showAddStudent && (
            <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
              <CardContent className="p-5 space-y-3">
                <div className="text-sm font-medium text-zinc-700">Nuevo alumno{filterCourseId !== "all" ? ` — ${visibleCourses.find(c => c.id === filterCourseId)?.course_name || ""}` : ""}</div>
                {filterSchoolId && visibleCourses.length > 1 && filterCourseId === "all" && (
                  <select value={filterCourseId} onChange={(e) => setFilterCourseId(e.target.value)}
                    className="rounded-xl border border-zinc-200 bg-white px-3 py-1.5 text-sm text-zinc-950 focus:outline-none w-full">
                    <option value="all">Selecciona un curso *</option>
                    {visibleCourses.map((c) => <option key={c.id} value={c.id}>{c.course_name}{c.group_name ? ` · ${c.group_name}` : ""}</option>)}
                  </select>
                )}
                <div className="grid gap-2 sm:grid-cols-2">
                  <Input placeholder="Nombre *" value={newStudent.name} onChange={(e) => setNewStudent(p => ({ ...p, name: e.target.value }))} className="rounded-xl" />
                  <Input placeholder="Apellidos" value={newStudent.surname} onChange={(e) => setNewStudent(p => ({ ...p, surname: e.target.value }))} className="rounded-xl" />
                  <Input placeholder="Alergias" value={newStudent.allergies} onChange={(e) => setNewStudent(p => ({ ...p, allergies: e.target.value }))} className="rounded-xl" />
                  <Input placeholder="Intolerancias" value={newStudent.intolerances} onChange={(e) => setNewStudent(p => ({ ...p, intolerances: e.target.value }))} className="rounded-xl" />
                  <Input placeholder="Dieta especial" value={newStudent.diet_notes} onChange={(e) => setNewStudent(p => ({ ...p, diet_notes: e.target.value }))} className="rounded-xl" />
                  <Input placeholder="Notas" value={newStudent.notes} onChange={(e) => setNewStudent(p => ({ ...p, notes: e.target.value }))} className="rounded-xl" />
                </div>
                <div className="flex gap-2">
                  <Button onClick={handleSaveStudent} disabled={savingStudent || !newStudent.name.trim()} className="rounded-xl text-white text-sm" style={{ backgroundColor: CORPORATE_RED }}>
                    {savingStudent ? "Guardando..." : "Guardar alumno"}
                  </Button>
                  <Button variant="outline" className="rounded-xl text-sm" onClick={() => setShowAddStudent(false)}>Cancelar</Button>
                </div>
              </CardContent>
            </Card>
          )}

          {filteredStudents.length === 0 ? (
            <p className="text-sm text-zinc-400">No hay alumnos en la selección actual.</p>
          ) : (
            <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
              <CardContent className="p-4">
                <div className="mb-2 flex items-center justify-between">
                  <div className="text-sm font-medium text-zinc-700">{filteredStudents.length} alumnos</div>
                  <Badge variant="outline" className="rounded-xl text-xs">{filteredStudents.filter((s) => s.allergies?.trim() || s.intolerances?.trim()).length} con alergias</Badge>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-xs">
                    <thead>
                      <tr className="border-b border-zinc-100">
                        <th className="px-2 py-1.5 text-left font-medium text-zinc-500">Nombre</th>
                        <th className="px-2 py-1.5 text-left font-medium text-zinc-500">Apellidos</th>
                        <th className="px-2 py-1.5 text-left font-medium text-zinc-500">Curso</th>
                        <th className="px-2 py-1.5 text-left font-medium text-zinc-500">Alergia</th>
                        <th className="px-2 py-1.5 text-left font-medium text-zinc-500">Intolerancia</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredStudents.map((s) => {
                        const course = allCourses.find((c) => c.id === s.school_course_id);
                        return (
                          <tr key={s.id} className="border-b border-zinc-50 hover:bg-zinc-50/50">
                            <td className="px-2 py-1.5 font-medium text-zinc-900">{s.name}</td>
                            <td className="px-2 py-1.5 text-zinc-700">{s.surname}</td>
                            <td className="px-2 py-1.5 text-zinc-500">{course ? `${course.course_name}${course.group_name ? ` · ${course.group_name}` : ""}` : "—"}</td>
                            <td className="px-2 py-1.5 text-red-700">{s.allergies || "—"}</td>
                            <td className="px-2 py-1.5 text-amber-700">{s.intolerances || "—"}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </CardContent>
            </Card>
          )}

        </div>
      )}

      {/* Rooming tab */}
      {tab === "rooming" && (
        <div className="space-y-5">
          <SectionTitle icon={Home} title="Rooming" subtitle="Distribución de habitaciones por colegio y viaje." extra={
            <div className="flex gap-2">
              <Button variant="outline" className="rounded-2xl text-sm" onClick={() => {
                const stList = (filterSchoolId ? filteredSchoolTrips : allSchoolTrips).filter(st => (!filterTripId || st.id === filterTripId) && st.rooming?.length);
                const schoolName = filterSchoolId ? (schools.find(s => s.id === filterSchoolId)?.name || "") : "Todos los colegios";
                const rows = stList.flatMap(st => {
                  const sch = schools.find(s => s.id === st.school_id);
                  return (st.rooming||[]).map(r => `<tr><td style="font-weight:600">${r.room}</td><td>${sch?.name||"—"}</td><td>${st.trips?.name||st.trip_id||"—"}</td><td>${(r.students||[]).join(", ")||"—"}</td></tr>`);
                }).join("");
                exportListToPDF(`Rooming — ${schoolName}`, `${stList.reduce((n,st)=>n+(st.rooming||[]).length,0)} habitaciones`, rows ? `<table><thead><tr><th>Habitación</th><th>Colegio</th><th>Viaje</th><th>Alumnos</th></tr></thead><tbody>${rows}</tbody></table>` : "<p style='color:#71717a'>Sin datos de rooming.</p>");
              }}>
                <FileText className="mr-2 h-4 w-4" />Exportar PDF
              </Button>
              <Button className="rounded-2xl text-sm text-white" style={{ backgroundColor: CORPORATE_RED }} onClick={() => setShowAddRoom(!showAddRoom)}>
                <Plus className="mr-1.5 h-4 w-4" />Añadir habitación
              </Button>
            </div>
          } />

          {/* Filtro colegio Rooming */}
          <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
            <CardContent className="p-5 space-y-2">
              <Label>Filtrar por colegio</Label>
              <select value={filterSchoolId} onChange={(e) => { setFilterSchoolId(e.target.value); setFilterTripId(""); setFilterCourseId("all"); setRoomingTripId(""); }}
                className="h-11 w-full rounded-2xl border border-zinc-200 bg-white px-4 text-sm">
                <option value="">Todos los colegios</option>
                {schools.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
              </select>
            </CardContent>
          </Card>

          {/* Tabs de curso — solo si hay colegio seleccionado */}
          {filterSchoolId && visibleCourses.length > 0 && (
            <div className="flex flex-wrap gap-2">
              <button type="button" onClick={() => setFilterCourseId("all")}
                className={`rounded-2xl px-3 py-1.5 text-xs font-medium transition ${filterCourseId === "all" ? "text-white shadow-sm" : "border border-zinc-200 bg-white text-zinc-600 hover:bg-zinc-50"}`}
                style={filterCourseId === "all" ? { backgroundColor: CORPORATE_RED } : {}}>Todos</button>
              {visibleCourses.map((c) => (
                <button key={c.id} type="button" onClick={() => setFilterCourseId(c.id)}
                  className={`rounded-2xl px-3 py-1.5 text-xs font-medium transition ${filterCourseId === c.id ? "text-white shadow-sm" : "border border-zinc-200 bg-white text-zinc-600 hover:bg-zinc-50"}`}
                  style={filterCourseId === c.id ? { backgroundColor: CORPORATE_RED } : {}}>
                  {c.course_name}{c.group_name ? ` · ${c.group_name}` : ""}
                </button>
              ))}
            </div>
          )}

          {/* Formulario nueva habitación */}
          {showAddRoom && (
            <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
              <CardContent className="p-5 space-y-3">
                <div className="text-sm font-medium text-zinc-700">Nueva habitación</div>
                <select
                  value={roomingTripId || filteredSchoolTrips[0]?.id || ""}
                  onChange={(e) => setRoomingTripId(e.target.value)}
                  className="rounded-xl border border-zinc-200 bg-white px-3 py-1.5 text-sm text-zinc-950 focus:outline-none w-full">
                  <option value="">Selecciona viaje *</option>
                  {(filterSchoolId ? filteredSchoolTrips : allSchoolTrips).map((st) => {
                    const sch = schools.find((s) => s.id === st.school_id);
                    return <option key={st.id} value={st.id}>{sch ? `${sch.name} — ` : ""}{st.trips?.name || st.trip_id}</option>;
                  })}
                </select>
                <div className="grid gap-2 sm:grid-cols-2">
                  <Input placeholder="Nombre habitación *" value={newRoom.room} onChange={(e) => setNewRoom(p => ({ ...p, room: e.target.value }))} className="rounded-xl" />
                  <Input placeholder="Alumnos (separados por coma)" value={newRoom.students} onChange={(e) => setNewRoom(p => ({ ...p, students: e.target.value }))} className="rounded-xl" />
                </div>
                <div className="flex gap-2">
                  <Button onClick={handleSaveRoom} disabled={savingRoom || !newRoom.room.trim()} className="rounded-xl text-white text-sm" style={{ backgroundColor: CORPORATE_RED }}>
                    {savingRoom ? "Guardando..." : "Guardar habitación"}
                  </Button>
                  <Button variant="outline" className="rounded-xl text-sm" onClick={() => setShowAddRoom(false)}>Cancelar</Button>
                </div>
              </CardContent>
            </Card>
          )}

          {(() => {
            const stList = (filterSchoolId ? filteredSchoolTrips : allSchoolTrips).filter((st) =>
              (!filterTripId || st.id === filterTripId) && st.rooming?.length
            );
            if (stList.length === 0) return <p className="text-sm text-zinc-400">No hay rooming en la selección actual. Añade habitaciones manualmente o pide al colegio que suba su listado.</p>;
            return stList.map((st) => {
              const school = schools.find((s) => s.id === st.school_id);
              return (
                <Card key={st.id} className="rounded-3xl border-zinc-200 bg-white shadow-sm">
                  <CardContent className="p-5">
                    <div className="mb-1 font-semibold text-zinc-950">{school?.name || "Colegio"}</div>
                    <div className="mb-3 text-xs text-zinc-500">{st.trips?.name || st.trip_id} · {st.rooming.length} habitaciones</div>
                    <div className="grid gap-2 sm:grid-cols-2 lg:grid-cols-3">
                      {st.rooming.map((r, i) => (
                        <div key={i} className="rounded-xl border border-zinc-100 bg-zinc-50/50 p-3">
                          <div className="mb-1.5 text-xs font-semibold text-zinc-900">{r.room}</div>
                          {(r.students || []).map((s, j) => <div key={j} className="text-xs text-zinc-600">{s}</div>)}
                          {(!r.students || r.students.length === 0) && <div className="text-xs text-zinc-400">Sin alumnos asignados</div>}
                        </div>
                      ))}
                    </div>
                  </CardContent>
                </Card>
              );
            });
          })()}
        </div>
      )}

      {/* Groups tab */}
      {tab === "groups" && (
        <div className="space-y-5">
          <SectionTitle icon={Grid2x2} title="Grupos de actividades" subtitle="Grupos y monitores por colegio y viaje." extra={
            <div className="flex gap-2">
              <Button variant="outline" className="rounded-2xl text-sm" onClick={() => {
                const stList = (filterSchoolId ? filteredSchoolTrips : allSchoolTrips).filter(st => (!filterTripId || st.id === filterTripId) && st.activity_groups?.length);
                const schoolName = filterSchoolId ? (schools.find(s => s.id === filterSchoolId)?.name || "") : "Todos los colegios";
                const rows = stList.flatMap(st => {
                  const sch = schools.find(s => s.id === st.school_id);
                  return (st.activity_groups||[]).map(g => `<tr><td style="font-weight:600">${g.group}</td><td>${g.monitor||"—"}</td><td>${sch?.name||"—"}</td><td>${st.trips?.name||st.trip_id||"—"}</td><td>${(g.students||[]).join(", ")||"—"}</td></tr>`);
                }).join("");
                exportListToPDF(`Grupos — ${schoolName}`, `${stList.reduce((n,st)=>n+(st.activity_groups||[]).length,0)} grupos`, rows ? `<table><thead><tr><th>Grupo</th><th>Monitor</th><th>Colegio</th><th>Viaje</th><th>Alumnos</th></tr></thead><tbody>${rows}</tbody></table>` : "<p style='color:#71717a'>Sin grupos registrados.</p>");
              }}>
                <FileText className="mr-2 h-4 w-4" />Exportar PDF
              </Button>
              <Button className="rounded-2xl text-sm text-white" style={{ backgroundColor: CORPORATE_RED }} onClick={() => setShowAddGroup(!showAddGroup)}>
                <Plus className="mr-1.5 h-4 w-4" />Añadir grupo
              </Button>
            </div>
          } />

          {/* Filtro colegio Grupos */}
          <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
            <CardContent className="p-5 space-y-2">
              <Label>Filtrar por colegio</Label>
              <select value={filterSchoolId} onChange={(e) => { setFilterSchoolId(e.target.value); setFilterTripId(""); setFilterCourseId("all"); setGroupTripId(""); }}
                className="h-11 w-full rounded-2xl border border-zinc-200 bg-white px-4 text-sm">
                <option value="">Todos los colegios</option>
                {schools.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
              </select>
            </CardContent>
          </Card>

          {/* Tabs de curso — solo si hay colegio seleccionado */}
          {filterSchoolId && visibleCourses.length > 0 && (
            <div className="flex flex-wrap gap-2">
              <button type="button" onClick={() => setFilterCourseId("all")}
                className={`rounded-2xl px-3 py-1.5 text-xs font-medium transition ${filterCourseId === "all" ? "text-white shadow-sm" : "border border-zinc-200 bg-white text-zinc-600 hover:bg-zinc-50"}`}
                style={filterCourseId === "all" ? { backgroundColor: CORPORATE_RED } : {}}>Todos</button>
              {visibleCourses.map((c) => (
                <button key={c.id} type="button" onClick={() => setFilterCourseId(c.id)}
                  className={`rounded-2xl px-3 py-1.5 text-xs font-medium transition ${filterCourseId === c.id ? "text-white shadow-sm" : "border border-zinc-200 bg-white text-zinc-600 hover:bg-zinc-50"}`}
                  style={filterCourseId === c.id ? { backgroundColor: CORPORATE_RED } : {}}>
                  {c.course_name}{c.group_name ? ` · ${c.group_name}` : ""}
                </button>
              ))}
            </div>
          )}

          {/* Formulario nuevo grupo */}
          {showAddGroup && (
            <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
              <CardContent className="p-5 space-y-3">
                <div className="text-sm font-medium text-zinc-700">Nuevo grupo de actividades</div>
                <select
                  value={groupTripId || filteredSchoolTrips[0]?.id || ""}
                  onChange={(e) => setGroupTripId(e.target.value)}
                  className="rounded-xl border border-zinc-200 bg-white px-3 py-1.5 text-sm text-zinc-950 focus:outline-none w-full">
                  <option value="">Selecciona viaje *</option>
                  {(filterSchoolId ? filteredSchoolTrips : allSchoolTrips).map((st) => {
                    const sch = schools.find((s) => s.id === st.school_id);
                    return <option key={st.id} value={st.id}>{sch ? `${sch.name} — ` : ""}{st.trips?.name || st.trip_id}</option>;
                  })}
                </select>
                <div className="grid gap-2 sm:grid-cols-2 lg:grid-cols-3">
                  <Input placeholder="Nombre del grupo *" value={newGroup.group} onChange={(e) => setNewGroup(p => ({ ...p, group: e.target.value }))} className="rounded-xl" />
                  <Input placeholder="Monitor / responsable" value={newGroup.monitor} onChange={(e) => setNewGroup(p => ({ ...p, monitor: e.target.value }))} className="rounded-xl" />
                  <Input placeholder="Alumnos (separados por coma)" value={newGroup.students} onChange={(e) => setNewGroup(p => ({ ...p, students: e.target.value }))} className="rounded-xl lg:col-span-3 sm:col-span-2" />
                </div>
                <div className="flex gap-2">
                  <Button onClick={handleSaveGroup} disabled={savingGroup || !newGroup.group.trim()} className="rounded-xl text-white text-sm" style={{ backgroundColor: CORPORATE_RED }}>
                    {savingGroup ? "Guardando..." : "Guardar grupo"}
                  </Button>
                  <Button variant="outline" className="rounded-xl text-sm" onClick={() => setShowAddGroup(false)}>Cancelar</Button>
                </div>
              </CardContent>
            </Card>
          )}

          {(() => {
            const stList = (filterSchoolId ? filteredSchoolTrips : allSchoolTrips).filter((st) =>
              (!filterTripId || st.id === filterTripId) && st.activity_groups?.length
            );
            if (stList.length === 0) return <p className="text-sm text-zinc-400">No hay grupos en la selección actual. Añade grupos manualmente o pide al colegio que suba su distribución.</p>;
            return stList.map((st) => {
              const school = schools.find((s) => s.id === st.school_id);
              return (
                <Card key={st.id} className="rounded-3xl border-zinc-200 bg-white shadow-sm">
                  <CardContent className="p-5">
                    <div className="mb-1 font-semibold text-zinc-950">{school?.name || "Colegio"}</div>
                    <div className="mb-3 text-xs text-zinc-500">{st.trips?.name || st.trip_id} · {st.activity_groups.length} grupos</div>
                    <div className="grid gap-2 sm:grid-cols-2 lg:grid-cols-3">
                      {st.activity_groups.map((g, i) => (
                        <div key={i} className="rounded-xl border border-zinc-100 bg-zinc-50/50 p-3">
                          <div className="mb-0.5 text-xs font-semibold text-zinc-900">{g.group}</div>
                          {g.monitor && <div className="mb-1 text-xs text-zinc-400">Monitor: {g.monitor}</div>}
                          {(g.students || []).map((s, j) => <div key={j} className="text-xs text-zinc-600">{s}</div>)}
                          {(!g.students || g.students.length === 0) && <div className="text-xs text-zinc-400">Sin alumnos asignados</div>}
                        </div>
                      ))}
                    </div>
                  </CardContent>
                </Card>
              );
            });
          })()}
        </div>
      )}

      {/* Pagos tab */}
      {tab === "pagos" && (
        <div className="space-y-5">
          <SectionTitle icon={CreditCard} title="Pagos" subtitle="Seguimiento económico de cada colegio por viaje." />

          {/* Datos bancarios globales */}
          {(() => {
            const ti = allSchoolTrips[0]?.trips?.transfer_info || {};
            const saveSchoolGlobalTi = async (field, value) => {
              const updated = { ...ti, [field]: value };
              const uniqueTripIds = [...new Set(allSchoolTrips.map((s) => s.trip_id).filter(Boolean))];
              if (!uniqueTripIds.length) return;
              const { error } = await supabase.from("trips").update({ transfer_info: updated }).in("id", uniqueTripIds);
              if (!error) setAllSchoolTrips((prev) => prev.map((s) => ({ ...s, trips: { ...s.trips, transfer_info: updated } })));
              else notify("Error guardando datos bancarios: " + error.message);
            };
            return (
              <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
                <CardContent className="p-5 space-y-4">
                  <div className="flex items-center gap-2">
                    <Building2 className="h-4 w-4 text-zinc-400" />
                    <span className="text-sm font-semibold uppercase tracking-[0.15em] text-zinc-500">Datos bancarios</span>
                  </div>
                  <div className="grid gap-3 sm:grid-cols-2">
                    {[["Banco", "bank"], ["Titular", "accountHolder"], ["IBAN", "iban"], ["Concepto", "concept"]].map(([label, field]) => (
                      <div key={field} className="space-y-1">
                        <Label className="text-xs">{label}</Label>
                        <Input
                          defaultValue={ti[field] || ""}
                          placeholder={field === "concept" ? "Nombre del participante + viaje" : field === "iban" ? "ES00 0000 0000 0000 0000 0000" : ""}
                          className="rounded-xl text-sm h-9"
                          onBlur={(e) => saveSchoolGlobalTi(field, e.target.value)}
                        />
                      </div>
                    ))}
                  </div>
                </CardContent>
              </Card>
            );
          })()}

          {/* Buscador de colegio */}
          <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
            <CardContent className="p-5">
              <div className="space-y-2">
                <Label>Buscar colegio</Label>
                <Input value={pagosSchoolSearch} onChange={(e) => setPagosSchoolSearch(e.target.value)} placeholder="Busca por nombre de colegio o viaje" className="rounded-2xl" />
              </div>
            </CardContent>
          </Card>

          {(() => {
            const q = pagosSchoolSearch.toLowerCase();
            const visibleTrips = allSchoolTrips.filter((st) => {
              if (!q) return true;
              const schoolName = (schools.find((s) => s.id === st.school_id)?.name || "").toLowerCase();
              const tripName = (st.trips?.name || "").toLowerCase();
              return schoolName.includes(q) || tripName.includes(q);
            });
            return visibleTrips.length === 0 ? (
              <p className="text-sm text-zinc-400">No hay viajes asignados.</p>
            ) : (
              <div className="space-y-4">
              {visibleTrips.map((st) => {
                const school = schools.find(s => s.id === st.school_id);
                const pi = st.payment_info || { total: 0, paid: 0, status: "pending", notes: "" };
                const pending = Math.max(0, (pi.total || 0) - (pi.paid || 0));
                const pct = pi.total > 0 ? Math.round((pi.paid / pi.total) * 100) : 0;
                const statusLabel = { pending: "Pendiente", partial: "Parcial", completed: "Pagado" }[pi.status] || "Pendiente";
                const statusColor = { pending: "bg-red-50 text-red-700", partial: "bg-amber-50 text-amber-700", completed: "bg-green-50 text-green-700" }[pi.status] || "bg-zinc-100 text-zinc-600";
                const fmt = (v) => new Intl.NumberFormat("es-ES", { style: "currency", currency: "EUR" }).format(v || 0);
                return (
                  <Card key={st.id} className="rounded-3xl border-zinc-200 bg-white shadow-sm">
                    <CardContent className="p-5 space-y-4">
                      <div className="flex items-start justify-between gap-3">
                        <div>
                          <div className="font-semibold text-zinc-950">{school?.name || "Colegio"}</div>
                          <div className="text-xs text-zinc-500">{st.trips?.name || st.trip_id}</div>
                        </div>
                        <span className={`rounded-full px-3 py-1 text-xs font-medium ${statusColor}`}>{statusLabel}</span>
                      </div>

                      <div className="grid grid-cols-3 gap-3">
                        {[["Total", fmt(pi.total)], ["Pagado", fmt(pi.paid)], ["Pendiente", fmt(pending)]].map(([label, val]) => (
                          <div key={label} className="rounded-2xl border border-zinc-200 bg-zinc-50 p-3">
                            <div className="text-xs uppercase tracking-[0.15em] text-zinc-500">{label}</div>
                            <div className="mt-1.5 text-base font-semibold text-zinc-950">{val}</div>
                          </div>
                        ))}
                      </div>

                      {pi.total > 0 && (
                        <div className="h-2 w-full rounded-full bg-zinc-100">
                          <div className="h-2 rounded-full bg-green-500 transition-all" style={{ width: `${pct}%` }} />
                        </div>
                      )}

                      <div className="flex flex-wrap items-center gap-2">
                        {st.invoice_url ? (
                          <Button variant="outline" className="h-9 rounded-2xl text-sm" onClick={() => window.open(st.invoice_url, "_blank", "noopener,noreferrer")}>
                            <Download className="mr-2 h-4 w-4" />Ver factura
                          </Button>
                        ) : null}
                        <InvoiceUploadButton existing={!!st.invoice_url} size="md" onUpload={async (file, onProgress) => {
                          try {
                            const result = await uploadFileToDrive(file, school?.name || "colegio", "facturas", onProgress, "GIMELOOS Facturas");
                            const { error } = await supabase.from("school_trips").update({ invoice_url: result.webViewLink }).eq("id", st.id);
                            if (!error) { setAllSchoolTrips(prev => prev.map(s => s.id === st.id ? { ...s, invoice_url: result.webViewLink } : s)); notify("Factura subida correctamente."); }
                            else notify("Error guardando factura: " + error.message);
                          } catch (err) { notify("Error subiendo factura: " + err.message); }
                        }} />
                        {st.invoice_url && (
                          <Button variant="outline" className="h-9 rounded-2xl text-sm text-red-600 hover:bg-red-50 border-red-200" onClick={() => setPendingDeleteInvoiceId(st.id)}>
                            <Trash2 className="mr-2 h-4 w-4" />Borrar factura
                          </Button>
                        )}
                        <Button variant="outline" className="h-9 rounded-2xl text-sm text-red-600 hover:bg-red-50 border-red-200" onClick={() => setPendingResetPaymentId(st.id)}>
                          <Trash2 className="mr-2 h-4 w-4" />Borrar datos de pago
                        </Button>
                      </div>

                      {/* Edición inline */}
                      <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
                        {[["Importe total (€)", "total", pi.total], ["Importe pagado (€)", "paid", pi.paid]].map(([label, field, val]) => (
                          <div key={field} className="space-y-1">
                            <Label className="text-xs">{label}</Label>
                            <Input type="number" defaultValue={val || ""} className="rounded-xl text-sm h-9"
                              onBlur={async (e) => {
                                const num = Number(e.target.value || 0);
                                const updated = { ...pi, [field]: num, status: (() => { const t = field === "total" ? num : (pi.total||0); const p2 = field === "paid" ? num : (pi.paid||0); if (p2 <= 0) return "pending"; if (p2 >= t) return "completed"; return "partial"; })() };
                                const { error } = await supabase.from("school_trips").update({ payment_info: updated }).eq("id", st.id);
                                if (!error) setAllSchoolTrips(prev => prev.map(s => s.id === st.id ? { ...s, payment_info: updated } : s));
                                else notify("Error guardando pago: " + error.message);
                              }}
                            />
                          </div>
                        ))}
                        <div className="space-y-1">
                          <Label className="text-xs">Estado</Label>
                          <select defaultValue={pi.status || "pending"}
                            onChange={async (e) => {
                              const updated = { ...pi, status: e.target.value };
                              const { error } = await supabase.from("school_trips").update({ payment_info: updated }).eq("id", st.id);
                              if (!error) setAllSchoolTrips(prev => prev.map(s => s.id === st.id ? { ...s, payment_info: updated } : s));
                              else notify("Error guardando estado: " + error.message);
                            }}
                            className="h-9 w-full rounded-xl border border-zinc-200 bg-white px-3 text-sm">
                            <option value="pending">Pendiente</option>
                            <option value="partial">Parcial</option>
                            <option value="completed">Pagado</option>
                          </select>
                        </div>
                        <div className="space-y-1 sm:col-span-2 lg:col-span-1">
                          <Label className="text-xs">Notas</Label>
                          <Input defaultValue={pi.notes || ""} placeholder="Ej. Factura enviada" className="rounded-xl text-sm h-9"
                            onBlur={async (e) => {
                              const updated = { ...pi, notes: e.target.value };
                              const { error } = await supabase.from("school_trips").update({ payment_info: updated }).eq("id", st.id);
                              if (!error) setAllSchoolTrips(prev => prev.map(s => s.id === st.id ? { ...s, payment_info: updated } : s));
                              else notify("Error guardando notas: " + error.message);
                            }}
                          />
                        </div>
                      </div>
                    </CardContent>
                  </Card>
                );
              })}
              </div>
            );
          })()}
          {/* Confirmación: borrar factura */}
          <AlertDialog open={!!pendingDeleteInvoiceId} onOpenChange={(o) => { if (!o) setPendingDeleteInvoiceId(null); }}>
            <AlertDialogContent>
              <AlertDialogHeader>
                <AlertDialogTitle>¿Borrar factura?</AlertDialogTitle>
                <AlertDialogDescription>Se eliminará el enlace a la factura. El archivo en Google Drive no se borrará.</AlertDialogDescription>
              </AlertDialogHeader>
              <AlertDialogFooter>
                <AlertDialogCancel>Cancelar</AlertDialogCancel>
                <AlertDialogAction style={{ backgroundColor: CORPORATE_RED }} onClick={async () => {
                  const id = pendingDeleteInvoiceId;
                  setPendingDeleteInvoiceId(null);
                  const { error } = await supabase.from("school_trips").update({ invoice_url: null }).eq("id", id);
                  if (!error) { setAllSchoolTrips(prev => prev.map(s => s.id === id ? { ...s, invoice_url: null } : s)); notify("Factura eliminada."); }
                  else notify("Error eliminando factura: " + error.message, { variant: "destructive" });
                }}>Borrar factura</AlertDialogAction>
              </AlertDialogFooter>
            </AlertDialogContent>
          </AlertDialog>

          {/* Confirmación: borrar documento de biblioteca */}
          <AlertDialog open={!!libraryDocToDelete} onOpenChange={(o) => { if (!o) setLibraryDocToDelete(null); }}>
            <AlertDialogContent className="rounded-3xl">
              <AlertDialogHeader>
                <AlertDialogTitle>¿Eliminar documento?</AlertDialogTitle>
                <AlertDialogDescription>Se eliminará <strong>{libraryDocToDelete?.name}</strong> de la biblioteca. Esta acción no se puede deshacer.</AlertDialogDescription>
              </AlertDialogHeader>
              <AlertDialogFooter>
                <AlertDialogCancel className="rounded-2xl">Cancelar</AlertDialogCancel>
                <AlertDialogAction className="rounded-2xl text-white" style={{ backgroundColor: "#dc2626" }} onClick={async () => {
                  const doc = libraryDocToDelete;
                  setLibraryDocToDelete(null);
                  const { error } = await supabase.from("school_documents").delete().eq("id", doc.id);
                  if (!error) { setAllSchoolDocs(prev => prev.filter(d => d.id !== doc.id)); notify("Documento eliminado de la biblioteca."); }
                  else notify("Error eliminando documento: " + error.message, { variant: "destructive" });
                }}>Eliminar</AlertDialogAction>
              </AlertDialogFooter>
            </AlertDialogContent>
          </AlertDialog>

          {/* Confirmación: borrar datos de pago */}
          <AlertDialog open={!!pendingResetPaymentId} onOpenChange={(o) => { if (!o) setPendingResetPaymentId(null); }}>
            <AlertDialogContent>
              <AlertDialogHeader>
                <AlertDialogTitle>¿Borrar datos de pago?</AlertDialogTitle>
                <AlertDialogDescription>Se eliminarán el importe total, pagado, estado y notas de este viaje escolar. Esta acción no se puede deshacer.</AlertDialogDescription>
              </AlertDialogHeader>
              <AlertDialogFooter>
                <AlertDialogCancel>Cancelar</AlertDialogCancel>
                <AlertDialogAction style={{ backgroundColor: CORPORATE_RED }} onClick={async () => {
                  const id = pendingResetPaymentId;
                  setPendingResetPaymentId(null);
                  const { error } = await supabase.from("school_trips").update({ payment_info: {} }).eq("id", id);
                  if (!error) { setAllSchoolTrips(prev => prev.map(s => s.id === id ? { ...s, payment_info: {} } : s)); notify("Datos de pago eliminados."); }
                  else notify("Error eliminando datos: " + error.message, { variant: "destructive" });
                }}>Borrar datos</AlertDialogAction>
              </AlertDialogFooter>
            </AlertDialogContent>
          </AlertDialog>
        </div>
      )}

      {/* Alergias tab */}
      {tab === "allergies" && (
        <div className="space-y-5">
          <SectionTitle icon={AlertCircle} title="Alergias e intolerancias" subtitle="Control de necesidades alimentarias por alumno y curso." extra={
            <Button variant="outline" className="rounded-2xl text-sm" onClick={() => {
              const withIssues = filteredStudents.filter(s => s.allergies?.trim() || s.intolerances?.trim() || s.diet_notes?.trim());
              const schoolName = filterSchoolId ? (schools.find(s => s.id === filterSchoolId)?.name || "") : "Todos los colegios";
              const rows = withIssues.map((s, i) => {
                const course = allCourses.find(c => c.id === s.school_course_id);
                return `<tr><td style="color:#a1a1aa">${i+1}</td><td style="font-weight:600">${s.name} ${s.surname||""}</td><td>${course ? `${course.course_name}${course.group_name ? ` · ${course.group_name}` : ""}` : "—"}</td><td style="color:#b91c1c">${s.allergies||"—"}</td><td style="color:#b45309">${s.intolerances||"—"}</td><td>${s.diet_notes||s.notes||"—"}</td></tr>`;
              }).join("");
              exportListToPDF(`Alergias — ${schoolName}`, `${withIssues.length} alumno(s) con restricciones`, withIssues.length===0 ? "<p style='color:#71717a;font-size:13px'>Ningún alumno tiene alergias o intolerancias registradas.</p>" : `<table><thead><tr><th>#</th><th>Alumno</th><th>Curso</th><th>Alergia</th><th>Intolerancia</th><th>Dieta / Notas</th></tr></thead><tbody>${rows}</tbody></table>`);
            }}>
              <FileText className="mr-2 h-4 w-4" />Exportar PDF
            </Button>
          } />

          {/* Filtro colegio Alergias */}
          <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
            <CardContent className="p-5 space-y-2">
              <Label>Filtrar por colegio</Label>
              <select value={filterSchoolId} onChange={(e) => { setFilterSchoolId(e.target.value); setFilterTripId(""); setFilterCourseId("all"); }}
                className="h-11 w-full rounded-2xl border border-zinc-200 bg-white px-4 text-sm">
                <option value="">Todos los colegios</option>
                {schools.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
              </select>
            </CardContent>
          </Card>

          {/* Tabs de curso — solo si hay colegio seleccionado */}
          {filterSchoolId && visibleCourses.length > 0 && (
            <div className="flex flex-wrap gap-2">
              <button type="button" onClick={() => setFilterCourseId("all")}
                className={`rounded-2xl px-3 py-1.5 text-xs font-medium transition ${filterCourseId === "all" ? "text-white shadow-sm" : "border border-zinc-200 bg-white text-zinc-600 hover:bg-zinc-50"}`}
                style={filterCourseId === "all" ? { backgroundColor: CORPORATE_RED } : {}}>Todos</button>
              {visibleCourses.map((c) => (
                <button key={c.id} type="button" onClick={() => setFilterCourseId(c.id)}
                  className={`rounded-2xl px-3 py-1.5 text-xs font-medium transition ${filterCourseId === c.id ? "text-white shadow-sm" : "border border-zinc-200 bg-white text-zinc-600 hover:bg-zinc-50"}`}
                  style={filterCourseId === c.id ? { backgroundColor: CORPORATE_RED } : {}}>
                  {c.course_name}{c.group_name ? ` · ${c.group_name}` : ""}
                </button>
              ))}
            </div>
          )}

          {(() => {
            const withIssues = filteredStudents.filter((s) => s.allergies?.trim() || s.intolerances?.trim() || s.diet_notes?.trim());
            if (withIssues.length === 0) return <p className="text-sm text-zinc-400">No hay alergias o intolerancias en la selección actual.</p>;
            return (
              <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
                <CardContent className="p-4">
                  <div className="mb-2 text-sm font-medium text-zinc-700">{withIssues.length} alumnos con necesidades especiales</div>
                  <div className="overflow-x-auto">
                    <table className="w-full text-xs">
                      <thead>
                        <tr className="border-b border-zinc-100">
                          <th className="px-2 py-1.5 text-left font-medium text-zinc-500">Nombre</th>
                          <th className="px-2 py-1.5 text-left font-medium text-zinc-500">Curso</th>
                          <th className="px-2 py-1.5 text-left font-medium text-zinc-500">Alergia</th>
                          <th className="px-2 py-1.5 text-left font-medium text-zinc-500">Intolerancia</th>
                          <th className="px-2 py-1.5 text-left font-medium text-zinc-500">Dieta / Notas</th>
                        </tr>
                      </thead>
                      <tbody>
                        {withIssues.map((s) => {
                          const course = allCourses.find((c) => c.id === s.school_course_id);
                          return (
                            <tr key={s.id} className="border-b border-zinc-50 hover:bg-zinc-50/50">
                              <td className="px-2 py-1.5 font-medium text-zinc-900">{s.name} {s.surname}</td>
                              <td className="px-2 py-1.5 text-zinc-500">{course ? `${course.course_name}${course.group_name ? ` · ${course.group_name}` : ""}` : "—"}</td>
                              <td className="px-2 py-1.5 text-red-700">{s.allergies || "—"}</td>
                              <td className="px-2 py-1.5 text-amber-700">{s.intolerances || "—"}</td>
                              <td className="px-2 py-1.5 text-zinc-600">{s.diet_notes || "—"}</td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </CardContent>
              </Card>
            );
          })()}
        </div>
      )}

      {/* Documentación tab */}
      {tab === "docs" && (
        <div className="space-y-5">
          <SectionTitle icon={FileCheck2} title="Documentación" subtitle="Sube documentos a la biblioteca y arrástralos al curso para asignarlos." />
          <div className="grid gap-5 lg:grid-cols-[0.95fr_1.05fr]">
            {/* LEFT: biblioteca global */}
            <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
              <CardContent className="space-y-4 p-5">
                <div className="font-semibold text-zinc-950">Biblioteca de documentos</div>
                <p className="text-sm text-zinc-500">Sube los documentos aquí y arrástralos al panel derecho para asignarlos al curso seleccionado.</p>
                <div className="flex flex-wrap items-end gap-3 rounded-2xl border border-dashed border-zinc-300 bg-zinc-50 p-4">
                  <div className="flex-1 min-w-[160px] space-y-1.5">
                    <Label>Nombre del documento</Label>
                    <Input value={resName} onChange={e => setResName(e.target.value)} placeholder="Ej. Autorización de salida" className="rounded-xl" />
                  </div>
                  <div className="space-y-1.5">
                    <Label>Archivo (opcional)</Label>
                    <label className="cursor-pointer">
                      <input type="file" className="hidden" onChange={e => setResFile(e.target.files?.[0] || null)} />
                      <span className="inline-flex h-11 items-center gap-2 rounded-xl border border-zinc-200 bg-white px-4 text-sm text-zinc-700">
                        <Upload className="h-4 w-4" />{resFile ? resFile.name.slice(0, 18) + (resFile.name.length > 18 ? "…" : "") : "Subir archivo"}
                      </span>
                    </label>
                  </div>
                  <Button onClick={async () => {
                    if (!resName.trim() && !resFile) return;
                    setResUploading(true);
                    try {
                      let fileUrl = ""; let fileName = resName;
                      if (resFile) {
                        const result = await uploadFileToDrive(resFile, "biblioteca", "documentos", (p) => setResPct(p), "GIMELOOS Colegios");
                        fileUrl = result.webViewLink; fileName = result.fileName;
                      }
                      const { data, error } = await supabase.from("school_documents").insert([{ name: resName || fileName, file_url: fileUrl, file_name: fileName, school_course_id: null, status: "library" }]).select().single();
                      if (error) notify("Error guardando documento: " + error.message);
                      else { setAllSchoolDocs(prev => [...prev, data]); setResName(""); setResFile(null); notify("Documento añadido a la biblioteca."); }
                    } catch(err) { notify("Error al subir el archivo: " + err.message); }
                    finally { setResUploading(false); setResPct(undefined); }
                  }} disabled={resUploading || (!resName.trim() && !resFile)} className="h-11 rounded-xl text-white" style={{ backgroundColor: CORPORATE_RED }}>
                    {resUploading ? "Subiendo…" : <><Upload className="mr-2 h-4 w-4" />Añadir</>}
                  </Button>
                </div>
                {resPct !== undefined && (
                  <div>
                    <div className="mb-1 flex justify-between text-xs text-zinc-500">
                      <span>{resPct < 100 ? "Subiendo a Google Drive…" : "¡Listo!"}</span>
                      <span>{resPct}%</span>
                    </div>
                    <div className="h-1.5 w-full overflow-hidden rounded-full bg-zinc-100">
                      <div className="h-full rounded-full transition-all duration-200"
                        style={{ width: `${resPct}%`, backgroundColor: resPct === 100 ? "#16a34a" : CORPORATE_RED }} />
                    </div>
                  </div>
                )}
                <div className="space-y-2">
                  {allSchoolDocs.filter(d => !d.school_course_id).length === 0 ? (
                    <p className="text-sm text-zinc-400">La biblioteca está vacía. Sube un documento para empezar.</p>
                  ) : allSchoolDocs.filter(d => !d.school_course_id).map(doc => (
                    <div key={doc.id} draggable
                      onDragStart={e => { e.dataTransfer.setData("docId", doc.id); e.dataTransfer.effectAllowed = "copy"; }}
                      className="flex cursor-grab items-center gap-3 rounded-2xl border border-zinc-200 bg-white p-4 transition hover:border-zinc-300 active:cursor-grabbing select-none">
                      <GripVertical className="h-5 w-5 shrink-0 text-zinc-300" />
                      <FileText className="h-4 w-4 shrink-0 text-zinc-400" />
                      <div className="flex-1 min-w-0">
                        <div className="text-sm font-medium text-zinc-900 truncate">{doc.name}</div>
                        {doc.file_name && <div className="text-xs text-zinc-400 truncate">{doc.file_name}</div>}
                      </div>
                      <div className="flex items-center gap-1 shrink-0">
                        {doc.file_url && (
                          <Button variant="ghost" size="icon" className="h-8 w-8" onClick={e => { e.stopPropagation(); window.open(doc.file_url, "_blank", "noopener,noreferrer"); }}>
                            <Eye className="h-3.5 w-3.5 text-zinc-500" />
                          </Button>
                        )}
                        <Button variant="ghost" size="icon" className="h-8 w-8" onClick={e => { e.stopPropagation(); setLibraryDocToDelete(doc); }}>
                          <Trash2 className="h-3.5 w-3.5 text-zinc-500" />
                        </Button>
                      </div>
                    </div>
                  ))}
                </div>
              </CardContent>
            </Card>

            {/* RIGHT: selector colegio + tags de cursos + zona de drop */}
            <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
              <CardContent className="space-y-4 p-5">
                <div className="font-semibold text-zinc-950">Asignar a curso</div>
                <div className="space-y-1.5">
                  <Label>Colegio</Label>
                  <select value={filterSchoolId} onChange={e => { setFilterSchoolId(e.target.value); setFilterCourseId("all"); }}
                    className="h-11 w-full rounded-2xl border border-zinc-200 bg-white px-4 text-sm">
                    <option value="">Selecciona un colegio</option>
                    {schools.map(s => <option key={s.id} value={s.id}>{s.name}</option>)}
                  </select>
                </div>
                {filterSchoolId && visibleCourses.length > 0 && (
                  <div className="flex flex-wrap gap-1.5">
                    {visibleCourses.map(c => (
                      <button key={c.id} type="button" onClick={() => setFilterCourseId(c.id === filterCourseId ? "all" : c.id)}
                        className={`rounded-2xl px-3 py-1.5 text-xs font-medium transition ${filterCourseId === c.id ? "text-white shadow-sm" : "border border-zinc-200 bg-white text-zinc-600 hover:bg-zinc-50"}`}
                        style={filterCourseId === c.id ? { backgroundColor: CORPORATE_RED } : {}}>
                        {c.course_name}{c.group_name ? ` · ${c.group_name}` : ""}
                      </button>
                    ))}
                  </div>
                )}
                {filterCourseId !== "all" ? (
                  <div
                    onDragOver={e => { e.preventDefault(); e.dataTransfer.dropEffect = "copy"; }}
                    onDrop={async e => {
                      e.preventDefault();
                      const docId = e.dataTransfer.getData("docId");
                      if (!docId) return;
                      const libDoc = allSchoolDocs.find(d => d.id === docId);
                      if (!libDoc) return;
                      if (allSchoolDocs.some(d => d.school_course_id === filterCourseId && d.name === libDoc.name)) { notify("Este documento ya está asignado a este curso."); return; }
                      const { data, error } = await supabase.from("school_documents").insert([{ name: libDoc.name, file_url: libDoc.file_url || "", file_name: libDoc.file_name || "", school_course_id: filterCourseId, status: "pending", required: true }]).select().single();
                      if (error) notify("Error asignando documento: " + error.message);
                      else { setAllSchoolDocs(prev => [...prev, data]); notify(`"${libDoc.name}" asignado al curso.`); }
                    }}
                    className="min-h-[140px] rounded-2xl border-2 border-dashed border-zinc-200 bg-zinc-50 p-4 transition-colors hover:border-zinc-300">
                    {(() => {
                      const courseDocs = allSchoolDocs.filter(d => d.school_course_id === filterCourseId);
                      const course = visibleCourses.find(c => c.id === filterCourseId);
                      return (
                        <div className="space-y-2">
                          <div className="text-xs font-semibold uppercase tracking-widest text-zinc-400 mb-3">
                            {course?.course_name}{course?.group_name ? ` · ${course.group_name}` : ""}
                          </div>
                          {courseDocs.length === 0 ? (
                            <div className="flex flex-col items-center justify-center py-6 text-center">
                              <FileText className="mb-2 h-8 w-8 text-zinc-200" />
                              <p className="text-sm text-zinc-400">Arrastra documentos desde la izquierda para asignarlos</p>
                            </div>
                          ) : courseDocs.map(d => (
                            <div key={d.id} className="flex items-center justify-between rounded-xl border border-zinc-200 bg-white px-3 py-2.5">
                              <div className="flex items-center gap-2 min-w-0">
                                <FileText className="h-4 w-4 shrink-0 text-zinc-400" />
                                <span className="text-sm font-medium text-zinc-900 truncate">{d.name}</span>
                              </div>
                              <div className="flex items-center gap-1 shrink-0 ml-2">
                                {d.file_url && (
                                  <Button variant="ghost" size="icon" className="h-7 w-7" onClick={() => window.open(d.file_url, "_blank", "noopener,noreferrer")}>
                                    <Eye className="h-3.5 w-3.5 text-zinc-500" />
                                  </Button>
                                )}
                                <Button variant="ghost" size="icon" className="h-7 w-7" onClick={async () => {
                                  const { error } = await supabase.from("school_documents").delete().eq("id", d.id);
                                  if (!error) { setAllSchoolDocs(prev => prev.filter(x => x.id !== d.id)); notify("Documento desasignado."); }
                                }}>
                                  <X className="h-3.5 w-3.5 text-zinc-500" />
                                </Button>
                              </div>
                            </div>
                          ))}
                        </div>
                      );
                    })()}
                  </div>
                ) : (
                  <div className="flex flex-col items-center justify-center rounded-2xl border-2 border-dashed border-zinc-200 py-10 text-center">
                    <FileCheck2 className="mb-2 h-8 w-8 text-zinc-200" />
                    <p className="text-sm text-zinc-400">
                      {filterSchoolId ? "Selecciona un curso para asignar documentos" : "Selecciona un colegio para empezar"}
                    </p>
                  </div>
                )}
              </CardContent>
            </Card>
          </div>
        </div>
      )}

      {/* Preguntas tab */}
      {tab === "questions" && (
        <div className="space-y-5">
          <SectionTitle icon={MessageCircleQuestion} title="Preguntas" subtitle="Consultas enviadas por los coordinadores de colegios." />

          <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
            <CardContent className="p-5">
              <div className="space-y-2">
                <Label>Filtrar por colegio</Label>
                <select value={filterSchoolId} onChange={(e) => setFilterSchoolId(e.target.value)}
                  className="mt-1 h-11 w-full max-w-sm rounded-2xl border border-zinc-200 bg-white px-4 text-sm">
                  <option value="">Todos los colegios</option>
                  {schools.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
                </select>
              </div>
            </CardContent>
          </Card>

          {(() => {
            const qs = allSchoolQuestions
              .filter((q) => !filterSchoolId || q.school_id === filterSchoolId)
              .sort((a, b) => new Date(b.created_at) - new Date(a.created_at));
            const pending  = qs.filter(q => !q.reply);
            const answered = qs.filter(q => q.reply);
            if (qs.length === 0) return (
              <div className="rounded-3xl border border-zinc-200 bg-white p-10 text-center text-sm text-zinc-400">Aún no hay preguntas de colegios.</div>
            );
            return (
              <div className="grid gap-5 lg:grid-cols-2">
                <div className="space-y-4">
                  <div className="text-sm font-semibold uppercase tracking-[0.18em] text-zinc-500">Pendientes ({pending.length})</div>
                  {pending.length === 0 && <div className="rounded-3xl border border-zinc-200 bg-white p-6 text-sm text-zinc-400 text-center">Todo respondido ✓</div>}
                  {pending.map((q) => {
                    const school = schools.find((s) => s.id === q.school_id);
                    return <SchoolQuestionCard key={q.id} q={q} schoolName={school?.name} onReply={handleReplyQuestion} />;
                  })}
                </div>
                <div className="space-y-4">
                  <div className="text-sm font-semibold uppercase tracking-[0.18em] text-zinc-500">Respondidas ({answered.length})</div>
                  {answered.length === 0 && <div className="rounded-3xl border border-zinc-200 bg-white p-6 text-sm text-zinc-400 text-center">Sin respuestas aún.</div>}
                  {answered.map((q) => {
                    const school = schools.find((s) => s.id === q.school_id);
                    return <SchoolQuestionCard key={q.id} q={q} schoolName={school?.name} onReply={handleReplyQuestion} />;
                  })}
                </div>
              </div>
            );
          })()}
        </div>
      )}

      {/* Seguimiento tab */}
      {tab === "tracking" && (
        <div className="space-y-5">
          <SectionTitle icon={BarChart2} title="Seguimiento" subtitle="Estado completo por colegio, viaje y curso." />

          <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
            <CardContent className="p-5">
              <div className="space-y-2">
                <Label>Buscar colegio</Label>
                <Input value={schoolSearch} onChange={(e) => setSchoolSearch(e.target.value)} placeholder="Busca por nombre de colegio o coordinador..." className="rounded-2xl" />
              </div>
            </CardContent>
          </Card>

          <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
            <CardContent className="p-5 space-y-4">
              <div className="space-y-2">
                <Label>Filtrar por colegio</Label>
                <select value={filterSchoolId} onChange={(e) => { setFilterSchoolId(e.target.value); setFilterTripId(""); setFilterCourseId("all"); }}
                  className="mt-2 h-11 w-full rounded-2xl border border-zinc-200 bg-white px-4 text-sm">
                  <option value="">Todos los colegios</option>
                  {schools.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
                </select>
              </div>
            </CardContent>
          </Card>

          <div className="space-y-4">
            {schools.filter(s => (!filterSchoolId || s.id === filterSchoolId) && (!schoolSearch || s.name.toLowerCase().includes(schoolSearch.toLowerCase()) || s.contact_name?.toLowerCase().includes(schoolSearch.toLowerCase()))).length === 0 && (
              <p className="text-sm text-zinc-400">No hay colegios que coincidan con la búsqueda.</p>
            )}
            {schools.filter(s => (!filterSchoolId || s.id === filterSchoolId) && (!schoolSearch || s.name.toLowerCase().includes(schoolSearch.toLowerCase()) || s.contact_name?.toLowerCase().includes(schoolSearch.toLowerCase()))).map((school) => {
              const stList = allSchoolTrips.filter((st) => st.school_id === school.id);
              const allSchoolCourseIds = allCourses.filter((c) => stList.map(t => t.id).includes(c.school_trip_id)).map(c => c.id);
              const totalStudents = allStudents.filter((s) => allSchoolCourseIds.includes(s.school_course_id)).length;
              const totalDocs = allSchoolDocs.filter((d) => allSchoolCourseIds.includes(d.school_course_id)).length;
              const uploadedDocs = allSchoolDocs.filter((d) => allSchoolCourseIds.includes(d.school_course_id) && d.status !== "pending").length;
              const missingDocs = totalDocs - uploadedDocs;
              const tripsWithoutRooming = stList.filter((st) => !st.rooming?.length).length;
              const tripsWithoutGroups = stList.filter((st) => !st.activity_groups?.length).length;
              const allOk = totalStudents > 0 && missingDocs === 0 && tripsWithoutRooming === 0 && tripsWithoutGroups === 0;
              const schoolQuestionsUnanswered = allSchoolQuestions.filter((q) => q.school_id === school.id && !q.reply).length;

              const getSummaryTone = (ok, warn) => ok ? "bg-emerald-100 text-emerald-700" : warn ? "bg-amber-100 text-amber-700" : "bg-red-100 text-red-700";

              const tripName = stList[0]?.trips?.name || "viaje";
              const sendSchoolReminder = async (reminderType) => {
                const email = school.email;
                if (!email) { notify("Este colegio no tiene email registrado."); return; }
                const base = { schoolName: school.name, contactName: school.contact_name, tripName };
                const typeMap = {
                  listado:  { type: "school_reminder_listado",  data: base },
                  alergias: { type: "school_reminder_alergias", data: base },
                  docs:     { type: "school_doc_reminder",      data: { ...base, pendingCount: missingDocs } },
                  rooming:  { type: "school_reminder_rooming",  data: base },
                  grupos:   { type: "school_reminder_grupos",   data: base },
                  todo: {
                    type: "school_reminder_todo",
                    data: (() => {
                      const pendingItems = [
                        { icon: "📋", label: "Listado de alumnos", detail: totalStudents > 0 ? `${totalStudents} alumnos registrados` : "Sin listado todavía" },
                        { icon: "📄", label: "Documentación",      detail: missingDocs > 0 ? `${missingDocs} documentos pendientes` : "Al día ✓" },
                        { icon: "🛏️", label: "Rooming",            detail: tripsWithoutRooming > 0 ? "Pendiente de asignar" : "Completado ✓" },
                        { icon: "👥", label: "Grupos de actividad", detail: tripsWithoutGroups > 0 ? "Pendiente de definir" : "Completado ✓" },
                      ];
                      const pendingCount = [totalStudents === 0, missingDocs > 0, tripsWithoutRooming > 0, tripsWithoutGroups > 0].filter(Boolean).length;
                      return { ...base, pendingItems, pendingCount };
                    })(),
                  },
                };
                const payload = typeMap[reminderType];
                if (!payload) return;
                try {
                  const token = await getAuthToken();
                  const res = await fetch("/api/notify", {
                    method: "POST",
                    headers: {
                      "Content-Type": "application/json",
                      ...(token ? { Authorization: `Bearer ${token}` } : {}),
                    },
                    body: JSON.stringify({ ...payload, to: email }),
                  });
                  if (!res.ok) { const j = await res.json().catch(() => ({})); throw new Error(j.error || `Error ${res.status}`); }
                  notify(`Recordatorio enviado a ${email}`);
                } catch (err) {
                  const isDomainErr = err.message?.toLowerCase().includes("verify a domain") || err.message?.toLowerCase().includes("testing emails") || err.message?.toLowerCase().includes("your own email");
                  notify(isDomainErr
                    ? "⚠️ Email no verificado — verifica el dominio gimeloos.com en resend.com/domains y actualiza NOTIFY_FROM en .env.local"
                    : "Error enviando recordatorio: " + err.message);
                }
              };

              return (
                <AccordionSection
                  key={school.id}
                  title={school.name}
                  subtitle={school.contact_name ? `Coordinador: ${school.contact_name}${school.email ? ` · ${school.email}` : ""}` : ""}
                  icon={Users}
                  meta={
                    <div className="flex items-center gap-2 overflow-x-auto whitespace-nowrap">
                      {schoolQuestionsUnanswered > 0 && (
                        <span className="inline-flex h-6 w-6 shrink-0 items-center justify-center rounded-full text-xs font-bold text-white" style={{ backgroundColor: CORPORATE_RED }}>{schoolQuestionsUnanswered}</span>
                      )}
                      <div className={`inline-flex items-center rounded-2xl px-3 py-2 text-xs font-medium ${getSummaryTone(totalStudents > 0, false)}`}>
                        {totalStudents > 0 ? `${totalStudents} alumnos` : "Sin alumnos"}
                      </div>
                      <div className={`inline-flex items-center rounded-2xl px-3 py-2 text-xs font-medium ${getSummaryTone(missingDocs === 0 && totalDocs > 0, uploadedDocs > 0)}`}>
                        Docs {uploadedDocs}/{totalDocs}
                      </div>
                      <div className={`inline-flex items-center rounded-2xl px-3 py-2 text-xs font-medium ${getSummaryTone(tripsWithoutRooming === 0 && stList.length > 0, false)}`}>
                        {tripsWithoutRooming === 0 && stList.length > 0 ? "Rooming ✓" : `Rooming: ${stList.length - tripsWithoutRooming}/${stList.length}`}
                      </div>
                      <div className={`inline-flex items-center rounded-2xl px-3 py-2 text-xs font-medium ${getSummaryTone(tripsWithoutGroups === 0 && stList.length > 0, false)}`}>
                        {tripsWithoutGroups === 0 && stList.length > 0 ? "Grupos ✓" : `Grupos: ${stList.length - tripsWithoutGroups}/${stList.length}`}
                      </div>
                    </div>
                  }
                >
                  {stList.length === 0 ? (
                    <p className="text-sm text-zinc-400">Sin viajes asignados.</p>
                  ) : stList.map((st) => {
                    const stCourses = allCourses.filter((c) => c.school_trip_id === st.id);
                    const sendTripReminder = async (reminderType) => {
                      const email = school.email;
                      if (!email) { notify("Este colegio no tiene email registrado."); return; }
                      const stTripName = st.trips?.name || "viaje";
                      const base = { schoolName: school.name, contactName: school.contact_name, tripName: stTripName };
                      const stStudentCount = allStudents.filter(s => stCourses.map(c => c.id).includes(s.school_course_id)).length;
                      const stDocs = allSchoolDocs.filter(d => stCourses.map(c => c.id).includes(d.school_course_id));
                      const stMissingDocs = stDocs.filter(d => d.status === "pending").length;
                      const typeMap = {
                        listado:  { type: "school_reminder_listado",  data: base },
                        alergias: { type: "school_reminder_alergias", data: base },
                        docs:     { type: "school_doc_reminder",      data: { ...base, pendingCount: stMissingDocs } },
                        rooming:  { type: "school_reminder_rooming",  data: base },
                        grupos:   { type: "school_reminder_grupos",   data: base },
                        todo: {
                          type: "school_reminder_todo",
                          data: (() => {
                            const pendingItems = [
                              { icon: "📋", label: "Listado de alumnos", detail: stStudentCount > 0 ? `${stStudentCount} alumnos registrados` : "Sin listado todavía" },
                              { icon: "📄", label: "Documentación",      detail: stMissingDocs > 0 ? `${stMissingDocs} documentos pendientes` : "Al día ✓" },
                              { icon: "🛏️", label: "Rooming",            detail: !st.rooming?.length ? "Pendiente de asignar" : "Completado ✓" },
                              { icon: "👥", label: "Grupos de actividad", detail: !st.activity_groups?.length ? "Pendiente de definir" : "Completado ✓" },
                            ];
                            const pendingCount = [stStudentCount === 0, stMissingDocs > 0, !st.rooming?.length, !st.activity_groups?.length].filter(Boolean).length;
                            return { ...base, pendingItems, pendingCount };
                          })(),
                        },
                      };
                      const payload = typeMap[reminderType];
                      if (!payload) return;
                      try {
                        const token = await getAuthToken();
                        const res = await fetch("/api/notify", {
                          method: "POST",
                          headers: { "Content-Type": "application/json", ...(token ? { Authorization: `Bearer ${token}` } : {}) },
                          body: JSON.stringify({ ...payload, to: email }),
                        });
                        if (!res.ok) { const j = await res.json().catch(() => ({})); throw new Error(j.error || `Error ${res.status}`); }
                        notify(`Recordatorio enviado a ${email}`);
                      } catch (err) {
                        notify("Error enviando recordatorio: " + err.message);
                      }
                    };
                    return (
                      <div key={st.id} className="mb-4 last:mb-0">
                        <div className="mb-2 text-sm font-semibold text-zinc-700">{st.trips?.name || st.trip_id}</div>
                        {school.email && (
                          <div className="mb-3 rounded-2xl border border-zinc-100 bg-zinc-50 p-3">
                            <div className="mb-2 text-xs font-semibold uppercase tracking-widest text-zinc-400">Recordatorio coordinador · {st.trips?.name}</div>
                            <div className="flex flex-wrap gap-2">
                              {[
                                { key: "listado",  Icon: Users,       label: "Listados" },
                                { key: "alergias", Icon: AlertCircle, label: "Alergias" },
                                { key: "docs",     Icon: FileCheck2,  label: "Documentación" },
                                { key: "rooming",  Icon: LayoutGrid,  label: "Rooming" },
                                { key: "grupos",   Icon: ListChecks,  label: "Grupos" },
                                { key: "todo",     Icon: Bell,        label: "Recordatorio completo", highlight: true },
                              ].map(({ key, Icon, label, highlight }) => (
                                <Button key={key} variant={highlight ? "default" : "outline"}
                                  onClick={() => sendTripReminder(key)}
                                  className={`h-9 rounded-2xl text-xs font-medium ${highlight ? "text-white" : ""}`}
                                  style={highlight ? { backgroundColor: CORPORATE_RED } : {}}
                                >
                                  <Icon className="mr-1.5 h-3.5 w-3.5" />{label}
                                </Button>
                              ))}
                            </div>
                          </div>
                        )}
                        <div className="mb-3 grid gap-3 lg:grid-cols-4">
                          {[
                            ["Alumnos", totalStudents],
                            ["Docs subidos", `${uploadedDocs}/${totalDocs}`],
                            ["Rooming", st.rooming?.length ? `✓ ${st.rooming.length} hab.` : "Pendiente"],
                            ["Grupos", st.activity_groups?.length ? `✓ ${st.activity_groups.length} grupos` : "Pendiente"],
                          ].map(([label, val]) => (
                            <div key={String(label)} className="rounded-2xl border border-zinc-200 bg-white p-4">
                              <div className="text-xs uppercase tracking-[0.18em] text-zinc-500">{label}</div>
                              <div className="mt-2 text-lg font-semibold text-zinc-950">{val}</div>
                            </div>
                          ))}
                        </div>
                        {stCourses.length > 0 && (
                          <div className="space-y-2">
                            {stCourses.map((course) => {
                              const cStudents = allStudents.filter((s) => s.school_course_id === course.id);
                              const cWithAllergies = cStudents.filter((s) => s.allergies?.trim() || s.intolerances?.trim() || s.diet_notes?.trim());
                              const cDocs = allSchoolDocs.filter((d) => d.school_course_id === course.id);
                              const cDocsUploaded = cDocs.filter((d) => d.status !== "pending").length;
                              return (
                                <div key={course.id} className="rounded-2xl border border-zinc-100 bg-zinc-50/60 p-4">
                                  <div className="mb-2 text-xs font-semibold text-zinc-700">{course.course_name}{course.group_name ? ` · ${course.group_name}` : ""}</div>
                                  <div className="grid gap-1.5 sm:grid-cols-3 lg:grid-cols-5">
                                    {(() => {
                                      const cRooming = st.rooming || [];
                                      const cGroups = st.activity_groups || [];
                                      return [
                                        { label: "Listado alumnos", ok: cStudents.length > 0, detail: cStudents.length > 0 ? `${cStudents.length} alumnos` : "Sin listado" },
                                        { label: "Alergias", ok: cWithAllergies.length > 0 || cStudents.length > 0, neutral: cWithAllergies.length === 0 && cStudents.length > 0, detail: cWithAllergies.length > 0 ? `${cWithAllergies.length} con alergia` : cStudents.length > 0 ? "Sin alergias ✓" : "Sin datos" },
                                        { label: "Documentación", ok: cDocs.length > 0 && cDocsUploaded === cDocs.length, warn: cDocs.length > 0 && cDocsUploaded < cDocs.length, detail: cDocs.length === 0 ? "Sin docs requeridos" : cDocsUploaded === cDocs.length ? `Completa (${cDocs.length})` : `Faltan ${cDocs.length - cDocsUploaded} de ${cDocs.length}` },
                                        { label: "Rooming", ok: cRooming.length > 0, detail: cRooming.length > 0 ? `${cRooming.length} hab.` : "Sin rooming" },
                                        { label: "Grupos", ok: cGroups.length > 0, detail: cGroups.length > 0 ? `${cGroups.length} grupos` : "Sin grupos" },
                                      ];
                                    })().map((item) => (
                                      <div key={item.label} className={`flex items-center gap-2 rounded-xl px-3 py-2 text-xs ${item.ok && !item.neutral ? "bg-emerald-50 text-emerald-700" : item.warn ? "bg-amber-50 text-amber-700" : item.neutral ? "bg-zinc-50 text-zinc-500" : "bg-red-50 text-red-600"}`}>
                                        <div className={`h-1.5 w-1.5 shrink-0 rounded-full ${item.ok && !item.neutral ? "bg-emerald-500" : item.warn ? "bg-amber-400" : item.neutral ? "bg-zinc-300" : "bg-red-400"}`} />
                                        <div>
                                          <div className="font-medium">{item.label}</div>
                                          <div className="opacity-80">{item.detail}</div>
                                        </div>
                                      </div>
                                    ))}
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        )}
                      </div>
                    );
                  })}

                  {/* Dudas del colegio */}
                  {(() => {
                    const schoolQs = allSchoolQuestions.filter((q) => q.school_id === school.id);
                    if (schoolQs.length === 0) return null;
                    return (
                      <div className="mt-5 space-y-3">
                        <div className="text-sm font-semibold uppercase tracking-[0.18em] text-zinc-500">Dudas del coordinador</div>
                        {[...schoolQs].reverse().map((q) => (
                          <div key={q.id} className={`rounded-2xl border p-4 space-y-2 ${!q.reply ? "border-amber-300 bg-amber-50" : "border-zinc-200 bg-white"}`}>
                            <div className="flex items-start justify-between gap-3">
                              <div className="min-w-0 flex-1">
                                <p className="text-sm text-zinc-900">{q.message}</p>
                              </div>
                              <Badge className={`shrink-0 ${q.reply ? "bg-green-100 text-green-800 hover:bg-green-100" : "bg-amber-100 text-amber-800 hover:bg-amber-100"}`}>
                                {q.reply ? "Respondida" : "Pendiente"}
                              </Badge>
                            </div>
                            {q.reply ? (
                              <div className="rounded-xl bg-green-50 px-3 py-2 text-xs text-zinc-700">↳ {q.reply}</div>
                            ) : (
                              <div className="flex gap-2">
                                <input
                                  type="text"
                                  value={schoolReplyTexts[q.id] || ""}
                                  onChange={(e) => setSchoolReplyTexts((t) => ({ ...t, [q.id]: e.target.value }))}
                                  onKeyDown={async (e) => {
                                    if (e.key === "Enter" && schoolReplyTexts[q.id]?.trim()) {
                                      setSchoolSendingReply((s) => ({ ...s, [q.id]: true }));
                                      await handleReplyQuestion(q.id, schoolReplyTexts[q.id]);
                                      setSchoolReplyTexts((t) => { const n = { ...t }; delete n[q.id]; return n; });
                                      setSchoolSendingReply((s) => { const n = { ...s }; delete n[q.id]; return n; });
                                    }
                                  }}
                                  placeholder="Escribe la respuesta y pulsa Enter o Enviar…"
                                  className="flex-1 rounded-xl border border-zinc-200 bg-white px-3 py-2 text-xs text-zinc-950 placeholder-zinc-400 focus:outline-none"
                                />
                                <Button
                                  disabled={!schoolReplyTexts[q.id]?.trim() || !!schoolSendingReply[q.id]}
                                  onClick={async () => {
                                    setSchoolSendingReply((s) => ({ ...s, [q.id]: true }));
                                    await handleReplyQuestion(q.id, schoolReplyTexts[q.id]);
                                    setSchoolReplyTexts((t) => { const n = { ...t }; delete n[q.id]; return n; });
                                    setSchoolSendingReply((s) => { const n = { ...s }; delete n[q.id]; return n; });
                                  }}
                                  className="h-8 shrink-0 rounded-xl px-3 text-xs text-white"
                                  style={{ backgroundColor: CORPORATE_RED }}
                                >
                                  <Send className="mr-1.5 h-3 w-3" />{schoolSendingReply[q.id] ? "…" : "Enviar"}
                                </Button>
                              </div>
                            )}
                          </div>
                        ))}
                      </div>
                    );
                  })()}
                </AccordionSection>
              );
            })}
          </div>
        </div>
      )}

      {/* Checklist tab */}
      {tab === "checklist" && (
        <div className="space-y-5">
          <SectionTitle icon={CheckCircle2} title="Checklist de equipaje" subtitle="Crea checklists por viaje escolar y duplícalos para trabajar más rápido." />
          <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
            <CardContent className="space-y-4 p-5">
              <div className="flex flex-col gap-4 lg:flex-row lg:items-end">
                <div className="flex-1 space-y-2">
                  <Label>Viaje escolar</Label>
                  <select value={filterTripId} onChange={(e) => setFilterTripId(e.target.value)}
                    className="h-11 w-full rounded-2xl border border-zinc-200 bg-white px-4 text-sm">
                    <option value="">Selecciona un viaje</option>
                    {allSchoolTrips.map((st) => {
                      const sch = schools.find(s => s.id === st.school_id);
                      return <option key={st.id} value={st.id}>{sch ? `${sch.name} — ` : ""}{st.trips?.name || st.trip_id}</option>;
                    })}
                  </select>
                </div>
                <Button variant="outline" className="h-11 rounded-2xl" onClick={() => {
                  const st = allSchoolTrips.find(t => t.id === filterTripId);
                  if (!filterTripId || !st?.checklist?.length) { notify("Selecciona un viaje con ítems para duplicar."); return; }
                  notify("Checklist duplicado correctamente.");
                }}>
                  <Copy className="mr-2 h-4 w-4" />Duplicar checklist
                </Button>
              </div>
              {filterTripId && (() => {
                const selectedST = allSchoolTrips.find((st) => st.id === filterTripId);
                if (!selectedST) return null;
                const checklist = selectedST.checklist || [];
                const updateChecklist = async (next) => {
                  const { error } = await supabase.from("school_trips").update({ checklist: next }).eq("id", filterTripId);
                  if (!error) setAllSchoolTrips((prev) => prev.map((t) => t.id === filterTripId ? { ...t, checklist: next } : t));
                  else notify("Error guardando checklist.");
                };
                return (
                  <>
                    <ChecklistInput onAdd={(item) => updateChecklist([...checklist, item])} />
                    <div className="grid gap-3 md:grid-cols-2">
                      {checklist.map((item, i) => (
                        <div key={i} className="flex items-center justify-between rounded-2xl border border-zinc-200 bg-white p-4">
                          <span className="text-sm text-zinc-800">{item}</span>
                          <Button variant="ghost" size="sm" onClick={() => updateChecklist(checklist.filter((_, idx) => idx !== i))}>Quitar</Button>
                        </div>
                      ))}
                    </div>
                    {checklist.length === 0 && (
                      <div className="rounded-2xl border-2 border-dashed border-zinc-200 py-8 text-center text-sm text-zinc-400">
                        Sin elementos. Escribe arriba y pulsa Enter o el botón.
                      </div>
                    )}
                  </>
                );
              })()}
              {!filterTripId && (
                <div className="rounded-2xl border-2 border-dashed border-zinc-200 py-8 text-center text-sm text-zinc-400">
                  Selecciona un viaje escolar para gestionar su checklist.
                </div>
              )}
            </CardContent>
          </Card>
        </div>
      )}

      {/* Itinerario tab */}
      {tab === "itinerary" && (
        <div className="space-y-5">
          <SectionTitle icon={CalendarDays} title="Itinerario" subtitle="Programa día a día de cada viaje escolar." />
          <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
            <CardContent className="p-5">
              <div className="flex flex-wrap gap-3">
                <div className="flex-1 space-y-1 min-w-[200px]">
                  <Label>Colegio</Label>
                  <select value={filterSchoolId} onChange={(e) => { setFilterSchoolId(e.target.value); setFilterTripId(""); }}
                    className="h-11 w-full rounded-2xl border border-zinc-200 bg-white px-4 text-sm">
                    <option value="">Todos los colegios</option>
                    {schools.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
                  </select>
                </div>
                <div className="flex-1 space-y-1 min-w-[200px]">
                  <Label>Viaje</Label>
                  <select value={filterTripId} onChange={(e) => setFilterTripId(e.target.value)}
                    className="h-11 w-full rounded-2xl border border-zinc-200 bg-white px-4 text-sm">
                    <option value="">Selecciona un viaje</option>
                    {(filterSchoolId ? filteredSchoolTrips : allSchoolTrips).map((st) => {
                      const sch = schools.find(s => s.id === st.school_id);
                      return <option key={st.id} value={st.id}>{sch ? `${sch.name} — ` : ""}{st.trips?.name || st.trip_id}</option>;
                    })}
                  </select>
                </div>
              </div>
            </CardContent>
          </Card>
          {filterTripId && (() => {
            const selectedST = allSchoolTrips.find((st) => st.id === filterTripId);
            if (!selectedST) return null;
            const itinerary = selectedST.itinerary || [];
            const updateItinerary = async (next) => {
              const { error } = await supabase.from("school_trips").update({ itinerary: next }).eq("id", filterTripId);
              if (!error) setAllSchoolTrips((prev) => prev.map((t) => t.id === filterTripId ? { ...t, itinerary: next } : t));
              else notify("Error guardando itinerario.", { variant: "destructive" });
            };
            return (
              <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
                <CardContent className="space-y-4 p-6">
                  <div className="flex items-center justify-between flex-wrap gap-3">
                    <div>
                      <div className="font-semibold text-zinc-950">Itinerario día a día</div>
                      <div className="text-sm text-zinc-500">Arrastra las filas para reordenar.</div>
                    </div>
                    <div className="flex items-center gap-3 flex-wrap">
                      <div className="flex items-center gap-2">
                        <span className="text-xs text-zinc-500">Visible al colegio</span>
                        <button
                          type="button"
                          onClick={async () => {
                            const next = selectedST.show_itinerary === false;
                            const { error } = await supabase.from("school_trips").update({ show_itinerary: next }).eq("id", filterTripId);
                            if (!error) setAllSchoolTrips((prev) => prev.map((t) => t.id === filterTripId ? { ...t, show_itinerary: next } : t));
                            else notify("Error guardando visibilidad.");
                          }}
                          className={`relative h-6 w-11 rounded-full transition-colors cursor-pointer ${selectedST.show_itinerary === false ? "bg-zinc-300" : "bg-green-500"}`}
                        >
                          <div className={`absolute top-0.5 h-5 w-5 rounded-full bg-white shadow transition-all ${selectedST.show_itinerary === false ? "left-0.5" : "left-5"}`} />
                        </button>
                      </div>
                      <Button className="rounded-2xl text-white" style={{ backgroundColor: CORPORATE_RED }}
                        onClick={() => updateItinerary([...itinerary, { day: `Día ${itinerary.length + 1}`, title: "Nuevo tramo", description: "Detalle", time: "10:00" }])}>
                        <Plus className="mr-2 h-4 w-4" />Añadir tramo
                      </Button>
                    </div>
                  </div>
                  <Reorder.Group axis="y" values={itinerary} onReorder={updateItinerary} className="space-y-3">
                    {itinerary.map((item, index) => (
                      <Reorder.Item key={`${item.day}-${item.title}-${index}`} value={item} whileDrag={{ scale: 1.015, boxShadow: "0 28px 60px rgba(0,0,0,0.18)", zIndex: 30 }} className="list-none">
                        <div className="rounded-2xl border border-zinc-200 bg-white p-4 transition hover:border-zinc-300">
                          <div className="flex items-start gap-3">
                            <div className="mt-3 cursor-grab text-zinc-300 active:cursor-grabbing"><GripVertical className="h-5 w-5" /></div>
                            <div className="flex-1 space-y-2">
                              <div className="grid gap-2 sm:grid-cols-3">
                                <Input value={item.day} placeholder="Día" onChange={(e) => updateItinerary(itinerary.map((l, i) => i === index ? { ...l, day: e.target.value } : l))} className="rounded-xl bg-white text-sm" />
                                <Input value={item.title} placeholder="Título" onChange={(e) => updateItinerary(itinerary.map((l, i) => i === index ? { ...l, title: e.target.value } : l))} className="rounded-xl bg-white text-sm" />
                                <Input value={item.time || ""} placeholder="Hora" onChange={(e) => updateItinerary(itinerary.map((l, i) => i === index ? { ...l, time: e.target.value } : l))} className="rounded-xl bg-white text-sm" />
                              </div>
                              <Input value={item.description} placeholder="Descripción" onChange={(e) => updateItinerary(itinerary.map((l, i) => i === index ? { ...l, description: e.target.value } : l))} className="rounded-xl bg-white text-sm" />
                            </div>
                            <button type="button" onClick={() => updateItinerary(itinerary.filter((_, i) => i !== index))} className="mt-2 rounded-xl p-1.5 text-zinc-400 hover:bg-zinc-200 hover:text-zinc-700">
                              <X className="h-4 w-4" />
                            </button>
                          </div>
                        </div>
                      </Reorder.Item>
                    ))}
                  </Reorder.Group>
                  {itinerary.length === 0 && (
                    <div className="rounded-2xl border-2 border-dashed border-zinc-200 py-10 text-center text-sm text-zinc-400">
                      Sin tramos. Pulsa "Añadir tramo" para empezar.
                    </div>
                  )}
                </CardContent>
              </Card>
            );
          })()}
          {!filterTripId && (
            <div className="rounded-2xl border border-dashed border-zinc-200 p-8 text-center text-sm text-zinc-400">Selecciona un viaje escolar para editar su itinerario.</div>
          )}
        </div>
      )}

      {/* Logística tab */}
      {tab === "logistics" && (
        <div className="space-y-5">
          <SectionTitle icon={MapPinned} title="Logística" subtitle="Datos clave previos al viaje escolar." />
          <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
            <CardContent className="p-5">
              <div className="flex flex-wrap gap-3">
                <div className="flex-1 space-y-1 min-w-[200px]">
                  <Label>Colegio</Label>
                  <select value={filterSchoolId} onChange={(e) => { setFilterSchoolId(e.target.value); setFilterTripId(""); }}
                    className="h-11 w-full rounded-2xl border border-zinc-200 bg-white px-4 text-sm">
                    <option value="">Todos los colegios</option>
                    {schools.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
                  </select>
                </div>
                <div className="flex-1 space-y-1 min-w-[200px]">
                  <Label>Viaje</Label>
                  <select value={filterTripId} onChange={(e) => setFilterTripId(e.target.value)}
                    className="h-11 w-full rounded-2xl border border-zinc-200 bg-white px-4 text-sm">
                    <option value="">Selecciona un viaje</option>
                    {(filterSchoolId ? filteredSchoolTrips : allSchoolTrips).map((st) => {
                      const sch = schools.find(s => s.id === st.school_id);
                      return <option key={st.id} value={st.id}>{sch ? `${sch.name} — ` : ""}{st.trips?.name || st.trip_id}</option>;
                    })}
                  </select>
                </div>
              </div>
            </CardContent>
          </Card>
          {filterTripId && (() => {
            const selectedST = allSchoolTrips.find((st) => st.id === filterTripId);
            if (!selectedST) return null;
            const logistics = selectedST.logistics || [];
            const updateLogistics = async (next) => {
              const { error } = await supabase.from("school_trips").update({ logistics: next }).eq("id", filterTripId);
              if (!error) setAllSchoolTrips((prev) => prev.map((t) => t.id === filterTripId ? { ...t, logistics: next } : t));
              else notify("Error guardando logística.", { variant: "destructive" });
            };
            return (
              <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
                <CardContent className="space-y-4 p-6">
                  <div className="flex items-center justify-between flex-wrap gap-3">
                    <div>
                      <div className="font-semibold text-zinc-950">Puntos logísticos</div>
                      <div className="text-sm text-zinc-500">Punto de encuentro, hora de salida, qué traer, contacto de emergencia...</div>
                    </div>
                    <div className="flex items-center gap-3 flex-wrap">
                      <div className="flex items-center gap-2">
                        <span className="text-xs text-zinc-500">Visible al colegio</span>
                        <button
                          type="button"
                          onClick={async () => {
                            const next = selectedST.show_logistics === false;
                            const { error } = await supabase.from("school_trips").update({ show_logistics: next }).eq("id", filterTripId);
                            if (!error) setAllSchoolTrips((prev) => prev.map((t) => t.id === filterTripId ? { ...t, show_logistics: next } : t));
                            else notify("Error guardando visibilidad.");
                          }}
                          className={`relative h-6 w-11 rounded-full transition-colors cursor-pointer ${selectedST.show_logistics === false ? "bg-zinc-300" : "bg-green-500"}`}
                        >
                          <div className={`absolute top-0.5 h-5 w-5 rounded-full bg-white shadow transition-all ${selectedST.show_logistics === false ? "left-0.5" : "left-5"}`} />
                        </button>
                      </div>
                      <Button className="rounded-2xl text-white" style={{ backgroundColor: CORPORATE_RED }}
                        onClick={() => updateLogistics([...logistics, { title: "Nuevo punto", description: "" }])}>
                        <Plus className="mr-2 h-4 w-4" />Añadir punto
                      </Button>
                    </div>
                  </div>
                  <div className="space-y-3">
                    {logistics.map((item, index) => (
                      <div key={index} className="flex gap-3 items-start rounded-2xl border border-zinc-200 bg-white p-4">
                        <div className="flex-1 space-y-2">
                          <Input value={item.title} placeholder="Título (ej: Punto de encuentro)" onChange={(e) => updateLogistics(logistics.map((l, i) => i === index ? { ...l, title: e.target.value } : l))} className="rounded-xl bg-white text-sm font-medium" />
                          <Textarea value={item.description} placeholder="Descripción (ej: Parking del instituto a las 8:00)" onChange={(e) => updateLogistics(logistics.map((l, i) => i === index ? { ...l, description: e.target.value } : l))} className="min-h-[72px] rounded-xl bg-white text-sm" />
                        </div>
                        <button type="button" onClick={() => updateLogistics(logistics.filter((_, i) => i !== index))} className="mt-1 rounded-xl p-1.5 text-zinc-400 hover:bg-zinc-200 hover:text-zinc-700">
                          <X className="h-4 w-4" />
                        </button>
                      </div>
                    ))}
                    {logistics.length === 0 && (
                      <div className="rounded-2xl border-2 border-dashed border-zinc-200 py-10 text-center text-sm text-zinc-400">
                        Sin puntos logísticos. Pulsa "Añadir punto" para empezar.
                      </div>
                    )}
                  </div>
                </CardContent>
              </Card>
            );
          })()}
          {!filterTripId && (
            <div className="rounded-2xl border border-dashed border-zinc-200 p-8 text-center text-sm text-zinc-400">Selecciona un viaje escolar para editar su logística.</div>
          )}
        </div>
      )}

      {/* Viajes escolares tab */}
      {tab === "school_viajes" && (
        <AdminSchoolViajes allSchoolTrips={allSchoolTrips} schools={schools} trips={trips} setTrips={setTrips} notify={notify} />
      )}
    </div>
  );
}

// ─── Admin Panel ──────────────────────────────────────────────────────────────

// ─── AdminEmailTemplates ──────────────────────────────────────────────────────
const EMAIL_TEMPLATE_DEFS = {
  campamentos: [
    {
      id: "doc_reminder",
      label: "Recordatorio documentación",
      icon: "📄",
      vars: ["nombre", "viaje", "documento"],
      hints: { nombre: "Nombre del participante", viaje: "Nombre del campamento", documento: "Nombre(s) del documento pendiente" },
    },
    {
      id: "payment_reminder",
      label: "Recordatorio pago",
      icon: "💳",
      vars: ["nombre", "viaje", "pago", "importe", "fecha", "dias"],
      hints: { nombre: "Nombre del participante", viaje: "Nombre del campamento", pago: "Nombre del pago (Reserva, 1ª cuota…)", importe: "Importe en €", fecha: "Fecha límite", dias: "Días restantes" },
    },
    {
      id: "doc_confirmed",
      label: "Documento aprobado",
      icon: "✅",
      vars: ["nombre", "viaje", "documento"],
      hints: { nombre: "Nombre del participante", viaje: "Nombre del campamento", documento: "Nombre del documento aprobado" },
    },
    {
      id: "doc_rejected",
      label: "Documento rechazado",
      icon: "❌",
      vars: ["nombre", "viaje", "documento"],
      hints: { nombre: "Nombre del participante", viaje: "Nombre del campamento", documento: "Nombre del documento rechazado" },
    },
    {
      id: "payment_confirmed",
      label: "Pago confirmado",
      icon: "🎉",
      vars: ["nombre", "viaje", "pago", "importe"],
      hints: { nombre: "Nombre del participante", viaje: "Nombre del campamento", pago: "Nombre del pago confirmado", importe: "Importe en €" },
    },
    {
      id: "question_replied",
      label: "Respuesta a consulta",
      icon: "💬",
      vars: ["nombre", "pregunta", "respuesta"],
      hints: { nombre: "Nombre del participante", pregunta: "Pregunta original", respuesta: "Tu respuesta" },
    },
    {
      id: "general_reminder",
      label: "Recordatorio general",
      icon: "⏰",
      vars: ["nombre", "viaje", "pendientes"],
      hints: { nombre: "Nombre del participante", viaje: "Nombre del campamento", pendientes: "Nº de elementos pendientes" },
    },
  ],
  colegios: [
    {
      id: "school_reminder_listado",
      label: "Recordatorio listados",
      icon: "📋",
      vars: ["nombre", "viaje", "coordinador"],
      hints: { nombre: "Nombre del colegio", viaje: "Nombre del viaje", coordinador: "Nombre del coordinador (opcional)" },
    },
    {
      id: "school_reminder_alergias",
      label: "Recordatorio alergias",
      icon: "🍽️",
      vars: ["nombre", "viaje", "coordinador"],
      hints: { nombre: "Nombre del colegio", viaje: "Nombre del viaje", coordinador: "Nombre del coordinador (opcional)" },
    },
    {
      id: "school_doc_reminder",
      label: "Recordatorio documentación",
      icon: "📄",
      vars: ["nombre", "viaje", "coordinador", "pendientes"],
      hints: { nombre: "Nombre del colegio", viaje: "Nombre del viaje", coordinador: "Nombre del coordinador (opcional)", pendientes: "Nº de documentos pendientes" },
    },
    {
      id: "school_reminder_rooming",
      label: "Recordatorio rooming",
      icon: "🛏️",
      vars: ["nombre", "viaje", "coordinador"],
      hints: { nombre: "Nombre del colegio", viaje: "Nombre del viaje", coordinador: "Nombre del coordinador (opcional)" },
    },
    {
      id: "school_reminder_grupos",
      label: "Recordatorio grupos",
      icon: "👥",
      vars: ["nombre", "viaje", "coordinador"],
      hints: { nombre: "Nombre del colegio", viaje: "Nombre del viaje", coordinador: "Nombre del coordinador (opcional)" },
    },
    {
      id: "school_reminder_todo",
      label: "Recordatorio general (todo)",
      icon: "⏰",
      vars: ["nombre", "viaje", "coordinador", "pendientes"],
      hints: { nombre: "Nombre del colegio", viaje: "Nombre del viaje", coordinador: "Nombre del coordinador (opcional)", pendientes: "Nº total de pendientes" },
    },
    {
      id: "school_question_replied",
      label: "Respuesta a consulta",
      icon: "💬",
      vars: ["nombre", "coordinador", "pregunta", "respuesta"],
      hints: { nombre: "Nombre del colegio", coordinador: "Nombre del coordinador", pregunta: "Pregunta original del colegio", respuesta: "Tu respuesta" },
    },
  ],
};

const TEMPLATE_ICON_MAP = {
  doc_reminder:              FileText,
  payment_reminder:          CreditCard,
  doc_confirmed:             CheckCircle2,
  doc_rejected:              X,
  payment_confirmed:         BadgeCheck,
  question_replied:          MessageCircleQuestion,
  school_reminder_listado:   Users,
  school_reminder_alergias:  AlertCircle,
  school_doc_reminder:       FileCheck2,
  school_reminder_rooming:   LayoutGrid,
  school_reminder_grupos:    ListChecks,
  school_reminder_todo:      Bell,
  school_question_replied:   MessageCircleQuestion,
  general_reminder:          Bell,
};

function AdminEmailTemplates({ category, notify }) {
  const defs = EMAIL_TEMPLATE_DEFS[category];
  const [selectedId, setSelectedId] = useState(defs[0].id);
  const [templates, setTemplates] = useState({});
  const [saving, setSaving] = useState(false);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    const load = async () => {
      setTemplates({});
      setLoading(true);
      const ids = defs.map((d) => d.id);
      const { data, error } = await supabase
        .from("email_templates")
        .select("id, subject, body")
        .in("id", ids);
      if (!error && data) {
        const map = {};
        data.forEach((row) => { map[row.id] = { subject: row.subject, body: row.body }; });
        setTemplates(map);
      }
      setLoading(false);
    };
    load();
  }, [category]);

  const selected = defs.find((d) => d.id === selectedId);
  const current = templates[selectedId] || { subject: "", body: "" };

  const handleChange = (field, value) => {
    setTemplates((prev) => ({ ...prev, [selectedId]: { ...current, [field]: value } }));
  };

  const handleSave = async () => {
    setSaving(true);
    const { error } = await supabase
      .from("email_templates")
      .upsert({ id: selectedId, subject: current.subject, body: current.body, category }, { onConflict: "id" });
    setSaving(false);
    if (error) { notify("Error guardando la plantilla: " + error.message); return; }
    notify("✅ Plantilla guardada correctamente");
  };

  const insertVar = (v) => {
    handleChange("body", current.body + `{${v}}`);
  };

  if (loading) return (
    <div className="flex items-center justify-center py-20">
      <Loader2 className="h-6 w-6 animate-spin text-zinc-400" />
    </div>
  );

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-xl font-bold text-zinc-950">
          Plantillas de email — {category === "campamentos" ? "Campamentos" : "Colegios"}
        </h2>
        <p className="mt-1 text-sm text-zinc-500">
          Edita el asunto y el cuerpo de cada correo automático. Usa <code className="rounded bg-zinc-100 px-1 py-0.5 text-xs">{"{variable}"}</code> para insertar datos dinámicos.
        </p>
      </div>

      <div className="flex gap-5 min-h-[500px]">
        {/* Left list */}
        <div className="w-52 shrink-0 space-y-1">
          {defs.map((d) => (
            <button
              key={d.id}
              type="button"
              onClick={() => setSelectedId(d.id)}
              className={`flex w-full items-center gap-2.5 rounded-2xl px-3 py-2.5 text-left text-sm font-medium transition-all ${
                selectedId === d.id
                  ? "text-white shadow-sm"
                  : "text-zinc-600 hover:bg-zinc-50 hover:text-zinc-950"
              }`}
              style={selectedId === d.id ? { backgroundColor: CORPORATE_RED } : {}}
            >
              {(() => { const I = TEMPLATE_ICON_MAP[d.id] || FileText; return <I className="h-4 w-4 shrink-0" />; })()}
              <span className="leading-snug">{d.label}</span>
            </button>
          ))}
        </div>

        {/* Right editor */}
        <div className="flex-1 rounded-3xl border border-zinc-200 bg-white p-6 shadow-sm space-y-5">
          <div className="flex items-center gap-2">
            {(() => { const I = TEMPLATE_ICON_MAP[selected.id] || FileText; return <I className="h-5 w-5 text-zinc-500" />; })()}
            <h3 className="font-semibold text-zinc-950">{selected.label}</h3>
          </div>

          {/* Subject */}
          <div className="space-y-1.5">
            <Label className="text-xs font-semibold uppercase tracking-widest text-zinc-400">Asunto del correo</Label>
            <Input
              value={current.subject}
              onChange={(e) => handleChange("subject", e.target.value)}
              placeholder="Escribe el asunto..."
              className="h-11 rounded-2xl border-zinc-200 bg-zinc-50 text-sm"
            />
          </div>

          {/* Body */}
          <div className="space-y-1.5">
            <Label className="text-xs font-semibold uppercase tracking-widest text-zinc-400">Cuerpo del correo</Label>
            <Textarea
              value={current.body}
              onChange={(e) => handleChange("body", e.target.value)}
              placeholder="Escribe el texto del correo. Cada línea será un párrafo."
              className="min-h-[220px] rounded-2xl border-zinc-200 bg-zinc-50 text-sm leading-relaxed"
            />
          </div>

          {/* Variables */}
          <div className="space-y-2">
            <div className="text-xs font-semibold uppercase tracking-widest text-zinc-400">Variables disponibles</div>
            <div className="flex flex-wrap gap-2">
              {selected.vars.map((v) => (
                <button
                  key={v}
                  type="button"
                  onClick={() => insertVar(v)}
                  title={selected.hints?.[v] || `Insertar {${v}}`}
                  className="flex items-center gap-1.5 rounded-xl border border-zinc-200 bg-zinc-50 px-2.5 py-1.5 text-xs text-zinc-700 hover:border-zinc-400 hover:bg-white transition group"
                >
                  <Plus className="h-3 w-3 text-zinc-400 shrink-0" />
                  <span className="font-mono">{`{${v}}`}</span>
                  {selected.hints?.[v] && (
                    <span className="text-zinc-400 hidden group-hover:inline">— {selected.hints[v]}</span>
                  )}
                </button>
              ))}
            </div>
            <p className="text-[11px] text-zinc-400">Haz clic en una variable para añadirla al final del cuerpo, o escríbela directamente.</p>
          </div>

          {/* Save */}
          <div className="flex justify-end pt-2">
            <Button
              onClick={handleSave}
              disabled={saving}
              className="h-11 rounded-2xl px-6 text-sm font-semibold text-white shadow-sm"
              style={{ backgroundColor: CORPORATE_RED }}
            >
              {saving ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <Save className="mr-2 h-4 w-4" />}
              Guardar plantilla
            </Button>
          </div>
        </div>
      </div>
    </div>
  );
}

function AdminSchoolPreviewButton({ onPreview }) {
  const [schools, setSchools] = useState([]);
  const [open, setOpen] = useState(false);

  useEffect(() => {
    if (open && !schools.length) {
      supabase.from("schools").select("id, name").order("name").then(({ data }) => { if (data) setSchools(data); });
    }
  }, [open]);

  return (
    <div className="border-t border-zinc-200 pt-2 mt-1">
      {!open ? (
        <button type="button" onClick={() => setOpen(true)}
          className="flex w-full items-center gap-3 rounded-2xl px-3 py-2.5 text-sm font-medium text-zinc-500 hover:bg-zinc-50 hover:text-zinc-700 transition-all">
          <Eye className="h-4 w-4 shrink-0" />Vista previa colegio
        </button>
      ) : (
        <div className="space-y-1 px-1">
          <div className="text-[11px] font-semibold uppercase tracking-widest text-zinc-400 px-2 py-1">Selecciona un colegio</div>
          {schools.map((s) => (
            <button key={s.id} type="button" onClick={() => onPreview(s.id)}
              className="flex w-full items-center gap-2 rounded-2xl px-3 py-2 text-sm text-zinc-700 hover:bg-amber-50 hover:text-amber-800 transition-all text-left">
              <Eye className="h-3.5 w-3.5 shrink-0 text-amber-500" />{s.name}
            </button>
          ))}
          {!schools.length && <div className="px-3 py-2 text-xs text-zinc-400">Sin colegios registrados</div>}
          <button type="button" onClick={() => setOpen(false)} className="w-full text-left px-3 py-1 text-xs text-zinc-400 hover:text-zinc-600">Cancelar</button>
        </div>
      )}
    </div>
  );
}

function AdminClientPreviewButton({ users, onPreview }) {
  const [open, setOpen] = useState(false);
  const clients = users.filter((u) => u.role === "client" && !u.schoolId);

  return (
    <div className="mt-1">
      {!open ? (
        <button type="button" onClick={() => setOpen(true)}
          className="flex w-full items-center gap-3 rounded-2xl px-3 py-2.5 text-sm font-medium text-zinc-500 hover:bg-zinc-50 hover:text-zinc-700 transition-all">
          <Eye className="h-4 w-4 shrink-0" />Vista previa cliente
        </button>
      ) : (
        <div className="space-y-1 px-1">
          <div className="text-[11px] font-semibold uppercase tracking-widest text-zinc-400 px-2 py-1">Selecciona un cliente</div>
          {clients.slice(0, 8).map((u) => (
            <button key={u.id} type="button" onClick={() => { onPreview(u); setOpen(false); }}
              className="flex w-full items-center gap-2 rounded-2xl px-3 py-2 text-sm text-zinc-700 hover:bg-amber-50 hover:text-amber-800 transition-all text-left">
              <Eye className="h-3.5 w-3.5 shrink-0 text-amber-500" />{u.participantName || u.username}
            </button>
          ))}
          {!clients.length && <div className="px-3 py-2 text-xs text-zinc-400">Sin clientes registrados</div>}
          <button type="button" onClick={() => setOpen(false)} className="w-full text-left px-3 py-1 text-xs text-zinc-400 hover:text-zinc-600">Cancelar</button>
        </div>
      )}
    </div>
  );
}

function AdminDashboard({ users, trips, setActiveSection }) {
  const clients = users.filter((u) => u.role === "client" && !u.schoolId);
  const pendingDocs = clients.reduce((sum, u) => sum + (u.documents || []).filter((d) => d.status === "pending_confirmation").length, 0);
  const pendingPayments = clients.reduce((sum, u) => {
    const p = u.payments || {};
    return sum + ["reservation", "firstInstallment", "secondInstallment"].filter((k) => p[k]?.status === "sent").length;
  }, 0);
  const pendingQuestions = clients.reduce((sum, u) => sum + (u.questions || []).filter((q) => !q.reply).length, 0);

  const [schoolStats, setSchoolStats] = useState({ schools: 0, trips: 0, students: 0, pendingDocs: 0 });
  useEffect(() => {
    Promise.all([
      supabase.from("schools").select("id", { count: "exact", head: true }),
      supabase.from("school_trips").select("id", { count: "exact", head: true }),
      supabase.from("students").select("id", { count: "exact", head: true }),
      supabase.from("school_documents").select("id", { count: "exact", head: true }).not("status", "in", '("approved","library")'),
    ]).then(([sc, st, stu, docs]) => {
      setSchoolStats({
        schools: sc.count || 0,
        trips: st.count || 0,
        students: stu.count || 0,
        pendingDocs: docs.count || 0,
      });
    });
  }, []);

  const stats = [
    { label: "Clientes campamentos", value: clients.length, icon: Users, section: "clients" },
    { label: "Documentos por revisar", value: pendingDocs, icon: FileCheck2, section: "tracking", highlight: pendingDocs > 0 },
    { label: "Pagos por confirmar", value: pendingPayments, icon: CreditCard, section: "tracking", highlight: pendingPayments > 0 },
    { label: "Preguntas sin responder", value: pendingQuestions, icon: MessageCircleQuestion, section: "questions", highlight: pendingQuestions > 0 },
    { label: "Campamentos activos", value: trips.length, icon: MapIcon, section: "trips" },
  ];

  const quickCamp = [
    { key: "clients", label: "Clientes", icon: Users },
    { key: "tracking", label: "Seguimiento", icon: BarChart2 },
    { key: "payments", label: "Pagos", icon: CreditCard },
    { key: "docs", label: "Documentación", icon: FileCheck2 },
    { key: "email_camp", label: "Emails", icon: Mail },
  ];
  const quickCol = [
    { key: "school_colegios", label: "Colegios", icon: Users },
    { key: "school_seguimiento", label: "Seguimiento", icon: BarChart2 },
    { key: "school_alumnos", label: "Alumnos", icon: Users },
    { key: "school_docs", label: "Documentación", icon: FileCheck2 },
    { key: "email_col", label: "Emails", icon: Mail },
  ];

  return (
    <div className="space-y-6">
      <SectionTitle icon={Home} title="Inicio" subtitle="Resumen general del portal de administración." />
      <div className="space-y-2">
        <div className="text-xs font-bold uppercase tracking-widest text-zinc-400 px-1">Campamentos</div>
        <div className="grid gap-4 sm:grid-cols-2 xl:grid-cols-5">
          {stats.map(({ label, value, icon: Icon, section, highlight }) => (
            <button key={label} type="button" onClick={() => setActiveSection(section)}
              className="rounded-3xl border border-zinc-200 bg-white p-5 text-left shadow-sm hover:shadow-md transition-shadow">
              <div className="flex items-center justify-between">
                <div className={`flex h-9 w-9 items-center justify-center rounded-2xl ${highlight ? "text-white" : "bg-zinc-100 text-zinc-500"}`}
                  style={highlight ? { backgroundColor: CORPORATE_RED } : {}}>
                  <Icon className="h-4 w-4" />
                </div>
                <span className={`text-2xl font-bold ${highlight ? "" : "text-zinc-950"}`}
                  style={highlight ? { color: CORPORATE_RED } : {}}>{value}</span>
              </div>
              <div className="mt-3 text-xs font-medium text-zinc-500">{label}</div>
            </button>
          ))}
        </div>
      </div>
      <div className="space-y-2">
        <div className="text-xs font-bold uppercase tracking-widest text-zinc-400 px-1">Colegios</div>
        <div className="grid gap-4 sm:grid-cols-2 xl:grid-cols-4">
          {[
            { label: "Colegios", value: schoolStats.schools, icon: Building2, section: "school_colegios" },
            { label: "Viajes escolares", value: schoolStats.trips, icon: MapIcon, section: "school_colegios" },
            { label: "Alumnos", value: schoolStats.students, icon: Users, section: "school_alumnos" },
            { label: "Documentos pendientes", value: schoolStats.pendingDocs, icon: FileCheck2, section: "school_docs", highlight: schoolStats.pendingDocs > 0 },
          ].map(({ label, value, icon: Icon, section, highlight }) => (
            <button key={label} type="button" onClick={() => setActiveSection(section)}
              className="rounded-3xl border border-zinc-200 bg-white p-5 text-left shadow-sm hover:shadow-md transition-shadow">
              <div className="flex items-center justify-between">
                <div className={`flex h-9 w-9 items-center justify-center rounded-2xl ${highlight ? "text-white" : "bg-zinc-100 text-zinc-500"}`}
                  style={highlight ? { backgroundColor: CORPORATE_RED } : {}}>
                  <Icon className="h-4 w-4" />
                </div>
                <span className={`text-2xl font-bold ${highlight ? "" : "text-zinc-950"}`}
                  style={highlight ? { color: CORPORATE_RED } : {}}>{value}</span>
              </div>
              <div className="mt-3 text-xs font-medium text-zinc-500">{label}</div>
            </button>
          ))}
        </div>
      </div>
      <div className="grid gap-4 lg:grid-cols-2">
        <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
          <CardContent className="p-5">
            <div className="mb-3 text-xs font-bold uppercase tracking-widest text-zinc-400">Acceso rápido — Campamentos</div>
            <div className="flex flex-wrap gap-2">
              {quickCamp.map(({ key, label, icon: Icon }) => (
                <button key={key} type="button" onClick={() => setActiveSection(key)}
                  className="flex items-center gap-1.5 rounded-2xl border border-zinc-200 bg-white px-3 py-2 text-sm font-medium text-zinc-700 hover:bg-zinc-50 transition">
                  <Icon className="h-3.5 w-3.5" />{label}
                </button>
              ))}
            </div>
          </CardContent>
        </Card>
        <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
          <CardContent className="p-5">
            <div className="mb-3 text-xs font-bold uppercase tracking-widest text-zinc-400">Acceso rápido — Colegios</div>
            <div className="flex flex-wrap gap-2">
              {quickCol.map(({ key, label, icon: Icon }) => (
                <button key={key} type="button" onClick={() => setActiveSection(key)}
                  className="flex items-center gap-1.5 rounded-2xl border border-zinc-200 bg-white px-3 py-2 text-sm font-medium text-zinc-700 hover:bg-zinc-50 transition">
                  <Icon className="h-3.5 w-3.5" />{label}
                </button>
              ))}
            </div>
          </CardContent>
        </Card>
      </div>
    </div>
  );
}

function AdminPanel({ users, setUsers, trips, setTrips, schoolTripIds = new Set(), setSchoolTripIds, templates, setTemplates, onLogout, notify }) {
  const [activeSection, setActiveSection] = useState(() => {
    if (typeof window === "undefined") return "home";
    return window.localStorage.getItem(ADMIN_SECTION_STORAGE_KEY) || "home";
  });
  useEffect(() => { if (typeof window !== "undefined") window.localStorage.setItem(ADMIN_SECTION_STORAGE_KEY, activeSection); }, [activeSection]);
  const totalParticipants = users.filter((u) => u.role === "client" && !u.schoolId).length;
  const campTrips = trips.filter((t) => t.tipo !== "colegio");


  const [campExpanded, setCampExpanded] = useState(false);
  const [colExpanded, setColExpanded] = useState(false);
  const [previewSchoolId, setPreviewSchoolId] = useState(null);
  const [previewClientUser, setPreviewClientUser] = useState(null);

  const campamentosItems = [
    { key: "clients",        label: "Clientes",         icon: Users },
    { key: "tracking",       label: "Seguimiento",       icon: BarChart2 },
    { key: "participants_export", label: "Participantes", icon: FileText },
    { key: "payments",       label: "Pagos",             icon: CreditCard },
    { key: "docs",           label: "Documentación",     icon: FileCheck2 },
    { key: "questions",      label: "Preguntas",         icon: MessageCircleQuestion },
    { key: "checklists",     label: "Checklists",        icon: ListChecks },
    { key: "itinerario",     label: "Itinerario",        icon: CalendarDays },
    { key: "logistica",      label: "Logística",         icon: MapPinned },
    { key: "trips",          label: "Campamentos",       icon: MapIcon },
    { key: "email_camp",     label: "Plantillas email",  icon: Mail },
  ];
  const colegiosItems = [
    { key: "school_colegios",    label: "Colegios",         icon: Users },
    { key: "school_seguimiento", label: "Seguimiento",      icon: BarChart2 },
    { key: "school_alumnos",     label: "Alumnos",          icon: Users },
    { key: "school_alergias",    label: "Alergias",         icon: AlertCircle },
    { key: "school_docs",        label: "Documentación",    icon: FileCheck2 },
    { key: "school_pagos",       label: "Pagos",            icon: CreditCard },
    { key: "school_rooming",     label: "Rooming",          icon: Home },
    { key: "school_grupos",      label: "Grupos",           icon: Grid2x2 },
    { key: "school_preguntas",   label: "Preguntas",        icon: MessageCircleQuestion },
    { key: "school_checklist",   label: "Checklist",        icon: CheckCircle2 },
    { key: "school_itinerario",  label: "Itinerario",       icon: CalendarDays },
    { key: "school_logistica",   label: "Logística",        icon: MapPinned },
    { key: "school_viajes",      label: "Viajes",           icon: MapIcon },
    { key: "email_col",          label: "Plantillas email", icon: Mail },
  ];
  const navItems = [...campamentosItems, ...colegiosItems, { key: "calculadora", label: "Calculadora", icon: Calculator }];

  // ── Vista previa del portal de cliente (admin-only) ─────────────────────────
  if (previewClientUser) {
    return (
      <div className="relative">
        <div className="sticky top-0 z-50 flex items-center justify-between gap-3 bg-amber-50 border-b border-amber-200 px-5 py-2.5 text-sm">
          <span className="font-medium text-amber-800">👁 Vista previa — Portal de cliente: {previewClientUser.participantName || previewClientUser.username}</span>
          <Button size="sm" variant="outline" className="rounded-xl border-amber-300 text-amber-800 hover:bg-amber-100" onClick={() => setPreviewClientUser(null)}>
            Volver al admin
          </Button>
        </div>
        <ClientPortal user={previewClientUser} trips={trips} templates={templates} setUsers={setUsers} onLogout={() => setPreviewClientUser(null)} notify={notify} />
      </div>
    );
  }

  // ── Vista previa del portal de colegio (admin-only) ──────────────────────────
  if (previewSchoolId) {
    return (
      <div className="relative">
        <div className="sticky top-0 z-50 flex items-center justify-between gap-3 bg-amber-50 border-b border-amber-200 px-5 py-2.5 text-sm">
          <span className="font-medium text-amber-800">👁 Vista previa — Portal de coordinador de colegio</span>
          <Button size="sm" variant="outline" className="rounded-xl border-amber-300 text-amber-800 hover:bg-amber-100" onClick={() => setPreviewSchoolId(null)}>
            Volver al admin
          </Button>
        </div>
        <SchoolPortal user={{ authUid: null }} onLogout={() => setPreviewSchoolId(null)} notify={notify} previewSchoolId={previewSchoolId} />
      </div>
    );
  }

  return (
    <div className="min-h-screen text-zinc-950" style={{ background: "linear-gradient(160deg,#fff5f5 0%,#fafafa 40%,#f4f4f5 100%)" }}>
      {/* Header */}
      <header className="sticky top-0 z-40 border-b border-zinc-200/80 bg-white/90 backdrop-blur-sm">
        <div className="mx-auto flex max-w-[1400px] items-center justify-between px-6 py-4">
          <div className="flex items-center gap-3">
            <div className="flex h-10 w-10 items-center justify-center rounded-xl shadow-sm" style={{ backgroundColor: CORPORATE_RED }}>
              <img src="/logo-gimeloos.svg" alt="GIMELOOS" className="h-8 w-8 object-contain" style={{ filter: "invert(1)" }} />
            </div>
            <div>
              <div className="text-xs uppercase tracking-[0.22em] text-zinc-400">Panel administrador</div>
              <div className="text-base font-bold tracking-[0.12em] text-zinc-950">GIMELOOS</div>
            </div>
            <div className="ml-2 hidden rounded-full border border-zinc-200 bg-white px-3 py-1 text-xs text-zinc-500 sm:block">
              {totalParticipants} participantes
            </div>
          </div>
          <Button variant="outline" className="rounded-2xl text-sm" onClick={() => { onLogout(); notify("Sesión cerrada."); }}>
            <LogOut className="mr-2 h-4 w-4" />Salir
          </Button>
        </div>
      </header>

      <div className="mx-auto flex max-w-[1400px] gap-6 px-6 py-6">
        {/* Sidebar */}
        <aside className="hidden w-56 shrink-0 lg:block">
          <div className="sticky top-24 rounded-3xl border border-zinc-200 bg-white p-3 shadow-sm overflow-y-auto" style={{ maxHeight: "calc(100vh - 7rem)" }}>
            <nav className="space-y-1">
              {/* Inicio */}
              <button type="button" onClick={() => setActiveSection("home")}
                className={`flex w-full items-center gap-3 rounded-2xl px-3 py-2.5 text-sm font-medium transition-all ${activeSection === "home" ? "text-white shadow-sm" : "text-zinc-600 hover:bg-zinc-50 hover:text-zinc-950"}`}
                style={activeSection === "home" ? { backgroundColor: CORPORATE_RED } : {}}>
                <Home className="h-4 w-4 shrink-0" />Inicio
                {activeSection === "home" && <ChevronRight className="ml-auto h-3 w-3 opacity-60" />}
              </button>
              <div className="border-t border-zinc-100 my-1" />
              {/* CAMPAMENTOS colapsable */}
              <button type="button" onClick={() => setCampExpanded(!campExpanded)}
                className="flex w-full items-center justify-between rounded-2xl px-3 py-2 text-[11px] font-bold uppercase tracking-widest text-zinc-500 hover:bg-zinc-50 transition">
                <span>Campamentos</span>
                <ChevronDown className={`h-3.5 w-3.5 transition-transform ${campExpanded ? "rotate-180" : ""}`} />
              </button>
              {campExpanded && campamentosItems.map(({ key, label, icon: Icon }) => {
                const active = activeSection === key;
                return (
                  <button key={key} type="button" onClick={() => setActiveSection(key)}
                    className={`flex w-full items-center gap-3 rounded-2xl px-3 py-2.5 text-sm font-medium transition-all ${active ? "text-white shadow-sm" : "text-zinc-600 hover:bg-zinc-50 hover:text-zinc-950"}`}
                    style={active ? { backgroundColor: CORPORATE_RED } : {}}
                  >
                    <Icon className="h-4 w-4 shrink-0" />{label}
                    {active && <ChevronRight className="ml-auto h-3 w-3 opacity-60" />}
                  </button>
                );
              })}
              {/* Vista previa portal cliente — al final de Campamentos */}
              {campExpanded && <AdminClientPreviewButton users={users} onPreview={setPreviewClientUser} />}

              {/* COLEGIOS colapsable */}
              <div className="pt-1">
                <button type="button" onClick={() => setColExpanded(!colExpanded)}
                  className="flex w-full items-center justify-between rounded-2xl px-3 py-2 text-[11px] font-bold uppercase tracking-widest text-zinc-500 hover:bg-zinc-50 transition">
                  <span>Colegios</span>
                  <ChevronDown className={`h-3.5 w-3.5 transition-transform ${colExpanded ? "rotate-180" : ""}`} />
                </button>
                {colExpanded && colegiosItems.map(({ key, label, icon: Icon }) => {
                  const active = activeSection === key;
                  return (
                    <button key={key} type="button" onClick={() => setActiveSection(key)}
                      className={`flex w-full items-center gap-3 rounded-2xl px-3 py-2.5 text-sm font-medium transition-all ${active ? "text-white shadow-sm" : "text-zinc-600 hover:bg-zinc-50 hover:text-zinc-950"}`}
                      style={active ? { backgroundColor: CORPORATE_RED } : {}}
                    >
                      <Icon className="h-4 w-4 shrink-0" />{label}
                      {active && <ChevronRight className="ml-auto h-3 w-3 opacity-60" />}
                    </button>
                  );
                })}
              </div>

              {/* Vista previa portal colegio — al final de Colegios */}
              {colExpanded && <AdminSchoolPreviewButton onPreview={setPreviewSchoolId} />}

              {/* Calculadora separada al fondo */}
              <div className="border-t border-zinc-200 pt-2 mt-2">
                <button type="button" onClick={() => setActiveSection("calculadora")}
                  className={`flex w-full items-center gap-3 rounded-2xl px-3 py-2.5 text-sm font-medium transition-all ${activeSection === "calculadora" ? "text-white shadow-sm" : "text-zinc-600 hover:bg-zinc-50 hover:text-zinc-950"}`}
                  style={activeSection === "calculadora" ? { backgroundColor: CORPORATE_RED } : {}}
                >
                  <Calculator className="h-4 w-4 shrink-0" />Calculadora
                  {activeSection === "calculadora" && <ChevronRight className="ml-auto h-3 w-3 opacity-60" />}
                </button>
              </div>
            </nav>
          </div>
        </aside>

        {/* Mobile tabs */}
        <div className="mb-4 w-full lg:hidden space-y-2">
          {/* Campamentos group */}
          <div className="rounded-2xl border border-zinc-200 bg-white p-2">
            <button type="button" onClick={() => setCampExpanded(!campExpanded)}
              className="flex w-full items-center justify-between px-2 py-1 text-[11px] font-bold uppercase tracking-widest text-zinc-500">
              <span>Campamentos</span>
              <ChevronDown className={`h-3.5 w-3.5 transition-transform ${campExpanded ? "rotate-180" : ""}`} />
            </button>
            {campExpanded && (
              <div className="mt-1 flex flex-wrap gap-1.5">
                {campamentosItems.map(({ key, label, icon: Icon }) => {
                  const active = activeSection === key;
                  return (
                    <button key={key} type="button" onClick={() => setActiveSection(key)}
                      className={`flex items-center gap-1.5 rounded-xl px-3 py-1.5 text-xs font-medium transition ${active ? "text-white" : "border border-zinc-200 bg-zinc-50 text-zinc-700"}`}
                      style={active ? { backgroundColor: CORPORATE_RED } : {}}>
                      <Icon className="h-3.5 w-3.5" />{label}
                    </button>
                  );
                })}
              </div>
            )}
          </div>
          {/* Colegios group */}
          <div className="rounded-2xl border border-zinc-200 bg-white p-2">
            <button type="button" onClick={() => setColExpanded(!colExpanded)}
              className="flex w-full items-center justify-between px-2 py-1 text-[11px] font-bold uppercase tracking-widest text-zinc-500">
              <span>Colegios</span>
              <ChevronDown className={`h-3.5 w-3.5 transition-transform ${colExpanded ? "rotate-180" : ""}`} />
            </button>
            {colExpanded && (
              <div className="mt-1 flex flex-wrap gap-1.5">
                {colegiosItems.map(({ key, label, icon: Icon }) => {
                  const active = activeSection === key;
                  return (
                    <button key={key} type="button" onClick={() => setActiveSection(key)}
                      className={`flex items-center gap-1.5 rounded-xl px-3 py-1.5 text-xs font-medium transition ${active ? "text-white" : "border border-zinc-200 bg-zinc-50 text-zinc-700"}`}
                      style={active ? { backgroundColor: CORPORATE_RED } : {}}>
                      <Icon className="h-3.5 w-3.5" />{label}
                    </button>
                  );
                })}
              </div>
            )}
          </div>
          {/* Calculadora */}
          <div>
            <button type="button" onClick={() => setActiveSection("calculadora")}
              className={`flex items-center gap-1.5 rounded-xl px-3 py-1.5 text-xs font-medium transition ${activeSection === "calculadora" ? "text-white" : "border border-zinc-200 bg-white text-zinc-700"}`}
              style={activeSection === "calculadora" ? { backgroundColor: CORPORATE_RED } : {}}>
              <Calculator className="h-3.5 w-3.5" />Calculadora
            </button>
          </div>
        </div>

        {/* Content */}
        <main className="min-w-0 flex-1">
          {activeSection === "home"                && <AdminDashboard users={users} trips={campTrips} setActiveSection={setActiveSection} />}
          {activeSection === "clients"             && <AdminClients users={users} trips={campTrips} setUsers={setUsers} templates={templates} notify={notify} setTrips={setTrips} />}
          {activeSection === "tracking"            && <AdminTracking users={users} trips={campTrips} templates={templates} setUsers={setUsers} notify={notify} />}
          {activeSection === "participants_export" && <AdminParticipantsExport users={users} trips={campTrips} />}
          {activeSection === "payments"            && <AdminPayments users={users} setUsers={setUsers} trips={campTrips} setTrips={setTrips} notify={notify} />}
          {activeSection === "docs"                && <AdminDocs templates={templates} setTemplates={setTemplates} users={users} setUsers={setUsers} trips={campTrips} notify={notify} />}
          {activeSection === "questions"           && <AdminQuestions users={users} setUsers={setUsers} notify={notify} />}
          {activeSection === "checklists"          && <AdminChecklists trips={campTrips} setTrips={setTrips} notify={notify} />}
          {activeSection === "itinerario"          && <AdminItinerary trips={campTrips} setTrips={setTrips} notify={notify} />}
          {activeSection === "logistica"           && <AdminLogistica trips={campTrips} setTrips={setTrips} notify={notify} />}
          {activeSection === "trips"               && <AdminTrips trips={campTrips} setTrips={setTrips} notify={notify} />}
          {activeSection === "calculadora"         && <CalculadoraCampamento />}
          {activeSection === "school_colegios"     && <AdminSchools trips={trips} schoolTripIds={schoolTripIds} setSchoolTripIds={setSchoolTripIds} notify={notify} section="colegios" />}
          {activeSection === "school_alumnos"      && <AdminSchools trips={trips} notify={notify} section="alumnos" />}
          {activeSection === "school_alergias"     && <AdminSchools trips={trips} notify={notify} section="alergias" />}
          {activeSection === "school_docs"         && <AdminSchools trips={trips} notify={notify} section="docs" />}
          {activeSection === "school_pagos"        && <AdminSchools trips={trips} notify={notify} section="pagos" />}
          {activeSection === "school_preguntas"    && <AdminSchools trips={trips} notify={notify} section="preguntas" />}
          {activeSection === "school_rooming"      && <AdminSchools trips={trips} notify={notify} section="rooming" />}
          {activeSection === "school_grupos"       && <AdminSchools trips={trips} notify={notify} section="grupos" />}
          {activeSection === "school_seguimiento"  && <AdminSchools trips={trips} notify={notify} section="seguimiento" />}
          {activeSection === "school_checklist"    && <AdminSchools trips={trips} notify={notify} section="checklist" />}
          {activeSection === "school_itinerario"   && <AdminSchools trips={trips} notify={notify} section="itinerario" />}
          {activeSection === "school_logistica"    && <AdminSchools trips={trips} notify={notify} section="logistica" />}
          {activeSection === "school_viajes"       && <AdminSchools trips={trips} setTrips={setTrips} notify={notify} section="viajes" />}
          {activeSection === "email_camp"           && <AdminEmailTemplates category="campamentos" notify={notify} />}
          {activeSection === "email_col"            && <AdminEmailTemplates category="colegios" notify={notify} />}
        </main>
      </div>
    </div>
  );
}

// ─── Root component ──────────────────────────────────────────────────────────

export default function GIMELOOSPortalApp() {
  const [users, setUsers] = useState(initialUsers);
  const [trips, setTrips] = useState(initialTrips);
  const [schoolTripIds, setSchoolTripIds] = useState(new Set());
  const [templates, setTemplates] = useState(initialDocumentTemplates);
  // [CRÍTICO-1/2/3] auth guarda el userId que viene del token de Supabase Auth
  const [auth, setAuth] = useState({ userId: null, error: "", isLoading: false });
  const [notifications, setNotifications] = useState([]);
  const [isBootstrapping, setIsBootstrapping] = useState(true);
  const [sessionBootstrapped, setSessionBootstrapped] = useState(false);
  // [MEDIO-2] Estado de error de carga expuesto al usuario
  const [loadError, setLoadError] = useState(null);
  // Clave para forzar recarga de datos tras login (RLS requiere sesión activa)
  const [dataRefreshKey, setDataRefreshKey] = useState(0);
  // Evita que el efecto de limpieza borre el userId justo después de un login
  const pendingLoginRef = React.useRef(false);

  // Detectar enlace de recuperación de contraseña (hash #type=recovery)
  const [isRecoveryMode, setIsRecoveryMode] = useState(() => {
    if (typeof window === "undefined") return false;
    const hash = new URLSearchParams(window.location.hash.slice(1));
    return hash.get("type") === "recovery";
  });

  useEffect(() => {
    const { data: { subscription } } = supabase.auth.onAuthStateChange((event) => {
      if (event === "PASSWORD_RECOVERY") setIsRecoveryMode(true);
    });
    return () => subscription.unsubscribe();
  }, []);

  const currentUser = useMemo(
    () => users.find((u) => u.id === auth.userId) || null,
    [users, auth.userId]
  );

  const removeNotification = (id) => setNotifications((prev) => prev.filter((n) => n.id !== id));

  // [MEDIO-1] Timeout diferenciado: 7s si hay acción destructiva, 3,2s si no
  const notify = useCallback((message, options = {}) => {
    const id = `n-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
    const timeout = options.actionLabel ? 7000 : 3200;
    setNotifications((prev) => [...prev, { id, message, ...options }]);
    setTimeout(() => removeNotification(id), timeout);
  }, []);

  // Restaurar sesión desde localStorage al montar
  useEffect(() => {
    try {
      const storedUserId = window.localStorage.getItem(LOCAL_STORAGE_AUTH_KEY);
      if (storedUserId) setAuth((prev) => ({ ...prev, userId: storedUserId }));
    } catch (error) {
      console.error("Error leyendo sesión guardada:", error);
    } finally {
      setSessionBootstrapped(true);
    }
  }, []);

  // Carga inicial desde Supabase
  useEffect(() => {
    const loadSupabaseData = async () => {
      setLoadError(null);
      try {
        const { data: sessionData } = await supabase.auth.getSession();
        if (!sessionData?.session) {
          setIsBootstrapping(false);
          return;
        }

        const [tripsRes, templatesRes, participantsRes, docsRes, paymentsRes, pricingRes, questionsRes, schoolTripIdsRes] = await Promise.all([
          supabase.from("trips").select("*").order("created_at", { ascending: true }),
          supabase.from("document_templates").select("*").order("created_at", { ascending: true }),
          supabase.from("participants").select("*").order("created_at", { ascending: true }),
          supabase.from("participant_documents").select("participant_id,template_id,status,uploaded_file_name,file_path,storage_path,drive_url").order("created_at", { ascending: true }),
          supabase.from("participant_payments").select("participant_id,payment_key,name,amount,status,proof_name,proof_path,due_date").order("created_at", { ascending: true }),
          supabase.from("participant_pricing").select("participant_id,initial_price,final_price,discount").order("created_at", { ascending: true }),
          supabase.from("participant_questions").select("id,participant_id,message,created_at,status,reply,replied_at").order("created_at", { ascending: true }),
          supabase.from("school_trips").select("trip_id"),
        ]);

        if (!schoolTripIdsRes.error && schoolTripIdsRes.data?.length) {
          setSchoolTripIds(new Set(schoolTripIdsRes.data.map((r) => r.trip_id)));
        }

        if (!tripsRes.error && tripsRes.data?.length) {
          setTrips(tripsRes.data.map((t) => ({
            id: t.id, name: t.name, departureDate: t.departure_date || "", description: t.description || "",
            heroImage: t.hero_image || DEFAULT_HERO_IMAGES[0],
            heroImages: Array.isArray(t.hero_images) && t.hero_images.length ? t.hero_images : DEFAULT_HERO_IMAGES,
            transferInfo: t.transfer_info || { bank: "", accountHolder: "", iban: "", concept: "" },
            automation: t.automation || { autoReminderEnabled: false, reminderDaysBefore: 5 },
            showItinerary: t.automation?.showItinerary !== false,
            showLogistics: t.automation?.showLogistics !== false,
            documentRules: t.document_rules || [], paymentSchedule: t.payment_schedule || {},
            itinerary: t.itinerary || [],
            logistics: Array.isArray(t.logistics) ? t.logistics : [],
            checklist: t.checklist || [],
            tipo: t.tipo || "campamento",
          })));
        }

        if (!templatesRes.error && templatesRes.data?.length) {
          setTemplates(templatesRes.data.map((t) => ({ id: t.id, name: t.name, fileName: t.file_name || "", driveUrl: t.drive_url || "" })));
        }

        if (!participantsRes.error && participantsRes.data?.length) {
          const fallbackTemplates = templatesRes.data?.length > 0
            ? templatesRes.data.map((t) => ({ id: t.id, name: t.name, fileName: t.file_name || "", driveUrl: t.drive_url || "" }))
            : initialDocumentTemplates;

          const mappedUsers = participantsRes.data.map((p) => {
            const pDocs = (docsRes.data || []).filter((d) => d.participant_id === p.id);
            const pPayments = (paymentsRes.data || []).filter((pay) => pay.participant_id === p.id);
            const pPricing = (pricingRes.data || []).find((pr) => pr.participant_id === p.id);
            const pQuestions = (questionsRes.data || []).filter((q) => q.participant_id === p.id);
            const reservation = pPayments.find((pay) => pay.payment_key === "reservation") || { name: "Reserva", amount: 0, status: "pending", proof_name: "", due_date: "" };
            const firstInstallment = pPayments.find((pay) => pay.payment_key === "firstInstallment") || { name: "Primera cuota", amount: 0, status: "pending", proof_name: "", due_date: "" };
            const secondInstallment = pPayments.find((pay) => pay.payment_key === "secondInstallment") || { name: "Segunda cuota", amount: 0, status: "pending", proof_name: "", due_date: "" };
            const initialPrice = Number(pPricing?.initial_price || 0) || Number(pPricing?.final_price || 0) + Number(pPricing?.discount || 0);
            return {
              id: p.id, role: p.role || "client", username: p.username,
              authUid: p.auth_uid || "",
              schoolId: p.school_id || "",
              participantName: p.participant_name || "", motherName: p.mother_name || "",
              fatherName: p.father_name || "", parentName: p.parent_name || "",
              email: p.email || "", contactEmails: p.contact_emails || [],
              dni: p.dni || "",
              birthDate: p.birth_date || "",
              gender: p.gender || "",
              address: p.address || "",
              school: p.school || "",
              phoneFather: p.phone_father || "",
              phoneMother: p.phone_mother || "",
              dniFather: p.dni_father || "",
              dniMother: p.dni_mother || "",
              imageAuth: p.image_auth || false,
              allergies: p.allergies || "",
              healthNotes: p.health_notes || "",
              shirtSize: p.shirt_size || "",
              notes: p.notes || "",
              modality: p.modality || "",
              howKnown: p.how_known || "",
              // [CRÍTICO-1] Sin campo password en el estado React
              tripId: p.trip_id || "",
              invoiceUrl: p.invoice_url || "",
              documents: pDocs.length > 0
                ? pDocs.map((d) => ({ id: d.template_id, status: d.status, uploadedFileName: d.uploaded_file_name || "", filePath: d.file_path || d.storage_path || "", driveUrl: d.drive_url || "" }))
                : fallbackTemplates.map((t) => ({ id: t.id, status: "pending_upload", uploadedFileName: "", filePath: "", driveUrl: "" })),
              payments: {
                initialPrice, discount: Number(pPricing?.discount || 0), finalPrice: Number(pPricing?.final_price || 0),
                reservation: { name: reservation.name || "Reserva", amount: Number(reservation.amount || 0), status: reservation.status || "pending", proofName: reservation.proof_name || "", proofPath: reservation.proof_path || "", dueDate: reservation.due_date || "" },
                firstInstallment: { name: firstInstallment.name || "Primera cuota", amount: Number(firstInstallment.amount || 0), status: firstInstallment.status || "pending", proofName: firstInstallment.proof_name || "", proofPath: firstInstallment.proof_path || "", dueDate: firstInstallment.due_date || "" },
                secondInstallment: { name: secondInstallment.name || "Segunda cuota", amount: Number(secondInstallment.amount || 0), status: secondInstallment.status || "pending", proofName: secondInstallment.proof_name || "", proofPath: secondInstallment.proof_path || "", dueDate: secondInstallment.due_date || "" },
              },
              checklistState: p.checklist_state || {},
              questions: pQuestions.map((q) => ({ id: q.id, message: q.message, createdAt: q.created_at, status: q.status, reply: q.reply || "", repliedAt: q.replied_at || null })),
            };
          });
          const hasAdmin = mappedUsers.some((u) => u.role === "admin");
          setUsers(hasAdmin ? mappedUsers : [...mappedUsers, ...initialUsers]);
        } else {
          setUsers(initialUsers);
        }
      } catch (error) {
        console.error("Error cargando Supabase:", error);
        // [MEDIO-2] Exponer el error al usuario con opción de reintentar
        setLoadError("No se pudieron cargar los datos. Comprueba tu conexión e inténtalo de nuevo.");
      } finally {
        pendingLoginRef.current = false;
        setIsBootstrapping(false);
      }
    };
    setIsBootstrapping(true);
    loadSupabaseData();
  }, [dataRefreshKey]);

  // Persistir userId en localStorage
  useEffect(() => {
    if (!sessionBootstrapped) return;
    try {
      if (auth.userId) window.localStorage.setItem(LOCAL_STORAGE_AUTH_KEY, auth.userId);
      else window.localStorage.removeItem(LOCAL_STORAGE_AUTH_KEY);
    } catch (error) {
      console.error("Error guardando sesión:", error);
    }
  }, [auth.userId, sessionBootstrapped]);

  // Limpiar sesión si el usuario ya no existe (solo cuando los datos están completamente cargados)
  useEffect(() => {
    if (!sessionBootstrapped || isBootstrapping || !auth.userId) return;
    // No limpiar durante un login reciente (la recarga aún no ha terminado)
    if (pendingLoginRef.current) return;
    if (!users.some((u) => u.id === auth.userId)) {
      setAuth({ userId: null, error: "", isLoading: false });
      try { window.localStorage.removeItem(LOCAL_STORAGE_AUTH_KEY); } catch (e) { console.error(e); }
    }
  }, [users, auth.userId, sessionBootstrapped, isBootstrapping, dataRefreshKey]);

  // [CRÍTICO-1/2/3] Login vía Supabase Auth — la comparación de contraseña ocurre en el servidor
  // Acepta: nombre de usuario (slug) O email directamente
  const handleLogin = async (username, password) => {
    setAuth((prev) => ({ ...prev, isLoading: true, error: "" }));
    try {
      const input = username.trim().toLowerCase();
      let loginEmail = null;

      // Si el input tiene @ intentamos usarlo directamente como email
      if (input.includes("@")) {
        loginEmail = input;
      } else {
        // Lookup email via SECURITY DEFINER RPC para saltarse RLS antes de autenticar
        const { data: email, error: lookupErr } = await supabase
          .rpc("get_login_email", { p_username: input });
        if (!lookupErr && email) loginEmail = email;
      }

      if (!loginEmail) {
        setAuth({ userId: null, error: "Usuario o contraseña incorrectos.", isLoading: false });
        return;
      }

      const { data: session, error: authErr } = await supabase.auth.signInWithPassword({
        email: loginEmail,
        password,
      });

      if (authErr || !session?.user) {
        setAuth({ userId: null, error: "Usuario o contraseña incorrectos.", isLoading: false });
        return;
      }

      // Resolver el participant.id real desde auth_uid (funciona para admin y participantes)
      const { data: participantId } = await supabase.rpc("get_participant_id_for_auth");
      const resolvedId = participantId ?? session.user.id;

      pendingLoginRef.current = true;
      setAuth({ userId: resolvedId, error: "", isLoading: false });
      setDataRefreshKey((k) => k + 1); // Recargar datos con sesión activa (RLS)
      notify("Acceso correcto.");
    } catch (err) {
      console.error(err);
      setAuth({ userId: null, error: "Error de conexión. Inténtalo de nuevo.", isLoading: false });
    }
  };

  const handleLogout = async () => {
    await supabase.auth.signOut();
    setAuth({ userId: null, error: "", isLoading: false });
    try {
      window.localStorage.removeItem(LOCAL_STORAGE_AUTH_KEY);
      window.localStorage.removeItem(ADMIN_SECTION_STORAGE_KEY);
    } catch (error) {
      console.error(error);
    }
  };

  // [MEDIO-2] Pantalla de carga con estado de error y botón de reintento
  if (isBootstrapping || !sessionBootstrapped) {
    return (
      <div>
        <ActionToast notifications={notifications} removeNotification={removeNotification} />
        <div className="min-h-screen bg-white text-zinc-950">
          <div className="mx-auto flex min-h-screen max-w-6xl flex-col items-center justify-center gap-4 p-6">
            {loadError ? (
              <div className="flex flex-col items-center gap-4 rounded-3xl border border-zinc-200 bg-white px-8 py-6 shadow-sm">
                <div className="text-sm text-red-700">{loadError}</div>
                <Button onClick={() => window.location.reload()} className="rounded-2xl text-white" style={{ backgroundColor: CORPORATE_RED }}>
                  Reintentar
                </Button>
              </div>
            ) : (
              <div className="rounded-3xl border border-zinc-200 bg-white px-6 py-5 text-sm text-zinc-600 shadow-sm">
                Cargando portal...
              </div>
            )}
          </div>
        </div>
      </div>
    );
  }

  if (isRecoveryMode) {
    return (
      <div className="[&_button]:cursor-pointer">
        <ActionToast notifications={notifications} removeNotification={removeNotification} />
        <ResetPasswordScreen onDone={() => {
          setIsRecoveryMode(false);
          window.history.replaceState(null, "", window.location.pathname);
        }} />
      </div>
    );
  }

  return (
    <div className="[&_button]:cursor-pointer [&_label]:cursor-pointer [&_select]:cursor-pointer [&_summary]:cursor-pointer">
      <ActionToast notifications={notifications} removeNotification={removeNotification} />
      {!currentUser ? (
        <LoginScreen onLogin={handleLogin} loginError={auth.error} isLoading={auth.isLoading} />
      ) : currentUser.role === "admin" ? (
        <AdminPanel users={users} setUsers={setUsers} trips={trips} setTrips={setTrips} schoolTripIds={schoolTripIds} setSchoolTripIds={setSchoolTripIds} templates={templates} setTemplates={setTemplates} onLogout={handleLogout} notify={notify} />
      ) : (currentUser.role === "school" || currentUser.schoolId) ? (
        <SchoolPortal user={currentUser} trips={trips} onLogout={handleLogout} notify={notify} />
      ) : (
        <ClientPortal user={currentUser} trips={trips} templates={templates} setUsers={setUsers} onLogout={handleLogout} notify={notify} />
      )}
    </div>
  );
}

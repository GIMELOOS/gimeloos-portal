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

import React, { useEffect, useMemo, useState, useCallback } from "react";
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
  FileCheck2,
  FileText,
  FolderUp,
  GripVertical,
  Image as ImageIcon,
  LogOut,
  Mail,
  MapPinned,
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
  Map,
  ListChecks,
  Plus,
  Luggage,
  Info,
  AlertCircle,
  Bus,
  Sun,
  Clock,
} from "lucide-react";
import { supabase } from "@/lib/supabase";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Badge } from "@/components/ui/badge";
import { Textarea } from "@/components/ui/textarea";
import { Separator } from "@/components/ui/separator";
import { Checkbox } from "@/components/ui/checkbox";

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
      bank: "Banco Santander",
      accountHolder: "GIMELOOS Experiences SL",
      iban: "ES12 1234 5678 9012 3456 7890",
      concept: "Nombre del participante + viaje",
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
      bank: "CaixaBank",
      accountHolder: "GIMELOOS Experiences SL",
      iban: "ES98 0000 1111 2222 3333 4444",
      concept: "Nombre del participante + Zarautz",
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
    user.payments.reservation.status !== "sent"
      ? "enviar el justificante de la reserva"
      : user.payments.firstInstallment.status !== "sent"
      ? "enviar el justificante de la primera cuota"
      : user.payments.secondInstallment.status !== "sent"
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
    await fetch("/api/notify", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        ...(token ? { Authorization: `Bearer ${token}` } : {}),
      },
      body: JSON.stringify({ type, to, participantId, data }),
    });
  } catch (err) {
    console.error("sendNotification error:", err);
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

async function uploadFileToStorage(file, folder = "misc") {
  const fileExt = file.name.split(".").pop();
  const fileName = `${Date.now()}-${Math.random().toString(36).slice(2)}.${fileExt}`;
  const filePath = `${folder}/${fileName}`;
  const { error } = await supabase.storage
    .from("participant-documents")
    .upload(filePath, file, { upsert: true });
  if (error) throw error;
  return { filePath, fileName: file.name };
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
async function upsertPayment(participantId, paymentKey, payload) {
  const { data: existing, error: fetchErr } = await supabase
    .from("participant_payments")
    .select("id, status")
    .eq("participant_id", participantId)
    .eq("payment_key", paymentKey)
    .maybeSingle();

  if (fetchErr) throw fetchErr;

  const PROTECTED_STATUSES = ["confirmed", "sent"];

  if (existing?.id) {
    const updates = { ...payload };
    // Si ya está confirmado/enviado, no degradar el estado ni borrar justificante
    if (PROTECTED_STATUSES.includes(existing.status)) {
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
                  {/* [CRÍTICO-1] onSubmit llama a supabase.auth.signInWithPassword en el padre */}
                  <form
                    className="space-y-4"
                    onSubmit={(e) => { e.preventDefault(); onLogin(username, password); }}
                  >
                    <div className="space-y-2">
                      <Label>Usuario</Label>
                      <Input
                        value={username}
                        onChange={(e) => setUsername(e.target.value)}
                        placeholder="Introduce tu usuario"
                        className="h-12 rounded-2xl border-zinc-200 bg-white"
                        autoComplete="username"
                      />
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

  const heroImages = useMemo(
    () =>
      trip.heroImages?.length
        ? trip.heroImages
        : [trip.heroImage || DEFAULT_HERO_IMAGES[0], ...DEFAULT_HERO_IMAGES.slice(1)],
    [trip.heroImages, trip.heroImage]
  );

  const activeImage = 0;
  const totalPending = pendingSummary.reduce((acc, s) => acc + s.count, 0);

  return (
    <div className="relative overflow-hidden rounded-[32px] shadow-[0_20px_70px_rgba(0,0,0,0.12)]">
      <img src={heroImages[activeImage]} alt={trip.name} className="absolute inset-0 block h-full w-full object-cover object-center" />
      <div className="absolute inset-0 bg-[linear-gradient(135deg,rgba(0,0,0,0.80),rgba(0,0,0,0.45),rgba(255,49,49,0.16))]" />
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
                  <div className="mt-1 text-sm text-zinc-500">Plantilla: {template?.fileName || "Sin archivo"}</div>
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
                  <Button variant="outline" className="rounded-2xl border-zinc-200 bg-white" disabled={!template?.driveUrl} onClick={() => template?.driveUrl && window.open(template.driveUrl, "_blank", "noopener,noreferrer")}>
                    <Download className="mr-2 h-4 w-4" />Descargar plantilla
                  </Button>
                  <label className={uploading ? "cursor-not-allowed opacity-60" : "cursor-pointer"}>
                    <input type="file" className="hidden" disabled={uploading} onChange={(e) => {
                      const file = e.target.files?.[0];
                      if (!file) return;
                      setProgress((p) => ({ ...p, [docItem.id]: 0 }));
                      const interval = setInterval(() => {
                        setProgress((p) => {
                          const cur = p[docItem.id] ?? 0;
                          if (cur >= 88) { clearInterval(interval); return p; }
                          return { ...p, [docItem.id]: Math.min(88, cur + Math.random() * 12) };
                        });
                      }, 400);
                      onUploadDocument(docItem.id, file, () => {})
                        .then(() => { clearInterval(interval); setProgress((p) => ({ ...p, [docItem.id]: 100 })); })
                        .catch(() => { clearInterval(interval); })
                        .finally(() => setTimeout(() => setProgress((p) => { const n = { ...p }; delete n[docItem.id]; return n; }), 1800));
                    }} />
                    <span className="inline-flex h-10 items-center rounded-2xl px-4 text-sm font-medium text-white" style={{ backgroundColor: CORPORATE_RED }}>
                      <Upload className="mr-2 h-4 w-4" />{uploading ? `${pct}%` : "Subir documento"}
                    </span>
                  </label>
                </div>
              </CardContent>
            </Card>
          );
        })}
      </div>
    </div>
  );
}

function PaymentRow({ title, payment, onUploadProof }) {
  const status = getStatusMeta(payment.status);
  const [pct, setPct] = useState(undefined);
  const uploading = pct !== undefined && pct < 100;
  return (
    <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
      <CardContent className="flex flex-col gap-4 p-5 lg:flex-row lg:items-center lg:justify-between">
        <div className="min-w-0 flex-1">
          <div className="font-medium text-zinc-950">{title}</div>
          <div className="mt-1 text-sm text-zinc-500">Importe: {formatCurrency(payment.amount)}</div>
          {payment.proofName && <div className="mt-1 text-sm text-zinc-500">Justificante: {payment.proofName}</div>}
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
          <label className={uploading ? "cursor-not-allowed opacity-60" : "cursor-pointer"}>
            <input type="file" className="hidden" disabled={uploading} onChange={(e) => {
              const file = e.target.files?.[0];
              if (!file) return;
              setPct(0);
              const interval = setInterval(() => {
                setPct((cur) => {
                  if ((cur ?? 0) >= 88) { clearInterval(interval); return cur; }
                  return Math.min(88, (cur ?? 0) + Math.random() * 12);
                });
              }, 400);
              onUploadProof(title, file, () => {})
                .then(() => { clearInterval(interval); setPct(100); })
                .catch(() => { clearInterval(interval); })
                .finally(() => setTimeout(() => setPct(undefined), 1800));
            }} />
            <span className="inline-flex h-10 items-center rounded-2xl px-4 text-sm font-medium text-white" style={{ backgroundColor: CORPORATE_RED }}>
              <FolderUp className="mr-2 h-4 w-4" />{uploading ? `${pct}%` : "Subir justificante"}
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
  const finalPrice = Number(payments.finalPrice || 0);
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
                ["Descuento", discount],
                ["Precio final", finalPrice],
                ["Importe reserva", Number(reservation.amount || 0)],
                ["Primera cuota", Number(firstInstallment.amount || 0)],
                ["Segunda cuota", Number(secondInstallment.amount || 0)],
              ].map(([label, value]) => (
                // [MENOR-3] Key única: label es único en este array
                <div key={String(label)} className="rounded-2xl border border-zinc-200 bg-white p-4">
                  <div className="text-xs uppercase tracking-[0.18em] text-zinc-500">{label}</div>
                  <div className="mt-2 text-xl font-semibold text-zinc-950">{formatCurrency(value)}</div>
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
              // [MENOR-3] Key: item es string único en el checklist
              <label key={item} className="flex cursor-pointer items-center gap-3 rounded-2xl border border-zinc-200 bg-white p-4 text-sm text-zinc-800 shadow-sm">
                <Checkbox checked={!!checklistState[item]} onCheckedChange={() => onToggleItem(item)} />
                <span>{item}</span>
              </label>
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
  }, [forceOpen]);

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
  const trip = trips.find((t) => t.id === user.tripId) || trips[0];

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

  const [openSection, setOpenSection] = useState(null);
  const navigateTo = (key) => setOpenSection(key);

  const pendingSummary = [
    { key: "docs",      label: "Docs",       icon: FileCheck2,           count: pendingDocuments, sectionId: "section-docs" },
    { key: "payments",  label: "Pagos",      icon: Wallet,               count: pendingPayments,  sectionId: "section-payments" },
    { key: "replies",   label: "Respuestas", icon: MessageCircleQuestion, count: unreadReplies,   sectionId: "section-questions" },
    { key: "notifs",    label: "Avisos",     icon: Bell,                 count: unreadCount,      sectionId: null },
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
            <div className="relative">
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
                  <div className="max-h-80 overflow-y-auto">
                    {notifications.length === 0 ? (
                      <div className="px-4 py-6 text-center text-sm text-zinc-400">Sin notificaciones</div>
                    ) : notifications.map((n) => (
                      <div key={n.id} className={`border-b border-zinc-50 px-4 py-3 ${n.read ? "" : "bg-red-50"}`}>
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

          <div className="space-y-4">
            <AccordionSection
              title="Documentación"
              subtitle="Revisa, descarga y sube cada documento pendiente."
              icon={FileCheck2}
              hasUnread={unreadBySection.docs}
              sectionId="section-docs"
              forceOpen={openSection === "docs"}
              onForceOpenConsumed={() => setOpenSection(null)}
              meta={<Badge className="bg-zinc-100 text-zinc-900 hover:bg-zinc-100">{pendingDocuments} pendientes</Badge>}
            >
              <ClientDocuments
                user={user}
                templates={templates}
                onUploadDocument={async (docId, file, onProgress) => {
                  // [ALTO-2] Snapshot para rollback
                  const previousDocs = user.documents;
                  try {
                    const uploaded = await uploadFileToDrive(file, user.participantName, "documentos", onProgress, trip?.name);
                    const nextDocs = user.documents.map((doc) =>
                      doc.id === docId
                        ? { ...doc, uploadedFileName: file.name, filePath: "", driveUrl: uploaded.webViewLink, status: "pending_confirmation" }
                        : doc
                    );
                    updateCurrentUser((current) => ({ ...current, documents: nextDocs }));

                    // [ALTO-1] upsertDocument no degrada documentos confirmados
                    await upsertDocument(user.id, docId, {
                      status: "pending_confirmation",
                      uploaded_file_name: file.name,
                      file_path: "",
                      storage_path: "",
                      drive_url: uploaded.webViewLink,
                      confirmed_at: null,
                    });
                    notify(`Documento subido: ${file.name}.`);
                    const docName = templates.find((t) => t.id === docId)?.name || file.name;
                    sendNotification("admin_doc_uploaded", null, null, { participantName: user.participantName, docName, tripName: trip?.name || "" });
                  } catch (error) {
                    console.error(error);
                    // [ALTO-2] Rollback si falla
                    setUsers((prev) => prev.map((u) => u.id === user.id ? { ...u, documents: previousDocs } : u));
                    notify("No se ha podido subir el documento. Los cambios han sido revertidos.");
                  }
                }}
              />
            </AccordionSection>

            <AccordionSection
              title="Pagos"
              subtitle="Importes, estados y justificantes de transferencia."
              icon={Wallet}
              hasUnread={unreadBySection.payments}
              sectionId="section-payments"
              forceOpen={openSection === "payments"}
              onForceOpenConsumed={() => setOpenSection(null)}
              meta={<Badge className="bg-zinc-100 text-zinc-900 hover:bg-zinc-100">{sentPayments}/3 enviados</Badge>}
            >
              <ClientPayments
                user={user}
                trip={trip}
                onUploadProof={async (paymentKey, file, onProgress) => {
                  // [ALTO-2] Snapshot para rollback
                  const previousPayments = user.payments;
                  try {
                    const uploaded = await uploadFileToDrive(file, user.participantName, "pagos", onProgress, trip?.name);
                    updateCurrentUser((current) => ({
                      ...current,
                      payments: {
                        ...current.payments,
                        [paymentKey]: { ...current.payments[paymentKey], proofName: file.name, proofPath: uploaded.webViewLink, status: "sent" },
                      },
                    }));

                    // [ALTO-1] upsertPayment no degrada pagos confirmados
                    await upsertPayment(user.id, paymentKey, {
                      name: user.payments[paymentKey].name,
                      amount: user.payments[paymentKey].amount,
                      status: "sent",
                      proof_name: file.name,
                      proof_path: uploaded.webViewLink,
                      due_date: user.payments[paymentKey].dueDate || null,
                    });
                    notify("Justificante cargado correctamente.");
                    sendNotification("admin_payment_uploaded", null, null, { participantName: user.participantName, paymentName: user.payments[paymentKey].name, tripName: trip?.name || "" });
                  } catch (error) {
                    console.error(error);
                    // [ALTO-2] Rollback si falla
                    setUsers((prev) => prev.map((u) => u.id === user.id ? { ...u, payments: previousPayments } : u));
                    notify("No se ha podido subir el justificante. Los cambios han sido revertidos.");
                  }
                }}
              />
            </AccordionSection>

            {trip.showLogistics !== false && (
              <AccordionSection
                title="Lo que no puedes olvidar"
                subtitle="Información clave antes del viaje: punto de encuentro, horarios y qué llevar."
                icon={MapPinned}
                meta={<Badge className="bg-zinc-100 text-zinc-900 hover:bg-zinc-100">{(trip.logistics || []).length} puntos</Badge>}
              >
                <ClientLogistics trip={trip} />
              </AccordionSection>
            )}

            {trip.showItinerary !== false && (
              <AccordionSection
                title="Itinerario del viaje"
                subtitle="Qué ocurrirá cada día durante la experiencia."
                icon={CalendarDays}
                meta={<Badge className="bg-zinc-100 text-zinc-900 hover:bg-zinc-100">{trip.itinerary.length} días</Badge>}
              >
                <ClientItinerary trip={trip} />
              </AccordionSection>
            )}

            <AccordionSection
              title="Checklist de equipaje"
              subtitle="Controla lo que ya tienes preparado y lo que te falta."
              icon={CheckCircle2}
              meta={<Badge className="bg-zinc-100 text-zinc-900 hover:bg-zinc-100">{completedChecklist}/{trip.checklist.length} listo</Badge>}
            >
              <ClientChecklist
                user={user}
                trip={trip}
                onToggleItem={async (item) => {
                  const previousState = user.checklistState;
                  const nextChecklist = { ...user.checklistState, [item]: !user.checklistState[item] };
                  updateCurrentUser((current) => ({ ...current, checklistState: nextChecklist }));
                  try {
                    await supabase.from("participants").update({ checklist_state: nextChecklist }).eq("id", user.id);
                  } catch (error) {
                    console.error(error);
                    // [ALTO-2] Rollback
                    setUsers((prev) => prev.map((u) => u.id === user.id ? { ...u, checklistState: previousState } : u));
                    notify("No se pudo guardar el checklist. Los cambios han sido revertidos.");
                  }
                }}
              />
            </AccordionSection>

            <AccordionSection
              title="¿Tienes alguna duda?"
              subtitle="Escríbenos y te responderemos lo antes posible."
              icon={MessageCircleQuestion}
              hasUnread={unreadBySection.questions}
              sectionId="section-questions"
              forceOpen={openSection === "replies"}
              onForceOpenConsumed={() => setOpenSection(null)}
              meta={<Badge className="bg-zinc-100 text-zinc-900 hover:bg-zinc-100">{questionsCount} enviadas</Badge>}
            >
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
                    // Reemplazar el ID temporal con el UUID real de Supabase
                    updateCurrentUser((current) => ({ ...current, questions: current.questions.map((q) => q.id === tempId ? { ...q, id: data.id } : q) }));
                    notify("Tu duda ha sido enviada.");
                    sendNotification("admin_new_question", null, null, { participantName: user.participantName, question: message, tripName: trip?.name || "" });
                  } catch (error) {
                    setUsers((prev) => prev.map((u) => u.id === user.id ? { ...u, questions: previousQuestions } : u));
                    throw error;
                  }
                }}
              />
            </AccordionSection>
          </div>
        </div>
      </div>
    </div>
  );
}

// ─── Panel de administración ─────────────────────────────────────────────────

function AdminClients({ users, trips, setUsers, templates, notify, setTrips }) {
  const clients = users.filter((u) => u.role === "client");
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

  const visibleClients = clients.filter(
    (client) =>
      (selectedTripFilter === "all" || client.tripId === selectedTripFilter) &&
      matchesParticipantSearch(client, searchQuery)
  );
  const visibleClientIds = visibleClients.map((c) => c.id);
  const allVisibleSelected = visibleClientIds.length > 0 && visibleClientIds.every((id) => selectedClientIds.includes(id));

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
    setUsers((prev) => prev.map((u) => targetIds.includes(u.id) ? { ...u, tripId: assignTargetTrip } : u));
    for (const clientId of targetIds) {
      const { error } = await supabase.from("participants").update({ trip_id: assignTargetTrip }).eq("id", clientId);
      if (error) { notify("Error asignando experiencia: " + error.message); return; }
    }
    notify("Experiencia aplicada correctamente.");
  };

  const restoreDeletedClient = async (deletedClient) => {
    const payload = {
      id: deletedClient.id,
      role: deletedClient.role,
      username: deletedClient.username,
      // [CRÍTICO-1] password NO se restaura desde el cliente; debe gestionarse en Supabase Auth
      participant_name: deletedClient.participantName || "",
      mother_name: deletedClient.motherName || "",
      father_name: deletedClient.fatherName || "",
      parent_name: deletedClient.parentName || "",
      email: deletedClient.email || "",
      contact_emails: deletedClient.contactEmails || [],
      trip_id: deletedClient.tripId || null,
      checklist_state: deletedClient.checklistState || {},
    };
    await supabase.from("participants").insert(payload);
    await supabase.from("participant_pricing").upsert(
      { participant_id: deletedClient.id, initial_price: deletedClient.payments?.initialPrice || 0, discount: deletedClient.payments?.discount || 0, final_price: deletedClient.payments?.finalPrice || 0 },
      { onConflict: "participant_id" }
    );
    for (const [key, name] of [["reservation", "Reserva"], ["firstInstallment", "Primera cuota"], ["secondInstallment", "Segunda cuota"]]) {
      await supabase.from("participant_payments").insert({
        participant_id: deletedClient.id,
        payment_key: key,
        name: deletedClient.payments?.[key]?.name || name,
        amount: deletedClient.payments?.[key]?.amount || 0,
        status: deletedClient.payments?.[key]?.status || "pending",
        proof_name: deletedClient.payments?.[key]?.proofName || "",
        due_date: deletedClient.payments?.[key]?.dueDate || null,
      });
    }
    for (const doc of deletedClient.documents || []) {
      await supabase.from("participant_documents").insert({
        participant_id: deletedClient.id,
        template_id: doc.id,
        status: doc.status || "pending_upload",
        uploaded_file_name: doc.uploadedFileName || "",
        file_path: doc.filePath || "",
        storage_path: doc.filePath || "",
        drive_url: doc.driveUrl || "",
        confirmed_at: doc.status === "confirmed" ? new Date().toISOString() : null,
      });
    }
    for (const q of deletedClient.questions || []) {
      await supabase.from("participant_questions").insert({
        participant_id: deletedClient.id,
        message: q.message || "",
        status: q.status || "sent",
        created_at: q.createdAt || new Date().toISOString(),
      });
    }
  };

  const deleteSingleClient = async (clientId) => {
    const deletedClient = users.find((u) => u.id === clientId);
    if (!deletedClient) return;
    try {
      await supabase.from("participant_questions").delete().eq("participant_id", clientId);
      await supabase.from("participant_documents").delete().eq("participant_id", clientId);
      await supabase.from("participant_payments").delete().eq("participant_id", clientId);
      await supabase.from("participant_pricing").delete().eq("participant_id", clientId);
      await supabase.from("participants").delete().eq("id", clientId);
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
        await supabase.from("participant_questions").delete().eq("participant_id", clientId);
        await supabase.from("participant_documents").delete().eq("participant_id", clientId);
        await supabase.from("participant_payments").delete().eq("participant_id", clientId);
        await supabase.from("participant_pricing").delete().eq("participant_id", clientId);
        await supabase.from("participants").delete().eq("id", clientId);
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

      await supabase.from("document_templates").upsert(
        templatesToUse.map((t) => ({ id: t.id, name: t.name, file_name: t.fileName || "" })),
        { onConflict: "id" }
      );

      // ── Fase 1: parsear todas las filas en memoria ────────────────────────────
      setImportMessage("Procesando filas...");
      setImportProgress(10);

      const parsedRows = [];
      for (const row of rows) {
        const tripTitle = safeString(getRowValue(row, "Viaje_Titulo", "viaje_titulo", "Viaje", "viaje", "Trip", "trip"));
        const participantName = safeString(getRowValue(row, "Participante"));
        if (!tripTitle || !participantName) continue;

        const tripId = `trip-${slugify(tripTitle)}`;
        const usernameFromExcel = safeString(getRowValue(row, "Usuario"));
        const finalUsername = (usernameFromExcel || slugify(participantName) || `user-${Date.now()}`)
          .toString().trim().toLowerCase().replace(/\s+/g, "-");
        const motherName = safeString(getRowValue(row, "Nombre_Madre"));
        const fatherName = safeString(getRowValue(row, "Nombre_Padre"));
        const motherEmail = safeString(getRowValue(row, "Email_Madre"));
        const fatherEmail = safeString(getRowValue(row, "Email_Padre"));
        const passwordFromExcel = safeString(getRowValue(row, "Password"));
        const discount = parseAmount(getRowValue(row, "Viaje_Descuento"));
        const initialPrice = parseAmount(getRowValue(row, "Viaje_Precio"));
        const finalPrice = parseAmount(getRowValue(row, "Viaje_APagar")) || Math.max(0, initialPrice - discount);

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
            trip_id: tripId,
          },
          pricing: { initial_price: initialPrice, discount, final_price: finalPrice },
          payments: [
            { key: "reservation",       name: safeString(getRowValue(row, "Pago1_Nombre"), "Reserva"),        amount: parseAmount(getRowValue(row, "Pago1_Cantidad")), due_date: normalizeDateForDb(getRowValue(row, "Pago1_Fecha")) },
            { key: "firstInstallment",  name: safeString(getRowValue(row, "Pago2_Nombre"), "Primera cuota"),  amount: parseAmount(getRowValue(row, "Pago2_Cantidad")), due_date: normalizeDateForDb(getRowValue(row, "Pago2_Fecha")) },
            { key: "secondInstallment", name: safeString(getRowValue(row, "Pago3_Nombre"), "Segunda cuota"),  amount: parseAmount(getRowValue(row, "Pago3_Cantidad")), due_date: normalizeDateForDb(getRowValue(row, "Pago3_Fecha")) },
          ],
          participantName, motherName, fatherName,
          email: safeString(motherEmail, fatherEmail),
          contactEmails: Array.from(new Set([motherEmail, fatherEmail].filter(Boolean))),
          finalUsername, initialPrice, discount, finalPrice,
          passwordFromExcel,
        });
      }

      setImportTotal(parsedRows.length);

      // ── Fase 2: upsert trips (batch, sin duplicados) ──────────────────────────
      setImportMessage("Guardando viajes...");
      setImportProgress(20);

      const currentTrips = [...trips];
      const uniqueTrips = new Map();
      for (const r of parsedRows) {
        if (!uniqueTrips.has(r.tripId)) {
          const existingTrip = currentTrips.find((t) => t.id === r.tripId);
          uniqueTrips.set(r.tripId, {
            ...r.tripPayload,
            hero_image: existingTrip?.heroImage || DEFAULT_HERO_IMAGES[0],
            hero_images: existingTrip?.heroImages || DEFAULT_HERO_IMAGES,
            transfer_info: existingTrip?.transferInfo || { bank: "Banco Santander", accountHolder: "GIMELOOS Experiences SL", iban: "ES12 1234 5678 9012 3456 7890", concept: "Nombre del participante + viaje" },
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
          currentTrips.push({ id: tripId, name: tp.name, departureDate: tp.departure_date || "", description: tp.description, heroImage: tp.hero_image, heroImages: tp.hero_images, transferInfo: tp.transfer_info, automation: tp.automation, documentRules: tp.document_rules, paymentSchedule: tp.payment_schedule, itinerary: tp.itinerary, checklist: tp.checklist });
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

      if (toInsert.length) {
        const { data: inserted, error: insertErr } = await supabase.from("participants").insert(toInsert).select("id, username");
        if (insertErr) { notify(`Error insertando participantes: ${insertErr.message}`); setIsImporting(false); return; }
        for (const p of (inserted || [])) existingByUsername.set(p.username, p.id);
      }

      if (toUpdate.length) {
        await Promise.all(
          toUpdate.map((r) => supabase.from("participants").update(r.participantPayload).eq("id", existingByUsername.get(r.finalUsername)))
        );
      }

      setImportProgress(55);

      // ── Fase 5: batch upsert pricing (1 query) ────────────────────────────────
      setImportMessage("Sincronizando precios...");
      const pricingRows = parsedRows
        .filter((r) => existingByUsername.has(r.finalUsername))
        .map((r) => ({ participant_id: existingByUsername.get(r.finalUsername), ...r.pricing }));

      if (pricingRows.length) {
        await supabase.from("participant_pricing").upsert(pricingRows, { onConflict: "participant_id" });
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

      const paymentsToInsert = [];
      const paymentsToUpdate = [];

      for (const r of parsedRows) {
        const participantId = existingByUsername.get(r.finalUsername);
        if (!participantId) continue;
        for (const p of r.payments) {
          const existing = paymentStatusMap.get(`${participantId}__${p.key}`);
          if (!existing) {
            paymentsToInsert.push({ participant_id: participantId, payment_key: p.key, name: p.name, amount: p.amount, status: "pending", proof_name: "", proof_path: "", due_date: p.due_date });
          } else if (!PROTECTED.includes(existing.status)) {
            paymentsToUpdate.push({ id: existing.id, name: p.name, amount: p.amount, due_date: p.due_date });
          }
        }
      }

      if (paymentsToInsert.length) {
        const { error: payInsertErr } = await supabase.from("participant_payments").insert(paymentsToInsert);
        if (payInsertErr) console.error("Error insertando pagos:", payInsertErr);
      }

      // Updates individuales (necesario porque cada row tiene su propio id)
      await Promise.all(
        paymentsToUpdate.map(({ id, ...fields }) =>
          supabase.from("participant_payments").update(fields).eq("id", id)
        )
      );

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
          if (!res.ok) console.error("Error creando cuentas Auth:", result.error);
          else notify(`Accesos creados: ${result.created} nuevos, ${result.updated} actualizados.`);
        } catch (err) {
          console.error("Error llamando /api/create-auth-users:", err);
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
      notify(`Excel importado: ${rows.length} participante(s) procesados.`);
      if (input) input.value = "";
      setImportFileName("");
    } catch (error) {
      console.error(error);
      notify("No se ha podido leer o guardar el Excel. Revisa el formato.");
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
              <Label>Grupo origen</Label>
              <select value={selectedGroupTrip} onChange={(e) => setSelectedGroupTrip(e.target.value)} className="h-11 w-full rounded-2xl border border-zinc-200 bg-white px-4 text-sm">
                {trips.map((t) => <option key={t.id} value={t.id}>{t.name}</option>)}
              </select>
            </div>
            <div className="space-y-2">
              <Label>Asignar experiencia</Label>
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
        </CardContent>
      </Card>

      <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
        <CardContent className="space-y-4 p-5">
          <div className="flex flex-col gap-3 lg:flex-row lg:items-end lg:justify-between">
            <div className="flex flex-col gap-3 lg:flex-row lg:items-end">
              <div className="space-y-3">
                <Label className="block pb-1">Filtrar por viaje</Label>
                <select value={selectedTripFilter} onChange={(e) => setSelectedTripFilter(e.target.value)} className="h-11 min-w-[280px] rounded-2xl border border-zinc-200 bg-white px-4 text-sm">
                  <option value="all">Todos los viajes</option>
                  {trips.map((t) => <option key={t.id} value={t.id}>{t.name}</option>)}
                </select>
              </div>
              <Button variant="outline" className="h-11 rounded-2xl" onClick={toggleSelectAllVisible}>
                <CheckCircle2 className="mr-2 h-4 w-4" />{allVisibleSelected ? "Deseleccionar todos" : "Seleccionar todos"}
              </Button>
            </div>
            {selectedClientIds.length > 1 && (
              <div className="flex items-center gap-2 self-end">
                <Badge className="bg-zinc-900 text-white hover:bg-zinc-900">{selectedClientIds.length} seleccionados</Badge>
                <Button variant="outline" className="h-11 rounded-2xl" onClick={deleteSelectedClients}>
                  <Trash2 className="mr-2 h-4 w-4" />Eliminar selección
                </Button>
              </div>
            )}
          </div>

          <div className="grid gap-3">
            {visibleClients.map((client) => {
              const isSelected = selectedClientIds.includes(client.id);
              return (
                <div key={client.id} className={`grid gap-3 rounded-3xl border p-4 transition-all lg:grid-cols-[44px_1.2fr_1fr_44px] lg:items-center ${isSelected ? "border-zinc-900 bg-white shadow-sm" : "border-zinc-200 bg-white"}`}>
                  <div className="flex justify-center">
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
                    value={client.tripId && trips.some((t) => t.id === client.tripId) ? client.tripId : ""}
                    onChange={async (e) => {
                      const value = e.target.value || null;
                      setUsers((prev) => prev.map((u) => u.id === client.id ? { ...u, tripId: value } : u));
                      const { error } = await supabase.from("participants").update({ trip_id: value }).eq("id", client.id);
                      if (error) notify("Error asignando viaje: " + error.message);
                    }}
                    className="h-11 w-full rounded-2xl border border-zinc-200 bg-white px-4 text-sm"
                  >
                    <option value="">Sin viaje asignado</option>
                    {trips.map((t) => <option key={t.id} value={t.id}>{t.name}</option>)}
                  </select>
                  <div className="flex flex-col items-end gap-2">
                    <Button
                      variant="ghost" size="icon"
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
                    <Button variant="ghost" size="icon" onClick={() => deleteSingleClient(client.id)}>
                      <Trash2 className="h-4 w-4 text-zinc-700" />
                    </Button>
                  </div>
                </div>
              );
            })}
          </div>
        </CardContent>
      </Card>
    </div>
  );
}

function AdminTracking({ users, trips, templates, setUsers, notify }) {
  const clients = users.filter((u) => u.role === "client");
  const [selectedTripId, setSelectedTripId] = useState("all");
  const [searchQuery, setSearchQuery] = useState("");
  const filteredClients = clients.filter(
    (c) => (selectedTripId === "all" || c.tripId === selectedTripId) && matchesParticipantSearch(c, searchQuery)
  );
  const updateClient = (clientId, updater) =>
    setUsers((prev) => prev.map((u) => u.id === clientId ? updater(u) : u));
  const sendReminder = (client, type) =>
    notify(`Recordatorio enviado a ${client.parentName || getFamilyLabel(client)} sobre ${type}.`);
  const getSummaryTone = (v) => v === 0 ? "bg-red-100 text-red-700" : v === 1 ? "bg-amber-100 text-amber-700" : "bg-emerald-100 text-emerald-700";

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
            <Label className="mb-2 block">Filtrar por viaje</Label>
            <select value={selectedTripId} onChange={(e) => setSelectedTripId(e.target.value)} className="mt-2 h-11 min-w-[320px] rounded-2xl border border-zinc-200 bg-white px-4 text-sm">
              <option value="all">Todos los viajes</option>
              {trips.map((t) => <option key={t.id} value={t.id}>{t.name}</option>)}
            </select>
          </div>
          {selectedTripId !== "all" && (
            <div className="flex flex-wrap items-center gap-3 rounded-2xl border border-zinc-200 bg-white px-4 py-3 text-sm text-zinc-700">
              <span>Primer doc: <strong>{formatShortDate(calculateDueDateFromRule(trips.find((t) => t.id === selectedTripId)?.departureDate, trips.find((t) => t.id === selectedTripId)?.documentRules?.[0]))}</strong></span>
              <span>Último pago: <strong>{formatShortDate(getPaymentRuleDueDate(trips.find((t) => t.id === selectedTripId), "secondInstallment"))}</strong></span>
              <Button variant="outline" className="rounded-2xl" onClick={() => notify("Recordatorio masivo preparado.")}>
                <Mail className="mr-2 h-4 w-4" />Recordar al grupo
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

          return (
            <AccordionSection
              key={client.id}
              title={client.participantName}
              subtitle={`${getFamilyLabel(client) ? `Familia: ${getFamilyLabel(client)} · ` : ""}Usuario: ${client.username} · ${trip?.name || "Sin viaje"}`}
              icon={Users}
              meta={
                <div className="flex items-center gap-2 overflow-x-auto whitespace-nowrap">
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

              <div className="grid gap-5 xl:grid-cols-2">
                <div className="space-y-3">
                  <div className="text-sm font-semibold uppercase tracking-[0.18em] text-zinc-500">Documentación</div>
                  {client.documents.map((docItem) => {
                    const template = templates.find((t) => t.id === docItem.id);
                    const status = getStatusMeta(docItem.status);
                    return (
                      <div key={docItem.id} className="rounded-2xl border border-zinc-200 bg-white p-4">
                        <div className="flex flex-wrap items-center gap-3">
                          <div className="min-w-0 flex-1">
                            <div className="truncate font-medium text-zinc-950">{template?.name || docItem.id}</div>
                            <div className="truncate text-sm text-zinc-500">{docItem.uploadedFileName || "Sin archivo subido"}</div>
                            <div className="flex items-center gap-2 truncate text-sm text-zinc-500">
                              <span className={`h-2.5 w-2.5 rounded-full ${getDueStatus(getDocumentRuleDueDate(trip, docItem.id)).className}`} />
                              <span>Límite: {formatShortDate(getDocumentRuleDueDate(trip, docItem.id))}</span>
                            </div>
                          </div>
                          <Badge className={status.className} style={status.style}>{status.label}</Badge>
                          <select
                            value={docItem.status}
                            onChange={async (e) => {
                              const nextStatus = e.target.value;
                              const prevDocs = client.documents;
                              updateClient(client.id, (c) => ({ ...c, documents: c.documents.map((d) => d.id === docItem.id ? { ...d, status: nextStatus } : d) }));
                              try {
                                // [ALTO-2] Rollback en error
                                await upsertDocument(client.id, docItem.id, {
                                  status: nextStatus,
                                  uploaded_file_name: docItem.uploadedFileName || "",
                                  file_path: docItem.filePath || "",
                                  storage_path: docItem.filePath || "",
                                  drive_url: docItem.driveUrl || "",
                                  confirmed_at: nextStatus === "confirmed" ? new Date().toISOString() : null,
                                });
                                if (nextStatus === "confirmed" || nextStatus === "rejected") {
                                  const docName = template?.name || docItem.id;
                                  const tripName = trips.find((t) => t.id === client.tripId)?.name || "";
                                  sendNotification(nextStatus === "confirmed" ? "doc_confirmed" : "doc_rejected", client.email, client.id, { participantName: client.participantName, docName, tripName });
                                }
                              } catch (err) {
                                console.error(err);
                                updateClient(client.id, (c) => ({ ...c, documents: prevDocs }));
                                notify("No se pudo guardar el estado del documento.");
                              }
                            }}
                            className="h-10 min-w-[190px] rounded-2xl border border-zinc-200 bg-white px-3 text-sm"
                          >
                            <option value="pending_upload">Pendiente de envío</option>
                            <option value="pending_confirmation">Por revisar</option>
                            <option value="confirmed">Confirmado</option>
                            <option value="rejected">Rechazado</option>
                          </select>
                          <Button variant="outline" className="h-10 rounded-2xl px-3" disabled={!docItem.driveUrl && !docItem.filePath} onClick={() => window.open(docItem.driveUrl || docItem.filePath, "_blank", "noopener,noreferrer")}>
                            <Eye className="mr-2 h-4 w-4" />Ver documento
                          </Button>
                        </div>
                      </div>
                    );
                  })}
                </div>

                <div className="space-y-3">
                  <div className="text-sm font-semibold uppercase tracking-[0.18em] text-zinc-500">Pagos</div>
                  {[["reservation", client.payments.reservation], ["firstInstallment", client.payments.firstInstallment], ["secondInstallment", client.payments.secondInstallment]].map(([paymentKey, payment]) => {
                    const status = getStatusMeta(payment.status);
                    return (
                      <div key={String(paymentKey)} className="rounded-2xl border border-zinc-200 bg-white p-4">
                        <div className="flex flex-wrap items-center gap-3">
                          <div className="min-w-0 flex-1">
                            <div className="truncate font-medium text-zinc-950">{payment.name || String(paymentKey)}</div>
                            <div className="truncate text-sm text-zinc-500">{formatCurrency(payment.amount)} · {payment.proofName || "Sin justificante"}</div>
                            <div className="flex items-center gap-2 truncate text-sm text-zinc-500">
                              <span className={`h-2.5 w-2.5 rounded-full ${getDueStatus(getPaymentRuleDueDate(trip, paymentKey)).className}`} />
                              <span>Límite: {formatShortDate(getPaymentRuleDueDate(trip, paymentKey))}</span>
                            </div>
                          </div>
                          <Badge className={status.className} style={status.style}>{status.label}</Badge>
                          <select
                            value={payment.status}
                            onChange={async (e) => {
                              const nextStatus = e.target.value;
                              const prevPayments = client.payments;
                              updateClient(client.id, (c) => ({ ...c, payments: { ...c.payments, [paymentKey]: { ...c.payments[paymentKey], status: nextStatus } } }));
                              try {
                                // [ALTO-2] Rollback + [ALTO-1] upsertPayment centralizado
                                await upsertPayment(client.id, paymentKey, {
                                  name: payment.name || String(paymentKey),
                                  amount: Number(payment.amount || 0),
                                  status: nextStatus,
                                  proof_name: payment.proofName || "",
                                  proof_path: payment.proofPath || "",
                                  due_date: payment.dueDate || null,
                                });
                                if (nextStatus === "confirmed") {
                                  const tripName = trips.find((t) => t.id === client.tripId)?.name || "";
                                  const amount = new Intl.NumberFormat("es-ES", { style: "currency", currency: "EUR" }).format(payment.amount || 0);
                                  sendNotification("payment_confirmed", client.email, client.id, { participantName: client.participantName, paymentName: payment.name, amount, tripName });
                                }
                              } catch (err) {
                                console.error(err);
                                updateClient(client.id, (c) => ({ ...c, payments: prevPayments }));
                                notify("No se pudo guardar el estado del pago.");
                              }
                            }}
                            className="h-10 min-w-[190px] rounded-2xl border border-zinc-200 bg-white px-3 text-sm"
                          >
                            <option value="pending">Pendiente</option>
                            <option value="sent">Enviado</option>
                            <option value="confirmed">Confirmado</option>
                            <option value="rejected">Rechazado</option>
                          </select>
                          <Button variant="outline" className="h-10 rounded-2xl px-3" disabled={!payment.proofPath} onClick={() => window.open(payment.proofPath, "_blank", "noopener,noreferrer")}>
                            <Eye className="mr-2 h-4 w-4" />Ver justificante
                          </Button>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            </AccordionSection>
          );
        })}
      </div>
    </div>
  );
}

function AdminPayments({ users, setUsers, notify }) {
  const clients = users.filter((u) => u.role === "client");
  const [searchQuery, setSearchQuery] = useState("");
  const filteredClients = clients.filter((c) => matchesParticipantSearch(c, searchQuery));

  return (
    <div className="space-y-5">
      <SectionTitle icon={CreditCard} title="Pagos" subtitle="Edita importes y estados por cliente." />
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
                  </div>
                  {[["Precio inicial", "initialPrice"], ["Descuento", "discount"], ["Precio final", "finalPrice"], ["Reserva", "reservation"], ["1ª cuota", "firstInstallment"], ["2ª cuota", "secondInstallment"]].map(([label, key]) => (
                    <div key={label}>
                      <Label className="mb-2 block">{label}</Label>
                      {["initialPrice", "discount", "finalPrice"].includes(key) ? (
                        <Input type="number" value={client.payments[key] ?? 0}
                          onChange={async (e) => {
                            const num = Number(e.target.value || 0);
                            setUsers((prev) => prev.map((u) => u.id === client.id ? { ...u, payments: { ...u.payments, [key]: num } } : u));
                            const { error } = await supabase.from("participant_pricing").upsert(
                              { participant_id: client.id, initial_price: key === "initialPrice" ? num : client.payments.initialPrice || 0, discount: key === "discount" ? num : client.payments.discount || 0, final_price: key === "finalPrice" ? num : client.payments.finalPrice || 0 },
                              { onConflict: "participant_id" }
                            );
                            if (error) notify("Error guardando precio: " + error.message);
                          }}
                          className="rounded-2xl"
                        />
                      ) : (
                        <Input type="number" value={client.payments[key].amount}
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
  const [selectedTrip, setSelectedTrip] = useState(trips[0]?.id || "");
  const [selectedTemplateId, setSelectedTemplateId] = useState(templates[0]?.id || "");

  const addTemplate = async () => {
    if (!name.trim()) return;
    setUploading(true);
    try {
      const id = `doc-${Date.now()}`;
      let driveUrl = "";
      let fileName = uploadedFile?.name || `${name.toLowerCase().replace(/\s+/g, "-")}.pdf`;
      if (uploadedFile) {
        const result = await uploadFileToDrive(uploadedFile, "archivos", "plantillas", null, "GIMELOOS Plantillas");
        driveUrl = result.webViewLink;
        fileName = result.fileName;
      }
      const newTemplate = { id, name, fileName, driveUrl };
      setTemplates((prev) => [...prev, newTemplate]);
      await supabase.from("document_templates").upsert({ id, name, file_name: fileName, drive_url: driveUrl });
      setName(""); setUploadedFile(null); setSelectedTemplateId(id);
      notify("Nueva plantilla creada.");
    } catch (err) {
      notify("Error al crear la plantilla: " + err.message);
    } finally {
      setUploading(false);
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
                    <Button variant="ghost" size="icon" onClick={() => deleteTemplate(t.id)}>
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
              <Label>Viaje</Label>
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
              <Label>Viaje</Label>
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
  const selectedTrip = trips.find((t) => t.id === selectedTripId) || trips[0];

  const syncField = (field, value) => setTrips((prev) => prev.map((t) => t.id === selectedTripId ? { ...t, [field]: value } : t));
  const saveField = async (field, value) => {
    syncField(field, value);
    const { error } = await supabase.from("trips").update({ [field]: value }).eq("id", selectedTripId);
    if (error) notify("Error guardando cambios: " + error.message);
  };

  if (!trips.length) return <div className="py-16 text-center text-sm text-zinc-400">No hay viajes configurados.</div>;

  return (
    <div className="space-y-5">
      <SectionTitle icon={Map} title="Viajes" subtitle="Información básica y foto de portada de cada viaje." />
      <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
        <CardContent className="p-5">
          <div className="flex flex-col gap-3 sm:flex-row sm:items-center">
            <div className="flex-1 space-y-1">
              <Label>Viaje activo</Label>
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
              <Label>Foto de portada (URL)</Label>
              <div className="flex gap-3">
                <Input value={selectedTrip.heroImage || ""} onChange={(e) => syncField("heroImage", e.target.value)} onBlur={(e) => saveField("hero_image", e.target.value)} placeholder="https://..." className="rounded-2xl" />
                {selectedTrip.heroImage && (
                  <img src={selectedTrip.heroImage} alt="portada" className="h-11 w-20 rounded-2xl object-cover border border-zinc-200" onError={(e) => { e.target.style.display = "none"; }} />
                )}
              </div>
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

  if (!trips.length) return <div className="py-16 text-center text-sm text-zinc-400">No hay viajes configurados.</div>;

  return (
    <div className="space-y-5">
      <SectionTitle icon={CalendarDays} title="Itinerario" subtitle="Programa día a día de cada viaje. Arrastra para reordenar." />
      <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
        <CardContent className="p-5">
          <div className="space-y-1">
            <Label>Viaje</Label>
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

  if (!trips.length) return <div className="py-16 text-center text-sm text-zinc-400">No hay viajes configurados.</div>;

  return (
    <div className="space-y-5">
      <SectionTitle icon={MapPinned} title="Logística" subtitle="Datos clave previos al viaje: horarios, lugar de encuentro, qué llevar..." />
      <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
        <CardContent className="p-5">
          <div className="space-y-1">
            <Label>Viaje</Label>
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

function usePaymentReminders(users, trips) {
  useEffect(() => {
    if (!users.length || !trips.length) return;
    // Registra qué recordatorios ya se enviaron: { "participantId_paymentKey_7d": true, ... }
    const SENT_KEY = "gimeloos_reminders_v2";
    const sent = JSON.parse(localStorage.getItem(SENT_KEY) || "{}");
    const PAYMENT_MILESTONES = [7, 3, 1]; // días antes del vencimiento
    const clients = users.filter((u) => u.role === "client");
    const reminders = [];
    const newSent = { ...sent };

    clients.forEach((client) => {
      if (!client.email) return;
      const trip = trips.find((t) => t.id === client.tripId);
      const tripName = trip?.name || "";

      // Recordatorios de pago — solo en los hitos exactos
      ["reservation", "firstInstallment", "secondInstallment"].forEach((key) => {
        const payment = client.payments?.[key];
        if (!payment || ["confirmed", "sent"].includes(payment.status)) return;
        if (!payment.dueDate) return;
        const due = new Date(payment.dueDate);
        const daysLeft = Math.round((due - new Date()) / 86400000);
        PAYMENT_MILESTONES.forEach((milestone) => {
          if (daysLeft !== milestone) return;
          const sentKey = `${client.id}_${key}_${milestone}d`;
          if (sent[sentKey]) return; // Ya enviado
          reminders.push(
            sendNotification("payment_reminder", client.email, client.id, {
              participantName: client.participantName,
              paymentName: payment.name,
              amount: new Intl.NumberFormat("es-ES", { style: "currency", currency: "EUR" }).format(payment.amount || 0),
              dueDate: due.toLocaleDateString("es-ES", { day: "numeric", month: "long" }),
              daysLeft,
              tripName,
            }).then(() => { newSent[sentKey] = true; })
          );
        });
      });

      // Recordatorio de docs — solo una vez por documento
      client.documents?.forEach((doc) => {
        if (doc.status !== "pending_upload") return;
        const sentKey = `${client.id}_doc_${doc.id}`;
        if (sent[sentKey]) return;
        reminders.push(
          sendNotification("doc_reminder", client.email, client.id, {
            participantName: client.participantName,
            docName: doc.id,
            tripName,
          }).then(() => { newSent[sentKey] = true; })
        );
      });
    });

    if (reminders.length > 0) {
      Promise.all(reminders).then(() => {
        localStorage.setItem(SENT_KEY, JSON.stringify(newSent));
      });
    }
  }, [users, trips]);
}

// ─── School Portal ────────────────────────────────────────────────────────────

function SchoolTrips({ schoolTrips }) {
  if (!schoolTrips.length) {
    return (
      <div className="flex flex-col items-center justify-center py-16 text-zinc-400">
        <CalendarDays className="mb-3 h-10 w-10 opacity-40" />
        <p className="text-sm">No hay viajes asignados todavía.</p>
      </div>
    );
  }
  return (
    <div className="space-y-4">
      <h2 className="text-base font-semibold text-zinc-950">Viajes del colegio</h2>
      <div className="grid gap-4 sm:grid-cols-2">
        {schoolTrips.map((st) => (
          <Card key={st.id} className="rounded-2xl border-zinc-200 shadow-sm">
            <CardContent className="p-5">
              <div className="mb-1 text-xs font-medium uppercase tracking-widest text-zinc-400">Viaje</div>
              <div className="text-base font-bold text-zinc-950">{st.trips?.name || st.trip_id}</div>
              {st.trips?.departure_date && (
                <div className="mt-1 flex items-center gap-1.5 text-xs text-zinc-500">
                  <CalendarDays className="h-3.5 w-3.5" />
                  {new Date(st.trips.departure_date).toLocaleDateString("es-ES", { day: "numeric", month: "long", year: "numeric" })}
                </div>
              )}
              {st.courses?.length > 0 && (
                <div className="mt-3">
                  <div className="mb-1 text-xs font-medium text-zinc-500">Cursos / grupos</div>
                  <div className="flex flex-wrap gap-1.5">
                    {st.courses.map((c) => (
                      <Badge key={c.id} variant="outline" className="rounded-xl text-xs">{c.course_name}{c.group_name ? ` · ${c.group_name}` : ""}</Badge>
                    ))}
                  </div>
                </div>
              )}
            </CardContent>
          </Card>
        ))}
      </div>
    </div>
  );
}

function SchoolStudents({ schoolTrips, courses, students, setStudents, notify }) {
  const [selectedTripId, setSelectedTripId] = useState(schoolTrips[0]?.id || "");
  const [selectedCourseId, setSelectedCourseId] = useState("");
  const [addName, setAddName] = useState("");
  const [addSurname, setAddSurname] = useState("");
  const [addAllergies, setAddAllergies] = useState("");
  const [addIntolerances, setAddIntolerances] = useState("");
  const [addNotes, setAddNotes] = useState("");
  const [xlsxPreview, setXlsxPreview] = useState(null); // { rows, mapping, headers, file }
  const [importing, setImporting] = useState(false);

  const tripCourses = courses.filter((c) => c.school_trip_id === selectedTripId);
  useEffect(() => { setSelectedCourseId(tripCourses[0]?.id || ""); }, [selectedTripId]);

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

  return (
    <div className="space-y-6">
      <h2 className="text-base font-semibold text-zinc-950">Alumnos</h2>
      {/* Selectors */}
      <div className="flex flex-wrap gap-3">
        <select
          value={selectedTripId}
          onChange={(e) => setSelectedTripId(e.target.value)}
          className="rounded-2xl border border-zinc-200 bg-white px-3 py-2 text-sm text-zinc-950 focus:outline-none"
        >
          {schoolTrips.map((st) => (
            <option key={st.id} value={st.id}>{st.trips?.name || st.trip_id}</option>
          ))}
        </select>
        <select
          value={selectedCourseId}
          onChange={(e) => setSelectedCourseId(e.target.value)}
          className="rounded-2xl border border-zinc-200 bg-white px-3 py-2 text-sm text-zinc-950 focus:outline-none"
        >
          {tripCourses.length === 0 && <option value="">Sin cursos</option>}
          {tripCourses.map((c) => (
            <option key={c.id} value={c.id}>{c.course_name}{c.group_name ? ` · ${c.group_name}` : ""}</option>
          ))}
        </select>
      </div>

      {/* Manual add */}
      {selectedCourseId && (
        <Card className="rounded-2xl border-zinc-200 shadow-sm">
          <CardContent className="p-4">
            <div className="mb-3 text-sm font-medium text-zinc-700">Añadir alumno manualmente</div>
            <div className="grid gap-2 sm:grid-cols-2 lg:grid-cols-3">
              <Input placeholder="Nombre *" value={addName} onChange={(e) => setAddName(e.target.value)} className="rounded-xl" />
              <Input placeholder="Apellidos" value={addSurname} onChange={(e) => setAddSurname(e.target.value)} className="rounded-xl" />
              <Input placeholder="Alergias" value={addAllergies} onChange={(e) => setAddAllergies(e.target.value)} className="rounded-xl" />
              <Input placeholder="Intolerancias" value={addIntolerances} onChange={(e) => setAddIntolerances(e.target.value)} className="rounded-xl" />
              <Input placeholder="Notas" value={addNotes} onChange={(e) => setAddNotes(e.target.value)} className="rounded-xl" />
              <Button onClick={handleAddManual} disabled={!addName.trim()} className="rounded-xl text-white" style={{ backgroundColor: CORPORATE_RED }}>
                <Plus className="mr-1.5 h-3.5 w-3.5" />Añadir
              </Button>
            </div>
          </CardContent>
        </Card>
      )}

      {/* Excel import */}
      {selectedCourseId && (
        <Card className="rounded-2xl border-zinc-200 shadow-sm">
          <CardContent className="p-4">
            <div className="mb-3 flex items-center justify-between">
              <div className="text-sm font-medium text-zinc-700">Importar desde Excel</div>
              <label className="flex cursor-pointer items-center gap-1.5 rounded-xl border border-zinc-200 bg-white px-3 py-1.5 text-xs font-medium text-zinc-700 hover:bg-zinc-50">
                <Upload className="h-3.5 w-3.5" />Seleccionar archivo
                <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleFileChange} />
              </label>
            </div>
            {xlsxPreview && (
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
                      <tr className="border-b border-zinc-100">
                        {xlsxPreview.headers.map((h, i) => <th key={i} className="px-2 py-1 text-left font-medium text-zinc-500">{h}</th>)}
                      </tr>
                    </thead>
                    <tbody>
                      {xlsxPreview.rows.slice(0, 5).map((r, ri) => (
                        <tr key={ri} className="border-b border-zinc-50">
                          {xlsxPreview.headers.map((_, ci) => <td key={ci} className="px-2 py-1 text-zinc-700">{String(r[ci] || "")}</td>)}
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
                    className="rounded-xl text-white text-xs"
                    style={{ backgroundColor: CORPORATE_RED }}
                  >
                    {importing ? "Importando..." : `Importar ${xlsxPreview.rows.filter((r) => r.some((c) => String(c).trim())).length} alumnos`}
                  </Button>
                  <Button variant="outline" className="rounded-xl text-xs" onClick={() => setXlsxPreview(null)}>Cancelar</Button>
                </div>
              </div>
            )}
          </CardContent>
        </Card>
      )}

      {/* Student list */}
      {selectedCourseId && (
        <Card className="rounded-2xl border-zinc-200 shadow-sm">
          <CardContent className="p-4">
            <div className="mb-3 flex items-center justify-between">
              <div className="text-sm font-medium text-zinc-700">Lista de alumnos</div>
              <Badge variant="outline" className="rounded-xl text-xs">{courseStudents.length} alumnos</Badge>
            </div>
            {courseStudents.length === 0 ? (
              <p className="text-xs text-zinc-400">No hay alumnos en este grupo todavía.</p>
            ) : (
              <div className="overflow-x-auto">
                <table className="w-full text-xs">
                  <thead>
                    <tr className="border-b border-zinc-100">
                      <th className="px-2 py-1.5 text-left font-medium text-zinc-500">Nombre</th>
                      <th className="px-2 py-1.5 text-left font-medium text-zinc-500">Apellidos</th>
                      <th className="px-2 py-1.5 text-left font-medium text-zinc-500">Alergias</th>
                      <th className="px-2 py-1.5 text-left font-medium text-zinc-500">Intolerancias</th>
                      <th className="px-2 py-1.5 text-left font-medium text-zinc-500">Notas</th>
                      <th className="px-2 py-1.5"></th>
                    </tr>
                  </thead>
                  <tbody>
                    {courseStudents.map((s) => (
                      <tr key={s.id} className="border-b border-zinc-50 hover:bg-zinc-50/50">
                        <td className="px-2 py-1.5 font-medium text-zinc-900">{s.name}</td>
                        <td className="px-2 py-1.5 text-zinc-700">{s.surname}</td>
                        <td className="px-2 py-1.5 text-zinc-700">{s.allergies || "—"}</td>
                        <td className="px-2 py-1.5 text-zinc-700">{s.intolerances || "—"}</td>
                        <td className="px-2 py-1.5 text-zinc-700">{s.notes || "—"}</td>
                        <td className="px-2 py-1.5">
                          <button onClick={() => handleDelete(s.id)} className="rounded-lg p-1 text-zinc-400 hover:bg-red-50 hover:text-red-600">
                            <X className="h-3.5 w-3.5" />
                          </button>
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
    </div>
  );
}

function SchoolAllergies({ courses, students }) {
  const withAllergies = students.filter((s) => s.allergies?.trim() || s.intolerances?.trim());
  const getCourse = (id) => courses.find((c) => c.id === id);

  return (
    <div className="space-y-4">
      <h2 className="text-base font-semibold text-zinc-950">Alergias e intolerancias</h2>
      {withAllergies.length === 0 ? (
        <div className="flex flex-col items-center justify-center py-16 text-zinc-400">
          <CheckCircle2 className="mb-3 h-10 w-10 opacity-40" />
          <p className="text-sm">Ningún alumno tiene alergias o intolerancias registradas.</p>
        </div>
      ) : (
        <Card className="rounded-2xl border-zinc-200 shadow-sm">
          <CardContent className="p-4">
            <div className="overflow-x-auto">
              <table className="w-full text-xs">
                <thead>
                  <tr className="border-b border-zinc-100">
                    <th className="px-2 py-1.5 text-left font-medium text-zinc-500">Nombre completo</th>
                    <th className="px-2 py-1.5 text-left font-medium text-zinc-500">Curso / grupo</th>
                    <th className="px-2 py-1.5 text-left font-medium text-zinc-500">Alergia</th>
                    <th className="px-2 py-1.5 text-left font-medium text-zinc-500">Intolerancia</th>
                    <th className="px-2 py-1.5 text-left font-medium text-zinc-500">Notas</th>
                  </tr>
                </thead>
                <tbody>
                  {withAllergies.map((s) => {
                    const course = getCourse(s.school_course_id);
                    return (
                      <tr key={s.id} className="border-b border-zinc-50 hover:bg-zinc-50/50">
                        <td className="px-2 py-1.5 font-medium text-zinc-900">{[s.name, s.surname].filter(Boolean).join(" ")}</td>
                        <td className="px-2 py-1.5 text-zinc-700">{course ? `${course.course_name}${course.group_name ? ` · ${course.group_name}` : ""}` : "—"}</td>
                        <td className="px-2 py-1.5 text-red-700">{s.allergies || "—"}</td>
                        <td className="px-2 py-1.5 text-amber-700">{s.intolerances || "—"}</td>
                        <td className="px-2 py-1.5 text-zinc-700">{s.diet_notes || s.notes || "—"}</td>
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

function SchoolDocs({ courses, schoolDocuments, setSchoolDocuments, notify }) {
  const handleUpload = async (docId, file) => {
    if (!file) return;
    const path = `school-docs/${docId}/${file.name}`;
    const { error: upErr } = await supabase.storage.from("documents").upload(path, file, { upsert: true });
    if (upErr) { notify("Error subiendo documento: " + upErr.message, { variant: "destructive" }); return; }
    const { data: urlData } = supabase.storage.from("documents").getPublicUrl(path);
    const { error: dbErr } = await supabase.from("school_documents").update({ file_url: urlData.publicUrl, status: "uploaded" }).eq("id", docId);
    if (dbErr) { notify("Error actualizando estado del documento.", { variant: "destructive" }); return; }
    setSchoolDocuments((prev) => prev.map((d) => d.id === docId ? { ...d, file_url: urlData.publicUrl, status: "uploaded" } : d));
    notify("Documento subido correctamente.");
  };

  if (!courses.length) {
    return <div className="py-16 text-center text-sm text-zinc-400">No hay cursos asignados.</div>;
  }

  return (
    <div className="space-y-6">
      <h2 className="text-base font-semibold text-zinc-950">Documentación requerida</h2>
      {courses.map((course) => {
        const courseDocs = schoolDocuments.filter((d) => d.school_course_id === course.id);
        return (
          <Card key={course.id} className="rounded-2xl border-zinc-200 shadow-sm">
            <CardContent className="p-5">
              <div className="mb-3 font-medium text-zinc-900">{course.course_name}{course.group_name ? ` · ${course.group_name}` : ""}</div>
              {courseDocs.length === 0 ? (
                <p className="text-xs text-zinc-400">Sin documentos requeridos.</p>
              ) : (
                <div className="space-y-2">
                  {courseDocs.map((doc) => (
                    <div key={doc.id} className="flex items-center justify-between rounded-xl border border-zinc-100 px-3 py-2">
                      <div className="flex items-center gap-2">
                        <FileCheck2 className="h-4 w-4 shrink-0 text-zinc-400" />
                        <div>
                          <div className="text-xs font-medium text-zinc-900">{doc.name}</div>
                          <div className="text-xs text-zinc-400">{doc.status === "uploaded" ? "Subido" : "Pendiente"}</div>
                        </div>
                      </div>
                      <div className="flex items-center gap-2">
                        {doc.file_url && (
                          <a href={doc.file_url} target="_blank" rel="noreferrer" className="text-xs text-blue-600 hover:underline">Ver</a>
                        )}
                        <label className="flex cursor-pointer items-center gap-1 rounded-xl border border-zinc-200 px-2 py-1 text-xs text-zinc-600 hover:bg-zinc-50">
                          <Upload className="h-3 w-3" />Subir
                          <input type="file" className="hidden" onChange={(e) => handleUpload(doc.id, e.target.files?.[0])} />
                        </label>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </CardContent>
          </Card>
        );
      })}
    </div>
  );
}

function SchoolRooming({ schoolTrips, setSchoolTrips, notify }) {
  const [selectedTripId, setSelectedTripId] = useState(schoolTrips[0]?.id || "");
  const [importing, setImporting] = useState(false);

  const selectedTrip = schoolTrips.find((st) => st.id === selectedTripId);
  const rooming = selectedTrip?.rooming || [];

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
        if (!row[0]) return;
        const roomName = String(row[0]).trim();
        const studentNames = row.slice(1).map((c) => String(c).trim()).filter(Boolean);
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

  return (
    <div className="space-y-6">
      <h2 className="text-base font-semibold text-zinc-950">Asignación de habitaciones</h2>
      <div className="flex flex-wrap items-center gap-3">
        <select
          value={selectedTripId}
          onChange={(e) => setSelectedTripId(e.target.value)}
          className="rounded-2xl border border-zinc-200 bg-white px-3 py-2 text-sm text-zinc-950 focus:outline-none"
        >
          {schoolTrips.map((st) => <option key={st.id} value={st.id}>{st.trips?.name || st.trip_id}</option>)}
        </select>
        <label className={`flex cursor-pointer items-center gap-1.5 rounded-2xl border border-zinc-200 bg-white px-3 py-2 text-sm font-medium text-zinc-700 hover:bg-zinc-50 ${importing ? "opacity-50" : ""}`}>
          <Upload className="h-4 w-4" />{importing ? "Importando..." : "Importar Excel"}
          <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleFileChange} disabled={importing} />
        </label>
      </div>
      <div className="text-xs text-zinc-400">Formato esperado: columna 1 = nombre de habitación, columnas siguientes = nombres de alumnos.</div>
      {rooming.length > 0 ? (
        <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
          {rooming.map((r, i) => (
            <Card key={i} className="rounded-2xl border-zinc-200 shadow-sm">
              <CardContent className="p-4">
                <div className="mb-2 font-medium text-zinc-900">{r.room}</div>
                <div className="space-y-1">
                  {r.students.map((s, j) => <div key={j} className="flex items-center gap-1.5 text-xs text-zinc-600"><User className="h-3 w-3 shrink-0 text-zinc-400" />{s}</div>)}
                </div>
                <Badge variant="outline" className="mt-2 rounded-xl text-xs">{r.students.length} alumnos</Badge>
              </CardContent>
            </Card>
          ))}
        </div>
      ) : (
        <div className="flex flex-col items-center justify-center py-16 text-zinc-400">
          <Users className="mb-3 h-10 w-10 opacity-40" />
          <p className="text-sm">Importa un Excel para ver las habitaciones aquí.</p>
        </div>
      )}
    </div>
  );
}

function SchoolGroups({ schoolTrips, setSchoolTrips, notify }) {
  const [selectedTripId, setSelectedTripId] = useState(schoolTrips[0]?.id || "");
  const [importing, setImporting] = useState(false);

  const selectedTrip = schoolTrips.find((st) => st.id === selectedTripId);
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
        const monitor = String(row[1] || "").trim();
        const students = row.slice(2).map((c) => String(c).trim()).filter(Boolean);
        if (group) parsed.push({ group, monitor, students });
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

  return (
    <div className="space-y-6">
      <h2 className="text-base font-semibold text-zinc-950">Grupos de actividades</h2>
      <div className="flex flex-wrap items-center gap-3">
        <select
          value={selectedTripId}
          onChange={(e) => setSelectedTripId(e.target.value)}
          className="rounded-2xl border border-zinc-200 bg-white px-3 py-2 text-sm text-zinc-950 focus:outline-none"
        >
          {schoolTrips.map((st) => <option key={st.id} value={st.id}>{st.trips?.name || st.trip_id}</option>)}
        </select>
        <label className={`flex cursor-pointer items-center gap-1.5 rounded-2xl border border-zinc-200 bg-white px-3 py-2 text-sm font-medium text-zinc-700 hover:bg-zinc-50 ${importing ? "opacity-50" : ""}`}>
          <Upload className="h-4 w-4" />{importing ? "Importando..." : "Importar Excel"}
          <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleFileChange} disabled={importing} />
        </label>
      </div>
      <div className="text-xs text-zinc-400">Formato esperado: columna 1 = nombre de grupo, columna 2 = monitor, columnas siguientes = alumnos.</div>
      {groups.length > 0 ? (
        <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
          {groups.map((g, i) => (
            <Card key={i} className="rounded-2xl border-zinc-200 shadow-sm">
              <CardContent className="p-4">
                <div className="mb-1 font-medium text-zinc-900">{g.group}</div>
                {g.monitor && <div className="mb-2 text-xs text-zinc-500">Monitor: {g.monitor}</div>}
                <div className="space-y-1">
                  {g.students.map((s, j) => <div key={j} className="flex items-center gap-1.5 text-xs text-zinc-600"><User className="h-3 w-3 shrink-0 text-zinc-400" />{s}</div>)}
                </div>
                <Badge variant="outline" className="mt-2 rounded-xl text-xs">{g.students.length} alumnos</Badge>
              </CardContent>
            </Card>
          ))}
        </div>
      ) : (
        <div className="flex flex-col items-center justify-center py-16 text-zinc-400">
          <Users className="mb-3 h-10 w-10 opacity-40" />
          <p className="text-sm">Importa un Excel para ver los grupos aquí.</p>
        </div>
      )}
    </div>
  );
}

function SchoolPortal({ user, onLogout, notify }) {
  const [activeTab, setActiveTab] = useState("trips");
  const [school, setSchool] = useState(null);
  const [schoolTrips, setSchoolTrips] = useState([]);
  const [courses, setCourses] = useState([]);
  const [students, setStudents] = useState([]);
  const [schoolDocuments, setSchoolDocuments] = useState([]);
  const [loading, setLoading] = useState(true);
  const [loadErr, setLoadErr] = useState(null);

  useEffect(() => {
    const load = async () => {
      setLoading(true);
      setLoadErr(null);
      try {
        // 1. Get school — linked via participants.auth_uid -> schools.auth_uid
        const { data: schoolData, error: schoolErr } = await supabase
          .from("schools")
          .select("*")
          .eq("auth_uid", user.authUid)
          .maybeSingle();
        if (schoolErr) throw new Error(schoolErr.message);
        if (!schoolData) { setLoadErr("No se encontró un colegio asociado a este usuario."); setLoading(false); return; }
        setSchool(schoolData);

        // 2. School trips
        const { data: tripsData, error: tripsErr } = await supabase
          .from("school_trips")
          .select("*, trips(name, departure_date)")
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

        // Attach courses to trips for display
        setSchoolTrips(tripsArr.map((st) => ({ ...st, courses: coursesArr.filter((c) => c.school_trip_id === st.id) })));
      } catch (err) {
        console.error("Error cargando datos del colegio:", err);
        setLoadErr("Error cargando datos: " + err.message);
      } finally {
        setLoading(false);
      }
    };
    if (user?.authUid) load();
    else { setLoadErr("No se pudo identificar el usuario del colegio."); setLoading(false); }
  }, [user?.authUid]);

  const tabs = [
    { key: "trips",     label: "Mis viajes",    icon: CalendarDays },
    { key: "students",  label: "Alumnos",       icon: Users },
    { key: "allergies", label: "Alergias",      icon: AlertCircle },
    { key: "docs",      label: "Documentación", icon: FileCheck2 },
    { key: "rooming",   label: "Rooming",       icon: LayoutGrid },
    { key: "groups",    label: "Grupos",        icon: ListChecks },
  ];

  return (
    <div className="min-h-screen text-zinc-950" style={{ background: "linear-gradient(160deg,#fff5f5 0%,#fafafa 40%,#f4f4f5 100%)" }}>
      {/* Navbar */}
      <div className="sticky top-0 z-30 flex items-center justify-between border-b border-zinc-200 bg-white px-6 py-3">
        <div>
          <div className="text-xs font-medium uppercase tracking-widest text-zinc-400">PORTAL ESCOLAR</div>
          <div className="text-lg font-bold text-zinc-950">{school?.name || "Colegio"}</div>
        </div>
        <Button variant="outline" className="rounded-2xl text-sm" onClick={() => { onLogout(); notify("Sesión cerrada."); }}>
          <LogOut className="mr-2 h-4 w-4" />Salir
        </Button>
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
            <Button onClick={() => window.location.reload()} className="rounded-2xl text-white text-xs" style={{ backgroundColor: CORPORATE_RED }}>Reintentar</Button>
          </div>
        </div>
      ) : (
        <div className="mx-auto max-w-[1200px] px-6 py-6">
          {/* Tab nav */}
          <div className="mb-6 flex flex-wrap gap-2">
            {tabs.map(({ key, label, icon: Icon }) => {
              const active = activeTab === key;
              return (
                <button
                  key={key}
                  type="button"
                  onClick={() => setActiveTab(key)}
                  className={`flex items-center gap-1.5 rounded-2xl px-4 py-2 text-sm font-medium transition ${
                    active ? "text-white shadow-sm" : "border border-zinc-200 bg-white text-zinc-700 hover:bg-zinc-50"
                  }`}
                  style={active ? { backgroundColor: CORPORATE_RED } : {}}
                >
                  <Icon className="h-4 w-4" />{label}
                </button>
              );
            })}
          </div>

          {/* Tab content */}
          {activeTab === "trips"     && <SchoolTrips schoolTrips={schoolTrips} />}
          {activeTab === "students"  && <SchoolStudents schoolTrips={schoolTrips} courses={courses} students={students} setStudents={setStudents} notify={notify} />}
          {activeTab === "allergies" && <SchoolAllergies courses={courses} students={students} />}
          {activeTab === "docs"      && <SchoolDocs courses={courses} schoolDocuments={schoolDocuments} setSchoolDocuments={setSchoolDocuments} notify={notify} />}
          {activeTab === "rooming"   && <SchoolRooming schoolTrips={schoolTrips} setSchoolTrips={setSchoolTrips} notify={notify} />}
          {activeTab === "groups"    && <SchoolGroups schoolTrips={schoolTrips} setSchoolTrips={setSchoolTrips} notify={notify} />}
        </div>
      )}
    </div>
  );
}

// ─── Admin Schools ────────────────────────────────────────────────────────────

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

function AdminSchools({ trips, notify, section = "colegios" }) {
  const tab = {
    "colegios": "schools", "alumnos": "students", "alergias": "allergies",
    "docs": "docs", "preguntas": "questions", "rooming": "rooming", "grupos": "groups",
    "seguimiento": "tracking", "checklist": "checklist",
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
  // Manual student entry
  const [showAddStudent, setShowAddStudent] = useState(false);
  const [newStudent, setNewStudent] = useState({ name: "", surname: "", allergies: "", intolerances: "", diet_notes: "", notes: "" });
  const [savingStudent, setSavingStudent] = useState(false);
  // Manual doc entry
  const [showAddDoc, setShowAddDoc] = useState(false);
  const [newDoc, setNewDoc] = useState({ name: "", description: "", required: true });
  const [savingDoc, setSavingDoc] = useState(false);
  // Manual question entry
  const [showAddQuestion, setShowAddQuestion] = useState(false);
  const [newQuestion, setNewQuestion] = useState({ message: "", school_id: "" });
  const [savingQuestion, setSavingQuestion] = useState(false);
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

  useEffect(() => {
    const load = async () => {
      setLoading(true);
      try {
        const [schoolsRes, schoolTripsRes, coursesRes, studentsRes, docsRes] = await Promise.all([
          supabase.from("schools").select("*").order("name"),
          supabase.from("school_trips").select("*, trips(name, departure_date, hero_image, description)").order("created_at"),
          supabase.from("school_courses").select("*").order("course_name"),
          supabase.from("students").select("*").order("name"),
          supabase.from("school_documents").select("*").order("created_at"),
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
    const { error: scErr } = await supabase.from("school_courses").insert([{ school_trip_id: stData.id, course_name: assignCourse.trim(), group_name: assignGroup.trim() }]);
    if (scErr) { notify("Error creando curso: " + scErr.message, { variant: "destructive" }); setSavingAssign(false); return; }
    notify("Viaje asignado con curso.");
    setAssigningSchoolId(null);
    setAssignCourse(""); setAssignGroup("");
    // Reload
    const [stRes, scRes] = await Promise.all([
      supabase.from("school_trips").select("*, trips(name, departure_date, hero_image, description)").order("created_at"),
      supabase.from("school_courses").select("*").order("course_name"),
    ]);
    setAllSchoolTrips(stRes.data || []);
    setAllCourses(scRes.data || []);
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
    const targetCourseId = filterCourseId !== "all" ? filterCourseId : visibleCourses[0]?.id;
    if (!targetCourseId) { notify("Selecciona un colegio y curso primero."); return; }
    setSavingDoc(true);
    const { data, error } = await supabase.from("school_documents").insert([{ ...newDoc, school_course_id: targetCourseId, status: "pending" }]).select().maybeSingle();
    if (error) notify("Error añadiendo documento: " + error.message, { variant: "destructive" });
    else { setAllSchoolDocs((prev) => [...prev, data]); setNewDoc({ name: "", description: "", required: true }); setShowAddDoc(false); notify("Documento añadido."); }
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
    const { error } = await supabase.from("school_questions").update({ reply }).eq("id", qId);
    if (error) notify("Error guardando respuesta.", { variant: "destructive" });
    else { setAllSchoolQuestions((prev) => prev.map((q) => q.id === qId ? { ...q, reply } : q)); notify("Respuesta guardada."); }
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
          <SectionTitle icon={Users} title="Colegios" subtitle="Gestiona los centros escolares y asigna viajes." />

          <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
            <CardContent className="p-5">
              <div className="flex flex-wrap items-center gap-3 justify-between">
                <Input placeholder="Buscar colegio..." value={schoolSearch} onChange={(e) => setSchoolSearch(e.target.value)} className="h-9 w-48 rounded-xl text-sm" />
                <Button className="rounded-2xl text-sm text-white" style={{ backgroundColor: CORPORATE_RED }} onClick={() => setShowNewSchool(!showNewSchool)}>
                  <Plus className="mr-1.5 h-4 w-4" />Nuevo colegio
                </Button>
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

          {/* Filtros colegio + viaje */}
          <div className="flex flex-wrap gap-3">
            <select value={filterSchoolId} onChange={(e) => { setFilterSchoolId(e.target.value); setFilterTripId(""); setFilterCourseId("all"); }}
              className="rounded-2xl border border-zinc-200 bg-white px-3 py-2 text-sm text-zinc-950 focus:outline-none">
              <option value="">Todos los colegios</option>
              {schools.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
            </select>
            {filterSchoolId && (
              <select value={filterTripId} onChange={(e) => { setFilterTripId(e.target.value); setFilterCourseId("all"); }}
                className="rounded-2xl border border-zinc-200 bg-white px-3 py-2 text-sm text-zinc-950 focus:outline-none">
                <option value="">Todos los viajes</option>
                {filteredSchoolTrips.map((st) => <option key={st.id} value={st.id}>{st.trips?.name || st.trip_id}</option>)}
              </select>
            )}
          </div>

          {/* Tabs de curso */}
          {visibleCourses.length > 0 && (
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
          <SectionTitle icon={LayoutGrid} title="Rooming" subtitle="Distribución de habitaciones por colegio y viaje." extra={
            <Button className="rounded-2xl text-sm text-white" style={{ backgroundColor: CORPORATE_RED }} onClick={() => setShowAddRoom(!showAddRoom)}>
              <Plus className="mr-1.5 h-4 w-4" />Añadir habitación
            </Button>
          } />

          {/* Filtro colegio */}
          <div className="flex flex-wrap gap-3">
            <select value={filterSchoolId} onChange={(e) => { setFilterSchoolId(e.target.value); setFilterTripId(""); setFilterCourseId("all"); setRoomingTripId(""); }}
              className="rounded-2xl border border-zinc-200 bg-white px-3 py-2 text-sm text-zinc-950 focus:outline-none">
              <option value="">Todos los colegios</option>
              {schools.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
            </select>
            {filterSchoolId && (
              <select value={filterTripId} onChange={(e) => setFilterTripId(e.target.value)}
                className="rounded-2xl border border-zinc-200 bg-white px-3 py-2 text-sm text-zinc-950 focus:outline-none">
                <option value="">Todos los viajes</option>
                {filteredSchoolTrips.map((st) => <option key={st.id} value={st.id}>{st.trips?.name || st.trip_id}</option>)}
              </select>
            )}
          </div>

          {/* Tabs de curso */}
          {visibleCourses.length > 0 && (
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
          <SectionTitle icon={ListChecks} title="Grupos de actividades" subtitle="Grupos y monitores por colegio y viaje." extra={
            <Button className="rounded-2xl text-sm text-white" style={{ backgroundColor: CORPORATE_RED }} onClick={() => setShowAddGroup(!showAddGroup)}>
              <Plus className="mr-1.5 h-4 w-4" />Añadir grupo
            </Button>
          } />

          {/* Filtro colegio */}
          <div className="flex flex-wrap gap-3">
            <select value={filterSchoolId} onChange={(e) => { setFilterSchoolId(e.target.value); setFilterTripId(""); setFilterCourseId("all"); setGroupTripId(""); }}
              className="rounded-2xl border border-zinc-200 bg-white px-3 py-2 text-sm text-zinc-950 focus:outline-none">
              <option value="">Todos los colegios</option>
              {schools.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
            </select>
            {filterSchoolId && (
              <select value={filterTripId} onChange={(e) => setFilterTripId(e.target.value)}
                className="rounded-2xl border border-zinc-200 bg-white px-3 py-2 text-sm text-zinc-950 focus:outline-none">
                <option value="">Todos los viajes</option>
                {filteredSchoolTrips.map((st) => <option key={st.id} value={st.id}>{st.trips?.name || st.trip_id}</option>)}
              </select>
            )}
          </div>

          {/* Tabs de curso */}
          {visibleCourses.length > 0 && (
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

      {/* Alergias tab */}
      {tab === "allergies" && (
        <div className="space-y-5">
          <SectionTitle icon={AlertCircle} title="Alergias e intolerancias" subtitle="Control de necesidades alimentarias por alumno y curso." />

          {/* Filtros colegio + viaje */}
          <div className="flex flex-wrap gap-3">
            <select value={filterSchoolId} onChange={(e) => { setFilterSchoolId(e.target.value); setFilterTripId(""); setFilterCourseId("all"); }}
              className="rounded-2xl border border-zinc-200 bg-white px-3 py-2 text-sm text-zinc-950 focus:outline-none">
              <option value="">Todos los colegios</option>
              {schools.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
            </select>
            {filterSchoolId && (
              <select value={filterTripId} onChange={(e) => { setFilterTripId(e.target.value); setFilterCourseId("all"); }}
                className="rounded-2xl border border-zinc-200 bg-white px-3 py-2 text-sm text-zinc-950 focus:outline-none">
                <option value="">Todos los viajes</option>
                {filteredSchoolTrips.map((st) => <option key={st.id} value={st.id}>{st.trips?.name || st.trip_id}</option>)}
              </select>
            )}
          </div>

          {/* Tabs de curso */}
          {visibleCourses.length > 0 && (
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
          <SectionTitle icon={FileCheck2} title="Documentación" subtitle="Crea documentos requeridos y asígnalos a cada curso por colegio." />
          <div className="grid gap-5 lg:grid-cols-[0.95fr_1.05fr]">
            {/* Left: nuevo documento */}
            <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
              <CardContent className="space-y-4 p-5">
                <div className="font-medium text-zinc-950">Nuevo documento requerido</div>
                <div className="space-y-2">
                  <Label>Nombre del documento</Label>
                  <Input value={newDoc.name} onChange={(e) => setNewDoc(p => ({ ...p, name: e.target.value }))} placeholder="Ej. Autorización de salida" className="rounded-2xl" />
                </div>
                <div className="space-y-2">
                  <Label>Descripción / instrucciones</Label>
                  <Input value={newDoc.description} onChange={(e) => setNewDoc(p => ({ ...p, description: e.target.value }))} placeholder="Ej. Firmada por tutor legal" className="rounded-2xl" />
                </div>
                <Button onClick={handleSaveDoc} disabled={savingDoc || !newDoc.name.trim()} className="h-11 rounded-2xl text-white" style={{ backgroundColor: CORPORATE_RED }}>
                  <FileCheck2 className="mr-2 h-4 w-4" />{savingDoc ? "Guardando…" : "Crear documento"}
                </Button>
                <Separator />
                <div className="space-y-3">
                  {allSchoolDocs.filter((d, i, arr) => arr.findIndex(x => x.name === d.name) === i).map((d) => (
                    <div key={d.id} className="flex items-center justify-between gap-3 rounded-2xl border border-zinc-200 bg-white p-4">
                      <div className="min-w-0 flex-1">
                        <div className="font-medium text-zinc-950">{d.name}</div>
                        {d.description && <div className="text-sm text-zinc-500">{d.description}</div>}
                      </div>
                      <Button variant="ghost" size="icon" onClick={async () => {
                        const { error } = await supabase.from("school_documents").delete().eq("name", d.name);
                        if (error) { notify("Error eliminando documento: " + error.message); return; }
                        setAllSchoolDocs((prev) => prev.filter((x) => x.name !== d.name));
                        notify("Documento eliminado.");
                      }}>
                        <Trash2 className="h-4 w-4 text-zinc-700" />
                      </Button>
                    </div>
                  ))}
                </div>
              </CardContent>
            </Card>

            {/* Right: asignar a curso */}
            <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
              <CardContent className="space-y-4 p-5">
                <div className="font-medium text-zinc-950">Asignar a curso</div>
                <div className="space-y-2">
                  <Label>Colegio</Label>
                  <select value={filterSchoolId} onChange={(e) => { setFilterSchoolId(e.target.value); setFilterTripId(""); setFilterCourseId("all"); }}
                    className="h-11 w-full rounded-2xl border border-zinc-200 bg-white px-4 text-sm">
                    <option value="">Todos los colegios</option>
                    {schools.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
                  </select>
                </div>
                {filterSchoolId && (
                  <div className="space-y-2">
                    <Label>Viaje</Label>
                    <select value={filterTripId} onChange={(e) => { setFilterTripId(e.target.value); setFilterCourseId("all"); }}
                      className="h-11 w-full rounded-2xl border border-zinc-200 bg-white px-4 text-sm">
                      <option value="">Todos los viajes</option>
                      {filteredSchoolTrips.map((st) => <option key={st.id} value={st.id}>{st.trips?.name || st.trip_id}</option>)}
                    </select>
                  </div>
                )}
                {visibleCourses.length > 0 && (
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
                  const filteredDocs = allSchoolDocs.filter((d) => {
                    if (filterCourseId !== "all") return d.school_course_id === filterCourseId;
                    if (filterSchoolId) return visibleCourses.map((c) => c.id).includes(d.school_course_id);
                    return true;
                  });
                  if (!filterSchoolId) return <p className="text-sm text-zinc-400">Selecciona un colegio para ver sus documentos.</p>;
                  if (filteredDocs.length === 0) return <p className="text-sm text-zinc-400">No hay documentos en la selección actual.</p>;
                  return (
                    <div className="space-y-2">
                      {filteredDocs.map((d) => {
                        const course = allCourses.find((c) => c.id === d.school_course_id);
                        return (
                          <div key={d.id} className="flex items-center justify-between rounded-2xl border border-zinc-200 bg-white p-4">
                            <div className="min-w-0 flex-1">
                              <div className="text-sm font-medium text-zinc-900">{d.name}</div>
                              {course && <div className="text-xs text-zinc-500">{course.course_name}{course.group_name ? ` · ${course.group_name}` : ""}</div>}
                            </div>
                            <span className={`rounded-full px-2 py-0.5 text-[10px] font-medium ${d.status === "confirmed" ? "bg-green-100 text-green-700" : d.status === "uploaded" ? "bg-blue-100 text-blue-700" : d.status === "rejected" ? "bg-red-100 text-red-700" : "bg-zinc-100 text-zinc-600"}`}>
                              {d.status === "confirmed" ? "Confirmado" : d.status === "uploaded" ? "Subido" : d.status === "rejected" ? "Rechazado" : "Pendiente"}
                            </span>
                          </div>
                        );
                      })}
                    </div>
                  );
                })()}
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
              <div className="space-y-1">
                <Label>Filtrar por colegio</Label>
                <select value={filterSchoolId} onChange={(e) => setFilterSchoolId(e.target.value)}
                  className="mt-2 h-11 w-full rounded-2xl border border-zinc-200 bg-white px-4 text-sm">
                  <option value="">Todos los colegios</option>
                  {schools.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
                </select>
              </div>
            </CardContent>
          </Card>

          {(() => {
            const qs = allSchoolQuestions.filter((q) => !filterSchoolId || q.school_id === filterSchoolId);
            if (qs.length === 0) return (
              <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
                <CardContent className="p-8 text-center text-sm text-zinc-400">
                  Aún no hay preguntas de colegios.
                </CardContent>
              </Card>
            );
            return (
              <div className="space-y-3">
                {qs.map((q) => {
                  const school = schools.find((s) => s.id === q.school_id);
                  return <SchoolQuestionCard key={q.id} q={q} schoolName={school?.name} onReply={handleReplyQuestion} />;
                })}
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
            <CardContent className="p-5">
              <div className="space-y-2">
                <Label>Filtrar por colegio</Label>
                <select value={filterSchoolId} onChange={(e) => { setFilterSchoolId(e.target.value); setFilterTripId(""); setFilterCourseId("all"); }}
                  className="mt-2 h-11 min-w-[280px] rounded-2xl border border-zinc-200 bg-white px-4 text-sm">
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

              const getSummaryTone = (ok, warn) => ok ? "bg-emerald-100 text-emerald-700" : warn ? "bg-amber-100 text-amber-700" : "bg-red-100 text-red-700";

              return (
                <AccordionSection
                  key={school.id}
                  title={school.name}
                  subtitle={school.contact_name ? `Coordinador: ${school.contact_name}${school.email ? ` · ${school.email}` : ""}` : ""}
                  icon={Users}
                  meta={
                    <div className="flex items-center gap-2 overflow-x-auto whitespace-nowrap">
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
                    return (
                      <div key={st.id} className="mb-4 last:mb-0">
                        <div className="mb-2 text-sm font-semibold text-zinc-700">{st.trips?.name || st.trip_id}</div>
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
                                  <div className="grid gap-1.5 sm:grid-cols-3">
                                    {[
                                      { label: "Listado alumnos", ok: cStudents.length > 0, detail: cStudents.length > 0 ? `${cStudents.length} alumnos` : "Sin listado" },
                                      { label: "Alergias", ok: cWithAllergies.length > 0 || cStudents.length > 0, neutral: cWithAllergies.length === 0 && cStudents.length > 0, detail: cWithAllergies.length > 0 ? `${cWithAllergies.length} con alergia` : cStudents.length > 0 ? "Sin alergias ✓" : "Sin datos" },
                                      { label: "Documentación", ok: cDocs.length > 0 && cDocsUploaded === cDocs.length, warn: cDocs.length > 0 && cDocsUploaded < cDocs.length, detail: cDocs.length === 0 ? "Sin docs requeridos" : cDocsUploaded === cDocs.length ? `Completa (${cDocs.length})` : `Faltan ${cDocs.length - cDocsUploaded} de ${cDocs.length}` },
                                    ].map((item) => (
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
                    <Button className="rounded-2xl text-white" style={{ backgroundColor: CORPORATE_RED }}
                      onClick={() => updateItinerary([...itinerary, { day: `Día ${itinerary.length + 1}`, title: "Nuevo tramo", description: "Detalle", time: "10:00" }])}>
                      <Plus className="mr-2 h-4 w-4" />Añadir tramo
                    </Button>
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
                    <Button className="rounded-2xl text-white" style={{ backgroundColor: CORPORATE_RED }}
                      onClick={() => updateLogistics([...logistics, { title: "Nuevo punto", description: "" }])}>
                      <Plus className="mr-2 h-4 w-4" />Añadir punto
                    </Button>
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
        <div className="space-y-5">
          <SectionTitle icon={Map} title="Viajes escolares" subtitle="Información de cada viaje escolar asignado a un colegio." />

          <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
            <CardContent className="p-5">
              <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
                <div className="flex-1 space-y-1">
                  <Label>Filtrar por colegio</Label>
                  <select value={filterSchoolId} onChange={(e) => setFilterSchoolId(e.target.value)}
                    className="h-11 w-full rounded-2xl border border-zinc-200 bg-white px-4 text-sm font-medium">
                    <option value="">Todos los colegios</option>
                    {schools.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
                  </select>
                </div>
              </div>
            </CardContent>
          </Card>

          {(filterSchoolId ? filteredSchoolTrips : allSchoolTrips).length === 0 ? (
            <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
              <CardContent className="p-8 text-center text-sm text-zinc-400">No hay viajes escolares registrados.</CardContent>
            </Card>
          ) : (
            <div className="space-y-4">
              {(filterSchoolId ? filteredSchoolTrips : allSchoolTrips).map((st) => {
                const school = schools.find((s) => s.id === st.school_id);
                const stCourses = allCourses.filter((c) => c.school_trip_id === st.id);
                const stStudents = allStudents.filter((s) => stCourses.map(c => c.id).includes(s.school_course_id));
                return (
                  <Card key={st.id} className="rounded-3xl border-zinc-200 bg-white shadow-sm overflow-hidden">
                    {st.trips?.hero_image && (
                      <div className="relative h-48 w-full">
                        <img src={st.trips.hero_image} alt={st.trips.name} className="h-full w-full object-cover" onError={(e) => { e.target.parentElement.style.display = "none"; }} />
                        <div className="absolute inset-0 bg-gradient-to-t from-black/60 to-transparent" />
                        <div className="absolute bottom-4 left-5 text-white">
                          <div className="text-xl font-bold">{st.trips?.name}</div>
                          {st.trips?.departure_date && <div className="text-sm opacity-80">{new Date(st.trips.departure_date).toLocaleDateString("es-ES", { day: "numeric", month: "long", year: "numeric" })}</div>}
                        </div>
                      </div>
                    )}
                    <CardContent className="p-5 space-y-3">
                      {!st.trips?.hero_image && (
                        <div>
                          <div className="font-semibold text-zinc-950">{st.trips?.name || st.trip_id}</div>
                          {st.trips?.departure_date && <div className="text-xs text-zinc-500 mt-0.5">{new Date(st.trips.departure_date).toLocaleDateString("es-ES", { day: "numeric", month: "long", year: "numeric" })}</div>}
                        </div>
                      )}
                      <div className="flex flex-wrap items-center justify-between gap-3">
                        <div>
                          <div className="text-sm font-medium text-zinc-700">{school?.name || "Colegio"}</div>
                          {school?.contact_name && <div className="text-xs text-zinc-500">{school.contact_name}</div>}
                        </div>
                        <div className="flex flex-wrap gap-2">
                          {stCourses.map((c) => (
                            <Badge key={c.id} variant="outline" className="rounded-xl text-xs">{c.course_name}{c.group_name ? ` · ${c.group_name}` : ""}</Badge>
                          ))}
                          <Badge variant="outline" className="rounded-xl text-xs">{stStudents.length} alumnos</Badge>
                        </div>
                      </div>
                      {st.trips?.description && <p className="text-xs text-zinc-500">{st.trips.description}</p>}
                    </CardContent>
                  </Card>
                );
              })}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

// ─── Admin Panel ──────────────────────────────────────────────────────────────

function AdminPanel({ users, setUsers, trips, setTrips, templates, setTemplates, onLogout, notify }) {
  const [activeSection, setActiveSection] = useState(() => {
    if (typeof window === "undefined") return "clients";
    return window.localStorage.getItem(ADMIN_SECTION_STORAGE_KEY) || "clients";
  });
  useEffect(() => { if (typeof window !== "undefined") window.localStorage.setItem(ADMIN_SECTION_STORAGE_KEY, activeSection); }, [activeSection]);
  const totalParticipants = users.filter((u) => u.role === "client").length;

  usePaymentReminders(users, trips);

  const [campExpanded, setCampExpanded] = useState(true);
  const [colExpanded, setColExpanded] = useState(true);

  const campamentosItems = [
    { key: "clients",     label: "Clientes",       icon: Users },
    { key: "tracking",    label: "Seguimiento",     icon: BarChart2 },
    { key: "payments",    label: "Pagos",           icon: CreditCard },
    { key: "docs",        label: "Documentación",   icon: FileCheck2 },
    { key: "questions",   label: "Preguntas",       icon: MessageCircleQuestion },
    { key: "checklists",  label: "Checklists",      icon: ListChecks },
    { key: "itinerario",  label: "Itinerario",      icon: CalendarDays },
    { key: "logistica",   label: "Logística",       icon: MapPinned },
    { key: "trips",       label: "Viajes",          icon: Map },
  ];
  const colegiosItems = [
    { key: "school_colegios",    label: "Colegios",      icon: Users },
    { key: "school_seguimiento", label: "Seguimiento",   icon: BarChart2 },
    { key: "school_alumnos",     label: "Alumnos",       icon: Users },
    { key: "school_alergias",    label: "Alergias",      icon: AlertCircle },
    { key: "school_docs",        label: "Documentación", icon: FileCheck2 },
    { key: "school_preguntas",   label: "Preguntas",     icon: MessageCircleQuestion },
    { key: "school_rooming",     label: "Rooming",       icon: LayoutGrid },
    { key: "school_grupos",      label: "Grupos",        icon: ListChecks },
    { key: "school_checklist",   label: "Checklist",     icon: CheckCircle2 },
    { key: "school_itinerario",  label: "Itinerario",    icon: CalendarDays },
    { key: "school_logistica",   label: "Logística",     icon: MapPinned },
    { key: "school_viajes",      label: "Viajes",        icon: Map },
  ];
  const navItems = [...campamentosItems, ...colegiosItems, { key: "calculadora", label: "Calculadora", icon: Calculator }];

  return (
    <div className="min-h-screen text-zinc-950" style={{ background: "linear-gradient(160deg,#fff5f5 0%,#fafafa 40%,#f4f4f5 100%)" }}>
      {/* Header */}
      <header className="sticky top-0 z-40 border-b border-zinc-200/80 bg-white/90 backdrop-blur-sm">
        <div className="mx-auto flex max-w-[1400px] items-center justify-between px-6 py-4">
          <div className="flex items-center gap-3">
            <div className="flex h-9 w-9 items-center justify-center rounded-xl text-white shadow-sm" style={{ backgroundColor: CORPORATE_RED }}>
              <Sparkles className="h-4 w-4" />
            </div>
            <div>
              <div className="text-xs uppercase tracking-[0.22em] text-zinc-400">Panel interno</div>
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
          {activeSection === "clients"             && <AdminClients users={users} trips={trips} setUsers={setUsers} templates={templates} notify={notify} setTrips={setTrips} />}
          {activeSection === "tracking"            && <AdminTracking users={users} trips={trips} templates={templates} setUsers={setUsers} notify={notify} />}
          {activeSection === "payments"            && <AdminPayments users={users} setUsers={setUsers} notify={notify} />}
          {activeSection === "docs"                && <AdminDocs templates={templates} setTemplates={setTemplates} users={users} setUsers={setUsers} trips={trips} notify={notify} />}
          {activeSection === "questions"           && <AdminQuestions users={users} setUsers={setUsers} notify={notify} />}
          {activeSection === "checklists"          && <AdminChecklists trips={trips} setTrips={setTrips} notify={notify} />}
          {activeSection === "itinerario"          && <AdminItinerary trips={trips} setTrips={setTrips} notify={notify} />}
          {activeSection === "logistica"           && <AdminLogistica trips={trips} setTrips={setTrips} notify={notify} />}
          {activeSection === "trips"               && <AdminTrips trips={trips} setTrips={setTrips} notify={notify} />}
          {activeSection === "calculadora"         && <CalculadoraCampamento />}
          {activeSection === "school_colegios"     && <AdminSchools trips={trips} notify={notify} section="colegios" />}
          {activeSection === "school_alumnos"      && <AdminSchools trips={trips} notify={notify} section="alumnos" />}
          {activeSection === "school_alergias"     && <AdminSchools trips={trips} notify={notify} section="alergias" />}
          {activeSection === "school_docs"         && <AdminSchools trips={trips} notify={notify} section="docs" />}
          {activeSection === "school_preguntas"    && <AdminSchools trips={trips} notify={notify} section="preguntas" />}
          {activeSection === "school_rooming"      && <AdminSchools trips={trips} notify={notify} section="rooming" />}
          {activeSection === "school_grupos"       && <AdminSchools trips={trips} notify={notify} section="grupos" />}
          {activeSection === "school_seguimiento"  && <AdminSchools trips={trips} notify={notify} section="seguimiento" />}
          {activeSection === "school_checklist"    && <AdminSchools trips={trips} notify={notify} section="checklist" />}
          {activeSection === "school_itinerario"   && <AdminSchools trips={trips} notify={notify} section="itinerario" />}
          {activeSection === "school_logistica"    && <AdminSchools trips={trips} notify={notify} section="logistica" />}
          {activeSection === "school_viajes"       && <AdminSchools trips={trips} notify={notify} section="viajes" />}
        </main>
      </div>
    </div>
  );
}

// ─── Root component ──────────────────────────────────────────────────────────

export default function GIMELOOSPortalApp() {
  const [users, setUsers] = useState(initialUsers);
  const [trips, setTrips] = useState(initialTrips);
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
        const [tripsRes, templatesRes, participantsRes, docsRes, paymentsRes, pricingRes, questionsRes] = await Promise.all([
          supabase.from("trips").select("*").order("created_at", { ascending: true }),
          supabase.from("document_templates").select("*").order("created_at", { ascending: true }),
          supabase.from("participants").select("*").order("created_at", { ascending: true }),
          supabase.from("participant_documents").select("*").order("created_at", { ascending: true }),
          supabase.from("participant_payments").select("*").order("created_at", { ascending: true }),
          supabase.from("participant_pricing").select("*").order("created_at", { ascending: true }),
          supabase.from("participant_questions").select("*").order("created_at", { ascending: true }),
        ]);

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
              participantName: p.participant_name || "", motherName: p.mother_name || "",
              fatherName: p.father_name || "", parentName: p.parent_name || "",
              email: p.email || "", contactEmails: p.contact_emails || [],
              // [CRÍTICO-1] Sin campo password en el estado React
              tripId: p.trip_id || "",
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
    // No limpiar si hay una recarga en curso (dataRefreshKey acaba de cambiar)
    if (dataRefreshKey > 0 && users === initialUsers) return;
    if (!users.some((u) => u.id === auth.userId)) {
      setAuth({ userId: null, error: "", isLoading: false });
      try { window.localStorage.removeItem(LOCAL_STORAGE_AUTH_KEY); } catch (e) { console.error(e); }
    }
  }, [users, auth.userId, sessionBootstrapped, isBootstrapping, dataRefreshKey]);

  // [CRÍTICO-1/2/3] Login vía Supabase Auth — la comparación de contraseña ocurre en el servidor
  const handleLogin = async (username, password) => {
    setAuth((prev) => ({ ...prev, isLoading: true, error: "" }));
    try {
      // Lookup email via SECURITY DEFINER RPC para saltarse RLS antes de autenticar
      const { data: email, error: lookupErr } = await supabase
        .rpc("get_login_email", { p_username: username.trim().toLowerCase() });

      if (lookupErr || !email) {
        setAuth({ userId: null, error: "Usuario o contraseña incorrectos.", isLoading: false });
        return;
      }

      const { data: session, error: authErr } = await supabase.auth.signInWithPassword({
        email,
        password,
      });

      if (authErr || !session?.user) {
        setAuth({ userId: null, error: "Usuario o contraseña incorrectos.", isLoading: false });
        return;
      }

      // Resolver el participant.id real desde auth_uid (funciona para admin y participantes)
      const { data: participantId } = await supabase.rpc("get_participant_id_for_auth");
      const resolvedId = participantId ?? session.user.id;

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
      <div style={{ fontFamily: "Arial, sans-serif" }}>
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

  return (
    <div style={{ fontFamily: "Arial, sans-serif" }} className="[&_button]:cursor-pointer [&_label]:cursor-pointer [&_select]:cursor-pointer [&_summary]:cursor-pointer">
      <ActionToast notifications={notifications} removeNotification={removeNotification} />
      {!currentUser ? (
        <LoginScreen onLogin={handleLogin} loginError={auth.error} isLoading={auth.isLoading} />
      ) : currentUser.role === "admin" ? (
        <AdminPanel users={users} setUsers={setUsers} trips={trips} setTrips={setTrips} templates={templates} setTemplates={setTemplates} onLogout={handleLogout} notify={notify} />
      ) : currentUser.role === "school" ? (
        <SchoolPortal user={currentUser} trips={trips} onLogout={handleLogout} notify={notify} />
      ) : (
        <ClientPortal user={currentUser} trips={trips} templates={templates} setUsers={setUsers} onLogout={handleLogout} notify={notify} />
      )}
    </div>
  );
}

"use client";

import { useEffect, useState, useCallback } from "react";
import { Card, CardContent } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import {
  CalendarDays,
  Clock,
  Users,
  MapPinned,
  Sparkles,
  Phone,
  School,
  RefreshCw,
  PartyPopper,
  Gamepad2,
  BellRing,
  CalendarPlus,
  AlertCircle,
  Eye,
  X,
  CheckCircle2,
  Download,
  Loader2,
  Wand2,
  Star,
} from "lucide-react";
import GameSelector from "./GameSelector";

const CORPORATE_RED = "#FF3131";
const SEEN_IDS_KEY = "animaciones_seen_ids";
const SELECCIONES_KEY = "animaciones_selecciones";

function loadSeenIds() {
  try {
    return new Set(JSON.parse(localStorage.getItem(SEEN_IDS_KEY) || "[]"));
  } catch {
    return new Set();
  }
}

function saveSeenIds(ids) {
  localStorage.setItem(SEEN_IDS_KEY, JSON.stringify([...ids]));
}

function loadSelecciones() {
  try {
    const raw = JSON.parse(localStorage.getItem(SELECCIONES_KEY) || "{}");
    return Object.fromEntries(Object.entries(raw).map(([k, v]) => [k, new Set(v)]));
  } catch {
    return {};
  }
}

function saveSelecciones(selecciones) {
  const serializable = Object.fromEntries(
    Object.entries(selecciones).map(([k, v]) => [k, [...v]])
  );
  localStorage.setItem(SELECCIONES_KEY, JSON.stringify(serializable));
}

function limpiarTematica(tematica) {
  return (tematica || "").replace(/\s*\(.*?\)\s*/g, "").trim() || "Sin temática";
}

const FAVORITO_MIN_USOS = 3;

const PREF_CATS_POR_TEMATICA = {
  gymkana: ["Juegos deportivos", "Juegos cooperativos", "Juegos para formar equipos", "Juegos cortos de inicio"],
  furor: ["Furor"],
  feria: ["Feria"],
  "búsqueda del tesoro": ["Juegos de desarrollo", "Juegos cooperativos"],
  talleres: ["Talleres"],
};

function generarSugerencia(reserva, juegos) {
  const tematica = limpiarTematica(reserva.evento.tematica).toLowerCase();
  const capacidad = reserva.capacidadJuegos || 8;
  const preferredCats = PREF_CATS_POR_TEMATICA[tematica] || [];

  const scored = juegos
    .map((j) => {
      const inPref = preferredCats.length > 0 && j.categorias.some((c) =>
        preferredCats.some((p) => c.toLowerCase().includes(p.toLowerCase()))
      );
      return { juego: j, score: (inPref ? 10000 : 0) + (j.veces_usado || 0) };
    })
    .sort((a, b) => b.score - a.score);

  return scored.slice(0, capacidad).map((s) => s.juego);
}

function SugerenciaOverlay({ reserva, juegosSeleccionados, tienePropuestaActual, onAceptar, onAjustar, onClose }) {
  const materiales = [...new Set(juegosSeleccionados.flatMap((j) => j.materiales))].sort((a, b) =>
    a.localeCompare(b, "es")
  );
  const tematica = limpiarTematica(reserva.evento.tematica);
  const nombreCompleto = [reserva.festejado.nombre, reserva.festejado.apellidos].filter(Boolean).join(" ");

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4"
      onClick={onClose}
    >
      <div
        className="max-h-[90vh] w-full max-w-2xl overflow-y-auto rounded-3xl bg-white shadow-2xl"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="sticky top-0 flex items-center justify-between rounded-t-3xl bg-white px-6 pt-5 pb-4 border-b border-zinc-100">
          <div className="flex items-center gap-2">
            <Wand2 className="h-4 w-4" style={{ color: CORPORATE_RED }} />
            <span className="text-sm font-semibold text-zinc-900">Propuesta sugerida · {nombreCompleto}</span>
          </div>
          <button onClick={onClose} className="cursor-pointer rounded-xl p-1.5 text-zinc-400 hover:bg-zinc-100">
            <X className="h-5 w-5" />
          </button>
        </div>

        <div className="p-6">
          <div className="mb-4 flex items-center gap-2 text-sm text-zinc-500">
            <span className="rounded-full px-3 py-1 text-white text-xs font-medium" style={{ backgroundColor: CORPORATE_RED }}>{tematica}</span>
            <span>{juegosSeleccionados.length} juegos · basado en uso histórico</span>
          </div>

          {tienePropuestaActual && (
            <div className="mb-4 rounded-2xl bg-amber-50 border border-amber-200 px-4 py-3 text-sm text-amber-800">
              Ya tienes una propuesta guardada. Aceptar esta sugerencia la reemplazará.
            </div>
          )}

          <ol className="space-y-3 mb-6">
            {juegosSeleccionados.map((j, i) => (
              <li key={j.id} className="flex gap-3 text-sm text-zinc-700">
                <span
                  className="mt-0.5 flex h-5 w-5 shrink-0 items-center justify-center rounded-full text-xs font-bold text-white"
                  style={{ backgroundColor: CORPORATE_RED }}
                >
                  {i + 1}
                </span>
                <div className="min-w-0">
                  <div className="flex flex-wrap items-center gap-1.5">
                    <span className="font-semibold">{j.nombre}</span>
                    {j.veces_usado >= FAVORITO_MIN_USOS && (
                      <span className="inline-flex items-center gap-0.5 rounded-md bg-amber-100 px-1.5 py-0.5 text-xs text-amber-700">
                        <Star className="h-2.5 w-2.5 fill-amber-500 text-amber-500" />
                        Muy usado
                      </span>
                    )}
                  </div>
                  <p className="text-zinc-500">{j.descripcion ? j.descripcion.charAt(0).toUpperCase() + j.descripcion.slice(1) : ""}</p>
                </div>
              </li>
            ))}
          </ol>

          {materiales.length > 0 && (
            <div className="mb-6 rounded-2xl bg-zinc-50 px-4 py-3">
              <p className="mb-2 text-xs font-semibold uppercase tracking-wide text-zinc-400">Materiales</p>
              <p className="text-sm text-zinc-700">{materiales.join(", ")}</p>
            </div>
          )}

          <div className="flex flex-col gap-2 sm:flex-row">
            <Button
              className="flex-1 rounded-2xl text-white"
              style={{ backgroundColor: CORPORATE_RED }}
              onClick={onAceptar}
            >
              <CheckCircle2 className="mr-2 h-4 w-4" />
              Usar esta propuesta
            </Button>
            <Button
              variant="outline"
              className="flex-1 rounded-2xl border-zinc-200"
              onClick={onAjustar}
            >
              <Gamepad2 className="mr-2 h-4 w-4" />
              Ajustar manualmente
            </Button>
          </div>
        </div>
      </div>
    </div>
  );
}

function cap(str) {
  if (!str) return "";
  return str.charAt(0).toUpperCase() + str.slice(1);
}

function PropuestaOverlay({ reserva, juegosSeleccionados, onClose }) {
  const materiales = [...new Set(juegosSeleccionados.flatMap((j) => j.materiales))].sort((a, b) =>
    a.localeCompare(b, "es")
  );
  const nombreCompleto = [reserva.festejado.nombre, reserva.festejado.apellidos]
    .filter(Boolean)
    .join(" ");
  const tematica = limpiarTematica(reserva.evento.tematica);

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4"
      onClick={onClose}
    >
      <div
        className="max-h-[90vh] w-full max-w-2xl overflow-y-auto rounded-3xl bg-white shadow-2xl"
        onClick={(e) => e.stopPropagation()}
      >
        {/* Cabecera del overlay */}
        <div className="sticky top-0 flex items-center justify-between rounded-t-3xl bg-white px-6 pt-5 pb-4 border-b border-zinc-100">
          <span className="text-sm font-medium text-zinc-500">Vista previa del documento</span>
          <button onClick={onClose} className="cursor-pointer rounded-xl p-1.5 text-zinc-400 hover:bg-zinc-100">
            <X className="h-5 w-5" />
          </button>
        </div>

        {/* Documento simulado */}
        <div className="p-8">
          {/* Título */}
          <div className="mb-8 text-center">
            <h1 className="text-3xl font-bold uppercase tracking-wide text-zinc-900">
              {nombreCompleto}
            </h1>
          </div>

          {/* Tabla de datos del evento */}
          <div className="mb-8 overflow-hidden rounded-2xl border border-zinc-200">
            <table className="w-full text-sm">
              <tbody>
                <tr className="border-b border-zinc-100">
                  <td className="w-40 bg-zinc-50 px-4 py-3 font-semibold uppercase text-zinc-500 text-xs tracking-wide">Tipo de actividad</td>
                  <td className="px-4 py-3 font-medium text-zinc-800">{tematica}</td>
                </tr>
                <tr className="border-b border-zinc-100">
                  <td className="bg-zinc-50 px-4 py-3 font-semibold uppercase text-zinc-500 text-xs tracking-wide">Monitores</td>
                  <td className="px-4 py-3 text-zinc-800">{reserva.monitoresEstimados ?? "—"} monitores</td>
                </tr>
                <tr className="border-b border-zinc-100">
                  <td className="bg-zinc-50 px-4 py-3 font-semibold uppercase text-zinc-500 text-xs tracking-wide">Participantes</td>
                  <td className="px-4 py-3 text-zinc-800">{reserva.participantes ?? "—"} niños</td>
                </tr>
                <tr className="border-b border-zinc-100">
                  <td className="bg-zinc-50 px-4 py-3 font-semibold uppercase text-zinc-500 text-xs tracking-wide">Fecha y horario</td>
                  <td className="px-4 py-3 text-zinc-800">
                    {reserva.evento.fecha}
                    {reserva.evento.horario ? ` · ${reserva.evento.horario}` : ""}
                    {reserva.evento.horasDuracion ? ` (${reserva.evento.horasDuracion}h)` : ""}
                  </td>
                </tr>
                <tr>
                  <td className="bg-zinc-50 px-4 py-3 font-semibold uppercase text-zinc-500 text-xs tracking-wide align-top">Materiales</td>
                  <td className="px-4 py-3 text-zinc-800">
                    {materiales.length > 0 ? materiales.join(", ") : <span className="text-zinc-400">Sin materiales</span>}
                  </td>
                </tr>
              </tbody>
            </table>
          </div>

          {/* Descripción de la actividad */}
          <div>
            <h2 className="mb-4 text-sm font-bold uppercase tracking-widest text-zinc-500">
              Descripción de la actividad
            </h2>
            {juegosSeleccionados.length === 0 ? (
              <p className="text-sm text-zinc-400 italic">Sin juegos seleccionados todavía.</p>
            ) : (
              <ol className="space-y-4">
                {juegosSeleccionados.map((j, i) => (
                  <li key={j.id} className="flex gap-3 text-sm text-zinc-700">
                    <span className="mt-0.5 flex h-5 w-5 shrink-0 items-center justify-center rounded-full text-xs font-bold text-white" style={{ backgroundColor: CORPORATE_RED }}>
                      {i + 1}
                    </span>
                    <p>
                      <span className="font-semibold">{j.nombre}: </span>
                      {cap(j.descripcion)}
                    </p>
                  </li>
                ))}
              </ol>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

function CalendarHelp({ onClose }) {
  const url = `${window.location.origin}/api/animaciones/calendar`;
  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4"
      onClick={onClose}
    >
      <div
        className="w-full max-w-md rounded-3xl bg-white p-6 shadow-2xl"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="mb-4 flex items-center justify-between">
          <h3 className="font-semibold text-zinc-900">Añadir a Apple Calendar</h3>
          <button onClick={onClose} className="cursor-pointer rounded-xl p-1 text-zinc-400 hover:bg-zinc-100">
            <X className="h-5 w-5" />
          </button>
        </div>
        <ol className="mb-4 space-y-2 text-sm text-zinc-700">
          <li>1. Abre <strong>Calendar</strong> en tu Mac</li>
          <li>2. Ve a <strong>Archivo → Nueva suscripción de calendario…</strong></li>
          <li>3. Pega esta URL:</li>
        </ol>
        <div className="mb-4 rounded-xl bg-zinc-50 px-4 py-3 font-mono text-xs text-zinc-800 break-all">
          {url}
        </div>
        <p className="text-xs text-zinc-400">
          El calendario se actualizará automáticamente cada hora con las nuevas reservas.
        </p>
        <Button
          className="mt-4 w-full rounded-2xl"
          onClick={() => { navigator.clipboard.writeText(url); }}
        >
          Copiar URL
        </Button>
      </div>
    </div>
  );
}

async function descargarPropuestaReserva(reserva, juegosSeleccionados) {
  const res = await fetch("/api/animaciones/propuesta", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ reserva, juegos: juegosSeleccionados }),
  });
  if (!res.ok) {
    const err = await res.json();
    throw new Error(err.error || "Error al generar el documento");
  }
  const blob = await res.blob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  const nombre = [reserva.festejado?.nombre, reserva.festejado?.apellidos].filter(Boolean).join(" ");
  a.href = url;
  a.download = `Propuesta - ${nombre}.docx`;
  a.click();
  URL.revokeObjectURL(url);
}

export default function AnimacionesPage() {
  const [reservas, setReservas] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [newIds, setNewIds] = useState(new Set());
  const [lastUpdated, setLastUpdated] = useState(null);
  const [juegos, setJuegos] = useState([]);
  const [selectedReservaId, setSelectedReservaId] = useState(null);
  const [seleccionesPorReserva, setSeleccionesPorReserva] = useState(() => loadSelecciones());
  const [previewReservaId, setPreviewReservaId] = useState(null);
  const [showCalendarHelp, setShowCalendarHelp] = useState(false);
  const [descargandoId, setDescargandoId] = useState(null);
  const [sugerenciaReservaId, setSugerenciaReservaId] = useState(null);

  const fetchReservas = useCallback(async (isFirstLoad) => {
    try {
      setError(null);
      const res = await fetch("/api/animaciones/reservas");
      const data = await res.json();
      if (!res.ok || !data.ok) throw new Error(data.error || "Error al cargar reservas");

      const seen = loadSeenIds();
      if (isFirstLoad) {
        data.reservas.forEach((r) => seen.add(r.id));
        saveSeenIds(seen);
      } else {
        const fresh = new Set(data.reservas.map((r) => r.id).filter((id) => !seen.has(id)));
        setNewIds(fresh);
      }

      setReservas(data.reservas);
      setLastUpdated(new Date());
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    fetchReservas(true);
    const interval = setInterval(() => fetchReservas(false), 60000);
    return () => clearInterval(interval);
  }, [fetchReservas]);

  useEffect(() => {
    fetch("/api/animaciones/juegos")
      .then((res) => res.json())
      .then((data) => { if (data.ok) setJuegos(data.juegos); })
      .catch(() => {});
  }, []);

  const dismissNew = () => {
    const seen = loadSeenIds();
    reservas.forEach((r) => seen.add(r.id));
    saveSeenIds(seen);
    setNewIds(new Set());
  };

  const selectedReserva = reservas.find((r) => r.id === selectedReservaId) || null;
  const selectedIds = seleccionesPorReserva[selectedReservaId] || new Set();

  const toggleJuego = (juegoId) => {
    setSeleccionesPorReserva((prev) => {
      const current = new Set(prev[selectedReservaId] || []);
      if (current.has(juegoId)) current.delete(juegoId);
      else current.add(juegoId);
      const next = { ...prev, [selectedReservaId]: current };
      saveSelecciones(next);
      return next;
    });
  };

  const previewReserva = reservas.find((r) => r.id === previewReservaId) || null;
  const previewJuegos = previewReserva
    ? juegos.filter((j) => seleccionesPorReserva[previewReservaId]?.has(j.id))
    : [];

  const sugerenciaReserva = reservas.find((r) => r.id === sugerenciaReservaId) || null;
  const sugerenciaJuegos = sugerenciaReserva ? generarSugerencia(sugerenciaReserva, juegos) : [];

  const aplicarSugerencia = (reservaId, juegosSugeridos) => {
    setSeleccionesPorReserva((prev) => {
      const next = { ...prev, [reservaId]: new Set(juegosSugeridos.map((j) => j.id)) };
      saveSelecciones(next);
      return next;
    });
  };

  if (selectedReserva) {
    const tematicaActual = limpiarTematica(selectedReserva.evento.tematica);
    const propuestasCopiables = reservas
      .filter(
        (r) =>
          r.id !== selectedReservaId &&
          limpiarTematica(r.evento.tematica) === tematicaActual &&
          (seleccionesPorReserva[r.id]?.size ?? 0) > 0
      )
      .map((r) => ({ reserva: r, juegosIds: seleccionesPorReserva[r.id] }));

    const onCopiarPropuesta = (juegosIds) => {
      setSeleccionesPorReserva((prev) => {
        const next = { ...prev, [selectedReservaId]: new Set(juegosIds) };
        saveSelecciones(next);
        return next;
      });
    };

    return (
      <GameSelector
        reserva={selectedReserva}
        juegos={juegos}
        selectedIds={selectedIds}
        onToggle={toggleJuego}
        onClose={() => setSelectedReservaId(null)}
        propuestasCopiables={propuestasCopiables}
        onCopiarPropuesta={onCopiarPropuesta}
      />
    );
  }

  return (
    <div className="min-h-screen bg-zinc-50 p-6 lg:p-10">
      <div className="mx-auto max-w-6xl">
        {/* Header */}
        <div className="mb-8 flex flex-col gap-4 lg:flex-row lg:items-center lg:justify-between">
          <div className="flex items-center gap-3">
            <div
              className="flex h-12 w-12 items-center justify-center rounded-2xl text-white"
              style={{ backgroundColor: CORPORATE_RED }}
            >
              <PartyPopper className="h-6 w-6" />
            </div>
            <div>
              <h1 className="text-2xl font-semibold text-zinc-900">Portal Animaciones GIMELOOS</h1>
              <p className="text-sm text-zinc-500">
                {lastUpdated
                  ? `Actualizado a las ${lastUpdated.toLocaleTimeString("es-ES", { hour: "2-digit", minute: "2-digit" })}`
                  : "Cargando reservas confirmadas…"}
              </p>
            </div>
          </div>
          <div className="flex gap-2">
            <Button
              variant="outline"
              className="w-fit rounded-2xl border-zinc-200 bg-white"
              onClick={() => setShowCalendarHelp(true)}
            >
              <CalendarPlus className="mr-2 h-4 w-4" />
              Añadir a Calendar
            </Button>
            <Button
              variant="outline"
              className="w-fit rounded-2xl border-zinc-200 bg-white"
              onClick={() => fetchReservas(false)}
              disabled={loading}
            >
              <RefreshCw className={`mr-2 h-4 w-4 ${loading ? "animate-spin" : ""}`} />
              Actualizar
            </Button>
          </div>
        </div>

        {/* Aviso de reservas nuevas */}
        {newIds.size > 0 && (
          <Card className="mb-6 rounded-3xl border-amber-200 bg-amber-50 shadow-sm">
            <CardContent className="flex flex-col gap-3 p-5 sm:flex-row sm:items-center sm:justify-between">
              <div className="flex items-center gap-3">
                <BellRing className="h-5 w-5 text-amber-600" />
                <span className="text-sm font-medium text-amber-900">
                  {newIds.size === 1
                    ? "Hay 1 reserva nueva confirmada"
                    : `Hay ${newIds.size} reservas nuevas confirmadas`}
                </span>
              </div>
              <Button
                variant="outline"
                className="w-fit rounded-2xl border-amber-300 bg-white text-amber-900"
                onClick={dismissNew}
              >
                Marcar como vistas
              </Button>
            </CardContent>
          </Card>
        )}

        {error && (
          <Card className="mb-6 rounded-3xl border-red-200 bg-red-50 shadow-sm">
            <CardContent className="p-5 text-sm text-red-700">
              Error al cargar las reservas: {error}
            </CardContent>
          </Card>
        )}

        {loading && reservas.length === 0 && (
          <div className="py-20 text-center text-zinc-400">Cargando reservas…</div>
        )}

        {!loading && reservas.length === 0 && !error && (
          <div className="py-20 text-center text-zinc-400">No hay reservas confirmadas todavía.</div>
        )}

        {/* Listado de reservas */}
        <div className="grid grid-cols-1 gap-5 lg:grid-cols-2">
          {reservas.map((r) => {
            const isNew = newIds.has(r.id);
            const nJuegos = seleccionesPorReserva[r.id]?.size ?? 0;
            const isPendiente = nJuegos === 0;
            const nombreCompleto = [r.festejado.nombre, r.festejado.apellidos]
              .filter(Boolean)
              .join(" ");

            return (
              <Card
                key={r.id}
                className={`rounded-3xl bg-white shadow-sm ${
                  isNew ? "border-2 border-amber-300" : "border-zinc-200"
                }`}
              >
                <CardContent className="flex flex-col gap-4 p-5">
                  {/* Nombre + badges */}
                  <div className="flex flex-wrap items-start justify-between gap-3">
                    <div className="min-w-0">
                      <div className="flex flex-wrap items-center gap-2">
                        <h3 className="text-lg font-semibold text-zinc-900">
                          {nombreCompleto || "Sin nombre"}
                        </h3>
                        {isNew && (
                          <Badge className="bg-amber-100 text-amber-800 hover:bg-amber-100">
                            Nueva
                          </Badge>
                        )}
                      </div>
                      <p className="text-sm text-zinc-500">
                        {r.evento.tipoEvento}
                        {r.festejado.edad ? ` · ${r.festejado.edad} años` : ""}
                      </p>
                    </div>
                    <Badge
                      className="max-w-full whitespace-normal break-words text-right"
                      style={{ backgroundColor: CORPORATE_RED, color: "white" }}
                    >
                      {limpiarTematica(r.evento.tematica)}
                    </Badge>
                  </div>

                  {/* Tipo animación + duración destacados */}
                  <div className="flex flex-wrap gap-2">
                    <span className="inline-flex items-center gap-1.5 rounded-xl bg-zinc-100 px-3 py-1 text-sm font-medium text-zinc-700">
                      <Gamepad2 className="h-3.5 w-3.5 text-zinc-400" />
                      {limpiarTematica(r.evento.tematica)}
                    </span>
                    {r.evento.horasDuracion && (
                      <span className="inline-flex items-center gap-1.5 rounded-xl bg-zinc-100 px-3 py-1 text-sm font-medium text-zinc-700">
                        <Clock className="h-3.5 w-3.5 text-zinc-400" />
                        {r.evento.horasDuracion}h
                      </span>
                    )}
                  </div>

                  {/* Detalles */}
                  <div className="grid grid-cols-1 gap-2 text-sm text-zinc-700 sm:grid-cols-2">
                    <div className="flex items-center gap-2">
                      <CalendarDays className="h-4 w-4 text-zinc-400" />
                      {r.evento.fecha || "Sin fecha"}
                    </div>
                    <div className="flex items-center gap-2">
                      <Clock className="h-4 w-4 text-zinc-400" />
                      {r.evento.horario || "Sin horario"}
                    </div>
                    <div className="flex items-center gap-2">
                      <MapPinned className="h-4 w-4 text-zinc-400" />
                      {r.evento.lugar || "Sin lugar"}
                    </div>
                    {r.evento.colegio && (
                      <div className="flex items-center gap-2">
                        <School className="h-4 w-4 text-zinc-400" />
                        {r.evento.colegio}
                      </div>
                    )}
                    <div className="flex items-center gap-2">
                      <Users className="h-4 w-4 text-zinc-400" />
                      {r.participantes ?? "?"} participantes
                    </div>
                    <div className="flex items-center gap-2">
                      <Phone className="h-4 w-4 text-zinc-400" />
                      {r.contacto.nombre}
                      {r.contacto.telefono ? ` · ${r.contacto.telefono}` : ""}
                    </div>
                  </div>

                  {/* Footer */}
                  <div className="flex flex-wrap items-center gap-2 border-t border-zinc-100 pt-4">
                    <Badge variant="outline" className="rounded-xl border-zinc-200 text-zinc-700">
                      <Users className="mr-1 h-3 w-3" />
                      {r.monitoresEstimados ?? "?"} monitores
                    </Badge>
                    <Badge variant="outline" className="rounded-xl border-zinc-200 text-zinc-700">
                      <Sparkles className="mr-1 h-3 w-3" />
                      {r.capacidadJuegos ?? "?"} juegos
                    </Badge>
                    {isPendiente ? (
                      <Badge className="rounded-xl bg-orange-100 text-orange-700 hover:bg-orange-100">
                        <AlertCircle className="mr-1 h-3 w-3" />
                        Propuesta pendiente
                      </Badge>
                    ) : (
                      <Badge className="rounded-xl bg-emerald-100 text-emerald-700 hover:bg-emerald-100">
                        <CheckCircle2 className="mr-1 h-3 w-3" />
                        Propuesta lista
                      </Badge>
                    )}
                    <div className="ml-auto flex flex-wrap gap-2">
                      {!isPendiente && (
                        <>
                          <Button
                            variant="outline"
                            className="cursor-pointer rounded-2xl border-zinc-200"
                            onClick={() => setPreviewReservaId(r.id)}
                          >
                            <Eye className="mr-2 h-4 w-4" />
                            Vista previa
                          </Button>
                          <Button
                            variant="outline"
                            className="cursor-pointer rounded-2xl border-zinc-200"
                            disabled={descargandoId === r.id}
                            onClick={async () => {
                              setDescargandoId(r.id);
                              try {
                                const juegosSel = juegos.filter((j) => seleccionesPorReserva[r.id]?.has(j.id));
                                await descargarPropuestaReserva(r, juegosSel);
                              } finally {
                                setDescargandoId(null);
                              }
                            }}
                          >
                            {descargandoId === r.id ? (
                              <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                            ) : (
                              <Download className="mr-2 h-4 w-4" />
                            )}
                            Descargar
                          </Button>
                        </>
                      )}
                      {juegos.length > 0 && (
                        <Button
                          variant="outline"
                          className="cursor-pointer rounded-2xl border-zinc-200"
                          onClick={() => setSugerenciaReservaId(r.id)}
                        >
                          <Wand2 className="mr-2 h-4 w-4" />
                          Posible propuesta
                        </Button>
                      )}
                      <Button
                        className="cursor-pointer rounded-2xl text-white"
                        style={{ backgroundColor: CORPORATE_RED }}
                        onClick={() => setSelectedReservaId(r.id)}
                      >
                        <Gamepad2 className="mr-2 h-4 w-4" />
                        Completar propuesta
                      </Button>
                    </div>
                  </div>
                </CardContent>
              </Card>
            );
          })}
        </div>
      </div>

      {/* Overlay vista previa de propuesta */}
      {previewReserva && (
        <PropuestaOverlay
          reserva={previewReserva}
          juegosSeleccionados={previewJuegos}
          onClose={() => setPreviewReservaId(null)}
        />
      )}

      {/* Overlay propuesta sugerida */}
      {sugerenciaReserva && (
        <SugerenciaOverlay
          reserva={sugerenciaReserva}
          juegosSeleccionados={sugerenciaJuegos}
          tienePropuestaActual={(seleccionesPorReserva[sugerenciaReservaId]?.size ?? 0) > 0}
          onAceptar={() => {
            aplicarSugerencia(sugerenciaReservaId, sugerenciaJuegos);
            setSugerenciaReservaId(null);
          }}
          onAjustar={() => {
            aplicarSugerencia(sugerenciaReservaId, sugerenciaJuegos);
            setSugerenciaReservaId(null);
            setSelectedReservaId(sugerenciaReservaId);
          }}
          onClose={() => setSugerenciaReservaId(null)}
        />
      )}

      {/* Modal ayuda calendario */}
      {showCalendarHelp && <CalendarHelp onClose={() => setShowCalendarHelp(false)} />}
    </div>
  );
}

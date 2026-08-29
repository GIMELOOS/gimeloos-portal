"use client";

import { useMemo, useState } from "react";
import { Card, CardContent } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { ArrowLeft, Search, Star, CheckCircle2, Circle, Sparkles, ListChecks, Download, Loader2, Copy, ChevronDown, ChevronUp } from "lucide-react";

const CORPORATE_RED = "#FF3131";
const FAVORITO_MIN_USOS = 3;

function cap(str) {
  if (!str) return "";
  return str.charAt(0).toUpperCase() + str.slice(1);
}

function limpiarNombre(reserva) {
  return [reserva.festejado?.nombre, reserva.festejado?.apellidos].filter(Boolean).join(" ") || "Sin nombre";
}

export default function GameSelector({ reserva, juegos, selectedIds, onToggle, onClose, propuestasCopiables = [], onCopiarPropuesta }) {
  const [query, setQuery] = useState("");
  const [categoria, setCategoria] = useState("Todas");
  const [descargando, setDescargando] = useState(false);
  const [errorDescarga, setErrorDescarga] = useState(null);
  const [mostrarCopiar, setMostrarCopiar] = useState(false);

  const categorias = useMemo(() => {
    const set = new Set();
    juegos.forEach((j) => j.categorias.forEach((c) => set.add(c)));
    return ["Todas", ...[...set].sort((a, b) => a.localeCompare(b, "es"))];
  }, [juegos]);

  const juegosFiltrados = useMemo(() => {
    const q = query.trim().toLowerCase();
    return juegos.filter((j) => {
      const matchCategoria = categoria === "Todas" || j.categorias.includes(categoria);
      const matchQuery = !q || j.nombre.toLowerCase().includes(q) || j.descripcion.toLowerCase().includes(q);
      return matchCategoria && matchQuery;
    });
  }, [juegos, query, categoria]);

  const seleccionados = useMemo(
    () => juegos.filter((j) => selectedIds.has(j.id)),
    [juegos, selectedIds]
  );

  const materialesUnicos = useMemo(() => {
    const set = new Set();
    seleccionados.forEach((j) => j.materiales.forEach((m) => set.add(m)));
    return [...set].sort((a, b) => a.localeCompare(b, "es"));
  }, [seleccionados]);

  const capacidad = reserva.capacidadJuegos;
  const overCapacity = capacidad != null && seleccionados.length > capacidad;

  const descargarPropuesta = async () => {
    if (seleccionados.length === 0) return;
    setDescargando(true);
    setErrorDescarga(null);
    try {
      const res = await fetch("/api/animaciones/propuesta", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ reserva, juegos: seleccionados }),
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
    } catch (e) {
      setErrorDescarga(e.message);
    } finally {
      setDescargando(false);
    }
  };

  return (
    <div className="flex h-screen flex-col bg-zinc-50">
      {/* Cabecera fija */}
      <div className="flex-shrink-0 bg-zinc-50 px-6 py-4 lg:px-10">
        <div className="mx-auto max-w-6xl flex items-center gap-3">
          <Button variant="outline" className="rounded-2xl border-zinc-200 bg-white" onClick={onClose}>
            <ArrowLeft className="mr-2 h-4 w-4" />
            Volver
          </Button>
          <div>
            <h2 className="text-xl font-semibold text-zinc-900">
              {[reserva.festejado.nombre, reserva.festejado.apellidos].filter(Boolean).join(" ")}
            </h2>
            <p className="text-sm text-zinc-500">
              {reserva.evento.tipoEvento}
              {reserva.evento.tematica ? ` · ${reserva.evento.tematica.replace(/\s*\(.*?\)\s*/g, "").trim()}` : ""}
              {reserva.evento.horasDuracion ? ` · ${reserva.evento.horasDuracion}h` : ""}
              {" · "}{reserva.evento.fecha} · {reserva.participantes} participantes
            </p>
          </div>
        </div>
      </div>

      {/* Contenido principal: dos columnas, sólo la izquierda scrollea */}
      <div className="flex-1 overflow-hidden px-6 pb-6 lg:px-10 lg:pb-10">
        <div className="mx-auto flex h-full max-w-6xl gap-6 pt-6">

          {/* Columna izquierda — scrolleable */}
          <div className="flex-1 overflow-y-auto px-1 py-1">
            <div className="mb-4 flex flex-col gap-3 sm:flex-row sm:items-center">
              <div className="relative flex-1">
                <Search className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-zinc-400" />
                <Input
                  placeholder="Buscar juego…"
                  className="rounded-2xl border-zinc-200 bg-white pl-9"
                  value={query}
                  onChange={(e) => setQuery(e.target.value)}
                />
              </div>
            </div>

            <div className="mb-4 flex flex-wrap gap-2">
              {categorias.map((c) => (
                <button
                  key={c}
                  onClick={() => setCategoria(c)}
                  className={`cursor-pointer rounded-xl border px-3 py-1.5 text-sm transition-colors ${
                    categoria === c
                      ? "border-transparent text-white"
                      : "border-zinc-200 bg-white text-zinc-600 hover:bg-zinc-100"
                  }`}
                  style={categoria === c ? { backgroundColor: CORPORATE_RED } : undefined}
                >
                  {c}
                </button>
              ))}
            </div>

            <div className="flex flex-col gap-2">
              {juegosFiltrados.map((j) => {
                const isSelected = selectedIds.has(j.id);
                const isFavorito = j.veces_usado >= FAVORITO_MIN_USOS;
                return (
                  <Card
                    key={j.id}
                    onClick={() => onToggle(j.id)}
                    className={`cursor-pointer rounded-2xl shadow-sm transition-colors ${
                      isSelected
                        ? "border-2 bg-red-50"
                        : isFavorito
                          ? "border-amber-300 bg-amber-50 hover:bg-amber-100"
                          : "border-zinc-200 bg-white hover:bg-zinc-50"
                    }`}
                    style={isSelected ? { borderColor: CORPORATE_RED } : undefined}
                  >
                    <CardContent className="flex items-start gap-3 p-4">
                      {isSelected ? (
                        <CheckCircle2 className="mt-0.5 h-5 w-5 shrink-0" style={{ color: CORPORATE_RED }} />
                      ) : (
                        <Circle className="mt-0.5 h-5 w-5 shrink-0 text-zinc-300" />
                      )}
                      <div className="min-w-0 flex-1">
                        <div className="flex flex-wrap items-center gap-2">
                          <h4 className="font-medium text-zinc-900">{j.nombre}</h4>
                          {isFavorito && (
                            <Badge className="bg-amber-100 text-amber-800 hover:bg-amber-100">
                              <Star className="mr-1 h-3 w-3 fill-amber-500 text-amber-500" />
                              Muy usado
                            </Badge>
                          )}
                        </div>
                        <p className="mt-1 text-sm text-zinc-500">{cap(j.descripcion)}</p>
                        {j.materiales.length > 0 && (
                          <p className="mt-1 text-xs text-zinc-400">
                            Materiales: {j.materiales.join(", ")}
                          </p>
                        )}
                      </div>
                    </CardContent>
                  </Card>
                );
              })}
              {juegosFiltrados.length === 0 && (
                <p className="py-10 text-center text-sm text-zinc-400">No hay juegos que coincidan.</p>
              )}
            </div>
          </div>{/* fin columna izquierda */}

          {/* Columna derecha — siempre visible, scrollea si hace falta */}
          <div className="hidden lg:block w-[360px] flex-shrink-0 overflow-y-auto px-1 py-1">
            <div className="flex flex-col gap-3">

            {/* Copiar propuesta de misma temática */}
            {propuestasCopiables.length > 0 && (
              <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm overflow-hidden">
                <button
                  className="w-full flex items-center justify-between gap-2 p-4 text-left hover:bg-zinc-50 transition-colors cursor-pointer"
                  onClick={() => setMostrarCopiar((v) => !v)}
                >
                  <div className="flex items-center gap-2">
                    <Copy className="h-4 w-4 shrink-0" style={{ color: CORPORATE_RED }} />
                    <span className="text-sm font-semibold text-zinc-900">
                      Copiar propuesta anterior
                    </span>
                    <span className="rounded-full bg-zinc-100 px-2 py-0.5 text-xs text-zinc-500">
                      {propuestasCopiables.length}
                    </span>
                  </div>
                  {mostrarCopiar ? (
                    <ChevronUp className="h-4 w-4 text-zinc-400" />
                  ) : (
                    <ChevronDown className="h-4 w-4 text-zinc-400" />
                  )}
                </button>
                {mostrarCopiar && (
                  <div className="border-t border-zinc-100 flex flex-col divide-y divide-zinc-100">
                    {propuestasCopiables.map(({ reserva: r, juegosIds }) => (
                      <div key={r.id} className="flex items-center justify-between gap-3 px-4 py-3">
                        <div className="min-w-0">
                          <p className="truncate text-sm font-medium text-zinc-800">{limpiarNombre(r)}</p>
                          <p className="text-xs text-zinc-400">
                            {r.evento.fecha} · {juegosIds.size} juegos
                          </p>
                        </div>
                        <Button
                          size="sm"
                          variant="outline"
                          className="shrink-0 rounded-xl border-zinc-200 text-xs"
                          onClick={() => {
                            if (
                              selectedIds.size > 0 &&
                              !window.confirm(`¿Reemplazar la selección actual (${selectedIds.size} juegos) con la propuesta de ${limpiarNombre(r)}?`)
                            ) return;
                            onCopiarPropuesta(juegosIds);
                            setMostrarCopiar(false);
                          }}
                        >
                          Copiar
                        </Button>
                      </div>
                    ))}
                  </div>
                )}
              </Card>
            )}

            <Card className="rounded-3xl border-zinc-200 bg-white shadow-sm">
              <CardContent className="flex flex-col gap-4 p-5">
                <div className="flex items-center gap-2">
                  <ListChecks className="h-5 w-5" style={{ color: CORPORATE_RED }} />
                  <h3 className="font-semibold text-zinc-900">Propuesta</h3>
                </div>

                <div
                  className={`rounded-xl px-3 py-2 text-sm ${
                    overCapacity ? "bg-red-50 text-red-700" : "bg-zinc-50 text-zinc-600"
                  }`}
                >
                  {seleccionados.length} juegos seleccionados
                  {capacidad != null && ` de ${capacidad} recomendados`}
                  {overCapacity && " — supera la capacidad estimada"}
                </div>

                <div>
                  <h4 className="mb-2 text-sm font-semibold text-zinc-700">Descripción de la actividad</h4>
                  {seleccionados.length === 0 ? (
                    <p className="text-sm text-zinc-400">Selecciona juegos para generar la descripción.</p>
                  ) : (
                    <ol className="list-decimal space-y-2 pl-4 text-sm text-zinc-700">
                      {seleccionados.map((j) => (
                        <li key={j.id}>
                          <span className="font-medium">{j.nombre}:</span> {cap(j.descripcion)}
                        </li>
                      ))}
                    </ol>
                  )}
                </div>

                <div>
                  <h4 className="mb-2 flex items-center gap-2 text-sm font-semibold text-zinc-700">
                    <Sparkles className="h-4 w-4" />
                    Materiales necesarios
                  </h4>
                  {materialesUnicos.length === 0 ? (
                    <p className="text-sm text-zinc-400">Sin materiales todavía.</p>
                  ) : (
                    <div className="flex flex-wrap gap-1.5">
                      {materialesUnicos.map((m) => (
                        <Badge key={m} variant="outline" className="rounded-lg border-zinc-200 text-zinc-600">
                          {m}
                        </Badge>
                      ))}
                    </div>
                  )}
                </div>
              </CardContent>
            </Card>

            {/* Botón descargar */}
            <Button
              className="w-full rounded-2xl text-white"
              style={{ backgroundColor: CORPORATE_RED }}
              disabled={seleccionados.length === 0 || descargando}
              onClick={descargarPropuesta}
            >
              {descargando ? (
                <>
                  <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                  Generando documento…
                </>
              ) : (
                <>
                  <Download className="mr-2 h-4 w-4" />
                  Descargar propuesta Word
                </>
              )}
            </Button>
            {errorDescarga && (
              <p className="text-center text-xs text-red-600">{errorDescarga}</p>
            )}
            {seleccionados.length === 0 && (
              <p className="text-center text-xs text-zinc-400">
                Selecciona al menos un juego para descargar
              </p>
            )}
            </div>{/* fin flex interno */}
          </div>{/* fin columna derecha */}

        </div>{/* fin flex principal */}
      </div>{/* fin overflow-hidden */}
    </div>
  );
}

"use client";

import { useState, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";

// ─── Constantes SS España 2024 ────────────────────────────────────────────────
const SS_TRABAJADOR = 0.0647;
const SS_EMPRESA = 0.305;

const CATEGORIAS = [
  { id: "alojamiento", label: "Alojamiento",           iva: 10 },
  { id: "transporte",  label: "Transporte",             iva: 10 },
  { id: "comida",      label: "Comida",                 iva: 10 },
  { id: "seguro",      label: "Seguro",                 iva:  0 },
  { id: "material",    label: "Material",               iva: 21 },
  { id: "gasolina",    label: "Gasolina",               iva: 21 },
  { id: "imprevisto",  label: "Imprevistos",            iva: 21 },
  { id: "actividad",   label: "Act. subcontratada",     iva: 21 },
  { id: "otro",        label: "Otros",                  iva: 21 },
];

// ─── Helpers ──────────────────────────────────────────────────────────────────
const uid = () => Math.random().toString(36).slice(2, 9);
const fmt = (n) =>
  new Intl.NumberFormat("es-ES", {
    style: "currency",
    currency: "EUR",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(n ?? 0);
const pct = (n) => `${(n ?? 0).toFixed(1)}%`;
const n2 = (v) => parseFloat(v) || 0;

// ─── Átomos de UI ─────────────────────────────────────────────────────────────

function SectionLabel({ children }) {
  return (
    <p className="px-4 pt-6 pb-1.5 text-[11px] font-semibold uppercase tracking-widest text-zinc-400">
      {children}
    </p>
  );
}

function Card({ children, className = "" }) {
  return (
    <div className={`mx-4 overflow-hidden rounded-2xl bg-white shadow-sm ${className}`}>
      {children}
    </div>
  );
}

function Row({ label, sub, value, valueClass = "", last = false }) {
  return (
    <div className={`flex items-center justify-between px-4 py-3 ${!last ? "border-b border-zinc-100" : ""}`}>
      <div>
        <div className="text-sm text-zinc-700">{label}</div>
        {sub && <div className="text-xs text-zinc-400">{sub}</div>}
      </div>
      <div className={`text-sm font-semibold tabular-nums ${valueClass}`}>{value}</div>
    </div>
  );
}

function Stepper({ value, onChange, min = 0, max = 999 }) {
  return (
    <div className="flex items-center gap-2">
      <button
        onClick={() => onChange(Math.max(min, value - 1))}
        className="flex h-7 w-7 items-center justify-center rounded-full bg-zinc-100 text-zinc-500 text-lg leading-none transition-colors hover:bg-zinc-200 active:scale-95"
      >
        −
      </button>
      <span className="w-8 text-center text-sm font-semibold tabular-nums">{value}</span>
      <button
        onClick={() => onChange(Math.min(max, value + 1))}
        className="flex h-7 w-7 items-center justify-center rounded-full bg-blue-500 text-white text-lg leading-none transition-colors hover:bg-blue-600 active:scale-95"
      >
        +
      </button>
    </div>
  );
}

function IVASeg({ value, onChange }) {
  return (
    <div className="flex gap-0.5 rounded-lg bg-zinc-100 p-0.5">
      {[0, 10, 21].map((r) => (
        <button
          key={r}
          onClick={() => onChange(r)}
          className={`rounded-md px-2 py-1 text-xs font-medium transition-all ${
            value === r ? "bg-white text-blue-600 shadow-sm" : "text-zinc-400 hover:text-zinc-600"
          }`}
        >
          {r}%
        </button>
      ))}
    </div>
  );
}

function NumInput({ value, onChange, prefix = "€", suffix, className = "" }) {
  return (
    <div className={`flex items-center gap-1 ${className}`}>
      {prefix && <span className="text-xs text-zinc-400">{prefix}</span>}
      <input
        type="number"
        min="0"
        step="0.01"
        value={value || ""}
        onChange={(e) => onChange(n2(e.target.value))}
        placeholder="0,00"
        className="w-28 bg-transparent text-right text-sm font-semibold text-zinc-900 placeholder:text-zinc-300 focus:outline-none"
      />
      {suffix && <span className="text-xs text-zinc-400">{suffix}</span>}
    </div>
  );
}

function TipoBadge({ value, onChange }) {
  return (
    <div className="flex gap-0.5 rounded-lg bg-zinc-100 p-0.5">
      {[
        { val: "grupo",   label: "Por grupo" },
        { val: "persona", label: "Por persona" },
      ].map((opt) => (
        <button
          key={opt.val}
          onClick={() => onChange(opt.val)}
          className={`rounded-md px-2 py-1 text-xs font-medium transition-all ${
            value === opt.val ? "bg-white text-blue-600 shadow-sm" : "text-zinc-400 hover:text-zinc-600"
          }`}
        >
          {opt.label}
        </button>
      ))}
    </div>
  );
}

// ─── Componente principal ─────────────────────────────────────────────────────
export default function CalculadoraCampamento() {
  // ── Estado básico ──────────────────────────────────────────────────────────
  const [nombre, setNombre]               = useState("");
  const [fechaInicio, setFechaInicio]     = useState("");
  const [fechaFin, setFechaFin]           = useState("");
  const [numParticipantes, setNumPart]    = useState(60);
  const [numDocentes, setNumDoc]          = useState(0);
  const [numMonitores, setNumMon]         = useState(1);

  // ── Personal ───────────────────────────────────────────────────────────────
  const [monitores, setMonitores] = useState([
    { id: uid(), nombre: "Monitor 1", dias: 5, neto: 0, irpf: 15 },
  ]);

  // ── Gastos ─────────────────────────────────────────────────────────────────
  const [gastos, setGastos] = useState([]);

  // ── Precio y margen ────────────────────────────────────────────────────────
  const [gimeloosFee, setGimeloosFee]     = useState(0);
  const [margen, setMargen]               = useState(15);

  // ── Tabla de escala ────────────────────────────────────────────────────────
  const [tablaMin, setTablaMin]           = useState(30);
  const [tablaMax, setTablaMax]           = useState(60);

  // ── UI ─────────────────────────────────────────────────────────────────────
  const [tab, setTab]                     = useState("datos");
  const [exportando, setExportando]       = useState(false);

  // ── Cálculo central ────────────────────────────────────────────────────────
  const totalPersonas = numParticipantes + numDocentes + numMonitores;

  const calcularCostesParaN = useCallback(
    (n) => {
      const totalP = n + numDocentes + numMonitores;
      let sinIVA = 0;
      let conIVA = 0;

      for (const g of gastos) {
        const base = g.tipo === "grupo" ? g.importe : g.importe * totalP;
        sinIVA += base;
        conIVA += base * (1 + g.iva / 100);
      }

      // Salarios monitores (sin IVA – son nóminas)
      const salarios = monitores.reduce((acc, m) => {
        if (!m.neto) return acc;
        const bruto = m.neto / (1 - SS_TRABAJADOR - m.irpf / 100);
        return acc + bruto * (1 + SS_EMPRESA) * m.dias;
      }, 0);

      // Fee Gimeloos (coste organizativo)
      const fee = gimeloosFee * n;

      const total = conIVA + salarios + fee;
      return { sinIVA, ivaGastos: conIVA - sinIVA, conIVA, salarios, fee, total };
    },
    [gastos, monitores, gimeloosFee, numDocentes, numMonitores]
  );

  const costes = useMemo(() => calcularCostesParaN(numParticipantes), [calcularCostesParaN, numParticipantes]);

  const costePorPart    = numParticipantes > 0 ? costes.total / numParticipantes : 0;
  const costeSinIVAPorP = numParticipantes > 0 ? (costes.sinIVA + costes.salarios + costes.fee) / numParticipantes : 0;
  const precioVenta     = margen < 100 ? costePorPart / (1 - margen / 100) : 0;
  const ingresosTotales = precioVenta * numParticipantes;
  const beneficioTotal  = ingresosTotales - costes.total;
  const margenReal      = ingresosTotales > 0 ? (beneficioTotal / ingresosTotales) * 100 : 0;
  const esViable        = beneficioTotal > 0;

  // ── Tabla de escala ────────────────────────────────────────────────────────
  const tablaEscala = useMemo(() => {
    const rows = [];
    for (let n = tablaMin; n <= tablaMax; n++) {
      const c = calcularCostesParaN(n);
      const costePP = n > 0 ? c.total / n : 0;
      const precio  = margen < 100 ? costePP / (1 - margen / 100) : 0;
      rows.push({ n, costeTotal: c.total, costePP, precio, beneficio: precio * n - c.total });
    }
    return rows;
  }, [calcularCostesParaN, tablaMin, tablaMax, margen]);

  // ── Gastos helpers ─────────────────────────────────────────────────────────
  const addGasto = () =>
    setGastos((prev) => [
      ...prev,
      { id: uid(), nombre: "", categoria: "otro", importe: 0, tipo: "grupo", iva: 21 },
    ]);

  const updGasto = (id, field, val) =>
    setGastos((prev) => prev.map((g) => (g.id === id ? { ...g, [field]: val } : g)));

  const delGasto = (id) => setGastos((prev) => prev.filter((g) => g.id !== id));

  // ── Monitor helpers ────────────────────────────────────────────────────────
  const addMonitor = () => {
    const newMon = { id: uid(), nombre: `Monitor ${monitores.length + 1}`, dias: 5, neto: 0, irpf: 15 };
    setMonitores((prev) => [...prev, newMon]);
    setNumMon((n) => n + 1);
  };

  const updMonitor = (id, field, val) =>
    setMonitores((prev) => prev.map((m) => (m.id === id ? { ...m, [field]: val } : m)));

  const delMonitor = (id) => {
    setMonitores((prev) => prev.filter((m) => m.id !== id));
    setNumMon((n) => Math.max(0, n - 1));
  };

  // ── Sync monitor count stepper ─────────────────────────────────────────────
  const handleNumMonitores = (v) => {
    setNumMon(v);
    setMonitores((prev) => {
      if (v > prev.length)
        return [
          ...prev,
          ...Array.from({ length: v - prev.length }, (_, i) => ({
            id: uid(),
            nombre: `Monitor ${prev.length + i + 1}`,
            dias: 5,
            neto: 0,
            irpf: 15,
          })),
        ];
      return prev.slice(0, v);
    });
  };

  // ── Export helpers ─────────────────────────────────────────────────────────
  const buildWorkbook = useCallback(() => {
    const wb = XLSX.utils.book_new();

    // Hoja 1: Resumen
    const resumenData = [
      ["PRESUPUESTO CAMPAMENTO GIMELOOS", nombre || "—"],
      ["Fecha inicio", fechaInicio || "—"],
      ["Fecha fin", fechaFin || "—"],
      [],
      ["PERSONAS", ""],
      ["Participantes (pagan)", numParticipantes],
      ["Docentes (no pagan)", numDocentes],
      ["Monitores Gimeloos (no pagan)", numMonitores],
      ["Total personas usando el servicio", totalPersonas],
      [],
      ["COSTES", ""],
      ["Gastos sin IVA", costes.sinIVA],
      ["IVA soportado en gastos", costes.ivaGastos],
      ["Gastos con IVA", costes.conIVA],
      ["Salarios monitores (coste empresa)", costes.salarios],
      ["Fee organizativo Gimeloos", costes.fee],
      ["COSTE TOTAL", costes.total],
      [],
      ["COSTE POR PARTICIPANTE", ""],
      ["Sin IVA", costeSinIVAPorP],
      ["Con IVA", costePorPart],
      [],
      ["PRECIO DE VENTA", ""],
      ["Margen objetivo (%)", margen],
      ["Precio por participante", precioVenta],
      ["Ingresos totales", ingresosTotales],
      ["Beneficio total", beneficioTotal],
      ["Margen real (%)", margenReal],
      ["¿Es viable?", esViable ? "SÍ" : "NO"],
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(resumenData), "Resumen");

    // Hoja 2: Gastos detallados
    const gastosData = [
      ["Nombre", "Categoría", "Tipo", "Importe base (€)", "IVA %", "Base total sin IVA", "Total con IVA"],
      ...gastos.map((g) => {
        const base = g.tipo === "grupo" ? g.importe : g.importe * totalPersonas;
        return [
          g.nombre || "(sin nombre)",
          CATEGORIAS.find((c) => c.id === g.categoria)?.label || g.categoria,
          g.tipo === "grupo" ? "Por grupo" : `Por persona (×${totalPersonas})`,
          g.importe,
          g.iva,
          base,
          base * (1 + g.iva / 100),
        ];
      }),
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(gastosData), "Gastos");

    // Hoja 3: Personal
    const personalData = [
      ["Monitor", "Días", "IRPF %", "Neto/día (€)", "Bruto/día (€)", "SS empresa/día (€)", "Coste empresa/día (€)", "Coste empresa total (€)"],
      ...monitores.map((m) => {
        const bruto   = m.neto > 0 ? m.neto / (1 - SS_TRABAJADOR - m.irpf / 100) : 0;
        const ssEmp   = bruto * SS_EMPRESA;
        const costeD  = bruto + ssEmp;
        const costeT  = costeD * m.dias;
        return [m.nombre, m.dias, m.irpf, m.neto, bruto, ssEmp, costeD, costeT];
      }),
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(personalData), "Personal");

    // Hoja 4: Tabla de precios por nº participantes
    const tablaData = [
      ["Nº participantes", "Coste total (€)", "Coste/plaza (€)", "Precio de venta/plaza (€)", "Beneficio total (€)"],
      ...tablaEscala.map((r) => [r.n, r.costeTotal, r.costePP, r.precio, r.beneficio]),
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(tablaData), "Tabla de precios");

    return wb;
  }, [nombre, fechaInicio, fechaFin, numParticipantes, numDocentes, numMonitores, totalPersonas, costes, costeSinIVAPorP, costePorPart, margen, precioVenta, ingresosTotales, beneficioTotal, margenReal, esViable, gastos, monitores, tablaEscala]);

  const downloadXLSX = () => {
    const wb = buildWorkbook();
    XLSX.writeFile(wb, `Presupuesto_${(nombre || "Campamento").replace(/\s+/g, "_")}.xlsx`);
  };

  const uploadToDrive = async () => {
    const clientId = process.env.NEXT_PUBLIC_GOOGLE_CLIENT_ID;
    if (!clientId) {
      alert("Para subir a Google Drive añade NEXT_PUBLIC_GOOGLE_CLIENT_ID en .env.local.\nVer instrucciones en /app/ui/calculadora-campamento.jsx");
      return;
    }
    setExportando(true);
    try {
      if (!window.google?.accounts?.oauth2) {
        await new Promise((resolve, reject) => {
          const s = document.createElement("script");
          s.src = "https://accounts.google.com/gsi/client";
          s.onload = resolve;
          s.onerror = reject;
          document.head.appendChild(s);
        });
      }
      const tokenClient = window.google.accounts.oauth2.initTokenClient({
        client_id: clientId,
        scope: "https://www.googleapis.com/auth/drive.file",
        callback: async (resp) => {
          if (resp.error) {
            alert("Error al conectar con Google: " + resp.error);
            setExportando(false);
            return;
          }
          try {
            const wb   = buildWorkbook();
            const buf  = XLSX.write(wb, { type: "array", bookType: "xlsx" });
            const blob = new Blob([buf], {
              type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            });
            const meta = JSON.stringify({
              name: `Presupuesto_${(nombre || "Campamento").replace(/\s+/g, "_")}.xlsx`,
              mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            });
            const form = new FormData();
            form.append("metadata", new Blob([meta], { type: "application/json" }));
            form.append("file", blob);
            const res  = await fetch(
              "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart",
              { method: "POST", headers: { Authorization: `Bearer ${resp.access_token}` }, body: form }
            );
            const data = await res.json();
            if (data.id) {
              alert(`✓ Subido a Google Drive: ${data.name}`);
            } else {
              console.error("Drive error:", data);
              alert("Error al subir a Drive. Ver consola.");
            }
          } finally {
            setExportando(false);
          }
        },
      });
      tokenClient.requestAccessToken();
    } catch (e) {
      console.error(e);
      alert("Error al conectar con Google.");
      setExportando(false);
    }
  };

  // ── Render ─────────────────────────────────────────────────────────────────
  const TABS = [
    { id: "datos",    label: "Datos" },
    { id: "gastos",   label: "Gastos" },
    { id: "personal", label: "Personal" },
    { id: "precio",   label: "Precio" },
    { id: "tabla",    label: "Tabla" },
  ];

  return (
    <div className="min-h-screen bg-zinc-100 pb-12">
      {/* ── Cabecera ─────────────────────────────────────────────────────── */}
      <div className="sticky top-0 z-10 border-b border-zinc-200 bg-white/90 backdrop-blur-md">
        <div className="flex items-center justify-between px-4 py-3">
          <div>
            <h2 className="text-base font-bold text-zinc-900">Calculadora de Campamentos</h2>
            <p className="text-xs text-zinc-400">{nombre || "Nuevo campamento"}</p>
          </div>
          <div className="flex gap-2">
            <button
              onClick={downloadXLSX}
              className="rounded-xl bg-zinc-100 px-3 py-2 text-xs font-semibold text-zinc-700 transition-colors hover:bg-zinc-200"
            >
              ↓ Excel
            </button>
            <button
              onClick={uploadToDrive}
              disabled={exportando}
              className="rounded-xl bg-blue-500 px-3 py-2 text-xs font-semibold text-white transition-colors hover:bg-blue-600 disabled:opacity-50"
            >
              {exportando ? "Subiendo…" : "↑ Drive"}
            </button>
          </div>
        </div>

        {/* Tabs */}
        <div className="flex gap-0 overflow-x-auto px-4">
          {TABS.map((t) => (
            <button
              key={t.id}
              onClick={() => setTab(t.id)}
              className={`whitespace-nowrap border-b-2 px-3 py-2.5 text-sm font-medium transition-colors ${
                tab === t.id
                  ? "border-blue-500 text-blue-500"
                  : "border-transparent text-zinc-400 hover:text-zinc-600"
              }`}
            >
              {t.label}
            </button>
          ))}
        </div>
      </div>

      {/* ── Tarjeta de resumen (siempre visible) ─────────────────────────── */}
      <div className="px-4 pt-4">
        <div
          className={`rounded-2xl border-2 bg-white p-4 shadow-sm ${
            costes.total === 0
              ? "border-zinc-200"
              : esViable
              ? "border-green-400"
              : "border-red-400"
          }`}
        >
          <div className="mb-3 flex items-center justify-between">
            <span className="text-[11px] font-semibold uppercase tracking-widest text-zinc-400">
              Resumen
            </span>
            {costes.total > 0 && (
              <span
                className={`rounded-full px-2.5 py-0.5 text-xs font-bold ${
                  esViable ? "bg-green-100 text-green-700" : "bg-red-100 text-red-700"
                }`}
              >
                {esViable ? "✓ Viable" : "✗ No viable"}
              </span>
            )}
          </div>
          <div className="grid grid-cols-2 gap-4 sm:grid-cols-4">
            <div>
              <div className="text-[11px] text-zinc-400">Precio/plaza</div>
              <div className="text-xl font-bold text-zinc-900">{fmt(precioVenta)}</div>
            </div>
            <div>
              <div className="text-[11px] text-zinc-400">Beneficio total</div>
              <div className={`text-xl font-bold ${beneficioTotal >= 0 ? "text-green-600" : "text-red-600"}`}>
                {fmt(beneficioTotal)}
              </div>
            </div>
            <div>
              <div className="text-[11px] text-zinc-400">Coste/plaza c/IVA</div>
              <div className="text-sm font-semibold text-zinc-700">{fmt(costePorPart)}</div>
            </div>
            <div>
              <div className="text-[11px] text-zinc-400">Margen real</div>
              <div className={`text-sm font-semibold ${margenReal >= margen ? "text-green-600" : "text-red-500"}`}>
                {pct(margenReal)}
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* ── Contenido de cada pestaña ─────────────────────────────────────── */}

      {/* ══ DATOS ══════════════════════════════════════════════════════════ */}
      {tab === "datos" && (
        <>
          <SectionLabel>Campamento</SectionLabel>
          <Card>
            <div className="border-b border-zinc-100 px-4 py-3">
              <input
                type="text"
                value={nombre}
                onChange={(e) => setNombre(e.target.value)}
                placeholder="Nombre del campamento"
                className="w-full bg-transparent text-sm font-semibold text-zinc-900 placeholder:text-zinc-300 focus:outline-none"
              />
            </div>
            <div className="flex items-center justify-between border-b border-zinc-100 px-4 py-3">
              <span className="text-sm text-zinc-500">Fecha inicio</span>
              <input
                type="date"
                value={fechaInicio}
                onChange={(e) => setFechaInicio(e.target.value)}
                className="bg-transparent text-sm text-zinc-900 focus:outline-none"
              />
            </div>
            <div className="flex items-center justify-between px-4 py-3">
              <span className="text-sm text-zinc-500">Fecha fin</span>
              <input
                type="date"
                value={fechaFin}
                onChange={(e) => setFechaFin(e.target.value)}
                className="bg-transparent text-sm text-zinc-900 focus:outline-none"
              />
            </div>
          </Card>

          <SectionLabel>Personas</SectionLabel>
          <Card>
            <div className="flex items-center justify-between border-b border-zinc-100 px-4 py-3.5">
              <div>
                <div className="text-sm font-medium text-zinc-900">Participantes</div>
                <div className="text-xs text-zinc-400">Pagan la actividad</div>
              </div>
              <Stepper value={numParticipantes} onChange={setNumPart} min={1} />
            </div>
            <div className="flex items-center justify-between border-b border-zinc-100 px-4 py-3.5">
              <div>
                <div className="text-sm font-medium text-zinc-900">Docentes</div>
                <div className="text-xs text-zinc-400">Usan el servicio, no pagan</div>
              </div>
              <Stepper value={numDocentes} onChange={setNumDoc} />
            </div>
            <div className="flex items-center justify-between px-4 py-3.5">
              <div>
                <div className="text-sm font-medium text-zinc-900">Monitores Gimeloos</div>
                <div className="text-xs text-zinc-400">Usan el servicio, no pagan</div>
              </div>
              <Stepper value={numMonitores} onChange={handleNumMonitores} />
            </div>
          </Card>

          <div className="mx-4 mt-3 rounded-xl bg-blue-50 p-3 text-xs text-blue-700">
            <span className="font-semibold">Total personas: {totalPersonas}</span> — Los gastos
            variables (por persona) se calculan para las {totalPersonas} personas. El coste total lo
            pagan solo los {numParticipantes} participantes.
          </div>
        </>
      )}

      {/* ══ GASTOS ═════════════════════════════════════════════════════════ */}
      {tab === "gastos" && (
        <>
          <SectionLabel>Gastos del campamento</SectionLabel>

          {gastos.length === 0 && (
            <Card>
              <div className="flex flex-col items-center gap-2 py-10 text-center">
                <span className="text-4xl">💶</span>
                <p className="text-sm text-zinc-400">Pulsa "Añadir gasto" para empezar</p>
              </div>
            </Card>
          )}

          {gastos.map((g) => {
            const base = g.tipo === "grupo" ? g.importe : g.importe * totalPersonas;
            const conIVA = base * (1 + g.iva / 100);
            return (
              <div key={g.id} className="mx-4 mb-2 overflow-hidden rounded-2xl bg-white shadow-sm">
                {/* Nombre */}
                <div className="flex items-center gap-2 border-b border-zinc-100 px-4 py-3">
                  <input
                    type="text"
                    value={g.nombre}
                    onChange={(e) => updGasto(g.id, "nombre", e.target.value)}
                    placeholder="Nombre del gasto"
                    className="flex-1 bg-transparent text-sm font-semibold text-zinc-900 placeholder:text-zinc-300 focus:outline-none"
                  />
                  <button
                    onClick={() => delGasto(g.id)}
                    className="text-lg leading-none text-zinc-300 hover:text-red-500"
                  >
                    ×
                  </button>
                </div>
                {/* Categoría */}
                <div className="flex items-center justify-between border-b border-zinc-100 px-4 py-2.5">
                  <span className="text-xs text-zinc-400">Categoría</span>
                  <select
                    value={g.categoria}
                    onChange={(e) => {
                      const cat = CATEGORIAS.find((c) => c.id === e.target.value);
                      updGasto(g.id, "categoria", e.target.value);
                      if (cat) updGasto(g.id, "iva", cat.iva);
                    }}
                    className="bg-transparent text-xs font-medium text-zinc-700 focus:outline-none"
                  >
                    {CATEGORIAS.map((c) => (
                      <option key={c.id} value={c.id}>{c.label}</option>
                    ))}
                  </select>
                </div>
                {/* Importe */}
                <div className="flex items-center justify-between border-b border-zinc-100 px-4 py-2.5">
                  <span className="text-xs text-zinc-400">Importe</span>
                  <NumInput value={g.importe} onChange={(v) => updGasto(g.id, "importe", v)} />
                </div>
                {/* Tipo */}
                <div className="flex items-center justify-between border-b border-zinc-100 px-4 py-2.5">
                  <span className="text-xs text-zinc-400">Se cobra…</span>
                  <TipoBadge value={g.tipo} onChange={(v) => updGasto(g.id, "tipo", v)} />
                </div>
                {/* IVA */}
                <div className="flex items-center justify-between border-b border-zinc-100 px-4 py-2.5">
                  <span className="text-xs text-zinc-400">IVA</span>
                  <IVASeg value={g.iva} onChange={(v) => updGasto(g.id, "iva", v)} />
                </div>
                {/* Totales calculados */}
                <div className="flex justify-between bg-zinc-50 px-4 py-2.5">
                  <span className="text-xs text-zinc-400">
                    {g.tipo === "persona" ? `Total grupo (×${totalPersonas})` : "Total grupo"}
                  </span>
                  <span className="text-xs">
                    <span className="text-zinc-400">{fmt(base)} s/IVA</span>
                    <span className="text-zinc-300"> → </span>
                    <span className="font-semibold text-zinc-700">{fmt(conIVA)} c/IVA</span>
                  </span>
                </div>
              </div>
            );
          })}

          <div className="mx-4 mt-2">
            <button
              onClick={addGasto}
              className="w-full rounded-2xl bg-blue-500 py-3.5 text-sm font-semibold text-white transition-colors hover:bg-blue-600 active:scale-[0.99]"
            >
              + Añadir gasto
            </button>
          </div>

          {gastos.length > 0 && (
            <>
              <SectionLabel>Total gastos</SectionLabel>
              <Card>
                <Row label="Sin IVA"        value={fmt(costes.sinIVA)} />
                <Row label="IVA soportado"  value={fmt(costes.ivaGastos)} />
                <Row label="Con IVA"        value={fmt(costes.conIVA)} valueClass="text-blue-600" last />
              </Card>
            </>
          )}
        </>
      )}

      {/* ══ PERSONAL ═══════════════════════════════════════════════════════ */}
      {tab === "personal" && (
        <>
          <SectionLabel>Monitores Gimeloos</SectionLabel>

          {monitores.length === 0 && (
            <Card>
              <div className="flex flex-col items-center gap-2 py-10 text-center">
                <span className="text-4xl">👤</span>
                <p className="text-sm text-zinc-400">Añade monitores para calcular salarios</p>
              </div>
            </Card>
          )}

          {monitores.map((m) => {
            const irpf      = m.irpf / 100;
            const bruto     = m.neto > 0 ? m.neto / (1 - SS_TRABAJADOR - irpf) : 0;
            const ssEmp     = bruto * SS_EMPRESA;
            const costeD    = bruto + ssEmp;
            const costeT    = costeD * m.dias;
            return (
              <div key={m.id} className="mx-4 mb-2 overflow-hidden rounded-2xl bg-white shadow-sm">
                <div className="flex items-center gap-2 border-b border-zinc-100 px-4 py-3">
                  <input
                    type="text"
                    value={m.nombre}
                    onChange={(e) => updMonitor(m.id, "nombre", e.target.value)}
                    className="flex-1 bg-transparent text-sm font-semibold text-zinc-900 focus:outline-none"
                  />
                  <button
                    onClick={() => delMonitor(m.id)}
                    className="text-lg leading-none text-zinc-300 hover:text-red-500"
                  >
                    ×
                  </button>
                </div>
                <div className="flex items-center justify-between border-b border-zinc-100 px-4 py-3">
                  <span className="text-xs text-zinc-500">Días trabajados</span>
                  <Stepper value={m.dias} onChange={(v) => updMonitor(m.id, "dias", v)} min={1} />
                </div>
                <div className="flex items-center justify-between border-b border-zinc-100 px-4 py-3">
                  <span className="text-xs text-zinc-500">Salario neto/día</span>
                  <NumInput value={m.neto} onChange={(v) => updMonitor(m.id, "neto", v)} />
                </div>
                <div className="border-b border-zinc-100 px-4 py-3">
                  <div className="mb-2 flex items-center justify-between">
                    <span className="text-xs text-zinc-500">Retención IRPF</span>
                    <span className="text-xs font-semibold text-zinc-700">{m.irpf}%</span>
                  </div>
                  <input
                    type="range"
                    min="0"
                    max="30"
                    value={m.irpf}
                    onChange={(e) => updMonitor(m.id, "irpf", parseInt(e.target.value))}
                    className="w-full accent-blue-500"
                  />
                </div>
                {/* Computed */}
                <div className="space-y-1.5 bg-zinc-50 px-4 py-3 text-xs">
                  <div className="flex justify-between">
                    <span className="text-zinc-400">Bruto/día</span>
                    <span className="font-medium text-zinc-700">{fmt(bruto)}</span>
                  </div>
                  <div className="flex justify-between">
                    <span className="text-zinc-400">SS trabajador ({(SS_TRABAJADOR * 100).toFixed(2)}%)</span>
                    <span className="font-medium text-zinc-700">{fmt(bruto * SS_TRABAJADOR)}</span>
                  </div>
                  <div className="flex justify-between">
                    <span className="text-zinc-400">SS empresa ({(SS_EMPRESA * 100).toFixed(1)}%)</span>
                    <span className="font-medium text-zinc-700">{fmt(ssEmp)}</span>
                  </div>
                  <div className="flex justify-between border-t border-zinc-200 pt-1.5">
                    <span className="font-medium text-zinc-600">Coste empresa/día</span>
                    <span className="font-semibold text-zinc-900">{fmt(costeD)}</span>
                  </div>
                  <div className="flex justify-between">
                    <span className="font-semibold text-zinc-700">Coste empresa total ({m.dias} días)</span>
                    <span className="font-bold text-blue-600">{fmt(costeT)}</span>
                  </div>
                </div>
              </div>
            );
          })}

          <div className="mx-4 mt-2">
            <button
              onClick={addMonitor}
              className="w-full rounded-2xl bg-blue-500 py-3.5 text-sm font-semibold text-white transition-colors hover:bg-blue-600"
            >
              + Añadir monitor
            </button>
          </div>

          {monitores.length > 0 && (
            <>
              <SectionLabel>Total personal</SectionLabel>
              <Card>
                <Row label="Coste empresa total monitores" value={fmt(costes.salarios)} valueClass="text-blue-600" last />
              </Card>
              <div className="mx-4 mt-3 rounded-xl bg-amber-50 p-3 text-xs text-amber-700">
                Los porcentajes SS son orientativos (2024). Consulta con tu gestor para datos exactos según convenio y categoría.
              </div>
            </>
          )}
        </>
      )}

      {/* ══ PRECIO ═════════════════════════════════════════════════════════ */}
      {tab === "precio" && (
        <>
          <SectionLabel>Configuración de precio</SectionLabel>
          <Card>
            <div className="flex items-center justify-between border-b border-zinc-100 px-4 py-3.5">
              <div>
                <div className="text-sm font-medium text-zinc-900">Fee Gimeloos</div>
                <div className="text-xs text-zinc-400">Coste organizativo por participante</div>
              </div>
              <NumInput value={gimeloosFee} onChange={setGimeloosFee} />
            </div>
            <div className="px-4 py-3.5">
              <div className="mb-3 flex items-center justify-between">
                <div>
                  <div className="text-sm font-medium text-zinc-900">Margen objetivo</div>
                  <div className="text-xs text-zinc-400">% de beneficio sobre precio final</div>
                </div>
                <span className="text-2xl font-bold text-blue-500">{margen}%</span>
              </div>
              <input
                type="range"
                min="0"
                max="50"
                value={margen}
                onChange={(e) => setMargen(parseInt(e.target.value))}
                className="w-full accent-blue-500"
              />
              <div className="mt-1 flex justify-between text-xs text-zinc-300">
                <span>0%</span><span>25%</span><span>50%</span>
              </div>
            </div>
          </Card>

          <SectionLabel>Desglose por participante</SectionLabel>
          <Card>
            <Row label="Gastos directos s/IVA"     value={fmt(numParticipantes > 0 ? costes.sinIVA / numParticipantes : 0)} sub="repartidos entre participantes" />
            <Row label="IVA soportado"              value={fmt(numParticipantes > 0 ? costes.ivaGastos / numParticipantes : 0)} />
            <Row label="Monitores (coste empresa)"  value={fmt(numParticipantes > 0 ? costes.salarios / numParticipantes : 0)} />
            <Row label="Fee Gimeloos"               value={fmt(gimeloosFee)} />
            <Row label="Coste/plaza sin IVA"        value={fmt(costeSinIVAPorP)} valueClass="text-zinc-700 font-bold" />
            <Row label="Coste/plaza con IVA"        value={fmt(costePorPart)}    valueClass="text-zinc-900 font-bold" />
          </Card>

          <SectionLabel>Resultado</SectionLabel>
          <Card>
            <Row label="Precio de venta por plaza"  value={fmt(precioVenta)}      valueClass="text-blue-600 text-base" />
            <Row label="Beneficio por plaza"        value={fmt(precioVenta - costePorPart)} valueClass={precioVenta - costePorPart >= 0 ? "text-green-600" : "text-red-600"} />
            <Row label="Ingresos totales"           value={fmt(ingresosTotales)} />
            <Row label="Coste total"                value={fmt(costes.total)} />
            <Row label="Beneficio total"            value={fmt(beneficioTotal)}   valueClass={beneficioTotal >= 0 ? "text-green-600" : "text-red-600"} />
            <Row label="Margen real"                value={pct(margenReal)}        valueClass={margenReal >= margen ? "text-green-600" : "text-orange-500"} last />
          </Card>

          <div className="mx-4 mt-4">
            <div
              className={`rounded-2xl border-2 p-4 ${
                esViable ? "border-green-400 bg-green-50" : "border-red-400 bg-red-50"
              }`}
            >
              <p className={`text-base font-bold ${esViable ? "text-green-700" : "text-red-700"}`}>
                {esViable ? "✓ Campamento viable" : "✗ Campamento no viable"}
              </p>
              <p className={`mt-1 text-sm ${esViable ? "text-green-600" : "text-red-600"}`}>
                {costes.total === 0
                  ? "Introduce gastos para calcular la viabilidad."
                  : esViable
                  ? `Ganas ${fmt(beneficioTotal)} con un margen del ${pct(margenReal)} sobre ingresos.`
                  : `Pierdes ${fmt(Math.abs(beneficioTotal))}. Sube el margen o reduce costes.`}
              </p>
            </div>
          </div>
        </>
      )}

      {/* ══ TABLA ══════════════════════════════════════════════════════════ */}
      {tab === "tabla" && (
        <>
          <SectionLabel>Rango de participantes</SectionLabel>
          <Card>
            <div className="flex items-center justify-between border-b border-zinc-100 px-4 py-3.5">
              <span className="text-sm text-zinc-600">Mínimo</span>
              <Stepper value={tablaMin} onChange={(v) => setTablaMin(Math.min(v, tablaMax))} min={1} />
            </div>
            <div className="flex items-center justify-between px-4 py-3.5">
              <span className="text-sm text-zinc-600">Máximo</span>
              <Stepper value={tablaMax} onChange={(v) => setTablaMax(Math.max(v, tablaMin))} />
            </div>
          </Card>

          <SectionLabel>Precio por plaza según asistentes</SectionLabel>
          <div className="mx-4 overflow-hidden rounded-2xl bg-white shadow-sm">
            {/* Cabecera tabla */}
            <div className="grid grid-cols-4 gap-0 bg-zinc-50 px-4 py-2">
              <span className="text-[10px] font-semibold uppercase tracking-wide text-zinc-400">Plazas</span>
              <span className="text-right text-[10px] font-semibold uppercase tracking-wide text-zinc-400">Coste total</span>
              <span className="text-right text-[10px] font-semibold uppercase tracking-wide text-zinc-400">Coste/plaza</span>
              <span className="text-right text-[10px] font-semibold uppercase tracking-wide text-zinc-400">Precio/plaza</span>
            </div>
            {tablaEscala.map((row) => {
              const isTarget = row.n === numParticipantes;
              return (
                <div
                  key={row.n}
                  className={`grid grid-cols-4 border-t border-zinc-100 px-4 py-2.5 text-xs ${
                    isTarget ? "bg-blue-50" : ""
                  }`}
                >
                  <span className={`font-medium ${isTarget ? "text-blue-600" : "text-zinc-700"}`}>
                    {row.n} {isTarget ? "◀" : ""}
                  </span>
                  <span className="text-right text-zinc-500">{fmt(row.costeTotal)}</span>
                  <span className="text-right text-zinc-500">{fmt(row.costePP)}</span>
                  <span
                    className={`text-right font-semibold ${
                      row.beneficio >= 0 ? "text-zinc-900" : "text-red-500"
                    }`}
                  >
                    {fmt(row.precio)}
                  </span>
                </div>
              );
            })}
          </div>

          <div className="mx-4 mt-3 rounded-xl bg-blue-50 p-3 text-xs text-blue-700">
            La tabla muestra el precio de venta por plaza para mantener el {margen}% de margen a
            distintos números de asistentes. La fila ◀ es tu estimación actual de{" "}
            {numParticipantes} participantes.
          </div>

          <SectionLabel>Análisis de break-even</SectionLabel>
          <Card>
            {(() => {
              // Encuentra el mínimo de participantes para obtener beneficio > 0
              const breakEvenRow = tablaEscala.find((r) => r.beneficio > 0);
              const minPart = breakEvenRow ? breakEvenRow.n : null;
              return (
                <>
                  <Row
                    label="Punto de equilibrio"
                    value={minPart ? `${minPart} participantes` : "—"}
                    sub="mínimo para no perder dinero"
                    valueClass="text-orange-500"
                  />
                  <Row
                    label="Con tu estimación actual"
                    value={beneficioTotal >= 0 ? `Ganas ${fmt(beneficioTotal)}` : `Pierdes ${fmt(Math.abs(beneficioTotal))}`}
                    sub={`${numParticipantes} participantes × ${fmt(precioVenta)}/plaza`}
                    valueClass={beneficioTotal >= 0 ? "text-green-600" : "text-red-600"}
                    last
                  />
                </>
              );
            })()}
          </Card>
        </>
      )}
    </div>
  );
}

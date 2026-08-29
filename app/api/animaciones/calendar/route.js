import { NextResponse } from "next/server";

export const runtime = "nodejs";

function parseHora(raw) {
  const s = (raw || "").trim().replace(",", ".");
  if (!s) return null;
  if (s.includes(":")) {
    const [h, m] = s.split(":").map(Number);
    return h + (m || 0) / 60;
  }
  if (s.includes(".")) {
    const [h, m] = s.split(".").map(Number);
    return h + (m || 0) / 60;
  }
  return Number(s);
}

function parseHorarioTimes(horarioStr) {
  if (!horarioStr) return null;
  const match = horarioStr.match(/^\s*([\d.,:]+)\s*(?:a|-)\s*([\d.,:]+)/i);
  if (!match) return null;
  let h1 = parseHora(match[1]);
  let h2 = parseHora(match[2]);
  if (h1 === null || h2 === null || isNaN(h1) || isNaN(h2)) return null;
  if (h1 < 8) h1 += 12;
  if (h2 <= h1) h2 += 12;
  return { inicio: h1, fin: h2 };
}

function toIcalDate(fechaStr, horaDecimal) {
  // fechaStr: "DD/MM/YYYY", horaDecimal: e.g. 15.5 = 15:30
  if (!fechaStr) return null;
  const parts = fechaStr.split("/");
  if (parts.length !== 3) return null;
  const [day, month, year] = parts.map((p) => p.padStart(2, "0"));
  const h = Math.floor(horaDecimal);
  const m = Math.round((horaDecimal - h) * 60);
  return `${year}${month}${day}T${String(h).padStart(2, "0")}${String(m).padStart(2, "0")}00`;
}

function escapeIcal(str) {
  return (str || "").replace(/\\/g, "\\\\").replace(/;/g, "\\;").replace(/,/g, "\\,").replace(/\n/g, "\\n");
}

function foldLine(line) {
  // iCal lines must be ≤75 octets; fold on character boundaries to avoid splitting UTF-8 sequences
  const bytes = Buffer.from(line, "utf-8");
  if (bytes.length <= 75) return line;
  const chunks = [];
  let pos = 0;
  while (pos < bytes.length) {
    let end = Math.min(pos + 75, bytes.length);
    // Step back until we're at a valid UTF-8 character boundary
    while (end > pos && (bytes[end] & 0xc0) === 0x80) end--;
    chunks.push(bytes.slice(pos, end).toString("utf-8"));
    pos = end;
  }
  return chunks.join("\r\n ");
}

export async function GET(request) {
  try {
    const origin = new URL(request.url).origin;
    const res = await fetch(`${origin}/api/animaciones/reservas`);
    if (!res.ok) throw new Error("No se pudieron cargar las reservas");
    const { reservas } = await res.json();

    const now = new Date().toISOString().replace(/[-:]/g, "").split(".")[0] + "Z";

    const events = reservas
      .map((r) => {
        const times = parseHorarioTimes(r.evento.horario);
        if (!times || !r.evento.fecha) return null;

        const dtStart = toIcalDate(r.evento.fecha, times.inicio);
        const dtEnd = toIcalDate(r.evento.fecha, times.fin);
        if (!dtStart || !dtEnd) return null;

        const nombre = [r.festejado.nombre, r.festejado.apellidos].filter(Boolean).join(" ") || "Sin nombre";
        const summary = `Animación · ${nombre}`;
        const desc = [
          r.evento.tipoEvento,
          r.festejado.edad ? `${r.festejado.edad} años` : null,
          r.participantes ? `${r.participantes} participantes` : null,
          r.monitoresEstimados ? `${r.monitoresEstimados} monitores` : null,
          r.contacto.nombre ? `Contacto: ${r.contacto.nombre}${r.contacto.telefono ? " · " + r.contacto.telefono : ""}` : null,
        ]
          .filter(Boolean)
          .join("\\n");

        const uid = `animacion-${r.id.replace(/[^a-zA-Z0-9]/g, "-")}@gimeloos.es`;

        return [
          "BEGIN:VEVENT",
          `DTSTART;TZID=Europe/Madrid:${dtStart}`,
          `DTEND;TZID=Europe/Madrid:${dtEnd}`,
          foldLine(`SUMMARY:${escapeIcal(summary)}`),
          foldLine(`DESCRIPTION:${escapeIcal(desc)}`),
          foldLine(`LOCATION:${escapeIcal(r.evento.lugar)}`),
          `UID:${uid}`,
          `DTSTAMP:${now}`,
          "END:VEVENT",
        ].join("\r\n");
      })
      .filter(Boolean);

    const ical = [
      "BEGIN:VCALENDAR",
      "VERSION:2.0",
      "PRODID:-//GIMELOOS//Portal Animaciones//ES",
      "CALSCALE:GREGORIAN",
      "METHOD:PUBLISH",
      "X-WR-CALNAME:Animaciones GIMELOOS",
      "X-WR-TIMEZONE:Europe/Madrid",
      "REFRESH-INTERVAL;VALUE=DURATION:PT1H",
      "X-PUBLISHED-TTL:PT1H",
      ...events,
      "END:VCALENDAR",
    ].join("\r\n");

    return new NextResponse(ical, {
      headers: {
        "Content-Type": "text/calendar; charset=utf-8",
        "Content-Disposition": 'attachment; filename="animaciones-gimeloos.ics"',
        "Cache-Control": "no-cache",
      },
    });
  } catch (err) {
    console.error("calendar route error:", err);
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}

import { NextResponse } from "next/server";
import { google } from "googleapis";

export const runtime = "nodejs";

const SPREADSHEET_ID = process.env.ANIMACIONES_SPREADSHEET_ID || "16o4zFk6K942l9_APkRMK0S-nueIdc2O95Kth4PPXdyM";
const SHEET_TAB = "CURSO 26-27";

// Índices de columna (0-indexado, A=0) confirmados manualmente con el usuario.
// Solo se leen estas columnas — el resto de la hoja no se toca.
const COL = {
  celebrada: 3, // D
  nombreContacto: 5, // F
  telefono: 7, // H
  fecha: 10, // K
  tipoEvento: 11, // L
  nombreFestejado: 12, // M
  apellidosFestejado: 13, // N
  edadFestejado: 14, // O
  colegio: 15, // P
  pack: 19, // T
  lugar: 22, // W
  participantes: 23, // X
  horario: 29, // AD
  tematica: 30, // AE
};
const MAX_COL_INDEX = COL.tematica;

function getAuth() {
  return new google.auth.JWT({
    email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
    key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, "\n"),
    scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"],
  });
}

// Convierte "3.30", "16.00", "7", "14" en horas decimales (3.5, 16, 7, 14)
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

// Calcula duración en horas a partir de cadenas tipo "3.30 a 7", "14-19", "16.00-20.00"
function parseHorario(horarioStr) {
  if (!horarioStr) return null;
  const match = horarioStr.match(/^\s*([\d.,:]+)\s*(?:a|-)\s*([\d.,:]+)/i);
  if (!match) return null;

  let h1 = parseHora(match[1]);
  let h2 = parseHora(match[2]);
  if (h1 === null || h2 === null || isNaN(h1) || isNaN(h2)) return null;

  // Los eventos son de tarde/noche: si la hora de inicio es muy baja, es formato 12h (ej. "3.30" = 15:30)
  if (h1 < 8) h1 += 12;
  if (h2 <= h1) h2 += 12;

  const horas = h2 - h1;
  return horas > 0 ? horas : null;
}

// Regla de negocio: 1h→5, 2h→8, 3h→10, +2 por cada hora adicional a partir de 2h
function calcularCapacidadJuegos(horas) {
  if (!horas || horas <= 0) return null;
  if (horas <= 1) return 5;
  if (horas <= 2) return Math.round(5 + 3 * (horas - 1));
  return Math.round(8 + 2 * (horas - 2));
}

function calcularMonitores(participantes) {
  if (!participantes || participantes <= 0) return null;
  return Math.ceil(participantes / 8);
}

export async function GET(request) {
  try {
    const auth = getAuth();
    await auth.authorize();
    const sheets = google.sheets({ version: "v4", auth });

    const { data } = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${SHEET_TAB}'!A2:AE1002`,
    });

    const rows = data.values || [];
    const get = (row, idx) => (idx < row.length ? (row[idx] || "").toString().trim() : "");

    const reservas = rows
      .filter((row) => row.length > COL.celebrada && get(row, COL.celebrada).toLowerCase() === "si")
      .map((row, i) => {
        const participantesRaw = get(row, COL.participantes);
        const participantes = parseInt(participantesRaw.replace(/[^\d]/g, ""), 10) || null;
        const horarioStr = get(row, COL.horario);
        const horas = parseHorario(horarioStr);

        return {
          id: `${i}-${get(row, COL.nombreFestejado)}-${get(row, COL.fecha)}`,
          contacto: {
            nombre: get(row, COL.nombreContacto),
            telefono: get(row, COL.telefono),
          },
          festejado: {
            nombre: get(row, COL.nombreFestejado),
            apellidos: get(row, COL.apellidosFestejado),
            edad: get(row, COL.edadFestejado) || null,
          },
          evento: {
            tipoEvento: get(row, COL.tipoEvento),
            fecha: get(row, COL.fecha),
            horario: horarioStr,
            horasDuracion: horas,
            lugar: get(row, COL.lugar),
            colegio: get(row, COL.colegio) || null,
            pack: get(row, COL.pack) || null,
            tematica: get(row, COL.tematica),
          },
          participantes,
          monitoresEstimados: calcularMonitores(participantes),
          capacidadJuegos: calcularCapacidadJuegos(horas),
        };
      });

    return NextResponse.json({ ok: true, total: reservas.length, reservas });
  } catch (err) {
    console.error("animaciones/reservas error:", err);
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}

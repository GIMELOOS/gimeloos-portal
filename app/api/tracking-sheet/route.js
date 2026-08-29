import { NextResponse } from "next/server";
import { google } from "googleapis";
import { readFileSync, existsSync, writeFileSync } from "fs";
import { join } from "path";
import { requireAdmin } from "@/lib/api-auth";

export const runtime = "nodejs";

const TOKEN_PATH = join(process.cwd(), ".google-token.json");
const PARENT_FOLDER_ID = process.env.GOOGLE_DRIVE_FOLDER_ID;

// Columnas: 0=Nombre, 1=Email, 2=Teléfono, 3=DNI, 4=Reserva, 5=1ª Cuota, 6=2ª Cuota
const HEADERS = ["Nombre", "Email", "Teléfono", "DNI", "Reserva", "1ª Cuota", "2ª Cuota"];
const PAY_COL = { reservation: 4, firstInstallment: 5, secondInstallment: 6 };
const COL_LETTER = (n) => String.fromCharCode(65 + n); // A=0, B=1, ...

function getClients() {
  if (!existsSync(TOKEN_PATH)) throw new Error("Google Drive no autorizado. Visita /api/auth/google-setup.");
  const tokens = JSON.parse(readFileSync(TOKEN_PATH, "utf-8"));
  const oauth2 = new google.auth.OAuth2(
    process.env.GOOGLE_OAUTH_CLIENT_ID,
    process.env.GOOGLE_OAUTH_CLIENT_SECRET,
    `${process.env.NEXT_PUBLIC_APP_URL || "http://localhost:3000"}/api/auth/google-callback`
  );
  oauth2.setCredentials(tokens);
  oauth2.on("tokens", (t) => writeFileSync(TOKEN_PATH, JSON.stringify({ ...tokens, ...t }, null, 2)));
  return {
    drive: google.drive({ version: "v3", auth: oauth2 }),
    sheets: google.sheets({ version: "v4", auth: oauth2 }),
  };
}

function bgColor(status) {
  if (status === "confirmed") return { red: 0.204, green: 0.659, blue: 0.325 }; // verde
  if (status === "sent")      return { red: 1,     green: 0.757, blue: 0.027 }; // ámbar
  if (status === "rejected")  return { red: 0.957, green: 0.263, blue: 0.212 }; // rojo oscuro
  return { red: 0.918, green: 0.345, blue: 0.192 }; // rojo pendiente
}

function cellLabel(status, amount) {
  const fmt = amount ? new Intl.NumberFormat("es-ES", { style: "currency", currency: "EUR" }).format(amount) : "";
  if (status === "confirmed") return fmt ? `✓ ${fmt}` : "✓ Confirmado";
  if (status === "sent")      return fmt ? `→ ${fmt}` : "→ Enviado";
  if (status === "rejected")  return "✗ Rechazado";
  return fmt ? fmt : "Pendiente";
}

async function findOrCreateSheet(drive, sheets, tripName) {
  const name = `GIMELOOS · Seguimiento — ${tripName}`;
  const escaped = name.replace(/'/g, "\\'");
  const { data } = await drive.files.list({
    q: `name='${escaped}' and mimeType='application/vnd.google-apps.spreadsheet' and '${PARENT_FOLDER_ID}' in parents and trashed=false`,
    fields: "files(id,webViewLink)",
  });
  if (data.files?.length) return { id: data.files[0].id, url: data.files[0].webViewLink };

  // Crear nueva hoja
  const created = await drive.files.create({
    requestBody: { name, mimeType: "application/vnd.google-apps.spreadsheet", parents: [PARENT_FOLDER_ID] },
    fields: "id,webViewLink",
  });
  const id = created.data.id;
  // Cabeceras
  await sheets.spreadsheets.values.update({
    spreadsheetId: id, range: `A1:${COL_LETTER(HEADERS.length - 1)}1`, valueInputOption: "USER_ENTERED",
    requestBody: { values: [HEADERS] },
  });
  // Nota: si la hoja ya existía con el formato anterior (6 columnas), se puede actualizar manualmente.
  // Formato cabecera
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: {
      requests: [{
        repeatCell: {
          range: { startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: HEADERS.length },
          cell: {
            userEnteredFormat: {
              backgroundColor: { red: 0.133, green: 0.133, blue: 0.133 },
              textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 } },
            },
          },
          fields: "userEnteredFormat(backgroundColor,textFormat)",
        },
      }],
    },
  });
  return { id, url: created.data.webViewLink };
}

async function getParticipantRow(sheets, spreadsheetId, participantName) {
  const { data } = await sheets.spreadsheets.values.get({ spreadsheetId, range: "A:A" });
  const rows = data.values || [];
  for (let i = 1; i < rows.length; i++) {
    if ((rows[i][0] || "").toLowerCase().trim() === participantName.toLowerCase().trim()) return i + 1; // 1-indexed
  }
  return null;
}

// POST /api/tracking-sheet
export async function POST(request) {
  const { error: authError } = await requireAdmin(request);
  if (authError) return authError;

  try {
    const body = await request.json();
    const { action, tripName, participants, participantName, participantEmail, participantPhone, participantDni, paymentKey, paymentStatus, paymentAmount } = body;

    if (!tripName) return NextResponse.json({ error: "tripName requerido" }, { status: 400 });

    const { drive, sheets } = getClients();
    const { id: spreadsheetId, url: sheetUrl } = await findOrCreateSheet(drive, sheets, tripName);

    // ── Actualizar un pago concreto ────────────────────────────────────────────
    if (action === "update_payment") {
      if (!participantName || !paymentKey || !paymentStatus) {
        return NextResponse.json({ error: "participantName, paymentKey, paymentStatus requeridos" }, { status: 400 });
      }

      let rowIndex = await getParticipantRow(sheets, spreadsheetId, participantName);

      // Si no existe aún, añadir fila
      if (!rowIndex) {
        await sheets.spreadsheets.values.append({
          spreadsheetId, range: "A:G", valueInputOption: "USER_ENTERED",
          requestBody: { values: [[participantName, participantEmail || "", participantPhone || "", participantDni || "", "", "", ""]] },
        });
        rowIndex = await getParticipantRow(sheets, spreadsheetId, participantName);
      }
      if (!rowIndex) return NextResponse.json({ error: "No se pudo localizar la fila" }, { status: 500 });

      const col = PAY_COL[paymentKey];
      if (col === undefined) return NextResponse.json({ error: "paymentKey inválido" }, { status: 400 });

      // Valor de la celda
      await sheets.spreadsheets.values.update({
        spreadsheetId, range: `${COL_LETTER(col)}${rowIndex}`, valueInputOption: "USER_ENTERED",
        requestBody: { values: [[cellLabel(paymentStatus, paymentAmount)]] },
      });

      // Color de fondo
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [{
            repeatCell: {
              range: { startRowIndex: rowIndex - 1, endRowIndex: rowIndex, startColumnIndex: col, endColumnIndex: col + 1 },
              cell: {
                userEnteredFormat: {
                  backgroundColor: bgColor(paymentStatus),
                  textFormat: { bold: paymentStatus === "confirmed" },
                },
              },
              fields: "userEnteredFormat(backgroundColor,textFormat)",
            },
          }],
        },
      });

      return NextResponse.json({ ok: true, sheetUrl });
    }

    // ── Sincronización completa de todos los participantes ─────────────────────
    if (action === "sync") {
      if (!participants?.length) return NextResponse.json({ ok: true, sheetUrl, message: "Sin participantes" });

      // Limpiar filas de datos (mantener cabecera)
      await sheets.spreadsheets.values.clear({ spreadsheetId, range: "A2:Z" });

      // Escribir valores
      const values = participants.map((p) => [
        p.participantName || "",
        p.email || "",
        p.phone || "",
        p.dni || "",
        cellLabel(p.payments?.reservation?.status || "pending", p.payments?.reservation?.amount),
        cellLabel(p.payments?.firstInstallment?.status || "pending", p.payments?.firstInstallment?.amount),
        cellLabel(p.payments?.secondInstallment?.status || "pending", p.payments?.secondInstallment?.amount),
      ]);

      await sheets.spreadsheets.values.update({
        spreadsheetId, range: `A2:G${1 + participants.length}`, valueInputOption: "USER_ENTERED",
        requestBody: { values },
      });

      // Colorear celdas de pago
      const requests = [];
      participants.forEach((p, i) => {
        Object.entries(PAY_COL).forEach(([key, col]) => {
          const status = p.payments?.[key]?.status || "pending";
          requests.push({
            repeatCell: {
              range: { startRowIndex: i + 1, endRowIndex: i + 2, startColumnIndex: col, endColumnIndex: col + 1 },
              cell: {
                userEnteredFormat: {
                  backgroundColor: bgColor(status),
                  textFormat: { bold: status === "confirmed" },
                },
              },
              fields: "userEnteredFormat(backgroundColor,textFormat)",
            },
          });
        });
      });

      if (requests.length) {
        await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests } });
      }

      return NextResponse.json({ ok: true, sheetUrl });
    }

    return NextResponse.json({ error: "Acción desconocida" }, { status: 400 });
  } catch (err) {
    console.error("tracking-sheet error:", err);
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}

import { NextResponse } from "next/server";
import { google } from "googleapis";
import { createClient } from "@supabase/supabase-js";
import { requireAdmin } from "@/lib/api-auth";

export const runtime = "nodejs";

const supabaseAdmin = createClient(
  process.env.NEXT_PUBLIC_SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY
);

async function getDriveClient() {
  const { data, error } = await supabaseAdmin
    .from("app_settings")
    .select("value")
    .eq("key", "google_drive_token")
    .single();
  if (error || !data) return null;

  const tokens = JSON.parse(data.value);
  const appUrl = process.env.NEXT_PUBLIC_APP_URL || "http://localhost:3000";
  const oauth2Client = new google.auth.OAuth2(
    process.env.GOOGLE_OAUTH_CLIENT_ID,
    process.env.GOOGLE_OAUTH_CLIENT_SECRET,
    `${appUrl}/api/auth/google-callback`
  );
  oauth2Client.setCredentials(tokens);
  oauth2Client.on("tokens", async (newTokens) => {
    const updated = { ...tokens, ...newTokens };
    await supabaseAdmin.from("app_settings").upsert(
      { key: "google_drive_token", value: JSON.stringify(updated) },
      { onConflict: "key" }
    );
  });
  return google.drive({ version: "v3", auth: oauth2Client });
}

export async function GET(request) {
  const { error: authError } = await requireAdmin(request);
  if (authError) return authError;

  const { searchParams } = new URL(request.url);
  const url = searchParams.get("url");

  if (!url) {
    return NextResponse.json({ error: "Falta el parámetro url" }, { status: 400 });
  }

  try {
    // Si es una URL de Google Sheets, usar la API de Drive con autenticación OAuth
    const gSheetMatch = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
    if (gSheetMatch) {
      const fileId = gSheetMatch[1];
      const drive = await getDriveClient();
      if (!drive) {
        return NextResponse.json(
          { error: "Google Drive no está conectado. Visita /api/auth/google-setup para conectarlo." },
          { status: 500 }
        );
      }
      const driveRes = await drive.files.export(
        { fileId, mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
        { responseType: "arraybuffer" }
      );
      return new NextResponse(driveRes.data, {
        headers: {
          "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          "Content-Disposition": "attachment; filename=sheet.xlsx",
        },
      });
    }

    // Para URLs de Excel normales, descarga directa
    const res = await fetch(url, { headers: { "User-Agent": "Mozilla/5.0" } });
    if (!res.ok) {
      return NextResponse.json(
        { error: `No se pudo descargar el archivo (${res.status}).` },
        { status: 400 }
      );
    }
    const buffer = await res.arrayBuffer();
    return new NextResponse(buffer, {
      headers: {
        "Content-Type": res.headers.get("Content-Type") || "application/octet-stream",
        "Content-Disposition": "attachment; filename=sheet.xlsx",
      },
    });
  } catch (err) {
    return NextResponse.json({ error: "Error al descargar: " + err.message }, { status: 500 });
  }
}

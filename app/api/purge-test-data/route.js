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
  const { data } = await supabaseAdmin
    .from("app_settings")
    .select("value")
    .eq("key", "google_drive_token")
    .single();
  if (!data) return null;

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

// Extrae el fileId de Drive de una URL (varios formatos)
function extractFileId(url) {
  if (!url || typeof url !== "string") return null;
  // https://drive.google.com/uc?export=view&id=FILEID
  const ucMatch = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (ucMatch) return ucMatch[1];
  // https://drive.google.com/file/d/FILEID/view
  const fileMatch = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (fileMatch) return fileMatch[1];
  // https://drive.google.com/open?id=FILEID
  const openMatch = url.match(/open\?id=([a-zA-Z0-9_-]+)/);
  if (openMatch) return openMatch[1];
  return null;
}

// Recoge todos los fileIds referenciados en la BD
async function collectDriveFileIds() {
  const ids = new Set();

  // Documentos y pruebas de pago de participantes
  const { data: participants } = await supabaseAdmin
    .from("participants")
    .select("documents, invoice_url, proof_reservation_url, proof_first_url, proof_second_url");

  for (const p of participants || []) {
    for (const field of ["invoice_url", "proof_reservation_url", "proof_first_url", "proof_second_url"]) {
      const id = extractFileId(p[field]);
      if (id) ids.add(id);
    }
    if (Array.isArray(p.documents)) {
      for (const doc of p.documents) {
        const id = doc.fileId || extractFileId(doc.driveUrl) || extractFileId(doc.url);
        if (id) ids.add(id);
      }
    }
  }

  // Documentos escolares
  const { data: schoolDocs } = await supabaseAdmin
    .from("school_documents")
    .select("file_url");
  for (const d of schoolDocs || []) {
    const id = extractFileId(d.file_url);
    if (id) ids.add(id);
  }

  // Fotos de portada de trips (solo las de Drive, no las de Unsplash)
  const { data: trips } = await supabaseAdmin
    .from("trips")
    .select("hero_image, hero_images");
  for (const t of trips || []) {
    const id = extractFileId(t.hero_image);
    if (id) ids.add(id);
    if (Array.isArray(t.hero_images)) {
      for (const img of t.hero_images) {
        const imgId = extractFileId(typeof img === "string" ? img : img?.url);
        if (imgId) ids.add(imgId);
      }
    }
  }

  // Fotos de portada de school_trips
  const { data: schoolTrips } = await supabaseAdmin
    .from("school_trips")
    .select("hero_image, hero_images");
  for (const t of schoolTrips || []) {
    const id = extractFileId(t.hero_image);
    if (id) ids.add(id);
    if (Array.isArray(t.hero_images)) {
      for (const img of t.hero_images) {
        const imgId = extractFileId(typeof img === "string" ? img : img?.url);
        if (imgId) ids.add(imgId);
      }
    }
  }

  return [...ids];
}

// Borra archivos de Drive (best-effort, no aborta si falla alguno)
async function deleteFromDrive(drive, fileIds) {
  const results = { deleted: [], failed: [] };
  for (const fileId of fileIds) {
    try {
      await drive.files.delete({ fileId });
      results.deleted.push(fileId);
    } catch (err) {
      results.failed.push({ fileId, error: err.message });
    }
  }
  return results;
}

// Borra datos de Supabase en orden correcto (hijos antes que padres)
async function purgeSupabase() {
  const errors = [];

  const tables = [
    "participant_questions",
    "notifications",
    "participants",
    "school_questions",
    "school_documents",
    "students",
    "school_courses",
    "school_trips",
    "schools",
    "trips",
  ];

  for (const table of tables) {
    const { error } = await supabaseAdmin.from(table).delete().neq("id", "00000000-0000-0000-0000-000000000000");
    if (error) errors.push({ table, error: error.message });
  }

  return errors;
}

export async function POST(request) {
  const { error: authError } = await requireAdmin(request);
  if (authError) return authError;

  const log = { driveFileIds: [], driveResults: null, supabaseErrors: [] };

  // 1. Recoger fileIds
  try {
    log.driveFileIds = await collectDriveFileIds();
  } catch (err) {
    return NextResponse.json({ error: "Error recogiendo fileIds: " + err.message }, { status: 500 });
  }

  // 2. Borrar de Drive (opcional — si no hay cliente Drive disponible, se omite)
  const drive = await getDriveClient();
  if (drive && log.driveFileIds.length > 0) {
    log.driveResults = await deleteFromDrive(drive, log.driveFileIds);
  } else {
    log.driveResults = { skipped: true, reason: drive ? "No hay archivos" : "Drive no conectado" };
  }

  // 3. Borrar datos de Supabase
  log.supabaseErrors = await purgeSupabase();

  return NextResponse.json({
    ok: true,
    driveFilesFound: log.driveFileIds.length,
    driveResults: log.driveResults,
    supabaseErrors: log.supabaseErrors,
  });
}

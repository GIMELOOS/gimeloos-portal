import { NextResponse } from "next/server";
import { google } from "googleapis";
import { readFileSync, existsSync } from "fs";
import { join } from "path";
import { requireAdmin } from "@/lib/api-auth";

export const runtime = "nodejs";

const TOKEN_PATH = join(process.cwd(), ".google-token.json");
const PARENT_FOLDER_ID = process.env.GOOGLE_DRIVE_FOLDER_ID;

function getDriveClient() {
  if (!existsSync(TOKEN_PATH)) {
    throw new Error("Google Drive no está autorizado. Visita /api/auth/google-setup para conectarlo.");
  }
  const tokens = JSON.parse(readFileSync(TOKEN_PATH, "utf-8"));
  const appUrl = process.env.NEXT_PUBLIC_APP_URL || "http://localhost:3000";
  const oauth2Client = new google.auth.OAuth2(
    process.env.GOOGLE_OAUTH_CLIENT_ID,
    process.env.GOOGLE_OAUTH_CLIENT_SECRET,
    `${appUrl}/api/auth/google-callback`
  );
  oauth2Client.setCredentials(tokens);

  // Actualiza el token si se refresca automáticamente
  oauth2Client.on("tokens", (newTokens) => {
    const updated = { ...tokens, ...newTokens };
    try { require("fs").writeFileSync(TOKEN_PATH, JSON.stringify(updated, null, 2)); } catch (_) {}
  });

  return google.drive({ version: "v3", auth: oauth2Client });
}

async function findOrCreateFolder(drive, name, parentId) {
  const escaped = name.replace(/'/g, "\\'");
  const res = await drive.files.list({
    q: `name='${escaped}' and mimeType='application/vnd.google-apps.folder' and '${parentId}' in parents and trashed=false`,
    fields: "files(id)",
    spaces: "drive",
  });
  if (res.data.files?.length > 0) return res.data.files[0].id;
  const created = await drive.files.create({
    requestBody: { name, mimeType: "application/vnd.google-apps.folder", parents: [parentId] },
    fields: "id",
  });
  return created.data.id;
}

// POST /api/upload-to-drive
// body: FormData { file, username, subfolder ("documentos" | "pagos") }
export async function POST(request) {
  const { error: authError } = await requireAdmin(request);
  if (authError) return authError;

  try {
    const formData = await request.formData();
    const file = formData.get("file");
    const participantNameRaw = formData.get("participantName") || formData.get("username");
    const subfolder = formData.get("subfolder") || "documentos";

    if (!file || !participantNameRaw) {
      return NextResponse.json({ error: "Faltan parámetros" }, { status: 400 });
    }

    const MAX_BYTES = 20 * 1024 * 1024; // 20 MB
    const ALLOWED_TYPES = new Set([
      "application/pdf",
      "image/jpeg", "image/png", "image/webp", "image/gif", "image/heic", "image/heif",
      "application/msword",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "application/vnd.ms-excel",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ]);

    if (!ALLOWED_TYPES.has(file.type)) {
      return NextResponse.json({ error: `Tipo de archivo no permitido: ${file.type}` }, { status: 400 });
    }
    if (file.size > MAX_BYTES) {
      return NextResponse.json({ error: "El archivo supera el límite de 20 MB" }, { status: 400 });
    }

    const drive = getDriveClient();

    // Carpeta del campamento (opcional) → carpeta del participante → subcarpeta
    const tripName = formData.get("tripName");
    const campFolderId = tripName
      ? await findOrCreateFolder(drive, tripName, PARENT_FOLDER_ID)
      : PARENT_FOLDER_ID;
    const participantName = participantNameRaw;
    const participantFolderId = await findOrCreateFolder(drive, participantName, campFolderId);
    const subFolderId = await findOrCreateFolder(drive, subfolder, participantFolderId);

    // Subir archivo
    const arrayBuffer = await file.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);
    const { Readable } = await import("stream");
    const stream = Readable.from(buffer);

    const uploaded = await drive.files.create({
      requestBody: {
        name: file.name,
        parents: [subFolderId],
      },
      media: {
        mimeType: file.type || "application/octet-stream",
        body: stream,
      },
      fields: "id, webViewLink, webContentLink, name",
    });

    const fileId = uploaded.data.id;

    // Hacer el archivo público para que la URL funcione como imagen directa
    await drive.permissions.create({
      fileId,
      requestBody: { role: "reader", type: "anyone" },
    });

    // URL directa de imagen (sustituye el &export=download por thumbnail o direct)
    const directUrl = `https://drive.google.com/uc?export=view&id=${fileId}`;

    return NextResponse.json({
      fileId,
      webViewLink: uploaded.data.webViewLink,
      webContentLink: directUrl,
      fileName: uploaded.data.name,
    });
  } catch (err) {
    console.error("upload-to-drive error:", err);
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}

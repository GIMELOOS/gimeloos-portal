import { NextResponse } from "next/server";
import { google } from "googleapis";
import { createClient } from "@supabase/supabase-js";

export const runtime = "nodejs";

const supabaseAdmin = createClient(
  process.env.NEXT_PUBLIC_SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY
);

export async function GET(request) {
  const { searchParams } = new URL(request.url);
  const code = searchParams.get("code");
  const state = searchParams.get("state");

  if (!code) {
    return NextResponse.json({ error: "No se recibió código de autorización" }, { status: 400 });
  }

  if (!process.env.GOOGLE_SETUP_SECRET || state !== process.env.GOOGLE_SETUP_SECRET) {
    return NextResponse.json({ error: "No autorizado" }, { status: 401 });
  }

  const appUrl = process.env.NEXT_PUBLIC_APP_URL || "http://localhost:3000";
  const oauth2Client = new google.auth.OAuth2(
    process.env.GOOGLE_OAUTH_CLIENT_ID,
    process.env.GOOGLE_OAUTH_CLIENT_SECRET,
    `${appUrl}/api/auth/google-callback`
  );

  const { tokens } = await oauth2Client.getToken(code);

  if (!tokens.refresh_token) {
    return NextResponse.json({
      error: "No se obtuvo refresh_token. Revoca el acceso en myaccount.google.com/permissions y vuelve a intentarlo.",
    }, { status: 400 });
  }

  // Guardar token en Supabase en vez de en disco
  await supabaseAdmin.from("app_settings").upsert(
    { key: "google_drive_token", value: JSON.stringify(tokens) },
    { onConflict: "key" }
  );

  return new NextResponse(
    `<html><body style="font-family:sans-serif;padding:40px">
      <h2>✅ Google Drive conectado correctamente</h2>
      <p>Ya puedes cerrar esta ventana y volver al portal.</p>
    </body></html>`,
    { headers: { "Content-Type": "text/html" } }
  );
}

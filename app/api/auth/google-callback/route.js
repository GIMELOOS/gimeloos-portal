import { NextResponse } from "next/server";
import { google } from "googleapis";
import { writeFileSync, readFileSync } from "fs";
import { join } from "path";

export const runtime = "nodejs";

const TOKEN_PATH = join(process.cwd(), ".google-token.json");

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

  writeFileSync(TOKEN_PATH, JSON.stringify(tokens, null, 2));

  return new NextResponse(
    `<html><body style="font-family:sans-serif;padding:40px">
      <h2>✅ Google Drive conectado correctamente</h2>
      <p>Ya puedes cerrar esta ventana y volver al portal.</p>
      <p style="color:#666;font-size:14px">Token guardado en <code>.google-token.json</code></p>
    </body></html>`,
    { headers: { "Content-Type": "text/html" } }
  );
}

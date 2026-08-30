import { NextResponse } from "next/server";
import { requireAdmin } from "@/lib/api-auth";

export const runtime = "nodejs";

export async function GET(request) {
  const { error: authError } = await requireAdmin(request);
  if (authError) return authError;

  const { searchParams } = new URL(request.url);
  const url = searchParams.get("url");

  if (!url) {
    return NextResponse.json({ error: "Falta el parámetro url" }, { status: 400 });
  }

  try {
    const res = await fetch(url, {
      headers: { "User-Agent": "Mozilla/5.0" },
    });

    if (!res.ok) {
      return NextResponse.json(
        { error: `No se pudo descargar el archivo (${res.status}). Asegúrate de que el documento es público.` },
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

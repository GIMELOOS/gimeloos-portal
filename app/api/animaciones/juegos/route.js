import { NextResponse } from "next/server";
import { readFileSync } from "fs";
import { join } from "path";

export const runtime = "nodejs";

const JUEGOS_PATH = join(process.cwd(), "data", "juegos.json");

export async function GET(request) {
  try {
    const juegos = JSON.parse(readFileSync(JUEGOS_PATH, "utf-8"))
      .filter((j) => j.activo)
      .sort((a, b) => b.veces_usado - a.veces_usado || a.nombre.localeCompare(b.nombre, "es"));

    return NextResponse.json({ ok: true, total: juegos.length, juegos });
  } catch (err) {
    console.error("animaciones/juegos error:", err);
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}

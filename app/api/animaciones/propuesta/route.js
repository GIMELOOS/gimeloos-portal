import { NextResponse } from "next/server";
import { execFile } from "child_process";
import { promisify } from "util";
import { readFileSync, existsSync } from "fs";
import { join } from "path";

export const runtime = "nodejs";

const execFileAsync = promisify(execFile);
const SCRIPT = join(process.cwd(), "scripts", "generar_propuesta.py");
const OUT_DIR = join(process.cwd(), "public", "propuestas");

function slugify(str) {
  return (str || "")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/[^a-zA-Z0-9]+/g, "-")
    .replace(/^-|-$/g, "")
    .toLowerCase()
    .slice(0, 40);
}

export async function POST(request) {
  try {
    const { reserva, juegos } = await request.json();

    if (!reserva || !juegos?.length) {
      return NextResponse.json({ error: "Faltan datos (reserva o juegos)" }, { status: 400 });
    }

    const nombre = [reserva.festejado?.nombre, reserva.festejado?.apellidos]
      .filter(Boolean)
      .join(" ");
    const filename = `${slugify(nombre)}-${slugify(reserva.evento?.fecha || "")}.docx`;
    const outputPath = join(OUT_DIR, filename);

    const payload = JSON.stringify({ reserva, juegos });

    const { stdout, stderr } = await execFileAsync("python3", [SCRIPT, payload, outputPath], {
      timeout: 30000,
    });

    if (stderr && !stdout.startsWith("OK:")) {
      console.error("propuesta script stderr:", stderr);
    }

    if (!existsSync(outputPath)) {
      throw new Error("El script no generó el archivo");
    }

    const fileBytes = readFileSync(outputPath);

    return new NextResponse(fileBytes, {
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "Content-Disposition": `attachment; filename="${filename}"`,
      },
    });
  } catch (err) {
    console.error("propuesta route error:", err);
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}

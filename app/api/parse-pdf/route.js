import { NextResponse } from "next/server";
import { requireAuth } from "@/lib/api-auth";

export const runtime = "nodejs";

// Formato columnas: <nº> | <nºCurso><Etapa><Grupo> | <Apellido1, Apellido2, Nombre>
// Ej: "12ESOBBAGNASCO RUDÓN, VIOLETA"  → row=1, curso=2ESOB, alumno=BAGNASCO RUDÓN, VIOLETA
const STUDENT_RE = /^(\d{1,3})(\d(?:ESO|BACH|Bachillerato|FP|Primaria|Infantil|Secundaria)[A-Z]?)([A-ZÁÉÍÓÚÜÑ].+,.+)$/i;

// Líneas de metadatos a ignorar (cabeceras, pie de página, info del colegio...)
const SKIP_RE = /^(colegio|cl[\s,]|cp[\s,]|\d{5}[\s-]|matr[ií]c|curso\s+escolar|listado\s+(emitido|de)|secci[oó]n[:\s]|tutor[:\s]|n[ºo°][\s.]?secci|alumno\/a|p[aá]gina|fecha|viaje|extraesco)/i;

// "Apellido1 Apellido2, Nombre" → { name, surname }
// El formato real del PDF es "APELLIDO1 APELLIDO2, NOMBRE" (un solo campo con coma)
// Los dos apellidos van juntos en el campo surname
function parseName(raw) {
  const str = raw.trim();
  const commaIdx = str.indexOf(",");
  if (commaIdx === -1) return { name: str, surname: "" };
  const surname = str.slice(0, commaIdx).trim();
  const name = str.slice(commaIdx + 1).trim();
  return { name, surname };
}

// Formatea el código de sección: "2ESOB" → "2º ESO B"
function formatCourse(code) {
  if (!code) return "";
  const m = code.match(/^(\d)(ESO|BACH(?:illerato)?|FP|Primaria|Infantil|Secundaria)([A-Z]?)$/i);
  if (!m) return code;
  const num = m[1];
  const stage = m[2].toUpperCase().replace("BACHILLERATO", "BACH");
  const group = m[3];
  return `${num}º ${stage}${group ? " " + group : ""}`;
}

export async function POST(request) {
  const { error: authError } = await requireAuth(request);
  if (authError) return authError;

  try {
    const formData = await request.formData();
    const file = formData.get("file");
    if (!file) return NextResponse.json({ error: "Falta el archivo" }, { status: 400 });

    const arrayBuffer = await file.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);

    const pdfParse = (await import("pdf-parse/lib/pdf-parse.js")).default;
    const data = await pdfParse(buffer);

    const lines = data.text.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
    const students = [];

    for (const line of lines) {
      if (SKIP_RE.test(line) || line.length < 5) continue;

      // Patrón principal: nº + sección académica + alumno pegados
      const m = line.match(STUDENT_RE);
      if (m) {
        const courseHint = formatCourse(m[2]);
        const { name, surname } = parseName(m[3]);
        if (name && surname) {
          students.push({ name, surname, courseHint });
          continue;
        }
      }

      // Fallback: línea es solo "APELLIDOS, Nombre" (ya sin número ni sección)
      if (line.includes(",") && /^[A-ZÁÉÍÓÚÜÑ]/.test(line) && !line.match(/^\d/)) {
        const { name, surname } = parseName(line);
        if (name && surname && name.length > 1 && surname.length > 2) {
          students.push({ name, surname, courseHint: "" });
        }
      }
    }

    return NextResponse.json({ students, pageCount: data.numpages });
  } catch (err) {
    console.error("parse-pdf error:", err);
    return NextResponse.json({ error: "No se pudo parsear el PDF: " + err.message }, { status: 500 });
  }
}

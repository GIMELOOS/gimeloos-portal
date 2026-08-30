import { NextResponse } from "next/server";
import { requireAdmin } from "@/lib/api-auth";
import { createClient } from "@supabase/supabase-js";

const supabaseAdmin = createClient(
  process.env.NEXT_PUBLIC_SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY
);

export async function POST(request) {
  const { error: authError } = await requireAdmin(request);
  if (authError) return authError;

  const { id, fields } = await request.json();
  if (!id || !fields) return NextResponse.json({ error: "Faltan parámetros" }, { status: 400 });

  const { data, error } = await supabaseAdmin
    .from("trips")
    .update(fields)
    .eq("id", id)
    .select("id");

  if (error) return NextResponse.json({ error: error.message }, { status: 500 });
  if (!data?.length) return NextResponse.json({ error: "No se encontró el viaje o no se actualizó" }, { status: 404 });

  return NextResponse.json({ ok: true });
}

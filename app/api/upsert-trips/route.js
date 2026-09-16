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

  const { trips } = await request.json();
  if (!Array.isArray(trips) || !trips.length) {
    return NextResponse.json({ error: "Faltan parámetros" }, { status: 400 });
  }

  const ALLOWED_FIELDS = new Set([
    "id", "name", "departure_date", "hero_image", "hero_images", "description",
    "checklist", "itinerary", "logistics", "automation", "transfer_info",
    "tipo", "document_rules", "payment_schedule",
  ]);

  const safe = trips.map((t) =>
    Object.fromEntries(Object.entries(t).filter(([k]) => ALLOWED_FIELDS.has(k)))
  );

  const { error } = await supabaseAdmin
    .from("trips")
    .upsert(safe, { onConflict: "id" });

  if (error) return NextResponse.json({ error: error.message }, { status: 500 });

  return NextResponse.json({ ok: true });
}

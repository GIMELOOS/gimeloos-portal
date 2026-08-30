import { NextResponse } from "next/server";
import { requireAuth } from "@/lib/api-auth";
import { createClient } from "@supabase/supabase-js";

const supabaseAdmin = createClient(
  process.env.NEXT_PUBLIC_SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY
);

// Devuelve datos de trips (con service role key) para una lista de IDs.
// Accesible para cualquier usuario autenticado — los datos de viaje no son sensibles.
export async function POST(request) {
  const { error: authError } = await requireAuth(request);
  if (authError) return authError;

  const { trip_ids } = await request.json();
  if (!Array.isArray(trip_ids) || !trip_ids.length) {
    return NextResponse.json({ trips: [] });
  }

  const { data, error } = await supabaseAdmin
    .from("trips")
    .select("id, name, departure_date, hero_image, hero_images, description, transfer_info")
    .in("id", trip_ids);

  if (error) return NextResponse.json({ error: error.message }, { status: 500 });

  return NextResponse.json({ trips: data || [] });
}

import { NextResponse } from "next/server";
import { createClient } from "@supabase/supabase-js";
import * as XLSX from "xlsx";

export const runtime = "nodejs";

const supabase = createClient(
  process.env.NEXT_PUBLIC_SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY
);

const DEFAULT_HERO_IMAGE =
  "https://images.unsplash.com/photo-1500375592092-40eb2168fd21?auto=format&fit=crop&w=1600&q=80";

function safeString(...values) {
  for (const value of values) {
    if (value !== undefined && value !== null && String(value).trim() !== "") {
      return String(value).trim();
    }
  }
  return "";
}

function normalizeHeaderKey(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "");
}

function getRowValue(row, ...keys) {
  const normalizedEntries = Object.entries(row || {}).map(([key, value]) => ({
    normalizedKey: normalizeHeaderKey(key),
    value,
  }));

  for (const requestedKey of keys) {
    const normalizedRequestedKey = normalizeHeaderKey(requestedKey);
    const exactMatch = normalizedEntries.find(
      (entry) => entry.normalizedKey === normalizedRequestedKey
    );
    if (exactMatch && safeString(exactMatch.value) !== "") return exactMatch.value;
  }

  for (const requestedKey of keys) {
    const normalizedRequestedKey = normalizeHeaderKey(requestedKey);
    const partialMatch = normalizedEntries.find(
      (entry) =>
        entry.normalizedKey.includes(normalizedRequestedKey) ||
        normalizedRequestedKey.includes(entry.normalizedKey)
    );
    if (partialMatch && safeString(partialMatch.value) !== "") return partialMatch.value;
  }

  return "";
}

function parseAmount(value) {
  if (value === undefined || value === null || value === "") return 0;
  if (typeof value === "number") return value;

  const cleaned = String(value)
    .replace(/€/g, "")
    .replace(/\./g, "")
    .replace(/,/g, ".")
    .trim();

  const num = Number(cleaned);
  return Number.isFinite(num) ? num : 0;
}

function parseDateTime(value) {
  if (!value) return null;

  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return null;

    const date = new Date(
      Date.UTC(
        parsed.y,
        parsed.m - 1,
        parsed.d,
        parsed.H || 0,
        parsed.M || 0,
        parsed.S || 0
      )
    );

    return date.toISOString();
  }

  const asDate = new Date(String(value));
  if (Number.isNaN(asDate.getTime())) return null;
  return asDate.toISOString();
}

function slugify(value) {
  return String(value || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-|-$/g, "");
}

function defaultDocuments() {
  return [
    { id: "doc-1", status: "pending_upload", uploadedFileName: "" },
    { id: "doc-2", status: "pending_upload", uploadedFileName: "" },
    { id: "doc-3", status: "pending_upload", uploadedFileName: "" },
  ];
}

function normalizeTripForClient(trip) {
  return {
    id: trip.id,
    name: trip.name,
    departureDate: trip.departure_at,
    description: trip.description || "",
    heroImage: trip.hero_image || DEFAULT_HERO_IMAGE,
  };
}

async function readBootstrap() {
  const { data: trips, error: tripsError } = await supabase
    .from("trips")
    .select("id, name, departure_at, description, hero_image")
    .order("departure_at", { ascending: true });

  if (tripsError) {
    throw new Error(`No se han podido leer los viajes: ${tripsError.message}`);
  }

  const { data: participants, error: participantsError } = await supabase
    .from("participants")
    .select("id, username, participant_name, mother_name, father_name")
    .order("participant_name", { ascending: true });

  if (participantsError) {
    throw new Error(
      `No se han podido leer los participantes: ${participantsError.message}`
    );
  }

  const participantIds = participants.map((item) => item.id);

  let payments = [];
  if (participantIds.length) {
    const { data: paymentsData, error: paymentsError } = await supabase
      .from("participant_payments")
      .select("participant_id, payment_key, name, amount, status")
      .in("participant_id", participantIds);

    if (paymentsError) {
      throw new Error(`No se han podido leer los pagos: ${paymentsError.message}`);
    }

    payments = paymentsData || [];
  }

  const normalizedTrips = (trips || []).map(normalizeTripForClient);

  const tripByName = new Map(
    normalizedTrips.map((trip) => [trip.name.trim().toLowerCase(), trip])
  );

  const normalizedUsers = (participants || []).map((participant) => {
    const participantPayments = payments.filter(
      (payment) => payment.participant_id === participant.id
    );

    const paymentMap = Object.fromEntries(
      participantPayments.map((payment) => [payment.payment_key, payment])
    );

    return {
      id: participant.id,
      role: "client",
      username: participant.username,
      participantName: participant.participant_name || "",
      motherName: participant.mother_name || "",
      fatherName: participant.father_name || "",
      parentName: "",
      email: "",
      contactEmails: [],
      tripId: "",
      tripName: "",
      documents: defaultDocuments(),
      payments: {
        discount: 0,
        finalPrice:
          Number(paymentMap.reservation?.amount || 0) +
          Number(paymentMap.firstInstallment?.amount || 0) +
          Number(paymentMap.secondInstallment?.amount || 0),
        reservation: {
          name: paymentMap.reservation?.name || "Reserva",
          amount: Number(paymentMap.reservation?.amount || 0),
          status: paymentMap.reservation?.status || "pending",
          proofName: "",
          dueDate: "",
        },
        firstInstallment: {
          name: paymentMap.firstInstallment?.name || "Primera cuota",
          amount: Number(paymentMap.firstInstallment?.amount || 0),
          status: paymentMap.firstInstallment?.status || "pending",
          proofName: "",
          dueDate: "",
        },
        secondInstallment: {
          name: paymentMap.secondInstallment?.name || "Segunda cuota",
          amount: Number(paymentMap.secondInstallment?.amount || 0),
          status: paymentMap.secondInstallment?.status || "pending",
          proofName: "",
          dueDate: "",
        },
      },
      checklistState: {},
      questions: [],
    };
  });

  return { trips: normalizedTrips, users: normalizedUsers, tripByName };
}

export async function GET() {
  try {
    if (!process.env.SUPABASE_SERVICE_ROLE_KEY) {
      return NextResponse.json(
        { error: "Falta SUPABASE_SERVICE_ROLE_KEY en .env.local" },
        { status: 500 }
      );
    }

    const payload = await readBootstrap();
    return NextResponse.json({
      trips: payload.trips,
      users: payload.users,
    });
  } catch (error) {
    return NextResponse.json(
      { error: error.message || "Error interno" },
      { status: 500 }
    );
  }
}

export async function POST(request) {
  try {
    if (!process.env.SUPABASE_SERVICE_ROLE_KEY) {
      return NextResponse.json(
        { error: "Falta SUPABASE_SERVICE_ROLE_KEY en .env.local" },
        { status: 500 }
      );
    }

    const formData = await request.formData();
    const file = formData.get("file");

    if (!file) {
      return NextResponse.json(
        { error: "No se ha recibido ningún archivo." },
        { status: 400 }
      );
    }

    const bytes = await file.arrayBuffer();
    const workbook = XLSX.read(bytes, { type: "array", cellDates: true });
    const firstSheetName = workbook.SheetNames?.[0];

    if (!firstSheetName) {
      return NextResponse.json(
        { error: "El archivo no contiene hojas válidas." },
        { status: 400 }
      );
    }

    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName], {
      defval: "",
    });

    if (!rows.length) {
      return NextResponse.json(
        { error: "El Excel está vacío." },
        { status: 400 }
      );
    }

    const { data: existingTrips, error: existingTripsError } = await supabase
      .from("trips")
      .select("id, name, departure_at, description, hero_image");

    if (existingTripsError) throw new Error(existingTripsError.message);

    const { data: existingParticipants, error: existingParticipantsError } =
      await supabase.from("participants").select("id, username");

    if (existingParticipantsError) throw new Error(existingParticipantsError.message);

    const tripsByName = new Map(
      (existingTrips || []).map((trip) => [trip.name.trim().toLowerCase(), trip])
    );

    const participantsByUsername = new Map(
      (existingParticipants || []).map((participant) => [
        participant.username.trim().toLowerCase(),
        participant,
      ])
    );

    const importedUsers = [];

    for (const row of rows) {
      const participantName = safeString(getRowValue(row, "Participante"));
      const username = safeString(
        getRowValue(row, "Usuario"),
        slugify(participantName)
      );
      const motherName = safeString(
        getRowValue(row, "Nombre_Madre", "Nombre _Madre")
      );
      const fatherName = safeString(getRowValue(row, "Nombre_Padre"));
      const motherEmail = safeString(getRowValue(row, "Email_Madre"));
      const fatherEmail = safeString(getRowValue(row, "Email_Padre"));

      const tripTitle = safeString(getRowValue(row, "Viaje_Titulo"));
      const tripDate = parseDateTime(getRowValue(row, "Viaje_Fecha"));
      const tripDestination = safeString(getRowValue(row, "Viaje_Destino"));
      const tripPrice = parseAmount(getRowValue(row, "Viaje_Precio"));
      const tripDiscount = parseAmount(getRowValue(row, "Viaje_Descuento"));
      const tripToPay =
        parseAmount(getRowValue(row, "Viaje_APagar")) ||
        Math.max(0, tripPrice - tripDiscount);

      if (!participantName || !tripTitle) continue;

      let trip = tripsByName.get(tripTitle.trim().toLowerCase());

      if (!trip) {
        const tripInsert = {
          id: crypto.randomUUID(),
          name: tripTitle,
          departure_at: tripDate,
          description: tripDestination || "Experiencia GIMELOOS",
          hero_image: DEFAULT_HERO_IMAGE,
        };

        const { data: insertedTrip, error: tripInsertError } = await supabase
          .from("trips")
          .insert(tripInsert)
          .select("id, name, departure_at, description, hero_image")
          .single();

        if (tripInsertError) {
          throw new Error(
            `No se ha podido crear el viaje ${tripTitle}: ${tripInsertError.message}`
          );
        }

        trip = insertedTrip;
        tripsByName.set(tripTitle.trim().toLowerCase(), trip);
      }

      let participant = participantsByUsername.get(username.trim().toLowerCase());

      if (!participant) {
        const participantInsert = {
          id: crypto.randomUUID(),
          username,
          participant_name: participantName,
          mother_name: motherName,
          father_name: fatherName,
        };

        const { data: insertedParticipant, error: participantInsertError } =
          await supabase
            .from("participants")
            .insert(participantInsert)
            .select("id, username")
            .single();

        if (participantInsertError) {
          throw new Error(
            `No se ha podido crear el participante ${participantName}: ${participantInsertError.message}`
          );
        }

        participant = insertedParticipant;
        participantsByUsername.set(username.trim().toLowerCase(), participant);
      } else {
        const { error: participantUpdateError } = await supabase
          .from("participants")
          .update({
            participant_name: participantName,
            mother_name: motherName,
            father_name: fatherName,
          })
          .eq("id", participant.id);

        if (participantUpdateError) {
          throw new Error(
            `No se ha podido actualizar el participante ${participantName}: ${participantUpdateError.message}`
          );
        }
      }

      const paymentRows = [
        {
          payment_key: "reservation",
          name: safeString(getRowValue(row, "Pago1_Nombre"), "Reserva"),
          amount: parseAmount(getRowValue(row, "Pago1_Cantidad")),
        },
        {
          payment_key: "firstInstallment",
          name: safeString(getRowValue(row, "Pago2_Nombre"), "Primera cuota"),
          amount: parseAmount(getRowValue(row, "Pago2_Cantidad")),
        },
        {
          payment_key: "secondInstallment",
          name: safeString(getRowValue(row, "Pago3_Nombre"), "Segunda cuota"),
          amount: parseAmount(getRowValue(row, "Pago3_Cantidad")),
        },
      ];

      for (const payment of paymentRows) {
        const { data: existingPayment, error: existingPaymentError } = await supabase
          .from("participant_payments")
          .select("id")
          .eq("participant_id", participant.id)
          .eq("payment_key", payment.payment_key)
          .maybeSingle();

        if (existingPaymentError) throw new Error(existingPaymentError.message);

        if (existingPayment?.id) {
          const { error: paymentUpdateError } = await supabase
            .from("participant_payments")
            .update({
              name: payment.name,
              amount: payment.amount,
            })
            .eq("id", existingPayment.id);

          if (paymentUpdateError) {
            throw new Error(
              `No se ha podido actualizar el pago ${payment.name}: ${paymentUpdateError.message}`
            );
          }
        } else {
          const { error: paymentInsertError } = await supabase
            .from("participant_payments")
            .insert({
              id: crypto.randomUUID(),
              participant_id: participant.id,
              payment_key: payment.payment_key,
              name: payment.name,
              amount: payment.amount,
              status: "pending",
            });

          if (paymentInsertError) {
            throw new Error(
              `No se ha podido crear el pago ${payment.name}: ${paymentInsertError.message}`
            );
          }
        }
      }

      importedUsers.push({
        id: participant.id,
        role: "client",
        username,
        participantName,
        motherName,
        fatherName,
        parentName: "",
        email: motherEmail || fatherEmail || "",
        contactEmails: [motherEmail, fatherEmail].filter(Boolean),
        tripId: trip.id,
        tripName: trip.name,
        documents: defaultDocuments(),
        payments: {
          discount: tripDiscount,
          finalPrice: tripToPay,
          reservation: {
            name: safeString(getRowValue(row, "Pago1_Nombre"), "Reserva"),
            amount: parseAmount(getRowValue(row, "Pago1_Cantidad")),
            status: "pending",
            proofName: "",
            dueDate: safeString(getRowValue(row, "Pago1_Fecha")),
          },
          firstInstallment: {
            name: safeString(getRowValue(row, "Pago2_Nombre"), "Primera cuota"),
            amount: parseAmount(getRowValue(row, "Pago2_Cantidad")),
            status: "pending",
            proofName: "",
            dueDate: safeString(getRowValue(row, "Pago2_Fecha")),
          },
          secondInstallment: {
            name: safeString(getRowValue(row, "Pago3_Nombre"), "Segunda cuota"),
            amount: parseAmount(getRowValue(row, "Pago3_Cantidad")),
            status: "pending",
            proofName: "",
            dueDate: safeString(getRowValue(row, "Pago3_Fecha")),
          },
        },
        checklistState: {},
        questions: [],
      });
    }

    const bootstrap = await readBootstrap();
    const usersById = new Map(bootstrap.users.map((user) => [user.id, user]));

    importedUsers.forEach((user) => {
      const existing = usersById.get(user.id);
      usersById.set(user.id, {
        ...(existing || {}),
        ...user,
        payments: {
          ...(existing?.payments || {}),
          ...(user.payments || {}),
        },
      });
    });

    return NextResponse.json({
      message: `Excel importado correctamente. ${importedUsers.length} participante(s) procesados.`,
      importedParticipants: importedUsers.length,
      trips: bootstrap.trips,
      users: Array.from(usersById.values()),
    });
  } catch (error) {
    return NextResponse.json(
      { error: error.message || "Error interno" },
      { status: 500 }
    );
  }
}
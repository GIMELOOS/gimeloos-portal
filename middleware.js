import { NextResponse } from "next/server";
import { timingSafeEqual } from "crypto";

const AUTH_COOKIE = "animaciones_auth";

function safeEqual(a, b) {
  try {
    const ba = Buffer.from(a);
    const bb = Buffer.from(b);
    if (ba.length !== bb.length) {
      // Still run timingSafeEqual on equal-length buffers to avoid length leak
      timingSafeEqual(ba, ba);
      return false;
    }
    return timingSafeEqual(ba, bb);
  } catch {
    return false;
  }
}

export function middleware(request) {
  const { pathname } = request.nextUrl;

  // La ruta de login y la API de autenticación siempre accesibles
  if (
    pathname.startsWith("/animaciones/login") ||
    pathname === "/api/animaciones/auth"
  ) {
    return NextResponse.next();
  }

  const password = process.env.ANIMACIONES_PASSWORD;

  // Si no está configurada la contraseña, no bloquear
  if (!password) return NextResponse.next();

  const token = request.cookies.get(AUTH_COOKIE)?.value ?? "";
  if (safeEqual(token, password)) return NextResponse.next();

  // Las rutas de API devuelven 401 en lugar de redirigir a login
  if (pathname.startsWith("/api/animaciones/")) {
    return NextResponse.json({ error: "No autorizado" }, { status: 401 });
  }

  return NextResponse.redirect(new URL("/animaciones/login", request.url));
}

export const config = {
  matcher: ["/animaciones/:path*", "/api/animaciones/:path*"],
};

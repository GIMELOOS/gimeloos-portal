import { NextResponse } from "next/server";

const AUTH_COOKIE = "animaciones_auth";

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

  const token = request.cookies.get(AUTH_COOKIE)?.value;
  if (token === password) return NextResponse.next();

  return NextResponse.redirect(new URL("/animaciones/login", request.url));
}

export const config = {
  matcher: "/animaciones/:path*",
};

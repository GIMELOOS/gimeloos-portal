import { Poppins, Open_Sans } from "next/font/google";
import "./globals.css";

const poppins = Poppins({
  subsets: ["latin"],
  weight: ["400", "500", "600", "700"],
  variable: "--font-poppins",
});

const openSans = Open_Sans({
  subsets: ["latin"],
  weight: ["400", "500", "600", "700"],
  variable: "--font-open-sans",
});

export const metadata = {
  title: "GIMELOOS",
  description: "Portal de clientes GIMELOOS",
};

export default function RootLayout({ children }) {
  return (
    <html lang="es">
      <body className={`${poppins.variable} ${openSans.variable}`}>
        {children}
      </body>
    </html>
  );
}
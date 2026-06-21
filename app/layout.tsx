import type { Metadata } from "next";
import { Geist, Geist_Mono } from "next/font/google";
import "./globals.css";

const geistSans = Geist({
  variable: "--font-geist-sans",
  subsets: ["latin"],
});

const geistMono = Geist_Mono({
  variable: "--font-geist-mono",
  subsets: ["latin"],
});

export const metadata: Metadata = {
  title: "Controle Financeiro",
  description: "Sistema de controle financeiro pessoal",
  manifest: "/manifest.webmanifest",
  appleWebApp: {
    capable: true,
    title: "Controle Financeiro",
    statusBarStyle: "black-translucent",
  },
  other: {
    google: "notranslate",
  },
};

export const viewport = {
  themeColor: "#09090b",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html
      lang="pt-BR"
      translate="no"
      className={`${geistSans.variable} ${geistMono.variable} h-full antialiased notranslate`}
    >
      <head>
        <meta name="google" content="notranslate" />
        <meta name="mobile-web-app-capable" content="yes" />
        <meta
          name="apple-mobile-web-app-capable"
          content="yes"
        />
        <meta
          name="apple-mobile-web-app-status-bar-style"
          content="black-translucent"
        />
        <meta
          name="apple-mobile-web-app-title"
          content="Controle Financeiro"
        />
      </head>

      <body className="min-h-full flex flex-col bg-zinc-950 text-white notranslate">
        {children}
      </body>
    </html>
  );
}
import type { Metadata } from "next";
import { JetBrains_Mono, Manrope } from "next/font/google";
import "katex/dist/katex.min.css";
import "./globals.css";

const manrope = Manrope({
  subsets: ["latin"],
  variable: "--font-sans",
  display: "swap",
});

const jetBrains = JetBrains_Mono({
  subsets: ["latin"],
  variable: "--font-mono",
  display: "swap",
});

export const metadata: Metadata = {
  title: "PromptPress Studio",
  description:
    "Transform LLM output into readable documents, with DOCX and PDF ready mathematical formulas.",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="it">
      <body
        className={`${manrope.variable} ${jetBrains.variable} bg-slate-950 text-slate-100 antialiased`}
      >
        {children}
      </body>
    </html>
  );
}

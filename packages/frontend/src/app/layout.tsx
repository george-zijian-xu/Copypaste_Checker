import type { Metadata } from "next";
import { Inter } from "next/font/google";
import "./globals.css";
import { Header } from "@/components/feature/layout/header";
import { Footer } from "@/components/feature/layout/footer";
import type React from "react";
import { Analytics } from "@vercel/analytics/react";
// @ts-ignore - types are provided at runtime
import { SpeedInsights } from "@vercel/speed-insights/next";
import { GoogleAnalytics } from "@/components/feature/analytics/google-analytics";

const inter = Inter({ subsets: ["latin"] });

export const metadata: Metadata = {
  title: "Copy Detector",
  description:
    "Uncover every single paste event with Word's own revision metadata—no black-box guesses, just clear evidence of what was pasted."
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="en">
      <body className={inter.className}>
        <div className="min-h-screen bg-white text-gray-800 flex flex-col">
          <Header />
          <main className="flex-grow">{children}</main>
          <Footer />
          <Analytics />
          <SpeedInsights />
        </div>
        <GoogleAnalytics />
      </body>
    </html>
  );
} 
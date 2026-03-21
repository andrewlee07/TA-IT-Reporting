import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "TeacherActive IT Reporting",
  description: "Internal workbook-driven app for TeacherActive IT reporting.",
  icons: {
    icon: "/icon.png",
    apple: "/apple-icon.png",
  },
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}

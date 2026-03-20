import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "TeacherActive Exec Reporting",
  description: "Internal executive reporting app for workbook-driven IT reporting.",
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

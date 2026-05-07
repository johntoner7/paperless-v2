import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: 'Paperless BE Lab',
  description: 'Upload DOCX files, convert to HTML, edit, save, and export to PDF.',
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

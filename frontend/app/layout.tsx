import type { Metadata, Viewport } from 'next';
import './globals.css';
import { Nav } from '@/components/nav';

export const metadata: Metadata = {
  title: 'Paperless',
  description: 'DOCX tools — convert, edit, and fill forms.',
};

export const viewport: Viewport = {
  width: 'device-width',
  initialScale: 1,
  userScalable: true,
  maximumScale: 5,
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body className="flex flex-col h-screen">
        <Nav />
        {children}
      </body>
    </html>
  );
}

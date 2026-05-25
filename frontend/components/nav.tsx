'use client';

import { FileText } from 'lucide-react';

export function Nav() {
  return (
    <nav className="flex items-center gap-4 px-4 h-12 border-b bg-white shrink-0">
      <div className="flex items-center gap-2 text-sm font-semibold">
        <FileText className="h-4 w-4 text-muted-foreground" />
        Paperless
      </div>
    </nav>
  );
}

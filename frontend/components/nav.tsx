'use client';

import Link from 'next/link';
import { usePathname } from 'next/navigation';
import { FileText } from 'lucide-react';
import { cn } from '@/lib/utils';

export function Nav() {
  const path = usePathname();
  return (
    <nav className="flex items-center gap-4 px-4 h-12 border-b bg-white shrink-0">
      <div className="flex items-center gap-2 text-sm font-semibold">
        <FileText className="h-4 w-4 text-muted-foreground" />
        Paperless
      </div>
      <div className="flex items-center gap-1">
        <Link
          href="/"
          className={cn(
            'px-3 py-1 rounded text-sm transition-colors',
            path === '/'
              ? 'bg-muted text-foreground font-medium'
              : 'text-muted-foreground hover:text-foreground hover:bg-muted/50',
          )}
        >
          Convert
        </Link>
        <Link
          href="/form"
          className={cn(
            'px-3 py-1 rounded text-sm transition-colors',
            path === '/form'
              ? 'bg-muted text-foreground font-medium'
              : 'text-muted-foreground hover:text-foreground hover:bg-muted/50',
          )}
        >
          Form Fill
        </Link>
      </div>
    </nav>
  );
}

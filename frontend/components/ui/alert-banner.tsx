'use client';

import { useEffect } from 'react';
import { X, AlertCircle, CheckCircle2 } from 'lucide-react';
import { cn } from '@/lib/utils';

type AlertType = 'error' | 'success';

export function AlertBanner({
  type,
  message,
  onDismiss,
}: {
  type: AlertType;
  message: string;
  onDismiss: () => void;
}) {
  useEffect(() => {
    if (type !== 'success') return;
    const t = setTimeout(onDismiss, 4000);
    return () => clearTimeout(t);
  }, [type, message, onDismiss]);

  return (
    <div
      className={cn(
        'flex items-center gap-2 px-4 py-2.5 border-b text-sm shrink-0',
        type === 'error'
          ? 'bg-red-50 border-red-200 text-red-700'
          : 'bg-emerald-50 border-emerald-200 text-emerald-700',
      )}
    >
      {type === 'error'
        ? <AlertCircle className="h-4 w-4 shrink-0" />
        : <CheckCircle2 className="h-4 w-4 shrink-0" />}
      <span className="flex-1">{message}</span>
      <button
        onClick={onDismiss}
        className="shrink-0 opacity-60 hover:opacity-100 transition-opacity"
        aria-label="Dismiss"
      >
        <X className="h-4 w-4" />
      </button>
    </div>
  );
}

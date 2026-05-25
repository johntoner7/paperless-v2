'use client';

import { useMemo } from 'react';
import { FileCheck, AlertCircle } from 'lucide-react';
import { Badge } from '@/components/ui/badge';
import { cn } from '@/lib/utils';

export type ExtractedField = {
  id?: string;
  label: string;
  type: 'text' | 'date' | 'checkbox' | 'choice' | 'number' | 'multiline' | string;
  value: string | boolean | null;
  confidence?: number;
  source?: { nodeIds: string[] };
};

const DATE_RE = /^\d{1,2}[/\-.]\d{1,2}[/\-.]\d{2,4}$|^\d{4}[/\-.]\d{1,2}[/\-.]\d{1,2}$/;

function validateField(field: ExtractedField): string | null {
  if (field.type === 'checkbox') return null;
  const val = String(field.value ?? '').trim();
  if (!val) return null;
  if (field.type === 'date' && !DATE_RE.test(val)) return 'Enter a date like DD/MM/YYYY';
  if (field.type === 'number' && (isNaN(Number(val)) || !isFinite(Number(val)))) return 'Must be a number';
  return null;
}

export function FieldPanel({
  fields,
  onFieldChange,
}: {
  fields: ExtractedField[];
  onFieldChange: (index: number, value: string | boolean) => void;
}) {
  const fieldErrors = useMemo(() => fields.map(validateField), [fields]);
  const hasErrors = fieldErrors.some(Boolean);

  const counts = useMemo(() => {
    const total = fields.length;
    const filled = fields.filter((f) => {
      if (f.type === 'checkbox') return typeof f.value === 'boolean';
      return Boolean(f.value && String(f.value).trim().length > 0);
    }).length;
    return { total, filled };
  }, [fields]);

  if (fields.length === 0) {
    return (
      <div className="flex flex-col items-center justify-center h-full gap-3 text-center p-8">
        <div className="rounded-full bg-muted p-4">
          <FileCheck className="h-8 w-8 text-muted-foreground" />
        </div>
        <p className="text-sm text-muted-foreground max-w-xs">
          Fields will appear here once extraction is complete.
        </p>
      </div>
    );
  }

  return (
    <div className="flex flex-col h-full">
      <div className="flex items-center gap-3 px-4 py-2 border-b bg-muted/20 shrink-0">
        <span className="text-xs text-muted-foreground">{counts.filled}/{counts.total} filled</span>
        {hasErrors && (
          <span className="text-xs text-red-500 flex items-center gap-1">
            <AlertCircle className="h-3 w-3" /> Fix errors before exporting
          </span>
        )}
      </div>

      <div className="flex-1 overflow-y-auto p-4">
        <div className="grid gap-3 max-w-5xl mx-auto sm:grid-cols-2 lg:grid-cols-3">
          {fields.map((field, index) => {
            const error = fieldErrors[index];
            const isAI = (field.confidence ?? 1) <= 0.5;
            return (
              <div
                key={field.id ?? `${field.label}-${index}`}
                className={cn('rounded-lg border bg-white p-4', error ? 'border-red-300' : 'border-border')}
              >
                <div className="flex items-start justify-between gap-2 mb-3">
                  <div className="min-w-0">
                    <div className="text-sm font-medium truncate">{field.label}</div>
                    <div className="text-xs text-muted-foreground">{field.type}</div>
                  </div>
                  <Badge
                    variant="outline"
                    className={cn('text-xs shrink-0', isAI ? 'border-purple-200 bg-purple-50 text-purple-600' : 'text-muted-foreground')}
                  >
                    {isAI ? 'AI' : 'rule'} · {Math.round((field.confidence ?? 0) * 100)}%
                  </Badge>
                </div>

                {field.type === 'checkbox' ? (
                  <label className="flex items-center gap-2 text-sm text-muted-foreground cursor-pointer">
                    <input
                      type="checkbox"
                      checked={Boolean(field.value)}
                      onChange={(e) => onFieldChange(index, e.target.checked)}
                      className="h-4 w-4 rounded border-input"
                    />
                    Checked
                  </label>
                ) : field.type === 'multiline' ? (
                  <textarea
                    className={cn(
                      'min-h-[72px] w-full rounded-md border px-3 py-2 text-sm bg-background focus:outline-none focus:ring-1 focus:ring-ring',
                      error ? 'border-red-300' : 'border-input',
                    )}
                    value={String(field.value ?? '')}
                    onChange={(e) => onFieldChange(index, e.target.value)}
                  />
                ) : (
                  <input
                    type="text"
                    className={cn(
                      'w-full rounded-md border px-3 py-2 text-sm bg-background focus:outline-none focus:ring-1 focus:ring-ring',
                      error ? 'border-red-300' : 'border-input',
                    )}
                    value={String(field.value ?? '')}
                    onChange={(e) => onFieldChange(index, e.target.value)}
                  />
                )}
                {error && <p className="mt-1 text-xs text-red-500">{error}</p>}
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

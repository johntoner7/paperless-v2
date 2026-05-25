'use client';

import { useMemo, useState } from 'react';
import { Upload, Wand2, FileCheck, Download, AlertCircle } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Badge } from '@/components/ui/badge';
import { Separator } from '@/components/ui/separator';
import { cn } from '@/lib/utils';

type ExtractedField = {
  id?: string;
  label: string;
  type: 'text' | 'date' | 'checkbox' | 'choice' | 'number' | 'multiline' | string;
  value: string | boolean | null;
  confidence?: number;
  source?: { nodeIds: string[] };
};

type ExtractionState = 'idle' | 'uploading' | 'extracting' | 'ready' | 'error';
type ExportState = 'idle' | 'exporting' | 'done' | 'error';

type ExtractResponse = {
  cached?: boolean;
  key?: string;
  fields?: ExtractedField[];
};

export default function FormExtractionPage() {
  const PRESIGN_URL = process.env.NEXT_PUBLIC_PRESIGN_URL ?? '';

  const [status, setStatus] = useState<ExtractionState>('idle');
  const [message, setMessage] = useState('');
  const [sourceFile, setSourceFile] = useState<File | null>(null);
  const [uploadKey, setUploadKey] = useState<string | null>(null);
  const [fields, setFields] = useState<ExtractedField[]>([]);
  const [exportStatus, setExportStatus] = useState<ExportState>('idle');
  const [exportMessage, setExportMessage] = useState('');
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);
  const [useAI, setUseAI] = useState(false);

  const isBusy = status === 'uploading' || status === 'extracting';

  const DATE_RE = /^\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4}$|^\d{4}[\/\-.]\d{1,2}[\/\-.]\d{1,2}$/;

  function validateField(field: ExtractedField): string | null {
    if (field.type === 'checkbox') return null;
    const val = String(field.value ?? '').trim();
    if (!val) return null;
    if (field.type === 'date' && !DATE_RE.test(val)) return 'Enter a date like DD/MM/YYYY';
    if (field.type === 'number' && (isNaN(Number(val)) || !isFinite(Number(val)))) return 'Must be a number';
    return null;
  }

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

  function handleFileChange(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0] ?? null;
    if (!file) return;
    if (!file.name.toLowerCase().endsWith('.docx')) {
      setStatus('error');
      setMessage('Only .docx files are supported.');
      return;
    }
    setSourceFile(file);
    setMessage('');
    setStatus('idle');
  }

  async function presign(fileName: string, contentType: string) {
    const res = await fetch(PRESIGN_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ fileName, contentType }),
    });
    if (!res.ok) throw new Error(`Presign failed: ${res.status}`);
    return res.json() as Promise<{ uploadUrl: string; key: string }>;
  }

  async function handleExtract() {
    if (!sourceFile || !PRESIGN_URL) return;
    setStatus('uploading');
    setMessage('Uploading document…');

    try {
      const { uploadUrl, key } = await presign(
        sourceFile.name,
        sourceFile.type || 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      );
      const upload = await fetch(uploadUrl, {
        method: 'PUT',
        body: sourceFile,
        headers: { 'Content-Type': sourceFile.type || 'application/octet-stream' },
      });
      if (!upload.ok) throw new Error(`Upload failed: ${upload.status}`);
      setUploadKey(key);

      setStatus('extracting');
      setMessage('Extracting fields…');
      const resp = await fetch(`${PRESIGN_URL}?action=extract`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ key, useAI }),
      });
      if (!resp.ok) throw new Error(`Extract failed: ${resp.status}`);
      const data = (await resp.json()) as ExtractResponse;
      setFields((data.fields ?? []).map((f, idx) => ({
        id: f.id ?? `field_${idx + 1}`,
        label: f.label,
        type: f.type,
        value: f.value ?? (f.type === 'checkbox' ? false : ''),
        confidence: f.confidence ?? 0,
        source: f.source,
      })));
      setExportStatus('idle');
      setDownloadUrl(null);
      setStatus('ready');
      setMessage(data.cached ? 'Loaded cached extraction.' : 'Extraction complete.');
    } catch (e) {
      setStatus('error');
      setMessage(e instanceof Error ? e.message : 'Extraction failed.');
    }
  }

  function updateFieldValue(index: number, value: string | boolean) {
    setFields((prev) =>
      prev.map((field, idx) => (idx === index ? { ...field, value } : field)),
    );
  }

  function buildPatches(fieldList: ExtractedField[]) {
    const patches: object[] = [];
    for (const field of fieldList) {
      const nodeIds = field.source?.nodeIds ?? [];
      if (field.type === 'checkbox') {
        const sdtId = nodeIds.find((id) => id.startsWith('sdt_'));
        if (sdtId) {
          patches.push({ node_id: sdtId, operation: 'toggle_checkbox', checked: Boolean(field.value) });
        }
      } else {
        const val = field.value;
        if (val === null || val === '') continue;
        const runIds = nodeIds.filter((id) => id.startsWith('r_'));
        const targetId = runIds[runIds.length - 1];
        if (targetId) {
          patches.push({ node_id: targetId, operation: 'replace_text', text: String(val) });
        }
      }
    }
    return patches;
  }

  async function handleExport() {
    if (!uploadKey || !PRESIGN_URL) return;
    setExportStatus('exporting');
    setExportMessage('');
    try {
      const patches = buildPatches(fields);
      const resp = await fetch(`${PRESIGN_URL}?action=export-docx`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ originalKey: uploadKey, patches }),
      });
      if (!resp.ok) throw new Error(`Export failed: ${resp.status}`);
      const data = (await resp.json()) as { url: string; key: string };
      setDownloadUrl(data.url);
      setExportStatus('done');
      setExportMessage(`${patches.length} patch${patches.length !== 1 ? 'es' : ''} applied.`);
    } catch (e) {
      setExportStatus('error');
      setExportMessage(e instanceof Error ? e.message : 'Export failed.');
    }
  }

  const statusBadge = status === 'error'
    ? <Badge variant="destructive" className="text-xs">{message}</Badge>
    : status === 'ready' && message
      ? <span className="text-xs text-muted-foreground">{message}</span>
      : null;

  return (
    <div className="flex flex-col flex-1 overflow-hidden bg-background">

      {/* ── Toolbar ── */}
      <div className="flex items-center gap-2 px-4 py-2.5 border-b bg-white shrink-0 flex-wrap">
        <label
          className={cn(
            'inline-flex items-center gap-1.5 h-8 px-3 rounded-md border text-sm cursor-pointer transition-colors',
            sourceFile
              ? 'border-primary/30 bg-primary/5 text-primary font-medium'
              : 'border-input bg-background text-muted-foreground hover:bg-accent',
          )}
        >
          <Upload className="h-3.5 w-3.5" />
          <span className="max-w-[160px] truncate">{sourceFile ? sourceFile.name : 'Choose DOCX'}</span>
          <input type="file" accept=".docx" className="sr-only" onChange={handleFileChange} />
        </label>

        <label className="inline-flex items-center gap-2 h-8 px-3 rounded-md border border-input bg-background text-sm cursor-pointer hover:bg-accent transition-colors">
          <input
            type="checkbox"
            checked={useAI}
            onChange={(e) => setUseAI(e.target.checked)}
            className="h-4 w-4 rounded"
          />
          <span className="text-muted-foreground">AI</span>
        </label>

        <Button
          size="sm"
          className="bg-sky-500 hover:bg-sky-600 text-white border-0"
          onClick={() => void handleExtract()}
          disabled={!sourceFile || !PRESIGN_URL || isBusy}
        >
          <Wand2 className="h-3.5 w-3.5" />
          {status === 'uploading' ? 'Uploading…' : status === 'extracting' ? 'Extracting…' : 'Extract Fields'}
        </Button>

        <Separator orientation="vertical" className="h-5 mx-1 hidden sm:block" />

        {statusBadge}

        <div className="ml-auto flex items-center gap-2">
          {fields.length > 0 && (
            <span className="text-xs text-muted-foreground hidden sm:inline">
              {counts.filled}/{counts.total} filled
            </span>
          )}

          {hasErrors && (
            <span className="text-xs text-red-500 flex items-center gap-1">
              <AlertCircle className="h-3 w-3" />
              <span className="hidden sm:inline">Fix errors</span>
            </span>
          )}

          {exportStatus === 'error' && (
            <span className="text-xs text-red-500 hidden sm:inline">{exportMessage}</span>
          )}
          {exportStatus === 'done' && !hasErrors && (
            <span className="text-xs text-muted-foreground hidden sm:inline">{exportMessage}</span>
          )}

          <Button
            variant="outline"
            size="sm"
            className="border-emerald-200 text-emerald-700 hover:bg-emerald-50"
            onClick={() => void handleExport()}
            disabled={fields.length === 0 || hasErrors || exportStatus === 'exporting'}
          >
            <Download className="h-3.5 w-3.5" />
            <span className="hidden sm:inline">
              {exportStatus === 'exporting' ? 'Exporting…' : 'Export DOCX'}
            </span>
          </Button>

          {downloadUrl && (
            <a
              href={downloadUrl}
              download
              className="inline-flex items-center gap-1.5 h-8 px-3 rounded-md border border-emerald-300 bg-emerald-50 text-sm text-emerald-700 hover:bg-emerald-100 transition-colors"
            >
              <FileCheck className="h-3.5 w-3.5" />
              <span className="hidden sm:inline">Download</span>
            </a>
          )}
        </div>
      </div>

      {/* ── Field grid ── */}
      <div className="flex-1 overflow-y-auto p-4 bg-muted/20">
        {fields.length === 0 ? (
          <div className="flex flex-col items-center justify-center h-full gap-3 text-center p-8">
            <div className="rounded-full bg-muted p-4">
              <FileCheck className="h-8 w-8 text-muted-foreground" />
            </div>
            <p className="text-sm text-muted-foreground max-w-xs">
              Choose a <strong>.docx</strong> file and click <strong>Extract Fields</strong> to list editable form fields.
            </p>
          </div>
        ) : (
          <div className="grid gap-3 max-w-5xl mx-auto sm:grid-cols-2 lg:grid-cols-3">
            {fields.map((field, index) => {
              const error = fieldErrors[index];
              const isAI = (field.confidence ?? 1) <= 0.5;
              return (
                <div
                  key={field.id ?? `${field.label}-${index}`}
                  className={cn(
                    'rounded-lg border bg-white p-4',
                    error ? 'border-red-300' : 'border-border',
                  )}
                >
                  <div className="flex items-start justify-between gap-2 mb-3">
                    <div className="min-w-0">
                      <div className="text-sm font-medium truncate">{field.label}</div>
                      <div className="text-xs text-muted-foreground">{field.type}</div>
                    </div>
                    <Badge
                      variant="outline"
                      className={cn(
                        'text-xs shrink-0',
                        isAI
                          ? 'border-purple-200 bg-purple-50 text-purple-600'
                          : 'text-muted-foreground',
                      )}
                    >
                      {isAI ? 'AI' : 'rule'} · {Math.round((field.confidence ?? 0) * 100)}%
                    </Badge>
                  </div>

                  {field.type === 'checkbox' ? (
                    <label className="flex items-center gap-2 text-sm text-muted-foreground cursor-pointer">
                      <input
                        type="checkbox"
                        checked={Boolean(field.value)}
                        onChange={(e) => updateFieldValue(index, e.target.checked)}
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
                      onChange={(e) => updateFieldValue(index, e.target.value)}
                    />
                  ) : (
                    <input
                      type="text"
                      className={cn(
                        'w-full rounded-md border px-3 py-2 text-sm bg-background focus:outline-none focus:ring-1 focus:ring-ring',
                        error ? 'border-red-300' : 'border-input',
                      )}
                      value={String(field.value ?? '')}
                      onChange={(e) => updateFieldValue(index, e.target.value)}
                    />
                  )}
                  {error && <p className="mt-1 text-xs text-red-500">{error}</p>}
                </div>
              );
            })}
          </div>
        )}
      </div>
    </div>
  );
}

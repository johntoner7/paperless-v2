'use client';

import { useMemo, useState } from 'react';
import { Space_Grotesk, Source_Sans_3 } from 'next/font/google';
import { UploadCloud, Sparkles, FileCheck } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Badge } from '@/components/ui/badge';
import { cn } from '@/lib/utils';

const spaceGrotesk = Space_Grotesk({ subsets: ['latin'], weight: ['400', '600', '700'] });
const sourceSans = Source_Sans_3({ subsets: ['latin'], weight: ['400', '600'] });

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
        body: JSON.stringify({ key, useAI: false }),
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
    setExportMessage('Building patches…');
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

  return (
    <div className={cn('min-h-screen text-slate-950', sourceSans.className)}>
      <div className="relative overflow-hidden bg-gradient-to-br from-amber-50 via-white to-sky-100">
        <div className="pointer-events-none absolute inset-0">
          <div className="absolute -top-24 right-[-10%] h-72 w-72 rounded-full bg-amber-200/40 blur-3xl" />
          <div className="absolute bottom-[-20%] left-[-10%] h-72 w-72 rounded-full bg-sky-200/50 blur-3xl" />
          <div className="absolute top-24 left-1/2 h-40 w-40 -translate-x-1/2 rounded-2xl bg-white/50 backdrop-blur-sm" />
        </div>

        <div className="relative mx-auto flex w-full max-w-6xl flex-col gap-10 px-6 py-10 lg:py-14">
          <header className="flex flex-col gap-4">
            <div className="inline-flex items-center gap-2 text-xs uppercase tracking-[0.2em] text-slate-500">
              <Sparkles className="h-4 w-4" />
              Form Extraction Lab
            </div>
            <div className="flex flex-col gap-3">
              <h1 className={cn('text-3xl font-semibold text-slate-900 md:text-4xl', spaceGrotesk.className)}>
                Upload a DOCX, extract fields, and start filling.
              </h1>
              <p className="max-w-2xl text-base text-slate-600">
                This page handles upload and automatic extraction (rules only). Fields are listed below so you can
                enter values without altering labels or types.
              </p>
            </div>
          </header>

          <div className="grid gap-6 lg:grid-cols-[360px_1fr]">
            <section className="rounded-2xl border border-slate-200 bg-white/90 p-6 shadow-sm backdrop-blur">
              <div className="flex items-center justify-between">
                <div>
                  <h2 className={cn('text-lg font-semibold', spaceGrotesk.className)}>Upload & Extract</h2>
                  <p className="text-sm text-slate-500">DOCX only · AI disabled</p>
                </div>
                <Badge variant="secondary" className="bg-slate-100 text-slate-600">Phase 1</Badge>
              </div>

              <div className="mt-5 flex flex-col gap-4">
                <label className="flex cursor-pointer flex-col gap-2 rounded-xl border border-dashed border-slate-300 bg-slate-50 px-4 py-5 text-sm text-slate-600 hover:border-slate-400">
                  <div className="flex items-center gap-2 text-slate-700">
                    <UploadCloud className="h-4 w-4" />
                    <span>{sourceFile ? sourceFile.name : 'Choose a DOCX file'}</span>
                  </div>
                  <span className="text-xs text-slate-500">Maximum size depends on your S3 settings.</span>
                  <input type="file" accept=".docx" className="sr-only" onChange={handleFileChange} />
                </label>

                <Button
                  className="h-10 bg-slate-900 text-white hover:bg-slate-800"
                  onClick={() => void handleExtract()}
                  disabled={!sourceFile || !PRESIGN_URL || isBusy}
                >
                  {status === 'extracting' ? 'Extracting…' : status === 'uploading' ? 'Uploading…' : 'Extract Fields'}
                </Button>

                <div className="rounded-xl border border-slate-200 bg-white px-4 py-3 text-sm text-slate-600">
                  <div className="flex items-center gap-2">
                    <FileCheck className="h-4 w-4 text-emerald-500" />
                    <span>{message || 'Ready to extract fields.'}</span>
                  </div>
                  {uploadKey && (
                    <div className="mt-2 text-xs text-slate-400">Upload key: {uploadKey}</div>
                  )}
                </div>
              </div>
            </section>

            <section className="rounded-2xl border border-slate-200 bg-white/95 p-6 shadow-sm backdrop-blur">
              <div className="flex flex-wrap items-center justify-between gap-3">
                <div>
                  <h2 className={cn('text-lg font-semibold', spaceGrotesk.className)}>Extracted Fields</h2>
                  <p className="text-sm text-slate-500">Edit values only. Labels and types stay locked.</p>
                </div>
                <div className="flex items-center gap-2 text-sm text-slate-500">
                  <Badge variant="outline" className="border-slate-200 text-slate-500">
                    {counts.filled}/{counts.total} filled
                  </Badge>
                  <Badge variant="secondary" className="bg-amber-100 text-amber-700">
                    Phase 2
                  </Badge>
                </div>
              </div>

              <div className="mt-5 grid gap-4">
                {fields.length === 0 && (
                  <div className="rounded-xl border border-dashed border-slate-200 bg-slate-50 px-4 py-8 text-center text-sm text-slate-500">
                    Upload a document to see extracted fields.
                  </div>
                )}

                {fields.map((field, index) => {
                  const error = fieldErrors[index];
                  const isAI = (field.confidence ?? 1) <= 0.5;
                  return (
                    <div
                      key={field.id ?? `${field.label}-${index}`}
                      className={cn('rounded-xl border bg-white px-4 py-4', error ? 'border-red-300' : 'border-slate-200')}
                    >
                      <div className="flex flex-wrap items-center justify-between gap-2">
                        <div>
                          <div className="text-sm font-semibold text-slate-900">{field.label}</div>
                          <div className="text-xs text-slate-500">{field.type}</div>
                        </div>
                        <div className="flex items-center gap-1.5">
                          <Badge
                            variant="outline"
                            className={cn(
                              'border text-xs',
                              isAI
                                ? 'border-purple-200 bg-purple-50 text-purple-600'
                                : 'border-slate-200 text-slate-500',
                            )}
                          >
                            {isAI ? 'AI' : 'rule'} · {Math.round((field.confidence ?? 0) * 100)}%
                          </Badge>
                        </div>
                      </div>

                      <div className="mt-3">
                        {field.type === 'checkbox' ? (
                          <label className="flex items-center gap-2 text-sm text-slate-600">
                            <input
                              type="checkbox"
                              checked={Boolean(field.value)}
                              onChange={(e) => updateFieldValue(index, e.target.checked)}
                              className="h-4 w-4 rounded border-slate-300"
                            />
                            Mark as checked
                          </label>
                        ) : field.type === 'multiline' ? (
                          <textarea
                            className={cn(
                              'min-h-[96px] w-full rounded-lg border px-3 py-2 text-sm focus:outline-none',
                              error ? 'border-red-300 focus:border-red-400' : 'border-slate-200 focus:border-slate-400',
                            )}
                            value={String(field.value ?? '')}
                            onChange={(e) => updateFieldValue(index, e.target.value)}
                          />
                        ) : (
                          <input
                            type="text"
                            className={cn(
                              'w-full rounded-lg border px-3 py-2 text-sm focus:outline-none',
                              error ? 'border-red-300 focus:border-red-400' : 'border-slate-200 focus:border-slate-400',
                            )}
                            value={String(field.value ?? '')}
                            onChange={(e) => updateFieldValue(index, e.target.value)}
                          />
                        )}
                        {error && <p className="mt-1 text-xs text-red-500">{error}</p>}
                      </div>
                    </div>
                  );
                })}
              </div>
            </section>
          </div>

          {/* Phase 3 + 4: Export */}
          <section className="rounded-2xl border border-slate-200 bg-white/90 p-6 shadow-sm backdrop-blur">
            <div className="flex flex-wrap items-center justify-between gap-3">
              <div>
                <h2 className={cn('text-lg font-semibold', spaceGrotesk.className)}>Export DOCX</h2>
                <p className="text-sm text-slate-500">Patches field values into the original document.</p>
              </div>
              <Badge variant="secondary" className="bg-sky-100 text-sky-700">Phase 3 + 4</Badge>
            </div>

            <div className="mt-5 flex flex-wrap items-center gap-4">
              <Button
                className="h-10 bg-slate-900 text-white hover:bg-slate-800"
                onClick={() => void handleExport()}
                disabled={fields.length === 0 || hasErrors || exportStatus === 'exporting'}
              >
                {exportStatus === 'exporting' ? 'Exporting…' : 'Export & Download'}
              </Button>

              {hasErrors && (
                <span className="text-sm text-red-500">Fix validation errors above before exporting.</span>
              )}

              {!hasErrors && exportMessage && (
                <span className={cn('text-sm', exportStatus === 'error' ? 'text-red-600' : 'text-slate-500')}>
                  {exportMessage}
                </span>
              )}

              {downloadUrl && (
                <a
                  href={downloadUrl}
                  download
                  className="inline-flex h-10 items-center gap-2 rounded-lg border border-emerald-300 bg-emerald-50 px-4 text-sm font-medium text-emerald-700 hover:bg-emerald-100"
                >
                  <FileCheck className="h-4 w-4" />
                  Download patched DOCX
                </a>
              )}
            </div>
          </section>
        </div>
      </div>
    </div>
  );
}

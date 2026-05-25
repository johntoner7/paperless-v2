"use client";

import { useEffect, useMemo, useRef, useState } from 'react';
import {
  FileText, Upload, Wand2, RefreshCw, SplitSquareHorizontal, X,
  Download, Save, BookOpen, ChevronDown,
} from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Badge } from '@/components/ui/badge';
import { Separator } from '@/components/ui/separator';
import { cn } from '@/lib/utils';

type ConversionState = 'idle' | 'converting' | 'ready' | 'error';
type LibraryUpload = {
  key: string;
  fileName: string;
  lastModified: string | null;
  status: 'ready' | 'pending';
};

function stripExtension(name: string) {
  return name.replace(/\.[^.]+$/, '');
}

function downloadBlob(fileName: string, content: string, mime: string) {
  const a = Object.assign(document.createElement('a'), {
    href: URL.createObjectURL(new Blob([content], { type: mime })),
    download: fileName,
  });
  document.body.appendChild(a);
  a.click();
  a.remove();
}

/* ── DOCX preview pane ── */
function DocxPane({ url }: { url: string }) {
  const ref = useRef<HTMLDivElement>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    setError(null);

    (async () => {
      if (!ref.current) return;
      try {
        const { renderAsync } = await import('docx-preview');
        const buf = await fetch(url).then(r => {
          if (!r.ok) throw new Error(`Fetch failed: ${r.status}`);
          return r.arrayBuffer();
        });
        if (cancelled || !ref.current) return;
        ref.current.innerHTML = '';
        await renderAsync(buf, ref.current, undefined, {
          inWrapper: true,
          ignoreWidth: false,
          ignoreFonts: false,
          breakPages: true,
          useBase64URL: true,
        });
      } catch (e) {
        if (!cancelled) setError(e instanceof Error ? e.message : 'Render failed');
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();

    return () => { cancelled = true; };
  }, [url]);

  return (
    <div className="document-pane h-full">
      {loading && (
        <div className="flex items-center justify-center h-40 text-sm text-muted-foreground">
          Rendering original document…
        </div>
      )}
      {error && <div className="text-sm text-red-500 p-4">{error}</div>}
      <div ref={ref} className="docx-wrapper" />
    </div>
  );
}

/* ── HTML editor ────────────────────────────── */
function HtmlEditor({ initialHtml, onChange }: { initialHtml: string; onChange: (h: string) => void }) {
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (ref.current && ref.current.innerHTML !== initialHtml) {
      ref.current.innerHTML = initialHtml;
    }
  }, [initialHtml]);

  return (
    <div
      ref={ref}
      className="editor-prose"
      contentEditable
      suppressContentEditableWarning
      onInput={() => ref.current && onChange(ref.current.innerHTML)}
    />
  );
}

/* ── Main workbench ─────────────────────────── */
export default function DocumentWorkbench() {
  const PRESIGN_URL = process.env.NEXT_PUBLIC_PRESIGN_URL ?? '';

  const [status, setStatus] = useState<ConversionState>('idle');
  const [message, setMessage] = useState('');
  const [sourceFile, setSourceFile] = useState<File | null>(null);
  const [docTitle, setDocTitle] = useState('');
  const [htmlDraft, setHtmlDraft] = useState('');
  const [savedAt, setSavedAt] = useState<string | null>(null);
  const [library, setLibrary] = useState<LibraryUpload[]>([]);
  const [selectedKey, setSelectedKey] = useState('');
  const [docxSource, setDocxSource] = useState<string | null>(null);
  const [originalKey, setOriginalKey] = useState<string | null>(null);
  const [showCompare, setShowCompare] = useState(false);
  const [libraryOpen, setLibraryOpen] = useState(false);
  const [useAI, setUseAI] = useState(false);
  const skipRestoreRef = useRef(false);

  const draftKey = useMemo(() => `pb:draft:${docTitle}`, [docTitle]);

  /* draft restore */
  useEffect(() => {
    if (skipRestoreRef.current) { skipRestoreRef.current = false; return; }
    if (!docTitle) return;
    const saved = localStorage.getItem(draftKey);
    if (saved) { setHtmlDraft(saved); setStatus('ready'); setMessage('Draft restored.'); }
  }, [draftKey, docTitle]);

  /* library */
  async function fetchLibrary() {
    if (!PRESIGN_URL) return;
    const r = await fetch(`${PRESIGN_URL}?action=list`);
    if (!r.ok) return;
    const d = await r.json() as { items?: LibraryUpload[] };
    setLibrary(d.items ?? []);
  }
  useEffect(() => { void fetchLibrary().catch(console.error); }, [PRESIGN_URL]);

  /* presign helpers */
  async function presign(fileName: string, contentType: string) {
    const r = await fetch(PRESIGN_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ fileName, contentType }),
    });
    if (!r.ok) throw new Error(`Presign failed: ${r.status}`);
    return r.json() as Promise<{ uploadUrl: string; key: string }>;
  }

  async function pollResult(key: string) {
    const r = await fetch(`${PRESIGN_URL}?key=${encodeURIComponent(key)}`);
    if (r.status === 404) return null;
    if (!r.ok) throw new Error(`Poll failed: ${r.status}`);
    return r.json() as Promise<{ status: string; resultUrl: string; originalUrl?: string; key: string }>;
  }

  /* convert */
  async function handleConvert() {
    if (!sourceFile) return;
    setStatus('converting'); setMessage('Uploading…');
    try {
      const { uploadUrl, key } = await presign(
        sourceFile.name,
        sourceFile.type || 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      );
      const up = await fetch(uploadUrl, {
        method: 'PUT', body: sourceFile,
        headers: { 'Content-Type': sourceFile.type || 'application/octet-stream' },
      });
      if (!up.ok) throw new Error(`Upload failed: ${up.status}`);
      
      // Invoke conversion with AI flag if enabled
      setMessage('Converting…');
      if (useAI) {
        const convResp = await fetch(PRESIGN_URL, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ action: 'convert', key, useAI: true }),
        });
        if (!convResp.ok) {
          console.warn('Failed to invoke AI conversion, will rely on S3 events');
        }
      }
      
      for (let i = 0; i < 60; i++) {
        const result = await pollResult(key);
        if (result?.status === 'ready' && result.resultUrl) {
          const html = await fetch(result.resultUrl).then(r => r.text());
          const title = stripExtension(sourceFile.name);
          skipRestoreRef.current = true;
          setDocTitle(title); setHtmlDraft(html);
          setDocxSource(result.originalUrl ?? null); setOriginalKey(key); setShowCompare(false);
          setStatus('ready'); setMessage('Conversion complete.');
          setSavedAt(null);
          localStorage.setItem(`pb:draft:${title}`, html);
          void fetchLibrary().catch(console.error);
          return;
        }
        await new Promise(r => setTimeout(r, 2000));
      }
      throw new Error('Timed out waiting for conversion.');
    } catch (e) {
      setStatus('error'); setMessage(e instanceof Error ? e.message : String(e));
    }
  }

  /* load from library */
  async function loadDoc(key: string) {
    const entry = library.find(i => i.key === key);
    if (!entry || entry.status !== 'ready') {
      setMessage(`${entry?.fileName ?? key} is still processing.`); return;
    }
    setStatus('converting'); setMessage(`Loading ${entry.fileName}…`);
    try {
      const result = await pollResult(key);
      if (!result?.resultUrl) { setStatus('idle'); setMessage('No HTML yet.'); return; }
      const html = await fetch(result.resultUrl).then(r => r.text());
      const title = stripExtension(entry.fileName);
      skipRestoreRef.current = true;
      setDocTitle(title); setSourceFile(null); setHtmlDraft(html);
      setDocxSource(result.originalUrl ?? null); setOriginalKey(key); setShowCompare(false);
      setSelectedKey(key); setStatus('ready'); setMessage(`Opened ${entry.fileName}`);
      setSavedAt(null);
      localStorage.setItem(`pb:draft:${title}`, html);
    } catch (e) {
      setStatus('error'); setMessage(e instanceof Error ? e.message : String(e));
    }
  }

  function handleFileChange(e: React.ChangeEvent<HTMLInputElement>) {
    const f = e.target.files?.[0];
    if (!f) return;
    if (!f.name.toLowerCase().endsWith('.docx')) {
      setStatus('error'); setMessage('Only .docx files are supported.'); return;
    }
    setSourceFile(f); setStatus('idle'); setMessage('');
  }

  async function handleSave() {
    if (status !== 'ready') return;
    const ts = new Date().toISOString().replace(/[:.]/g, '-');
    downloadBlob(`${docTitle || 'document'}-${ts}.html`, htmlDraft, 'text/html;charset=utf-8');
    localStorage.setItem(draftKey, htmlDraft);
    setSavedAt(new Date().toLocaleTimeString()); setMessage('Saved.');
  }

  async function handlePdf() {
    if (status !== 'ready') return;
    setMessage('Generating PDF…');
    try {
      const h2p = (await import('html2pdf.js')).default;
      const el = Object.assign(document.createElement('div'), { innerHTML: htmlDraft });
      await h2p().set({
        margin: 18, filename: `${docTitle}.pdf`,
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2, useCORS: true },
        jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' },
      }).from(el).save();
      setMessage('PDF downloaded.');
    } catch (e) {
      setStatus('error'); setMessage(e instanceof Error ? e.message : 'PDF failed.');
    }
  }

  async function handleExportDocx() {
    if (!originalKey || status !== 'ready') return;
    setMessage('Exporting DOCX…');
    try {
      // Parse current HTML and extract stable node-identified run elements.
      const parser = new DOMParser();
      const doc = parser.parseFromString(htmlDraft, 'text/html');
      const patches: Array<Record<string, unknown>> = [];

      doc.querySelectorAll('[data-node-kind="run"][data-node-id]').forEach(el => {
        const nodeId = el.getAttribute('data-node-id');
        if (!nodeId) return;

        const patch: Record<string, unknown> = {
          node_id: nodeId,
          operation: el.querySelector('[data-type="checkbox"]') ? 'toggle_checkbox' : 'replace_text',
        };

        if (patch.operation === 'toggle_checkbox') {
          patch.checked = el.textContent?.includes('☒') ?? false;
        } else {
          patch.text = el.textContent ?? '';
        }
        patches.push(patch);
      });

      const resp = await fetch(`${PRESIGN_URL}?action=export-docx`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ originalKey, patches }),
      });
      if (!resp.ok) throw new Error(`Export failed: ${resp.status}`);
      const { url } = await resp.json() as { url: string };

      // Trigger download
      const a = Object.assign(document.createElement('a'), {
        href: url,
        download: `${docTitle || 'document'}-exported.docx`,
      });
      document.body.appendChild(a);
      a.click();
      a.remove();
      setMessage('DOCX exported.');
    } catch (e) {
      setStatus('error');
      setMessage(e instanceof Error ? e.message : 'Export failed.');
    }
  }

  const isReady = status === 'ready';
  const canCompare = isReady && docxSource !== null;

  /* ── status badge ── */
  const statusBadge = {
    idle: null,
    converting: <Badge variant="warning" className="animate-pulse">Converting…</Badge>,
    ready: message ? <span className="text-xs text-muted-foreground">{message}</span> : null,
    error: <Badge variant="destructive">{message}</Badge>,
  }[status];

  return (
    <div className="flex flex-col flex-1 overflow-hidden bg-background">

      {/* ── Document actions bar ── */}
      <header className="flex items-center gap-3 px-4 h-10 border-b bg-white shrink-0">
        {docTitle
          ? <span className="text-sm text-muted-foreground truncate max-w-[160px] sm:max-w-xs">{docTitle}</span>
          : <span className="text-sm text-muted-foreground/40 italic">No document open</span>
        }
        <div className="ml-auto flex items-center gap-2">
          {statusBadge}
          {savedAt && <span className="text-xs text-muted-foreground hidden sm:block">Saved {savedAt}</span>}
          <Button variant="outline" size="sm" className="border-sky-200 text-sky-700 hover:bg-sky-50 h-7 text-xs" onClick={() => void handleSave()} disabled={!isReady}>
            <Save className="h-3 w-3" />
            <span className="hidden sm:inline">Save HTML</span>
          </Button>
          <Button
            variant="outline" size="sm"
            className="border-emerald-200 text-emerald-700 hover:bg-emerald-50 h-7 text-xs"
            onClick={() => void handleExportDocx()}
            disabled={!isReady || !originalKey}
          >
            <Download className="h-3 w-3" />
            <span className="hidden sm:inline">Export DOCX</span>
          </Button>
          <Button size="sm" className="bg-sky-500 hover:bg-sky-600 text-white border-0 h-7 text-xs" onClick={() => void handlePdf()} disabled={!isReady}>
            <Download className="h-3 w-3" />
            <span className="hidden sm:inline">PDF</span>
          </Button>
        </div>
      </header>

      {/* ── Toolbar ── */}
      <div className="flex items-center gap-2 px-4 py-2.5 border-b bg-white shrink-0 flex-wrap">
        {/* file pick + convert */}
        <label
          htmlFor="docx-upload"
          className={cn(
            'inline-flex items-center gap-1.5 h-8 px-3 rounded-md border text-sm cursor-pointer transition-colors',
            sourceFile
              ? 'border-primary/30 bg-primary/5 text-primary font-medium'
              : 'border-input bg-background text-muted-foreground hover:bg-accent',
          )}
        >
          <Upload className="h-3.5 w-3.5" />
          <span className="max-w-[160px] truncate">{sourceFile ? sourceFile.name : 'Choose DOCX'}</span>
        </label>
        <input id="docx-upload" type="file" accept=".docx" className="sr-only" onChange={handleFileChange} />

        <label className="inline-flex items-center gap-2 h-8 px-3 rounded-md border border-input bg-background text-sm cursor-pointer hover:bg-accent transition-colors">
          <input
            type="checkbox"
            checked={useAI}
            onChange={(e) => setUseAI(e.target.checked)}
            className="w-4 h-4 rounded"
          />
          <span className="text-muted-foreground">AI Annotations</span>
        </label>

        <Button
          size="sm"
          className="bg-sky-500 hover:bg-sky-600 text-white border-0"
          onClick={() => void handleConvert()}
          disabled={!sourceFile || status === 'converting'}
        >
          <Wand2 className="h-3.5 w-3.5" />
          Convert
        </Button>

        <Separator orientation="vertical" className="h-5 mx-1 hidden sm:block" />

        {/* library picker */}
        <div className="relative">
          <button
            className="inline-flex items-center gap-1.5 h-8 px-3 rounded-md border border-input bg-background text-sm text-muted-foreground hover:bg-accent transition-colors"
            onClick={() => setLibraryOpen(v => !v)}
          >
            <BookOpen className="h-3.5 w-3.5" />
            {selectedKey
              ? (library.find(i => i.key === selectedKey)?.fileName ?? 'Library')
              : 'Library'}
            <ChevronDown className="h-3 w-3 opacity-50" />
          </button>
          {libraryOpen && (
            <div className="absolute top-full left-0 mt-1 z-50 min-w-[260px] rounded-md border bg-white shadow-lg py-1">
              {library.length === 0 && (
                <p className="px-3 py-2 text-sm text-muted-foreground">No documents yet.</p>
              )}
              {library.map(item => (
                <button
                  key={item.key}
                  className="w-full flex items-center gap-2 px-3 py-2 text-sm hover:bg-accent text-left"
                  onClick={() => { setLibraryOpen(false); void loadDoc(item.key); }}
                >
                  <FileText className="h-3.5 w-3.5 shrink-0 text-muted-foreground" />
                  <span className="truncate">{item.fileName}</span>
                  {item.status !== 'ready' && (
                    <Badge variant="warning" className="ml-auto shrink-0 text-[10px]">processing</Badge>
                  )}
                </button>
              ))}
            </div>
          )}
        </div>

        <Button
          variant="ghost" size="icon"
          className="h-8 w-8"
          title="Refresh library"
          onClick={() => void fetchLibrary().catch(console.error)}
        >
          <RefreshCw className="h-3.5 w-3.5" />
        </Button>

        {canCompare && (
          <>
            <Separator orientation="vertical" className="h-5 mx-1 hidden sm:block" />
            <Button
              variant={showCompare ? 'secondary' : 'outline'}
              size="sm"
              className="hidden sm:inline-flex"
              onClick={() => setShowCompare(v => !v)}
            >
              {showCompare
                ? <><X className="h-3.5 w-3.5" /> Close original</>
                : <><SplitSquareHorizontal className="h-3.5 w-3.5" /> Compare</>
              }
            </Button>
          </>
        )}
      </div>

      {/* ── Workspace ── */}
      <div className="flex flex-1 overflow-hidden gap-0">

        {/* original DOCX pane — hidden on mobile even if showCompare is true */}
        {showCompare && docxSource && (
          <>
            <div className="hidden sm:flex flex-col flex-1 overflow-hidden border-r">
              <div className="flex items-center gap-2 px-4 h-10 border-b bg-muted/30 shrink-0">
                <span className="text-xs font-semibold uppercase tracking-wider text-muted-foreground">Original DOCX</span>
                {docTitle && <span className="text-xs text-muted-foreground">· {docTitle}</span>}
              </div>
              <div className="flex-1 overflow-hidden">
                <DocxPane url={docxSource} />
              </div>
            </div>
          </>
        )}

        {/* HTML editor pane */}
        <div className="flex flex-col flex-1 overflow-hidden">
          <div className="flex items-center gap-2 px-4 h-10 border-b bg-muted/30 shrink-0">
            <span className="text-xs font-semibold uppercase tracking-wider text-muted-foreground">Converted HTML</span>
            {docTitle && <span className="text-xs text-muted-foreground">· {docTitle}</span>}
            {isReady && <Badge variant="outline" className="ml-auto text-[10px] text-muted-foreground">editable</Badge>}
          </div>

          <div className="flex-1 overflow-hidden">
            {!isReady && !htmlDraft ? (
              <div className="flex flex-col items-center justify-center h-full gap-3 text-center p-8">
                <div className="rounded-full bg-muted p-4">
                  <FileText className="h-8 w-8 text-muted-foreground" />
                </div>
                <p className="text-sm text-muted-foreground max-w-xs">
                  Choose a <strong>.docx</strong> file and click <strong>Convert</strong>, or open a document from the Library.
                </p>
              </div>
            ) : (
              <div className="document-pane h-full">
                <HtmlEditor
                  key={docTitle}
                  initialHtml={htmlDraft}
                  onChange={h => {
                    setHtmlDraft(h);
                    localStorage.setItem(draftKey, h);
                  }}
                />
              </div>
            )}
          </div>
        </div>
      </div>

      {/* close library dropdown on outside click */}
      {libraryOpen && (
        <div className="fixed inset-0 z-40" onClick={() => setLibraryOpen(false)} />
      )}
    </div>
  );
}

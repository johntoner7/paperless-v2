"use client";

import { forwardRef, useEffect, useImperativeHandle, useMemo, useRef, useState } from 'react';
import {
  FileText, Upload, Wand2, RefreshCw, SplitSquareHorizontal,
  Download, Save, ChevronDown, MoreHorizontal,
} from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Badge } from '@/components/ui/badge';
import { AlertBanner } from '@/components/ui/alert-banner';
import { FieldPanel } from '@/components/field-panel';
import type { ExtractedField } from '@/components/field-panel';
import { cn } from '@/lib/utils';

type ConversionState = 'idle' | 'converting' | 'ready' | 'error';
type ExtractStatus = 'idle' | 'extracting' | 'ready' | 'error';
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

type HtmlEditorHandle = {
  updateNode: (nodeId: string, text: string) => void;
  updateCheckbox: (nodeId: string, checked: boolean) => void;
  getHtml: () => string;
};

/* ── HTML editor ────────────────────────────── */
const HtmlEditor = forwardRef<HtmlEditorHandle, {
  initialHtml: string;
  onChange: (h: string) => void;
  onNodeChange?: (nodeId: string, text: string) => void;
}>(
  function HtmlEditor({ initialHtml, onChange, onNodeChange }, ref) {
    const divRef = useRef<HTMLDivElement>(null);
    const isSyncingRef = useRef(false);
    const debounceRef = useRef<ReturnType<typeof setTimeout> | null>(null);
    // Stable ref so the observer never needs to be torn down when the callback changes
    const onNodeChangeRef = useRef(onNodeChange);
    useEffect(() => { onNodeChangeRef.current = onNodeChange; });

    useImperativeHandle(ref, () => ({
      updateNode(nodeId: string, text: string) {
        const el = divRef.current?.querySelector(`[data-node-id="${nodeId}"]`);
        if (!el) return;
        isSyncingRef.current = true;
        el.textContent = text;
        // Promise microtask queues after the MutationObserver microtask
        void Promise.resolve().then(() => { isSyncingRef.current = false; });
      },
      updateCheckbox(nodeId: string, checked: boolean) {
        const runEl = divRef.current?.querySelector(`[data-node-id="${nodeId}"]`);
        if (!runEl) return;
        const cb = runEl.querySelector('[data-type="checkbox"]') as HTMLElement | null;
        if (!cb) return;
        isSyncingRef.current = true;
        cb.textContent = checked ? '☒' : '☐';
        void Promise.resolve().then(() => { isSyncingRef.current = false; });
      },
      getHtml() {
        return divRef.current?.innerHTML ?? '';
      },
    }));

    useEffect(() => {
      if (divRef.current && divRef.current.innerHTML !== initialHtml) {
        divRef.current.innerHTML = initialHtml;
      }
    }, [initialHtml]);

    // Mount/unmount only — callback is read via ref so no reconnection on re-renders
    useEffect(() => {
      const el = divRef.current;
      if (!el) return;

      const observer = new MutationObserver((mutations) => {
        if (isSyncingRef.current) return;
        for (const mutation of mutations) {
          const target = mutation.type === 'characterData'
            ? mutation.target.parentElement
            : (mutation.target instanceof Element ? mutation.target : null);
          if (!target) continue;

          // Run-level node (pre-filled value fields)
          const runEl = target.closest('[data-node-kind="run"]');
          if (runEl) {
            const nodeId = runEl.getAttribute('data-node-id');
            if (!nodeId) continue;
            const text = runEl.textContent ?? '';
            if (debounceRef.current) clearTimeout(debounceRef.current);
            debounceRef.current = setTimeout(() => onNodeChangeRef.current?.(nodeId, text), 300);
            continue;
          }

          // Paragraph-level node (empty value cells in table fields)
          const paraEl = target.closest('[data-node-kind="paragraph"]');
          if (paraEl) {
            const nodeId = paraEl.getAttribute('data-node-id');
            if (!nodeId) continue;
            const text = paraEl.textContent ?? '';
            if (debounceRef.current) clearTimeout(debounceRef.current);
            debounceRef.current = setTimeout(() => onNodeChangeRef.current?.(nodeId, text), 300);
          }
        }
      });

      observer.observe(el, { characterData: true, childList: true, subtree: true });

      // Prevent cursor from landing inside checkbox spans (they are click-only)
      const handleMouseDown = (e: MouseEvent) => {
        if ((e.target as Element).closest('[data-type="checkbox"]')) {
          e.preventDefault();
        }
      };

      // Toggle ☒ ↔ ☐ on click without letting the MutationObserver re-fire
      const handleClick = (e: MouseEvent) => {
        const cb = (e.target as Element).closest('[data-type="checkbox"]') as HTMLElement | null;
        if (!cb) return;
        const nowChecked = cb.textContent === '☒';
        const next = nowChecked ? '☐' : '☒';
        isSyncingRef.current = true;
        cb.textContent = next;
        void Promise.resolve().then(() => { isSyncingRef.current = false; });
        const runEl = cb.closest('[data-node-kind="run"]');
        if (runEl) {
          const nodeId = runEl.getAttribute('data-node-id');
          if (nodeId) onNodeChangeRef.current?.(nodeId, next);
        }
        onChange(el.innerHTML);
      };

      el.addEventListener('mousedown', handleMouseDown);
      el.addEventListener('click', handleClick);

      return () => {
        observer.disconnect();
        el.removeEventListener('mousedown', handleMouseDown);
        el.removeEventListener('click', handleClick);
        if (debounceRef.current) clearTimeout(debounceRef.current);
      };
    }, []); // eslint-disable-line react-hooks/exhaustive-deps

    return (
      <div
        ref={divRef}
        className="editor-prose"
        contentEditable
        suppressContentEditableWarning
        onInput={() => divRef.current && onChange(divRef.current.innerHTML)}
      />
    );
  }
);

/* ── Main workbench ─────────────────────────── */
export default function DocumentWorkbench() {
  const PRESIGN_URL = process.env.NEXT_PUBLIC_PRESIGN_URL ?? '';

  const [status, setStatus] = useState<ConversionState>('idle');
  const [sourceFile, setSourceFile] = useState<File | null>(null);
  const [docTitle, setDocTitle] = useState('');
  const [htmlDraft, setHtmlDraft] = useState('');
  const [library, setLibrary] = useState<LibraryUpload[]>([]);
  const [docxSource, setDocxSource] = useState<string | null>(null);
  const [originalKey, setOriginalKey] = useState<string | null>(null);
  const [showCompare, setShowCompare] = useState(false);
  const [libraryOpen, setLibraryOpen] = useState(false);
  const [useAIConvert, setUseAIConvert] = useState(false);
  const [useAIExtract, setUseAIExtract] = useState(true);
  const [notification, setNotification] = useState<{ type: 'error' | 'success'; message: string } | null>(null);
  const [fields, setFields] = useState<ExtractedField[]>([]);
  const [extractStatus, setExtractStatus] = useState<ExtractStatus>('idle');
  const [mode, setMode] = useState<'html' | 'fields'>('html');
  const [exportBusy, setExportBusy] = useState(false);
  const [exportDownloadUrl, setExportDownloadUrl] = useState<string | null>(null);
  const [openMenuOpen, setOpenMenuOpen] = useState(false);
  const [overflowOpen, setOverflowOpen] = useState(false);
  const skipRestoreRef = useRef(false);
  const statusRef = useRef(status);
  const editorRef = useRef<HtmlEditorHandle>(null);
  const paneRef = useRef<HTMLDivElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [docZoom, setDocZoom] = useState(1);
  const zoomLiveRef = useRef(1);   // current zoom without triggering renders mid-gesture
  const zoomBaseRef = useRef(1);   // zoom at the start of each new pinch
  const pinchStartDist = useRef(0);
  useEffect(() => { statusRef.current = status; }, [status]);

  // Reset zoom when a new document loads
  useEffect(() => { setDocZoom(1); zoomLiveRef.current = 1; zoomBaseRef.current = 1; }, [docTitle]);

  // Pinch-to-zoom: attach imperatively so we can call preventDefault() on touchmove
  useEffect(() => {
    const pane = paneRef.current;
    if (!pane) return;

    const pinchDist = (t: TouchList) => {
      const dx = t[0].clientX - t[1].clientX;
      const dy = t[0].clientY - t[1].clientY;
      return Math.sqrt(dx * dx + dy * dy);
    };

    const onStart = (e: TouchEvent) => {
      if (e.touches.length === 2) pinchStartDist.current = pinchDist(e.touches);
    };

    const onMove = (e: TouchEvent) => {
      if (e.touches.length !== 2) return;
      e.preventDefault(); // block viewport zoom + scroll during pinch
      const scale = zoomBaseRef.current * (pinchDist(e.touches) / pinchStartDist.current);
      const clamped = Math.min(Math.max(scale, 0.4), 3);
      zoomLiveRef.current = clamped;
      setDocZoom(clamped);
    };

    const onEnd = () => { zoomBaseRef.current = zoomLiveRef.current; };

    pane.addEventListener('touchstart', onStart, { passive: true });
    pane.addEventListener('touchmove', onMove, { passive: false });
    pane.addEventListener('touchend', onEnd);
    return () => {
      pane.removeEventListener('touchstart', onStart);
      pane.removeEventListener('touchmove', onMove);
      pane.removeEventListener('touchend', onEnd);
    };
  }, []);

  async function runExtract(key: string, force = false) {
    setExtractStatus('extracting');
    setFields([]);
    setExportDownloadUrl(null);
    try {
      const resp = await fetch(`${PRESIGN_URL}?action=extract`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ key, useAI: useAIExtract, force }),
      });
      if (!resp.ok) throw new Error(`Extract failed (${resp.status})`);
      const data = await resp.json() as { fields?: ExtractedField[] };
      setFields((data.fields ?? []).map((f, idx) => ({
        id: f.id ?? `field_${idx + 1}`,
        label: f.label,
        type: f.type,
        value: f.value ?? (f.type === 'checkbox' ? false : ''),
        confidence: f.confidence ?? 0,
        source: f.source,
      })));
      setExtractStatus('ready');
      if (statusRef.current !== 'ready') setMode('fields');
    } catch {
      setExtractStatus('error');
    }
  }

  // Determine the value target node for a field.
  // If the last node in nodeIds is a p_* (empty value cell), target that paragraph.
  // Otherwise target the last r_* run (pre-filled value or label:value paragraph).
  function getValueTarget(field: ExtractedField): string | undefined {
    const nodeIds = field.source?.nodeIds ?? [];
    if (nodeIds.length === 0) return undefined;
    const last = nodeIds[nodeIds.length - 1];
    if (last.startsWith('p_')) return last;
    return nodeIds.filter(id => id.startsWith('r_')).at(-1);
  }

  function updateFieldValue(index: number, value: string | boolean) {
    setFields(prev => prev.map((f, i) => i === index ? { ...f, value } : f));

    const field = fields[index];
    if (field.type === 'checkbox' && typeof value === 'boolean') {
      const sdtId = field.source?.nodeIds.find(id => id.startsWith('sdt_'));
      if (sdtId && editorRef.current) {
        editorRef.current.updateCheckbox(sdtId, value);
        setHtmlDraft(editorRef.current.getHtml());
      }
    } else if (field.type !== 'checkbox' && typeof value === 'string') {
      const targetId = getValueTarget(field);
      if (targetId && editorRef.current) {
        editorRef.current.updateNode(targetId, value);
        setHtmlDraft(editorRef.current.getHtml());
      }
    }
  }

  function handleNodeChange(nodeId: string, text: string) {
    setFields(prev => prev.map(f => {
      // Checkbox runs carry sdt_* node IDs — match by presence in source.nodeIds
      if (f.type === 'checkbox' && f.source?.nodeIds.includes(nodeId)) {
        return { ...f, value: text === '☒' };
      }
      return getValueTarget(f) === nodeId ? { ...f, value: text } : f;
    }));
  }

  function buildPatchesFromHtml(html: string) {
    const doc = new DOMParser().parseFromString(html, 'text/html');
    const patches: Array<Record<string, unknown>> = [];

    // Run-level nodes — text or checkbox
    doc.querySelectorAll('[data-node-kind="run"][data-node-id]').forEach(el => {
      const nodeId = el.getAttribute('data-node-id')!;
      if (el.querySelector('[data-type="checkbox"]')) {
        patches.push({ node_id: nodeId, operation: 'toggle_checkbox', checked: el.textContent?.includes('☒') ?? false });
      } else {
        patches.push({ node_id: nodeId, operation: 'replace_text', text: el.textContent ?? '' });
      }
    });

    // Paragraph-level nodes that have no child run nodes (empty value-cell paragraphs
    // where the user typed directly into the paragraph rather than into a run)
    doc.querySelectorAll('[data-node-kind="paragraph"][data-node-id]').forEach(el => {
      const nodeId = el.getAttribute('data-node-id')!;
      if (el.querySelector('[data-node-kind="run"]')) return; // runs already captured above
      patches.push({ node_id: nodeId, operation: 'replace_text', text: el.textContent ?? '' });
    });

    return patches;
  }

  async function responseError(resp: Response, fallback: string): Promise<string> {
    try {
      const body = await resp.json() as { error?: string };
      return body.error ?? fallback;
    } catch { return fallback; }
  }

  const draftKey = useMemo(() => `pb:draft:${docTitle}`, [docTitle]);

  /* draft restore */
  useEffect(() => {
    if (skipRestoreRef.current) { skipRestoreRef.current = false; return; }
    if (!docTitle) return;
    const saved = localStorage.getItem(draftKey);
    if (saved) { setHtmlDraft(saved); setStatus('ready'); setNotification({ type: 'success', message: 'Draft restored.' }); }
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
    if (!r.ok) throw new Error(await responseError(r, `Upload authorisation failed (${r.status})`));
    return r.json() as Promise<{ uploadUrl: string; key: string }>;
  }

  async function pollResult(key: string) {
    const r = await fetch(`${PRESIGN_URL}?key=${encodeURIComponent(key)}`);
    if (r.status === 404) return null;
    if (!r.ok) throw new Error(`Poll failed: ${r.status}`);
    return r.json() as Promise<{ status: string; resultUrl: string; originalUrl?: string; key: string }>;
  }

  /* shared html-conversion helper: invoke renderer + poll */
  async function pollHtml(key: string): Promise<{ html: string; originalUrl: string | null }> {
    const convResp = await fetch(`${PRESIGN_URL}?action=convert`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ key, useAI: useAIConvert }),
    });
    if (!convResp.ok) console.warn('Failed to invoke renderer, will rely on S3 events');
    for (let i = 0; i < 60; i++) {
      const result = await pollResult(key);
      if (result?.status === 'ready' && result.resultUrl) {
        const html = await fetch(result.resultUrl).then(r => r.text());
        return { html, originalUrl: result.originalUrl ?? null };
      }
      await new Promise(r => setTimeout(r, 2000));
    }
    throw new Error('Timed out waiting for conversion.');
  }

  /* convert */
  async function handleConvert() {
    if (!sourceFile) return;
    setNotification(null);
    setStatus('converting');
    try {
      const contentType = sourceFile.type
        || (sourceFile.name.toLowerCase().endsWith('.pdf') ? 'application/pdf' : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
      const { uploadUrl, key } = await presign(sourceFile.name, contentType);
      const up = await fetch(uploadUrl, {
        method: 'PUT', body: sourceFile,
        headers: { 'Content-Type': sourceFile.type || 'application/octet-stream' },
      });
      if (!up.ok) throw new Error(`Upload failed (${up.status})`);

      const title = stripExtension(sourceFile.name);
      setDocTitle(title);
      setOriginalKey(key);

      void runExtract(key);

      const { html, originalUrl } = await pollHtml(key);
      skipRestoreRef.current = true;
      setHtmlDraft(html);
      setDocxSource(originalUrl); setShowCompare(false);
      setStatus('ready');
      setNotification({ type: 'success', message: 'Conversion complete.' });

      localStorage.setItem(`pb:draft:${title}`, html);
      void fetchLibrary().catch(console.error);
    } catch (e) {
      setStatus('idle');
      setNotification({ type: 'error', message: e instanceof Error ? e.message : String(e) });
    }
  }

  async function retryConvert() {
    if (!originalKey) return;
    setStatus('converting');
    setNotification(null);
    try {
      const { html, originalUrl } = await pollHtml(originalKey);
      skipRestoreRef.current = true;
      setHtmlDraft(html);
      setDocxSource(originalUrl); setShowCompare(false);
      setStatus('ready');
      setNotification({ type: 'success', message: 'Conversion complete.' });

      if (docTitle) localStorage.setItem(`pb:draft:${docTitle}`, html);
      void fetchLibrary().catch(console.error);
    } catch (e) {
      setStatus('idle');
      setNotification({ type: 'error', message: e instanceof Error ? e.message : String(e) });
    }
  }

  async function retryExtract() {
    if (!originalKey) return;
    await runExtract(originalKey, true);
  }

  /* load from library */
  async function loadDoc(key: string) {
    const entry = library.find(i => i.key === key);
    if (!entry || entry.status !== 'ready') {
      setNotification({ type: 'error', message: `${entry?.fileName ?? key} is still processing.` });
      return;
    }
    setNotification(null);
    setStatus('converting');
    try {
      const result = await pollResult(key);
      if (!result?.resultUrl) {
        setStatus('idle');
        setNotification({ type: 'error', message: 'No converted HTML found — try re-uploading.' });
        return;
      }
      const html = await fetch(result.resultUrl).then(r => r.text());
      const title = stripExtension(entry.fileName);
      skipRestoreRef.current = true;
      setDocTitle(title); setSourceFile(null); setHtmlDraft(html);
      setDocxSource(result.originalUrl ?? null); setOriginalKey(key); setShowCompare(false);
      setStatus('ready');
      setNotification({ type: 'success', message: `Opened ${entry.fileName}` });
      void runExtract(key);

      localStorage.setItem(`pb:draft:${title}`, html);
    } catch (e) {
      setStatus('idle');
      setNotification({ type: 'error', message: e instanceof Error ? e.message : String(e) });
    }
  }

  function handleFileChange(e: React.ChangeEvent<HTMLInputElement>) {
    const f = e.target.files?.[0];
    if (!f) return;
    const ext = f.name.toLowerCase();
    if (!ext.endsWith('.docx') && !ext.endsWith('.pdf')) {
      setNotification({ type: 'error', message: 'Only .docx and .pdf files are supported.' });
      return;
    }
    setSourceFile(f); setStatus('idle'); setNotification(null);
  }

  async function handleSave() {
    if (status !== 'ready') return;
    const ts = new Date().toISOString().replace(/[:.]/g, '-');
    downloadBlob(`${docTitle || 'document'}-${ts}.html`, htmlDraft, 'text/html;charset=utf-8');
    localStorage.setItem(draftKey, htmlDraft);

    setNotification({ type: 'success', message: 'HTML saved.' });
  }

  async function handlePdf() {
    if (status !== 'ready') return;
    setNotification(null);
    try {
      const h2p = (await import('html2pdf.js')).default;
      const el = Object.assign(document.createElement('div'), { innerHTML: htmlDraft });
      await h2p().set({
        margin: 18, filename: `${docTitle}.pdf`,
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2, useCORS: true },
        jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' },
      }).from(el).save();
      setNotification({ type: 'success', message: 'PDF downloaded.' });
    } catch (e) {
      setNotification({ type: 'error', message: e instanceof Error ? e.message : 'PDF generation failed.' });
    }
  }

  async function handleExportDocx() {
    if (!originalKey) return;
    setNotification(null);
    setExportBusy(true);
    setExportDownloadUrl(null);
    try {
      const patches = buildPatchesFromHtml(htmlDraft);
      const resp = await fetch(`${PRESIGN_URL}?action=export-docx`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ originalKey, patches }),
      });
      if (!resp.ok) throw new Error(await responseError(resp, `Export failed (${resp.status})`));
      const { url } = await resp.json() as { url: string };
      setExportDownloadUrl(url);
      setNotification({ type: 'success', message: `${patches.length} patch${patches.length !== 1 ? 'es' : ''} applied — ready to download.` });
    } catch (e) {
      setNotification({ type: 'error', message: e instanceof Error ? e.message : 'Export failed.' });
    } finally {
      setExportBusy(false);
    }
  }

  const isReady = status === 'ready';
  const canCompare = isReady && docxSource !== null;

  return (
    <div className="flex flex-col flex-1 overflow-hidden bg-background">

      {/* ── Single unified header ── */}
      <header className="flex items-center gap-2 px-3 h-12 border-b bg-white shrink-0 relative z-50">

        {/* Open dropdown */}
        <div className="relative">
          <button
            onClick={() => setOpenMenuOpen(v => !v)}
            className="inline-flex items-center gap-1.5 h-8 px-3 rounded-md bg-sky-500 hover:bg-sky-600 text-white text-sm font-medium transition-colors"
          >
            <Upload className="h-3.5 w-3.5" />
            <span>Open</span>
            <ChevronDown className="h-3 w-3 opacity-70" />
          </button>
          {openMenuOpen && (
            <div className="absolute top-full left-0 mt-1 z-50 w-72 rounded-md border bg-white shadow-lg py-1">
              <label
                className="flex items-center gap-2 px-3 py-2 text-sm hover:bg-accent cursor-pointer"
                onClick={(e) => {
                  e.stopPropagation();
                  fileInputRef.current?.click();
                  setOpenMenuOpen(false);
                }}
              >
                <Upload className="h-3.5 w-3.5 text-muted-foreground shrink-0" />
                Upload document (DOCX/PDF)…
              </label>
              <input ref={fileInputRef} id="docx-upload" type="file" accept=".docx,.pdf" className="sr-only" onChange={handleFileChange} />
              {library.length > 0 && (
                <>
                  <div className="my-1 border-t" />
                  <p className="px-3 pt-1 pb-0.5 text-[11px] font-medium uppercase tracking-wide text-muted-foreground">Recent</p>
                  {library.map(item => (
                    <button
                      key={item.key}
                      className="w-full flex items-center gap-2 px-3 py-2 text-sm hover:bg-accent text-left"
                      onClick={() => { setOpenMenuOpen(false); void loadDoc(item.key); }}
                    >
                      <FileText className="h-3.5 w-3.5 shrink-0 text-muted-foreground" />
                      <span className="truncate">{item.fileName}</span>
                      {item.status !== 'ready' && (
                        <Badge variant="warning" className="ml-auto shrink-0 text-[10px]">converting</Badge>
                      )}
                    </button>
                  ))}
                </>
              )}
            </div>
          )}
        </div>

        {/* Status / title */}
        <div className="flex-1 min-w-0 flex items-center gap-2 overflow-hidden">
          {status === 'converting' ? (
            <Badge variant="warning" className="animate-pulse shrink-0">Converting…</Badge>
          ) : docTitle ? (
            <span className="text-sm text-muted-foreground truncate">{docTitle}</span>
          ) : null}
          {extractStatus === 'extracting' && (
            <Badge variant="outline" className="animate-pulse text-purple-600 border-purple-200 bg-purple-50 shrink-0 text-[10px]">Extracting…</Badge>
          )}
        </div>

        {/* Mode toggle */}
        <div className="flex rounded-md border border-input overflow-hidden text-sm shrink-0">
          <button
            className={cn('h-8 px-3 transition-colors', mode === 'html' ? 'bg-sky-500 text-white' : 'bg-background text-muted-foreground hover:bg-accent')}
            onClick={() => setMode('html')}
          >
            Doc
          </button>
          <button
            className={cn('h-8 px-3 transition-colors flex items-center gap-1 border-l border-input', mode === 'fields' ? 'bg-sky-500 text-white' : 'bg-background text-muted-foreground hover:bg-accent')}
            onClick={() => setMode('fields')}
          >
            Fields
            {fields.length > 0 && (
              <span className={cn('text-xs rounded-full min-w-[16px] text-center px-0.5', mode === 'fields' ? 'bg-white/20' : 'bg-muted')}>
                {fields.length}
              </span>
            )}
          </button>
        </div>

        {/* Export / Download */}
        {exportDownloadUrl ? (
          <a
            href={exportDownloadUrl}
            download
            className="inline-flex items-center gap-1.5 h-8 px-3 rounded-md bg-emerald-500 hover:bg-emerald-600 text-white text-sm font-medium transition-colors shrink-0"
          >
            <Download className="h-3.5 w-3.5" />
            <span className="hidden sm:inline">Download</span>
          </a>
        ) : (
          <button
            onClick={() => void handleExportDocx()}
            disabled={!originalKey || exportBusy}
            className="inline-flex items-center gap-1.5 h-8 px-3 rounded-md border border-emerald-200 text-emerald-700 hover:bg-emerald-50 disabled:opacity-40 text-sm transition-colors shrink-0"
          >
            <Download className="h-3.5 w-3.5" />
            <span className="hidden sm:inline">{exportBusy ? 'Exporting…' : 'Export'}</span>
          </button>
        )}

        {/* Overflow menu */}
        <div className="relative">
          <button
            onClick={() => setOverflowOpen(v => !v)}
            className="h-8 w-8 flex items-center justify-center rounded-md text-muted-foreground hover:bg-accent transition-colors"
            aria-label="More options"
          >
            <MoreHorizontal className="h-4 w-4" />
          </button>
          {overflowOpen && (
            <div className="absolute top-full right-0 mt-1 z-50 w-56 rounded-md border bg-white shadow-lg py-1">
              <button
                className="w-full flex items-center gap-2 px-3 py-2 text-sm hover:bg-accent text-left disabled:opacity-40"
                onClick={() => { void handleSave(); setOverflowOpen(false); }}
                disabled={!isReady}
              >
                <Save className="h-3.5 w-3.5 text-muted-foreground" />
                Save HTML
              </button>
              {canCompare && mode === 'html' && (
                <button
                  className="w-full flex items-center gap-2 px-3 py-2 text-sm hover:bg-accent text-left"
                  onClick={() => { setShowCompare(v => !v); setOverflowOpen(false); }}
                >
                  <SplitSquareHorizontal className="h-3.5 w-3.5 text-muted-foreground" />
                  {showCompare ? 'Close original' : 'Compare original'}
                </button>
              )}
              <div className="my-1 border-t" />
              <p className="px-3 pt-1 pb-0.5 text-[11px] font-medium uppercase tracking-wide text-muted-foreground">AI</p>
              <label className="flex items-center gap-2 px-3 py-2 text-sm hover:bg-accent cursor-pointer">
                <input type="checkbox" checked={useAIConvert} onChange={e => setUseAIConvert(e.target.checked)} className="w-3.5 h-3.5 rounded" />
                AI Convert
              </label>
              <label className="flex items-center gap-2 px-3 py-2 text-sm hover:bg-accent cursor-pointer">
                <input type="checkbox" checked={useAIExtract} onChange={e => setUseAIExtract(e.target.checked)} className="w-3.5 h-3.5 rounded" />
                AI Extract
              </label>
              {originalKey && (
                <>
                  <div className="my-1 border-t" />
                  {status !== 'converting' && (
                    <button
                      className="w-full flex items-center gap-2 px-3 py-2 text-sm hover:bg-accent text-left"
                      onClick={() => { void retryConvert(); setOverflowOpen(false); }}
                    >
                      <RefreshCw className="h-3.5 w-3.5 text-muted-foreground" />
                      Retry conversion
                    </button>
                  )}
                  {extractStatus !== 'extracting' && (
                    <button
                      className="w-full flex items-center gap-2 px-3 py-2 text-sm hover:bg-accent text-left"
                      onClick={() => { void retryExtract(); setOverflowOpen(false); }}
                    >
                      <RefreshCw className="h-3.5 w-3.5 text-muted-foreground" />
                      Retry extraction
                    </button>
                  )}
                </>
              )}
            </div>
          )}
        </div>
      </header>

      {/* Convert banner — appears when a file is chosen but not yet converted */}
      {sourceFile && !isReady && status !== 'converting' && (
        <div className="flex items-center gap-3 px-4 py-2 border-b bg-sky-50 shrink-0">
          <FileText className="h-4 w-4 text-sky-600 shrink-0" />
          <span className="text-sm text-sky-700 flex-1 truncate">{sourceFile.name}</span>
          <Button
            size="sm"
            className="bg-sky-500 hover:bg-sky-600 text-white border-0 shrink-0"
            onClick={() => void handleConvert()}
          >
            <Wand2 className="h-3.5 w-3.5" />
            Convert
          </Button>
        </div>
      )}

      {/* ── Notification banner ── */}
      {notification && (
        <AlertBanner
          type={notification.type}
          message={notification.message}
          onDismiss={() => setNotification(null)}
        />
      )}

      {/* ── Workspace ── */}
      <div className="flex flex-1 overflow-hidden">

        {/* HTML view */}
        <div className={cn('flex-1 overflow-hidden', mode === 'html' ? 'flex' : 'hidden')}>
          {/* original DOCX pane (compare mode, desktop only) */}
          {showCompare && docxSource && (
            <div className="hidden sm:flex flex-col flex-1 overflow-hidden border-r">
              <div className="flex items-center px-4 h-9 border-b bg-muted/30 shrink-0">
                <span className="text-xs text-muted-foreground">Original</span>
              </div>
              <div className="flex-1 overflow-hidden">
                <DocxPane url={docxSource} />
              </div>
            </div>
          )}

          {/* HTML editor pane */}
          <div className="flex flex-col flex-1 overflow-hidden">
            {/* Disclaimer */}
            {isReady && (
              <div className="shrink-0 flex items-start gap-2 px-4 py-2 border-b bg-amber-50 text-xs text-amber-700">
                <span className="shrink-0 mt-0.5">⚠</span>
                <span>Approximate preview — the exported DOCX will match the original more closely.</span>
              </div>
            )}

            <div className="flex-1 overflow-hidden">
              {!isReady && !htmlDraft ? (
                <div className="flex flex-col items-center justify-center h-full gap-3 text-center p-8">
                  <div className="rounded-full bg-muted p-4">
                    <FileText className="h-8 w-8 text-muted-foreground" />
                  </div>
                  <p className="text-sm text-muted-foreground max-w-xs">
                    Open a <strong>.docx</strong> or <strong>.pdf</strong> file to get started.
                  </p>
                </div>
              ) : (
                <div className="document-pane h-full">
                  <HtmlEditor
                    ref={editorRef}
                    key={docTitle}
                    initialHtml={htmlDraft}
                    onChange={h => {
                      setHtmlDraft(h);
                      localStorage.setItem(draftKey, h);
                    }}
                    onNodeChange={handleNodeChange}
                  />
                </div>
              )}
            </div>
          </div>
        </div>

        {/* Fields view */}
        <div className={cn('flex-1 overflow-hidden', mode === 'fields' ? 'flex flex-col' : 'hidden')}>
          <FieldPanel fields={fields} onFieldChange={updateFieldValue} />
        </div>
      </div>

      {/* Backdrop for open/overflow menus */}
      {(openMenuOpen || overflowOpen || libraryOpen) && (
        <div
          className="fixed inset-0 z-40"
          onClick={() => { setOpenMenuOpen(false); setOverflowOpen(false); setLibraryOpen(false); }}
        />
      )}
    </div>
  );
}

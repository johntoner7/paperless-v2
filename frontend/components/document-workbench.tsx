"use client";

import { useEffect, useMemo, useRef, useState } from 'react';

type ConversionState = 'idle' | 'converting' | 'ready' | 'error';

type LibraryUpload = {
  key: string;
  fileName: string;
  lastModified: string | null;
  status: 'ready' | 'pending';
};

function stripExtension(fileName: string) {
  return fileName.replace(/\.[^.]+$/, '');
}

function downloadTextFile(fileName: string, content: string, mimeType: string) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}


function EditableDocument({
  initialHtml,
  onChange,
  documentTitle,
}: {
  initialHtml: string;
  onChange: (html: string) => void;
  documentTitle: string;
}) {
  const editorRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    if (editorRef.current && editorRef.current.innerHTML !== initialHtml) {
      editorRef.current.innerHTML = initialHtml;
    }
  }, [initialHtml]);

  return (
    <article className="editor-card">
      <div className="editor-header">
        <h2>{documentTitle}</h2>
      </div>
      <div
        ref={editorRef}
        className="editor-body"
        contentEditable
        suppressContentEditableWarning
        onInput={() => {
          if (editorRef.current) {
            onChange(editorRef.current.innerHTML);
          }
        }}
      />
    </article>
  );
}

export default function DocumentWorkbench() {
  const PRESIGN_URL = process.env.NEXT_PUBLIC_PRESIGN_URL || '';
  const [status, setStatus] = useState<ConversionState>('idle');
  const [message, setMessage] = useState('Upload a .docx file.');
  const [sourceFile, setSourceFile] = useState<File | null>(null);
  const [documentTitle, setDocumentTitle] = useState('Untitled document');
  const [htmlDraft, setHtmlDraft] = useState('<p>Your converted HTML will appear here.</p>');
  const [savedAt, setSavedAt] = useState<string | null>(null);
  const [libraryUploads, setLibraryUploads] = useState<LibraryUpload[]>([]);
  const [selectedLibraryKey, setSelectedLibraryKey] = useState('');
  const skipDraftRestoreRef = useRef(false);

  const draftStorageKey = useMemo(
    () => `paperless-be:draft:${documentTitle}`,
    [documentTitle],
  );

  useEffect(() => {
    if (skipDraftRestoreRef.current) {
      skipDraftRestoreRef.current = false;
      return;
    }

    const savedDraft = window.localStorage.getItem(draftStorageKey);
    if (savedDraft) {
      setHtmlDraft(savedDraft);
      setStatus('ready');
      setMessage('Draft restored.');
    }
  }, [draftStorageKey]);

  async function refreshLibraryUploads() {
    if (!PRESIGN_URL) return;

    const url = `${PRESIGN_URL}?action=list`;
    const res = await fetch(url, { method: 'GET' });
    if (!res.ok) throw new Error(`Failed to load uploaded files: ${res.status}`);

    const payload = await res.json() as { items?: LibraryUpload[] };
    setLibraryUploads(payload.items || []);
  }

  useEffect(() => {
    void refreshLibraryUploads().catch((err: unknown) => {
      console.error(err);
    });
  }, [PRESIGN_URL]);

  async function requestPresign(fileName: string, contentType: string) {
    if (!PRESIGN_URL) throw new Error('Presign endpoint not configured (NEXT_PUBLIC_PRESIGN_URL)');
    const res = await fetch(PRESIGN_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ fileName, contentType }),
    });
    if (!res.ok) throw new Error(`Presign request failed: ${res.status}`);
    return res.json(); // { uploadUrl, key }
  }

  async function checkResult(key: string) {
    if (!PRESIGN_URL) throw new Error('Presign endpoint not configured (NEXT_PUBLIC_PRESIGN_URL)');
    const url = `${PRESIGN_URL}?key=${encodeURIComponent(key)}`;
    const res = await fetch(url, { method: 'GET' });
    if (res.status === 404) return null;
    if (!res.ok) throw new Error(`Status check failed: ${res.status}`);
    return res.json(); // { status: 'ready', resultUrl }
  }

  async function loadLibraryUpload(key: string) {
    setSelectedLibraryKey(key);

    const entry = libraryUploads.find((item) => item.key === key);
    if (!entry) {
      setMessage('Selected file is no longer available.');
      return;
    }

    if (entry.status !== 'ready') {
      setMessage(`${entry.fileName} is uploaded, but converted HTML is not ready yet.`);
      return;
    }

    setStatus('converting');
    setMessage(`Loading ${entry.fileName}...`);

    try {
      const result = await checkResult(entry.key);
      if (!result || result.status !== 'ready' || !result.resultUrl) {
        setStatus('idle');
        setMessage(`${entry.fileName} has no converted HTML yet.`);
        return;
      }

      const htmlRes = await fetch(result.resultUrl);
      if (!htmlRes.ok) throw new Error(`Failed to fetch result: ${htmlRes.status}`);

      const html = await htmlRes.text();
      const nextTitle = stripExtension(entry.fileName);
      skipDraftRestoreRef.current = true;
      setDocumentTitle(nextTitle);
      setSourceFile(null);
      setHtmlDraft(html);
      setStatus('ready');
      setMessage(`Loaded ${entry.fileName}.`);
      setSavedAt(null);
      window.localStorage.setItem(`paperless-be:draft:${nextTitle}`, html);
    } catch (err: unknown) {
      console.error(err);
      setStatus('error');
      setMessage(err instanceof Error ? err.message : String(err));
    }
  }

  async function handleConvert() {
    if (!sourceFile) {
      setMessage('No source file selected.');
      return;
    }

    setStatus('converting');
    setMessage('Requesting upload URL...');

    try {
      const { uploadUrl, key } = await requestPresign(
        sourceFile.name,
        sourceFile.type || 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      );

      setMessage('Uploading...');
      const putRes = await fetch(uploadUrl, {
        method: 'PUT',
        body: sourceFile,
        headers: { 'Content-Type': sourceFile.type || 'application/octet-stream' },
      });
      if (!putRes.ok) throw new Error(`Upload failed: ${putRes.status}`);

      setMessage('Waiting for conversion...');

      const maxAttempts = 60;
      const intervalMs = 2000;
      for (let attempt = 0; attempt < maxAttempts; attempt++) {
        const res = await checkResult(key);
        if (res && res.status === 'ready' && res.resultUrl) {
          const htmlRes = await fetch(res.resultUrl);
          if (!htmlRes.ok) throw new Error(`Failed to fetch result: ${htmlRes.status}`);
          const html = await htmlRes.text();
          const nextTitle = stripExtension(sourceFile.name);
          const storageKey = `paperless-be:draft:${nextTitle}`;
          setDocumentTitle(nextTitle);
          setHtmlDraft(html);
          setStatus('ready');
          setMessage('Conversion complete.');
          setSavedAt(null);
          window.localStorage.setItem(storageKey, html);
          void refreshLibraryUploads().catch((refreshErr: unknown) => {
            console.error(refreshErr);
          });
          return;
        }
        await new Promise((r) => setTimeout(r, intervalMs));
      }

      setStatus('error');
      setMessage('Timed out waiting for conversion.');
    } catch (err: unknown) {
      console.error(err);
      setStatus('error');
      setMessage(err instanceof Error ? err.message : String(err));
    }
  }

  function handleFileUpload(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    if (!file.name.toLowerCase().endsWith('.docx')) {
      setStatus('error');
      setMessage('Please upload a .docx file.');
      return;
    }
    setSourceFile(file);
    setStatus('idle');
    setMessage(`${file.name} selected. Click Convert to process.`);
  }

  function handleSaveHtml() {
    if (status !== 'ready') return;
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const fileName = `${documentTitle || 'document'}-edited-${timestamp}.html`;
    downloadTextFile(fileName, htmlDraft, 'text/html;charset=utf-8');
    window.localStorage.setItem(draftStorageKey, htmlDraft);
    setSavedAt(new Date().toLocaleString());
    setMessage('Saved.');
  }

  async function handleExportPdf() {
    if (status !== 'ready') return;
    setMessage('Generating PDF…');
    try {
      const html2pdf = (await import('html2pdf.js')).default;
      const el = document.createElement('div');
      el.innerHTML = htmlDraft;
      await html2pdf()
        .set({
          margin: 18,
          filename: `${documentTitle}.pdf`,
          image: { type: 'jpeg', quality: 0.98 },
          html2canvas: { scale: 2, useCORS: true },
          jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' },
        })
        .from(el)
        .save();
      setMessage('PDF downloaded.');
    } catch (err: unknown) {
      setStatus('error');
      setMessage(err instanceof Error ? err.message : 'PDF export failed.');
    }
  }

  function handleRestoreSample() {
    const sample = '<h1>Sample document</h1><p>Edit this text.</p>';
    setStatus('ready');
    setMessage('Sample loaded.');
    skipDraftRestoreRef.current = true;
    setDocumentTitle('sample-document');
    setHtmlDraft(sample);
    setSourceFile(null);
    setSavedAt(null);
  }

  return (
    <main className="shell">
      <section className="toolbar">
        <div className="toolbar-row">
          <label className="upload-button" htmlFor="docx-upload">
            <span>{sourceFile ? sourceFile.name : 'Upload DOCX'}</span>
          </label>
          <input id="docx-upload" type="file" accept=".docx" onChange={handleFileUpload} />
          <button type="button" className="secondary-button" onClick={handleConvert} disabled={!sourceFile || status === 'converting'}>
            Convert
          </button>
          <button type="button" className="secondary-button" onClick={handleRestoreSample}>
            Sample
          </button>
          <button type="button" className="primary-button" onClick={handleSaveHtml} disabled={status !== 'ready'}>
            Save
          </button>
          <button type="button" className="primary-button alt" onClick={() => { void handleExportPdf(); }} disabled={status !== 'ready'}>
            PDF
          </button>
        </div>

        <div className="toolbar-row toolbar-row-picker">
          <label className="picker-label" htmlFor="uploaded-pdf-picker">
            Previously uploaded documents
          </label>
          <select
            id="uploaded-pdf-picker"
            className="picker-select"
            value={selectedLibraryKey}
            onChange={(event) => {
              void loadLibraryUpload(event.target.value);
            }}
            disabled={libraryUploads.length === 0}
          >
            <option value="">{libraryUploads.length ? 'Choose a document' : 'No uploaded documents found'}</option>
            {libraryUploads.map((item) => (
              <option key={item.key} value={item.key}>
                {item.fileName}{item.status !== 'ready' ? ' (processing)' : ''}
              </option>
            ))}
          </select>
          <button
            type="button"
            className="secondary-button"
            onClick={() => {
              void refreshLibraryUploads().catch((err: unknown) => {
                console.error(err);
              });
            }}
          >
            Refresh
          </button>
        </div>
      </section>

      <p className={`status status-${status}`}>{message}</p>

      <section className="workspace-grid">
        <EditableDocument
          key={documentTitle}
          initialHtml={htmlDraft}
          onChange={(nextHtml) => {
            setHtmlDraft(nextHtml);
            window.localStorage.setItem(draftStorageKey, nextHtml);
          }}
          documentTitle={documentTitle}
        />
      </section>

      <p className="footer-note">{savedAt ? `Saved ${savedAt}` : ' '}</p>
    </main>
  );
}

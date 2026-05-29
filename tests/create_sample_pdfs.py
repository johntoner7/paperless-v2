"""Generate representative sample PDFs for the pdf2htmlEX POC.

Requires: weasyprint (already in backend/requirements.txt)

Usage:
    pip install weasyprint
    python tests/create_sample_pdfs.py

Outputs to tests/fixtures/:
    sample_text.pdf    -- text-heavy (headings, paragraphs, a table)
    sample_layout.pdf  -- layout-heavy (two-column, coloured boxes, an image placeholder)
"""
import os
import sys

try:
    import weasyprint
except ImportError:
    sys.exit('weasyprint not installed. Run: pip install weasyprint')

FIXTURES = os.path.join(os.path.dirname(__file__), 'fixtures')
os.makedirs(FIXTURES, exist_ok=True)

# --- text-heavy ---
TEXT_HTML = """<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  body { font-family: Georgia, serif; margin: 2cm; font-size: 11pt; }
  h1   { font-size: 20pt; margin-bottom: 0.3cm; }
  h2   { font-size: 14pt; margin-top: 0.8cm; }
  p    { line-height: 1.6; }
  table{ border-collapse: collapse; width: 100%; margin-top: 0.5cm; }
  th, td { border: 1px solid #444; padding: 4px 8px; }
  th   { background: #ddd; }
</style>
</head>
<body>
<h1>Sample Text-Heavy Document</h1>
<p>This document is used as a representative input for the pdf2htmlEX proof-of-concept.
It contains typical prose content with headings, paragraphs, and a table so that
text extraction fidelity can be evaluated.</p>

<h2>Section 1 — Introduction</h2>
<p>Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor
incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud
exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.</p>

<h2>Section 2 — Data Summary</h2>
<table>
  <thead><tr><th>Item</th><th>Quantity</th><th>Unit Price</th><th>Total</th></tr></thead>
  <tbody>
    <tr><td>Widget A</td><td>10</td><td>$5.00</td><td>$50.00</td></tr>
    <tr><td>Widget B</td><td>3</td><td>$20.00</td><td>$60.00</td></tr>
    <tr><td>Widget C</td><td>7</td><td>$8.50</td><td>$59.50</td></tr>
  </tbody>
</table>

<h2>Section 3 — Conclusion</h2>
<p>Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu
fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa
qui officia deserunt mollit anim id est laborum.</p>
</body>
</html>"""

# --- layout-heavy ---
LAYOUT_HTML = """<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  @page { size: A4; margin: 1.5cm; }
  body  { font-family: Arial, sans-serif; font-size: 10pt; }
  .header { background: #2c3e50; color: white; padding: 12px 16px; margin-bottom: 12px; }
  .header h1 { margin: 0; font-size: 18pt; }
  .cols { display: flex; gap: 12px; }
  .col  { flex: 1; }
  .box  { border: 1px solid #ccc; border-radius: 4px; padding: 10px; margin-bottom: 10px; }
  .box.blue { background: #ebf5fb; border-color: #3498db; }
  .box.green{ background: #eafaf1; border-color: #27ae60; }
  .box.red  { background: #fdedec; border-color: #e74c3c; }
  .box h3   { margin: 0 0 6px 0; font-size: 11pt; }
  .img-placeholder {
    width: 100%; height: 80px;
    background: #bdc3c7;
    display: flex; align-items: center; justify-content: center;
    color: #7f8c8d; font-size: 9pt;
  }
  table { border-collapse: collapse; width: 100%; font-size: 9pt; }
  th, td { border: 1px solid #ccc; padding: 3px 6px; }
  th { background: #ecf0f1; }
</style>
</head>
<body>
<div class="header"><h1>Layout-Heavy Sample Document</h1></div>
<div class="cols">
  <div class="col">
    <div class="box blue">
      <h3>Status</h3>
      <p>All systems operational. Last checked: 2026-05-29.</p>
    </div>
    <div class="box green">
      <h3>Metrics</h3>
      <table>
        <tr><th>Metric</th><th>Value</th></tr>
        <tr><td>Uptime</td><td>99.97 %</td></tr>
        <tr><td>Latency</td><td>42 ms</td></tr>
        <tr><td>Errors</td><td>0.03 %</td></tr>
      </table>
    </div>
  </div>
  <div class="col">
    <div class="box red">
      <h3>Alerts</h3>
      <p>No active alerts.</p>
    </div>
    <div class="box">
      <h3>Image area</h3>
      <div class="img-placeholder">[ chart placeholder ]</div>
    </div>
  </div>
</div>
</body>
</html>"""


def write(name: str, html: str) -> None:
    path = os.path.join(FIXTURES, name)
    print(f'Writing {path} ...', end=' ', flush=True)
    weasyprint.HTML(string=html).write_pdf(path)
    size = os.path.getsize(path)
    print(f'{size:,} bytes')


write('sample_text.pdf', TEXT_HTML)
write('sample_layout.pdf', LAYOUT_HTML)
print('Done. Sample PDFs are in tests/fixtures/')

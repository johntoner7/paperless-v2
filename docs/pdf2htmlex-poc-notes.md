# pdf2htmlEX POC — Notes and Decisions

## Phase 1 — Conversion path

| Decision | Choice | Reason |
|---|---|---|
| Input path | S3-triggered (same as DOCX flow) | Reuses existing Lambda event shape |
| Output shape | Single self-contained HTML file | All assets embedded via `--embed-*` flags |
| Output storage | S3 under `converted/<uuid>/` prefix | Matches existing presign_handler pattern |
| Sample PDFs | `tests/fixtures/sample_text.pdf`, `tests/fixtures/sample_layout.pdf` | Generated with weasyprint from representative HTML |

A third "scanned/image-heavy" PDF should be added manually from a real scanned document
when available. The two generated samples cover text-layout and complex CSS layout.

### Repeatable conversion command

With pdf2htmlEX installed locally:

```sh
pdf2htmlEX \
  --dest-dir /tmp/out \
  --embed-css 1 --embed-font 1 --embed-image 1 \
  --embed-javascript 1 --embed-outline 1 \
  --process-outline 0 \
  tests/fixtures/sample_text.pdf sample_text.html
```

Via the wrapper script:

```sh
python tests/pdf_poc_test.py tests/fixtures/sample_text.pdf /tmp/pdf_poc_out
```

---

## Phase 2 — Converter

`backend/pdf_converter.py` is the converter module.

- Accepts `input_pdf` (path) and `output_dir`.
- Returns a dict with `html_path`, `duration_sec`, `success`, `exit_code`, `stderr`.
- Binary location controlled by `PDF2HTML_BIN` env var (default: `pdf2htmlEX`).
- Timeout controlled by caller (default: 120 s).
- Temp directory cleanup is the caller's responsibility (use `tempfile.TemporaryDirectory`).

---

## Phase 3 — Lambda packaging

`Dockerfile.lambda` installs pdf2htmlEX via miniconda + conda-forge. Expected image size
increase: ~600 MB (miniconda ~200 MB + pdf2htmlex package ~400 MB). Lambda supports images
up to 10 GB, so this is within limits.

Build and test locally:

```sh
docker build -f Dockerfile.lambda -t paperless-lambda-poc .
docker run --rm \
  -v "$PWD/tests/fixtures:/fixtures:ro" \
  -v "/tmp/poc-out:/tmp/poc-out" \
  -e PDF2HTML_BIN=/usr/local/bin/pdf2htmlEX \
  paperless-lambda-poc \
  python /var/task/pdf_converter.py /fixtures/sample_text.pdf /tmp/poc-out
```

`lambda_function.py` routes `.pdf` inputs to `pdf_converter.convert()` and all other
inputs (`.docx`, etc.) through the existing `run_renderer` path.

---

## Known limitations / open questions

- pdf2htmlEX has not had an official release since 2020; conda-forge maintains a patched build.
- Very large PDFs (>50 pages) may exceed the Lambda 15-minute timeout.
- Scanned PDFs (image-only) will produce HTML with embedded images but no selectable text.
  OCR is a separate concern (not in scope for this POC).
- Font embedding increases HTML file size significantly for font-heavy PDFs.

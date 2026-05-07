# Paperless BE Frontend

A simple Next.js + TypeScript frontend for testing the document conversion flow in a browser.

What it does:
- Upload a `.docx` file.
- Convert it to editable HTML in the browser using Mammoth.
- Edit the rendered document directly.
- Save changes as an HTML download and browser draft.
- Export to PDF using the browser print dialog.

## Local development

```bash
cd frontend
npm install
npm run dev
```

Then open `http://localhost:3000`.

## Deploy to Vercel

- Import the `frontend/` folder as the project root in Vercel.
- Vercel will detect this as a Next.js app.
- No special build step is needed beyond `npm run build`.

## Notes

- This frontend converts DOCX files in the browser, so it works without the AWS backend while you're testing.
- If you want server-side conversion later, point the UI at an API Gateway or Lambda Function URL and replace the browser conversion step.

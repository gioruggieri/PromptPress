# PromptPress Studio

Next.js web app for pasting LLM output (Gemini, GPT, Claude), previewing Markdown with LaTeX, and downloading DOCX or PDF already laid out.

## Stack
- Next.js 16 (App Router) + React 19
- Tailwind CSS (v4) for UI
- `react-markdown` + `remark-math` + `rehype-katex` for math rendering
- Server-side DOCX export with `docx` + KaTeX/MathML -> OMML conversion (`mathml2omml`) for editable Word equations
- Server-side PDF export with `puppeteer` (Chromium) to avoid `html2canvas` limitations with modern CSS (e.g. `oklab()`)

## Local setup
```bash
npm install
npm run dev
```
Open `http://localhost:3000`.

## Key features
- Paste or load a sample LLM response.
- Markdown/GFM support, inline math `$E=mc^2$`, and `$$ ... $$` blocks.
- Auto-fix mode for LLM output: removes indentation that turns into code blocks and converts `\(...\)` / `\[...\]` to `$...$` / `$$...$$`.
- Live preview with KaTeX, modern styling, and separate input/output sections.
- Export to `promptpress-export.docx` or `promptpress-export.pdf` directly from the browser.

## Notes
- DOCX: `/api/docx` generates a file with Word Equation (OMML) math. If a formula fails to convert, it stays visible as monospaced TeX (it does not disappear).
- PDF: `/api/pdf` generates a multi-page HTML-to-PDF via Chromium (faithful render, no `oklab()` issues).

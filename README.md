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

## Quick start (Tampermonkey + ChatGPT copy)
1) Install Tampermonkey in your browser.
2) Create a new userscript and paste the script below.
3) Open `https://chatgpt.com`, click "Copy clean" on a response, and paste into PromptPress.
4) Keep "Auto-fix for LLM output" enabled for best results.

```javascript
// ==UserScript==
// @name         ChatGPT Copy Clean (tables + math)
// @match        https://chatgpt.com/*
// @grant        GM_setClipboard
// ==/UserScript==
(function () {
  const stripZeroWidth = (text) =>
    text.replace(/[\u200B\u200C\u200D\u2060\uFEFF\u00AD\u202A-\u202E\u2066-\u2069]/g, "");

  const normalizeMatrixRows = (tex) => {
    if (!/\\begin\{(bmatrix|pmatrix|matrix|aligned|align\*?|cases|array)\}/.test(tex)) {
      return tex;
    }
    return tex.replace(/(?<!\\)\\\s*(?=[0-9+\-\s\]])/g, "\\\\");
  };

  const flattenMath = (tex) =>
    tex
      .replace(/\r\n?/g, "\n")
      .split("\n")
      .map((l) => l.trim())
      .filter(Boolean)
      .join(" ");

  const wrapMath = (tex, displayHint) => {
    let body = normalizeMatrixRows(tex);
    body = flattenMath(body);

    const wantsDisplay =
      displayHint ||
      /\\begin\{[^}]+\}/.test(body) ||
      /\\end\{[^}]+\}/.test(body);

    if (wantsDisplay) return `\n\n$$\n${body}\n$$\n\n`;
    return `$${body}$`;
  };

  const isNoise = (el) => {
    const tag = el.tagName.toLowerCase();
    if (tag === "button") return true;
    const aria = (el.getAttribute("aria-label") || "").toLowerCase();
    if (aria.includes("copy code") || aria.includes("copia codice")) return true;
    const testId = (el.getAttribute("data-testid") || "").toLowerCase();
    if (testId.includes("copy") && testId.includes("code")) return true;
    return false;
  };

  const renderTable = (tableEl) => {
    const rows = Array.from(tableEl.querySelectorAll("tr"));
    if (!rows.length) return "";

    const headerRow = tableEl.querySelector("thead tr") || rows[0];
    const headerCells = Array.from(headerRow.querySelectorAll("th, td"));
    if (!headerCells.length) return "";

    const colCount = headerCells.length;
    const bodyRows =
      headerRow === rows[0] ? rows.slice(1) : rows.filter((r) => r !== headerRow);

    const renderCell = (cell) => {
      const walkCell = (node) => {
        if (node.nodeType === Node.TEXT_NODE) {
          return stripZeroWidth(node.textContent || "");
        }
        if (node.nodeType !== Node.ELEMENT_NODE) return "";
        const el = node;

        if (el.classList.contains("katex") || el.classList.contains("katex-display")) {
          const ann = el.querySelector('annotation[encoding="application/x-tex"]');
          const tex = ann ? ann.textContent.trim() : (el.textContent || "").trim();
          const display = el.classList.contains("katex-display");
          return wrapMath(tex, display);
        }

        const tag = el.tagName.toLowerCase();
        if (tag === "br") return "\n";

        return Array.from(el.childNodes).map(walkCell).join("");
      };

      const raw = Array.from(cell.childNodes).map(walkCell).join("");
      let text = raw
        .replace(/\n\s*\$\$\s*\n([\s\S]*?)\n\s*\$\$\s*\n/g, (_, math) => `$${String(math).trim()}$`)
        .replace(/\n+/g, " ")
        .replace(/\|/g, "\\|")
        .trim();

      return text.length ? text : " ";
    };

    const buildRow = (cells) => {
      const values = [];
      for (let i = 0; i < colCount; i += 1) {
        values.push(cells[i] ? renderCell(cells[i]) : " ");
      }
      return `| ${values.join(" | ")} |`;
    };

    const headerLine = buildRow(headerCells);
    const separator = `| ${Array(colCount).fill("---").join(" | ")} |`;
    const bodyLines = bodyRows
      .map((row) => buildRow(Array.from(row.querySelectorAll("th, td"))))
      .join("\n");

    return `\n\n${headerLine}\n${separator}${bodyLines ? `\n${bodyLines}` : ""}\n\n`;
  };

  const htmlToMarkdown = (root) => {
    const out = [];
    const walk = (node) => {
      if (node.nodeType === Node.TEXT_NODE) {
        out.push(stripZeroWidth(node.textContent || ""));
        return;
      }
      if (node.nodeType !== Node.ELEMENT_NODE) return;
      const el = node;

      if (isNoise(el)) return;

      if (el.classList.contains("katex") || el.classList.contains("katex-display")) {
        const ann = el.querySelector('annotation[encoding="application/x-tex"]');
        const tex = ann ? ann.textContent.trim() : (el.textContent || "").trim();
        const display = el.classList.contains("katex-display");
        out.push(wrapMath(tex, display));
        return;
      }

      const tag = el.tagName.toLowerCase();
      if (tag === "table") {
        out.push(renderTable(el));
        return;
      }
      if (tag === "pre") {
        const code = el.textContent || "";
        out.push(`\n\n\`\`\`\n${code.replace(/\n$/, "")}\n\`\`\`\n\n`);
        return;
      }
      if (tag === "br") {
        out.push("\n");
        return;
      }

      const children = Array.from(el.childNodes);
      children.forEach(walk);

      if (tag === "p" || tag === "div") out.push("\n\n");
      if (/^h[1-6]$/.test(tag)) out.push("\n\n");
    };

    walk(root);
    return out.join("").replace(/\n{3,}/g, "\n\n").trim();
  };

  const addButtons = () => {
    document.querySelectorAll('[data-message-id]').forEach((msg) => {
      if (msg.querySelector('.copy-clean-btn')) return;
      const btn = document.createElement('button');
      btn.textContent = 'Copy clean';
      btn.className = 'copy-clean-btn';
      btn.style.marginLeft = '8px';
      btn.onclick = () => copyClean(msg);
      const toolbar = msg.querySelector('[data-testid="toolbox"]') || msg;
      toolbar.appendChild(btn);
    });
  };

  const copyClean = (msg) => {
    const html = msg.innerHTML;
    const div = document.createElement('div');
    div.innerHTML = html;

    const md = htmlToMarkdown(div);
    GM_setClipboard(md, 'text');
    alert('Copiato pulito');
  };

  setInterval(addButtons, 1000);
})();

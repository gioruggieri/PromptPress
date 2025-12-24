import { NextResponse } from "next/server";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

export const runtime = "nodejs";

let katexCssPromise: Promise<string> | null = null;

const isVercel = Boolean(process.env.VERCEL);

const getKatexCss = async () => {
  if (!katexCssPromise) {
    katexCssPromise = (async () => {
      const katexDistDir = path.join(
        process.cwd(),
        "node_modules",
        "katex",
        "dist",
      );
      const katexCssPath = path.join(katexDistDir, "katex.min.css");
      const fontsDir = path.join(katexDistDir, "fonts");
      const fontsBaseUrl = pathToFileURL(fontsDir + path.sep).toString();

      const css = await readFile(katexCssPath, "utf8");

      // Rewrite KaTeX font URLs so Chromium can load fonts from disk when rendering
      // via `page.setContent()` (otherwise it falls back to generic fonts).
      return css.replace(
        /url\((['"]?)fonts\/([^'")]+)\1\)/g,
        (_, quote: string, file: string) =>
          `url(${quote}${fontsBaseUrl}${file}${quote})`,
      );
    })();
  }
  return katexCssPromise;
};

const exportCss = `
  :root {
    color-scheme: light;
  }

  html, body {
    margin: 0;
    padding: 0;
    background: #ffffff;
    color: #0f172a;
    font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Arial, "Apple Color Emoji", "Segoe UI Emoji";
  }

  .page {
    padding: 22mm 18mm;
  }

  .markdown {
    max-width: 100%;
    font-size: 12pt;
    line-height: 1.45;
  }

  .markdown h1,
  .markdown h2,
  .markdown h3,
  .markdown h4 {
    font-weight: 700;
    line-height: 1.2;
    margin: 18pt 0 10pt;
    color: #0b1224;
  }

  .markdown h1 { font-size: 22pt; }
  .markdown h2 { font-size: 18pt; }
  .markdown h3 { font-size: 15pt; }

  .markdown p {
    margin: 10pt 0;
  }

  .markdown ul,
  .markdown ol {
    margin: 8pt 0 10pt 18pt;
    padding: 0 0 0 8pt;
  }

  .markdown li {
    margin: 4pt 0;
  }

  .markdown code {
    font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace;
    background: rgba(15, 23, 42, 0.06);
    padding: 0.15em 0.35em;
    border-radius: 6px;
  }

  .markdown pre {
    background: #0b1224;
    color: #f8fafc;
    border-radius: 12px;
    padding: 12pt;
    overflow-x: auto;
  }

  .markdown pre code {
    background: transparent;
    padding: 0;
  }

  .markdown blockquote {
    margin: 12pt 0;
    padding: 10pt 12pt;
    border-left: 4px solid #0ea5e9;
    background: rgba(14, 165, 233, 0.08);
    border-radius: 10px;
    color: #0f172a;
  }

  .markdown table {
    border-collapse: collapse;
    width: 100%;
    margin: 12pt 0;
  }

  .markdown th,
  .markdown td {
    border: 1px solid rgba(15, 23, 42, 0.16);
    padding: 7pt 9pt;
    vertical-align: top;
  }

  .markdown th {
    background: rgba(99, 102, 241, 0.14);
    text-align: left;
  }

  .markdown hr {
    margin: 14pt 0;
    border: none;
    height: 1px;
    background: rgba(15, 23, 42, 0.16);
  }

  /* KaTeX tweaks */
  .katex {
    font-size: 1.05em;
    color: #0f172a;
  }
  .katex-display {
    margin: 12pt 0;
    padding: 10pt 12pt;
    overflow-x: auto;
    border-radius: 12px;
    border: 1px solid rgba(15, 23, 42, 0.12);
    background: rgba(15, 23, 42, 0.03);
  }

  /* Avoid breaking a KaTeX block across pages when possible */
  .katex-display {
    break-inside: avoid;
    page-break-inside: avoid;
  }
`;

export async function POST(request: Request) {
  let browser: import("puppeteer-core").Browser | null = null;
  try {
    const { html } = await request.json();

    if (!html || typeof html !== "string") {
      return NextResponse.json({ error: "Missing HTML" }, { status: 400 });
    }

    const [katexCss, chromium, puppeteer] = await Promise.all([
      getKatexCss(),
      isVercel ? import("@sparticuz/chromium") : Promise.resolve(null),
      isVercel ? import("puppeteer-core") : import("puppeteer"),
    ]);

    const fullHtml = `<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <style>${katexCss}\n${exportCss}</style>
  </head>
  <body>
    <div class="page">
      <div class="markdown">${html}</div>
    </div>
  </body>
</html>`;

    if (isVercel && chromium) {
      const executablePath = await chromium.executablePath();
      browser = await puppeteer.launch({
        args: chromium.args,
        defaultViewport: chromium.defaultViewport,
        executablePath,
        headless: chromium.headless,
      });
    } else {
      browser = await puppeteer.launch({
        headless: true,
        args: ["--no-sandbox", "--disable-setuid-sandbox"],
      });
    }

    const page = await browser.newPage();
    await page.setContent(fullHtml, { waitUntil: ["domcontentloaded"] });
    await page.emulateMediaType("screen");

    // Wait a tick to allow KaTeX layout to settle.
    await page.evaluate(() => new Promise((resolve) => requestAnimationFrame(() => resolve(null))));

    const pdf = await page.pdf({
      format: "A4",
      printBackground: true,
      preferCSSPageSize: true,
      margin: { top: "10mm", right: "10mm", bottom: "12mm", left: "10mm" },
    });

    const body = new Blob([Buffer.from(pdf)], { type: "application/pdf" });

    return new NextResponse(body, {
      status: 200,
      headers: {
        "Content-Type": "application/pdf",
        "Content-Disposition": 'attachment; filename="promptpress-export.pdf"',
      },
    });
  } catch (error) {
    console.error("PDF export error", error);
    return NextResponse.json({ error: "Error generating PDF" }, { status: 500 });
  } finally {
    if (browser) {
      try {
        await browser.close();
      } catch {
        // ignore
      }
    }
  }
}

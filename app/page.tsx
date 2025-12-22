"use client";

import { useEffect, useMemo, useRef, useState, type ClipboardEvent } from "react";
import ReactMarkdown, { type Components } from "react-markdown";
import rehypeKatex from "rehype-katex";
import { MathMLToLaTeX } from "mathml-to-latex";
import remarkGfm from "remark-gfm";
import remarkMath from "remark-math";
import clsx from "clsx";

const starter = `# PromptPress style
Copied from an LLM, formatted to be readable and ready to export.

## Sample content
- **Inline formulas:** $E = mc^2$, \(\pi r^2\)
- **Math block:**
$$
\\int_0^{\\infty} e^{-x^2} dx = \\frac{\\sqrt{\\pi}}{2}
$$
- **Hybrid list**
  1. Title + short abstract
  2. Table with units
  3. Final notes section

### Code snippet
\`\`\`python
def softmax(logits):
    exps = [pow(2.71828, x) for x in logits]
    s = sum(exps)
    return [v/s for v in exps]
\`\`\`

### Table
| Variable | Description | Value |
|-----------|-------------|--------|
| $r$ | Radius | $4.2$ |
| $A$ | Area | $55.4\,m^2$ |

> Tip: paste raw output from Gemini or your LLM here.
`;

const looksLikeLatex = (text: string) =>
  /\\[a-zA-Z]+/.test(text) || /[_^]\{?[\w(]/.test(text);

type ExtractedMath = { tex: string; display: boolean };

const hasMathMarkers = (text: string) =>
  /(\$\$|\\\(|\\\[|\\begin\{)/.test(text) || looksLikeLatex(text);

const stripZeroWidth = (text: string) =>
  text.replace(/[\u200B\u200C\u200D\u2060\u2062\uFEFF\u00AD\u202A-\u202E\u2066-\u2069]/g, "");

const normalizeMathMlLatex = (texRaw: string) => {
  const tex = texRaw.trim();
  if (!tex) return tex;

  const wrappedMatrices: Array<{ open: string; close: string; env: string }> = [
    { open: "\\left[\\right.", close: "\\left]\\right.", env: "bmatrix" },
    { open: "\\left(\\right.", close: "\\left)\\right.", env: "pmatrix" },
    { open: "\\left\\{\\right.", close: "\\left\\}\\right.", env: "Bmatrix" },
  ];

  for (const { open, close, env } of wrappedMatrices) {
    if (!tex.startsWith(open) || !tex.endsWith(close)) continue;
    const inner = tex.slice(open.length, tex.length - close.length).trim();
    if (!inner) break;
    return `\\begin{${env}}\n${inner}\n\\end{${env}}`;
  }

  if ((tex.includes("&") || tex.includes("\\\\")) && !tex.includes("\\begin{")) {
    return `\\begin{matrix}\n${tex}\n\\end{matrix}`;
  }

  return tex;
};

const clipboardHtmlToMarkdown = (html: string) => {
  try {
    const fragmentMatch = html.match(
      /<!--StartFragment-->([\s\S]*?)<!--EndFragment-->/i,
    );
    const fragment = fragmentMatch ? fragmentMatch[1] : html;
    const doc = new DOMParser().parseFromString(fragment, "text/html");

    const normalizeText = (value: string) =>
      stripZeroWidth(value).replace(/\u00A0/g, " ");

    const isMathContainer = (el: Element) => {
      const tag = el.tagName.toLowerCase();
      if (tag === "math") return true;
      if (tag === "mjx-container" || tag === "mjx-assistive-mml") return true;
      if (tag === "script" && /^math\/tex/i.test(el.getAttribute("type") ?? ""))
        return true;
      if (tag === "img" && Boolean(el.getAttribute("alt"))) return true;
      if (el.classList.contains("katex-display")) return true;
      if (el.classList.contains("katex")) return true;
      if (el.classList.contains("MathJax")) return true;
      return false;
    };

    const isClipboardNoiseElement = (el: Element) => {
      const tag = el.tagName.toLowerCase();
      if (tag === "button") return true;

      const aria = (el.getAttribute("aria-label") ?? "").toLowerCase();
      if (aria.includes("copy code") || aria.includes("copia codice")) return true;

      const title = (el.getAttribute("title") ?? "").toLowerCase();
      if (title.includes("copy code") || title.includes("copia codice")) return true;

      const testId = (el.getAttribute("data-testid") ?? "").toLowerCase();
      if (testId.includes("copy") && testId.includes("code")) return true;
      if (testId.includes("code-block-header")) return true;

      const className = el.className?.toString().toLowerCase() ?? "";
      if (className.includes("copy-code")) return true;
      if (className.includes("code-block-header")) return true;
      if (className.includes("code-header") && tag !== "pre" && tag !== "code")
        return true;

      return false;
    };

    const extractTexFromElement = (el: Element): ExtractedMath | null => {
      const tag = el.tagName.toLowerCase();
      const display =
        el.classList.contains("katex-display") ||
        Boolean(el.closest(".katex-display")) ||
        (tag === "math" && el.getAttribute("display") === "block") ||
        (tag === "mjx-container" &&
          /^(true|block)$/i.test(el.getAttribute("display") ?? ""));

      const annotation = el.querySelector?.(
        'annotation[encoding="application/x-tex"], annotation[encoding="application/x-latex"], annotation[encoding="application/tex"]',
      );
      if (annotation?.textContent) {
        return { tex: annotation.textContent, display };
      }

      if (tag === "script") {
        const type = el.getAttribute("type") ?? "";
        const tex = el.textContent ?? "";
        const displayMode = /mode=display/i.test(type);
        return { tex, display: displayMode || display };
      }

      if (tag === "img") {
        const alt = el.getAttribute("alt") ?? "";
        if (alt) return { tex: alt, display };
      }

      if (tag === "math") {
        try {
          const tex = normalizeMathMlLatex(MathMLToLaTeX.convert(el.outerHTML));
          return { tex, display };
        } catch {
          // ignore conversion failures
        }
      }

      const mathDescendant = el.querySelector?.("math");
      if (mathDescendant) {
        const inner = extractTexFromElement(mathDescendant);
        if (!inner) return null;
        return { tex: inner.tex, display: display || inner.display };
      }

      const nodeText = el.textContent?.trim();
      if (nodeText && looksLikeLatex(nodeText)) return { tex: nodeText, display };
      return null;
    };

    const renderMath = (el: Element) => {
      const extracted = extractTexFromElement(el);
      if (!extracted) {
        const fallback = el.textContent?.trim();
        if (fallback) return `$${fallback}$`;
        return "```html\n" + el.outerHTML + "\n```";
      }

      let tex = extracted.tex.trim();
      tex = tex.replace(/^\$+\s*/, "").replace(/\s*\$+$/, "");
      if (!tex) {
        const fallback = el.textContent?.trim();
        if (fallback) return `$${fallback}$`;
        return "```html\n" + el.outerHTML + "\n```";
      }

      if (extracted.display) return `\n\n$$\n${tex}\n$$\n\n`;
      return `$${tex}$`;
    };

    type ListContext = { ordered: boolean; index: number };

    const formatListItem = (prefix: string, raw: string) => {
      const normalized = raw.replace(/\r\n?/g, "\n").trimEnd();
      const lines = normalized.split("\n");

      while (lines.length && !lines[0].trim()) lines.shift();
      while (lines.length && !lines[lines.length - 1].trim()) lines.pop();

      if (!lines.length) return prefix.trimEnd();

      const first = (lines.shift() ?? "").trim();
      if (!lines.length) return `${prefix}${first}`.trimEnd();

      const rest = lines
        .map((line) => (line.length ? `  ${line}` : ""))
        .join("\n");

      return `${prefix}${first}\n${rest}`.trimEnd();
    };

    const renderInline = (node: Node, list?: ListContext): string => {
      if (node.nodeType === Node.TEXT_NODE) {
        return normalizeText(node.textContent ?? "");
      }
      if (node.nodeType !== Node.ELEMENT_NODE) return "";
      const el = node as Element;

      if (isClipboardNoiseElement(el)) return "";
      if (isMathContainer(el)) return renderMath(el);

      const tag = el.tagName.toLowerCase();
      if (tag === "br") return "\n";
      if (tag === "ul" || tag === "ol") return "";

      const content = Array.from(el.childNodes)
        .map((child) => renderInline(child, list))
        .join("");

      if (tag === "strong" || tag === "b") return `**${content}**`;
      if (tag === "em" || tag === "i") return `*${content}*`;
      if (tag === "code") return "`" + content + "`";
      if (tag === "a") return content;
      return content;
    };

    const renderTable = (tableEl: Element) => {
      const rows = Array.from(tableEl.querySelectorAll("tr"));
      if (!rows.length) return "";

      const headerRow = tableEl.querySelector("thead tr") ?? rows[0];
      const headerCells = Array.from(headerRow.querySelectorAll("th, td"));
      if (!headerCells.length) return "";

      const colCount = headerCells.length;
      const bodyRows =
        headerRow === rows[0]
          ? rows.slice(1)
          : rows.filter((row) => !row.closest("thead") && row !== headerRow);

      const renderCell = (cell: Element) => {
        const raw = Array.from(cell.childNodes)
          .map((child) => renderInline(child))
          .join("");
        let text = raw.replace(
          /\n\s*\$\$\s*\n([\s\S]*?)\n\s*\$\$\s*\n/g,
          (_, math) => `$${String(math).trim()}$`,
        );
        text = text.replace(/\n+/g, "<br />").replace(/\|/g, "\\|").trim();
        return text.length ? text : " ";
      };

      const buildRow = (cells: Element[]) => {
        const values = [];
        for (let i = 0; i < colCount; i += 1) {
          const cell = cells[i];
          values.push(cell ? renderCell(cell) : " ");
        }
        return `| ${values.join(" | ")} |`;
      };

      const headerLine = buildRow(headerCells);
      const separatorLine = `| ${Array(colCount).fill("---").join(" | ")} |`;
      const bodyLines = bodyRows
        .map((row) => buildRow(Array.from(row.querySelectorAll("th, td"))))
        .join("\n");

      return `\n\n${headerLine}\n${separatorLine}${
        bodyLines ? `\n${bodyLines}` : ""
      }\n\n`;
    };

    const renderBlock = (node: Node, list?: ListContext): string => {
      if (node.nodeType === Node.TEXT_NODE) {
        const text = normalizeText(node.textContent ?? "");
        return text.trim().length ? text : "";
      }
      if (node.nodeType !== Node.ELEMENT_NODE) return "";
      const el = node as Element;

      if (isClipboardNoiseElement(el)) return "";
      if (isMathContainer(el)) return `${renderMath(el)}\n\n`;

      const tag = el.tagName.toLowerCase();

      if (tag === "br") return "\n";

      if (tag === "pre") {
        const codeEl = el.querySelector("code");
        const className = codeEl?.getAttribute("class") ?? "";
        const langMatch =
          className.match(/language-([a-zA-Z0-9+-]+)/) ??
          className.match(/lang(?:uage)?-([a-zA-Z0-9+-]+)/);
        const lang = langMatch ? langMatch[1] : "";
        const code = normalizeText(
          codeEl?.textContent ?? el.textContent ?? "",
        ).replace(/\n$/, "");
        return `\n\n\`\`\`${lang}\n${code}\n\`\`\`\n\n`;
      }

      if (tag === "table") return renderTable(el);

      if (tag === "ul" || tag === "ol") {
        const ordered = tag === "ol";
        const items = Array.from(el.children).filter(
          (child) => child.tagName.toLowerCase() === "li",
        );

        const lines = items
          .map((li, idx) => {
            const prefix = ordered ? `${idx + 1}. ` : "- ";
            const body = renderInline(li).trimEnd();
            const nestedLists = Array.from(li.children).filter((c) =>
              ["ul", "ol"].includes(c.tagName.toLowerCase()),
            );
            const nestedText = nestedLists
              .map((nested) => renderBlock(nested).trimEnd())
              .filter(Boolean)
              .join("\n");

            const combined = nestedText ? `${body}\n\n${nestedText}` : body;
            return formatListItem(prefix, combined);
          })
          .join("\n");

        return `${lines}\n\n`;
      }

      if (tag === "li") {
        const prefix = list
          ? list.ordered
            ? `${list.index}. `
            : "- "
          : "- ";
        const content = renderInline(el);
        return `${formatListItem(prefix, content)}\n`;
      }

      if (/^h[1-6]$/.test(tag)) {
        const level = Number.parseInt(tag.slice(1), 10);
        const hashes = "#".repeat(Math.max(1, Math.min(6, level)));
        const text = renderInline(el).replace(/\n+/g, " ").trim();
        return `${hashes} ${text}\n\n`;
      }

      if (tag === "p") {
        const text = renderInline(el).trim();
        return text ? `${text}\n\n` : "";
      }

      if (tag === "div" || tag === "section" || tag === "article") {
        const content = Array.from(el.childNodes)
          .map((child) => renderBlock(child))
          .join("")
          .trim();
        return content ? `${content}\n\n` : "";
      }

      const content = Array.from(el.childNodes)
        .map((child) => renderBlock(child))
        .join("");
      return content;
    };

    const output = Array.from(doc.body.childNodes)
      .map((node) => renderBlock(node))
      .join("")
      .replace(/\r\n?/g, "\n")
      .replace(/[ \t]+\n/g, "\n")
      .replace(/\n{3,}/g, "\n\n")
      .trim();

    return output;
  } catch {
    return "";
  }
};

const extractMathFromClipboardHtml = (html: string): ExtractedMath[] => {
  try {
    const fragmentMatch = html.match(
      /<!--StartFragment-->([\s\S]*?)<!--EndFragment-->/i,
    );
    const fragment = fragmentMatch ? fragmentMatch[1] : html;

    const doc = new DOMParser().parseFromString(fragment, "text/html");
    const out: ExtractedMath[] = [];
    const seen = new Set<string>();

    const push = (texRaw: string, display: boolean) => {
      let tex = texRaw.trim();
      if (!tex) return;

      tex = tex.replace(/^\$+\s*/, "").replace(/\s*\$+$/, "");

      const key = `${display ? "d" : "i"}:${tex}`;
      if (seen.has(key)) return;
      seen.add(key);
      out.push({ tex, display });
    };

    const annotations = Array.from(
      doc.querySelectorAll(
        'annotation[encoding="application/x-tex"], annotation[encoding="application/x-latex"], annotation[encoding="application/tex"]',
      ),
    );
    for (const node of annotations) {
      const tex = node.textContent ?? "";
      const display =
        Boolean(node.closest(".katex-display")) ||
        node.closest("math")?.getAttribute("display") === "block";
      push(tex, display);
    }

    const scripts = Array.from(doc.querySelectorAll('script[type^="math/tex"]'));
    for (const node of scripts) {
      const type = node.getAttribute("type") ?? "";
      const tex = node.textContent ?? "";
      const display = /mode=display/i.test(type);
      push(tex, display);
    }

    const images = Array.from(doc.querySelectorAll("img[alt]"));
    for (const node of images) {
      const alt = node.getAttribute("alt") ?? "";
      if (!alt) continue;
      if (!hasMathMarkers(alt) && !/(\$|\\)/.test(alt)) continue;
      const display =
        alt.includes("$$") || alt.includes("\\[") || alt.includes("\\begin{");
      push(alt, display);
    }

    const mathNodes = Array.from(doc.querySelectorAll("math"));
    for (const node of mathNodes) {
      if (
        node.querySelector(
          'annotation[encoding="application/x-tex"], annotation[encoding="application/x-latex"], annotation[encoding="application/tex"]',
        )
      ) {
        continue;
      }

      try {
        const tex = normalizeMathMlLatex(MathMLToLaTeX.convert(node.outerHTML));
        const display = node.getAttribute("display") === "block";
        push(tex, display);
      } catch {
        // ignore conversion failures
      }
    }

    const attrNodes = Array.from(
      doc.querySelectorAll(
        "[data-latex],[data-tex],[data-math],[aria-label],[title],[alt]",
      ),
    );
    const attrNames = [
      "data-latex",
      "data-tex",
      "data-math",
      "aria-label",
      "title",
      "alt",
    ];
    for (const node of attrNodes) {
      const display =
        Boolean(node.closest(".katex-display")) ||
        node.closest("math")?.getAttribute("display") === "block" ||
        node.className.toString().includes("display");

      for (const attr of attrNames) {
        const value = node.getAttribute?.(attr);
        if (!value) continue;
        if (!hasMathMarkers(value) && !value.includes("\\begin{")) continue;
        push(value, display);
      }
    }

    const text = doc.body?.textContent ?? "";
    for (const match of text.matchAll(/\$\$([\s\S]*?)\$\$/g)) {
      push(match[1], true);
    }
    for (const match of text.matchAll(/\\\[([\s\S]*?)\\\]/g)) {
      push(match[1], true);
    }
    for (const match of text.matchAll(/\\\(([\s\S]*?)\\\)/g)) {
      push(match[1], false);
    }
    for (const match of text.matchAll(
      /\\begin\{([a-zA-Z*]+)\}[\s\S]*?\\end\{\1\}/g,
    )) {
      push(match[0], true);
    }

    return out;
  } catch {
    return [];
  }
};

const rtfToText = (rtf: string) => {
  let out = "";
  let index = 0;
  let unicodeSkipCount = 1;
  let skipChars = 0;

  const pushChar = (char: string) => {
    out += char;
  };

  while (index < rtf.length) {
    const char = rtf[index];

    if (skipChars > 0) {
      skipChars -= 1;
      index += 1;
      continue;
    }

    if (char === "{" || char === "}") {
      index += 1;
      continue;
    }

    if (char !== "\\") {
      pushChar(char);
      index += 1;
      continue;
    }

    const next = rtf[index + 1];
    if (!next) {
      index += 1;
      continue;
    }

    // Control symbols.
    if (next === "\\" || next === "{" || next === "}") {
      pushChar(next);
      index += 2;
      continue;
    }

    if (next === "~") {
      pushChar(" ");
      index += 2;
      continue;
    }

    if (next === "_" || next === "-") {
      pushChar("-");
      index += 2;
      continue;
    }

    if (next === "'") {
      const hex = rtf.slice(index + 2, index + 4);
      if (/^[0-9a-fA-F]{2}$/.test(hex)) {
        pushChar(String.fromCharCode(parseInt(hex, 16)));
        index += 4;
        continue;
      }
      index += 2;
      continue;
    }

    if (next === "*") {
      // Destination control symbol: skip.
      index += 2;
      continue;
    }

    // Control words.
    let cursor = index + 1;
    let word = "";
    while (cursor < rtf.length && /[a-zA-Z]/.test(rtf[cursor])) {
      word += rtf[cursor];
      cursor += 1;
    }

    let sign = 1;
    if (rtf[cursor] === "-") {
      sign = -1;
      cursor += 1;
    }

    let numberText = "";
    while (cursor < rtf.length && /[0-9]/.test(rtf[cursor])) {
      numberText += rtf[cursor];
      cursor += 1;
    }

    const numberValue = numberText ? sign * Number.parseInt(numberText, 10) : null;

    if (rtf[cursor] === " ") {
      cursor += 1;
    }

    switch (word) {
      case "par":
      case "line":
        pushChar("\n");
        break;
      case "tab":
        pushChar("\t");
        break;
      case "uc":
        if (numberValue !== null) {
          unicodeSkipCount = Math.max(0, numberValue);
        }
        break;
      case "u":
        if (numberValue !== null) {
          const codePoint = numberValue < 0 ? numberValue + 65536 : numberValue;
          pushChar(String.fromCharCode(codePoint));
          skipChars = unicodeSkipCount;
        }
        break;
      default:
        break;
    }

    index = cursor;
  }

  return out
    .replace(/\r\n?/g, "\n")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
};

const extractMathFromClipboardRtf = (rtf: string): ExtractedMath[] => {
  try {
    const text = rtfToText(rtf);
    const out: ExtractedMath[] = [];
    const seen = new Set<string>();
    const push = (texRaw: string, display: boolean) => {
      const tex = texRaw.trim();
      if (!tex) return;
      const key = `${display ? "d" : "i"}:${tex}`;
      if (seen.has(key)) return;
      seen.add(key);
      out.push({ tex, display });
    };

    for (const match of text.matchAll(/\$\$([\s\S]*?)\$\$/g)) {
      push(match[1], true);
    }

    for (const match of text.matchAll(/\\begin\{([a-zA-Z*]+)\}[\s\S]*?\\end\{\1\}/g)) {
      push(match[0], true);
    }

    return out;
  } catch {
    return [];
  }
};

const wantsDisplayMath = ({ tex, display }: ExtractedMath) =>
  display ||
  tex.includes("\\begin{") ||
  tex.includes("&") ||
  tex.includes("\\\\") ||
  tex.includes("=") ||
  tex.length > 80;

const fillMathPlaceholders = (plain: string, math: ExtractedMath[]) => {
  const normalized = plain.replace(/\r\n?/g, "\n");
  const displayMath = math.filter(wantsDisplayMath);
  if (!displayMath.length) return { text: normalized, usedKeys: new Set<string>() };

  const makeKey = (m: ExtractedMath) => `${m.display ? "d" : "i"}:${m.tex.trim()}`;
  const usedKeys = new Set<string>();

  const lines = normalized.split("\n");
  const out: string[] = [];
  let cursor = 0;

  const cueRegex =
    /(definit|matric|matrix|matriciale|sistema|moltiplicazione|diventano|diventa|componenti|vettor|vector|frequenz|pari|dispari)/i;

  while (cursor < lines.length) {
    const line = lines[cursor];
    out.push(line);

    if (usedKeys.size < displayMath.length) {
      let lookahead = cursor + 1;
      let blankCount = 0;
      while (lookahead < lines.length && lines[lookahead].trim() === "") {
        blankCount += 1;
        lookahead += 1;
      }
      const hasBlankPlaceholder = blankCount >= 1;

      const trimmed = line.trim();
      const isLikelyCue =
        cueRegex.test(trimmed) || /^\s*[*+-]\s+\*\*/.test(line);

      if (hasBlankPlaceholder && isLikelyCue) {
        const formula = displayMath[usedKeys.size];
        usedKeys.add(makeKey(formula));

        const indent = /^\s*(?:\d+\.|[*+-])\s+/.test(line) ? "  " : "";
        out.push("");
        out.push(`${indent}$$`);
        out.push(`${indent}${formula.tex.trim()}`);
        out.push(`${indent}$$`);
        out.push("");

        // Skip the existing blank placeholder lines.
        cursor = lookahead;
        continue;
      }
    }

    cursor += 1;
  }

  const text = out
    .join("\n")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trimEnd();

  return { text, usedKeys };
};

const injectMathIntoPlainText = (plain: string, math: ExtractedMath[]) => {
  const normalized = plain.replace(/\r\n?/g, "\n");
  if (!math.length) return normalized;

  const blocks = math
    .map(({ tex, display }) => {
      const wantsDisplay = wantsDisplayMath({ tex, display });
      if (wantsDisplay) return `$$\n${tex}\n$$`;
      return `$${tex}$`;
    })
    .join("\n\n");

  const cueRegex =
    /(matric|matrix|equaz|equation|formula|sistema|moltiplicazione)/i;
  const lines = normalized.split("\n");
  let insertionIndex = -1;
  let offset = 0;

  for (const line of lines) {
    const trimmed = line.trim();
    if (trimmed.length && trimmed.endsWith(":") && cueRegex.test(trimmed)) {
      const afterLine = offset + line.length;
      const gap = normalized.indexOf("\n\n", afterLine);
      if (gap !== -1) insertionIndex = gap + 2;
      else insertionIndex = Math.min(afterLine + 1, normalized.length);
      break;
    }
    offset += line.length + 1;
  }

  if (insertionIndex === -1) {
    const firstGap = normalized.indexOf("\n\n");
    insertionIndex = firstGap !== -1 ? firstGap + 2 : normalized.length;
  }

  let head = normalized.slice(0, insertionIndex);
  let tail = normalized.slice(insertionIndex);

  head = head.replace(/\n*$/, "\n\n");
  tail = tail.replace(/^\n+/, "");

  return `${head}${blocks}\n\n${tail}`.trimEnd();
};

const buildSmartPasteText = (plain: string, html?: string, rtf?: string) => {
  const plainNormalized = plain.replace(/\r\n?/g, "\n");
  const plainClean = stripZeroWidth(plainNormalized);

  const fromHtml = html ? clipboardHtmlToMarkdown(html) : "";
  const fromHtmlClean = stripZeroWidth(fromHtml);

  const looksLikeRenderedCopyCorruption = (text: string, htmlText?: string) => {
    const normalized = text.replace(/\r\n?/g, "\n");
    const hasIndentLikeCodeBlock = /(?:^|\n)[ \t]{4,}\S/.test(normalized);

    // Typical symptom of copying KaTeX-rendered HTML: the MathML + visual layer
    // both contribute to plain text, producing immediate duplicates (e.g. `2i2i`, `pospos`).
    const duplicateTokens =
      normalized.match(/([0-9]+[a-zA-Z]|[a-zA-Z]{2,})\1/g) ?? [];
    const hasDuplicateTokens = duplicateTokens.length >= 1;

    const htmlHasList = Boolean(htmlText && /<(ul|ol|li)\b/i.test(htmlText));
    const htmlHasMath = Boolean(
      htmlText && /(katex|mathjax|mjx-|<math\b)/i.test(htmlText),
    );

    return (htmlHasMath && hasDuplicateTokens) || (htmlHasList && hasIndentLikeCodeBlock);
  };

  const htmlHasMath = Boolean(fromHtmlClean && hasMathMarkers(fromHtmlClean));
  const htmlHasRichBlocks = Boolean(
    html && /<(table|pre|code|ul|ol|blockquote|h[1-6])\b/i.test(html),
  );
  const plainHasMath = hasMathMarkers(plainClean);
  const plainHasEmptyDisplayMath = /\$\$\s*\$\$/.test(plainClean);

  // If the clipboard plain text looks like it came from rendered math (duplicates/ZW chars)
  // and the HTML contains recoverable TeX/MathML, prefer rebuilding from HTML instead of
  // injecting formulas at the end (which causes duplicates like `pospos`, `jj`, ...).
  const extractedFromHtml = html ? extractMathFromClipboardHtml(html) : [];
  const extractedFromRtf = rtf ? extractMathFromClipboardRtf(rtf) : [];
  const extracted =
    extractedFromHtml.length > 0 ? extractedFromHtml : extractedFromRtf;

  const plainSeemsToMissMostMath =
    htmlHasMath &&
    extractedFromHtml.length > 0 &&
    (() => {
      const important = extractedFromHtml.filter(
        (m) => wantsDisplayMath(m) || m.tex.includes("\\frac"),
      );
      if (important.length < 2) return false;
      let contained = 0;
      for (const m of important) {
        const t = m.tex.trim();
        if (!t) continue;
        if (
          plainClean.includes(t) ||
          plainClean.includes(`$${t}$`) ||
          plainClean.includes(`$$${t}$$`)
        ) {
          contained += 1;
        }
      }
      return contained / important.length < 0.35;
    })();

  // If HTML has math or rich blocks (tables/code), prefer the HTML version to avoid
  // flattened text/plain copies that lose structure.
  if ((htmlHasMath || htmlHasRichBlocks) && fromHtml) return fromHtml;

  if (!extracted.length) return plainNormalized;

  const normalizedPlain = plainNormalized;
  const displayMath = extracted.filter(wantsDisplayMath);

  // Some sources (LLMs/web UIs) paste `$$ $$` placeholders but omit the inner formula in `text/plain`.
  // If we detect empty display-math blocks, fill them in order using extracted formulas.
  if (displayMath.length > 0 && /\$\$\s*\$\$/.test(normalizedPlain)) {
    let index = 0;
    const filled = normalizedPlain.replace(/\$\$\s*\$\$/g, () => {
      const next = displayMath[index++];
      if (!next) return "$$\n\n$$";
      return `$$\n${next.tex}\n$$`;
    });
    if (filled !== normalizedPlain) return filled;
  }

  const missing = extracted.filter(({ tex }) => {
    const trimmed = tex.trim();
    return (
      !plainNormalized.includes(trimmed) &&
      !plainNormalized.includes(`$${trimmed}$`) &&
      !plainNormalized.includes(`$$${trimmed}$$`)
    );
  });

  if (!missing.length) return plainNormalized;

  const filled = fillMathPlaceholders(plainNormalized, missing);
  const remaining = missing.filter(
    (m) => !filled.usedKeys.has(`${m.display ? "d" : "i"}:${m.tex.trim()}`),
  );

  if (!remaining.length) return filled.text;
  if (filled.text !== plainNormalized)
    return injectMathIntoPlainText(filled.text, remaining);
  return injectMathIntoPlainText(plainNormalized, remaining);
};

const escapeRegExp = (value: string) =>
  value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

const splitDisplayDelimiters = (line: string) => {
  const parts: string[] = [];
  let rest = line;

  while (rest.length > 0) {
    const matchIndex = rest.search(/\\\[|\\\]/);
    if (matchIndex === -1) {
      parts.push(rest);
      break;
    }

    const before = rest.slice(0, matchIndex);
    if (before.trim().length > 0) parts.push(before);

    parts.push("$$");

    rest = rest.slice(matchIndex + 2);
  }

  return parts.length ? parts : [line];
};

const splitDoubleDollarMath = (line: string) => {
  const first = line.indexOf("$$");
  if (first === -1) return null;

  const second = line.indexOf("$$", first + 2);
  if (second === -1) {
    // Handle non-standard but common single-line fences like `$$E=mc^2` or `E=mc^2$$`
    // so our `inDisplayMath` state stays aligned with what remark-math will parse.
    const openMatch = line.match(/^(\s*)\$\$(\S[\s\S]*)$/);
    if (openMatch) {
      const indent = openMatch[1] ?? "";
      const rest = (openMatch[2] ?? "").trim();
      if (!rest) return [`${indent}$$`];
      return [`${indent}$$`, `${indent}${rest}`];
    }

    const closeMatch = line.match(/^(\s*)([\s\S]*\S)\$\$(\s*)$/);
    if (closeMatch) {
      const indent = closeMatch[1] ?? "";
      const rest = (closeMatch[2] ?? "").trimEnd();
      if (!rest) return [`${indent}$$`];
      return [`${indent}${rest}`, `${indent}$$`];
    }

    return null;
  }

  const beforeRaw = line.slice(0, first);
  const before = beforeRaw.trimEnd();
  const math = line.slice(first + 2, second).trim();
  const after = line.slice(second + 2).trimStart();

  if (!math) return null;

  const listPrefixMatch = beforeRaw.match(/^(\s*(?:\d+\.|[-*+])\s+)$/);

  const parts: string[] = [];
  if (!before) {
    parts.push("$$");
    parts.push(math);
    parts.push("$$");
    if (after) parts.push(after);
    return parts;
  }

  if (listPrefixMatch) {
    const prefix = before;
    const indent = " ".repeat(listPrefixMatch[1].length);
    parts.push(`${prefix}$$`);
    parts.push(`${indent}${math}`);
    parts.push(`${indent}$$`);
    if (after) parts.push(after);
    return parts;
  }

  parts.push(before);
  parts.push("$$");
  parts.push(math);
  parts.push("$$");
  if (after) parts.push(after);
  return parts;
};

const normalizeCopyCodeBlocks = (input: string) => {
  const normalized = input.replace(/\r\n?/g, "\n");
  const lines = normalized.split("\n");
  const out: string[] = [];
  let inFence = false;
  let fenceToken: "```" | "~~~" | null = null;
  let implicitFence = false;
  let implicitFenceHasContent = false;
  let pendingLangLine: string | null = null;

  const pushArrowLines = (value: string) => {
    if (!value) return;
    const parts = value.split(/\s*↓\s*/);
    if (parts.length === 1) {
      out.push(value);
      return;
    }
    const first = parts[0].trim();
    if (first) out.push(first);
    for (let i = 1; i < parts.length; i += 1) {
      out.push("↓");
      const chunk = parts[i].trim();
      if (chunk) out.push(chunk);
    }
  };

  const deriveLangFromLabel = (label: string) => {
    const tokens = label
      .replace(/[^a-zA-Z0-9#+.-]+/g, " ")
      .trim()
      .split(/\s+/)
      .filter(Boolean);
    if (!tokens.length) return "";
    let langRaw = "";
    if (tokens.length >= 2) {
      const last = tokens[tokens.length - 1];
      const prev = tokens[tokens.length - 2];
      if (last.length <= 10 && prev.length <= 12) {
        langRaw = `${prev}-${last}`;
      }
    }
    if (!langRaw) {
      langRaw = tokens[tokens.length - 1] ?? "";
    }
    return /^[a-zA-Z0-9#+.-]{1,32}$/.test(langRaw) ? langRaw : "";
  };

  const isPotentialLangLine = (line: string) => {
    const trimmed = stripZeroWidth(line).trim();
    if (!trimmed) return false;
    if (trimmed.length > 32) return false;
    if (/[:/]/.test(trimmed)) return false;
    if (/[.,;!?]/.test(trimmed)) return false;
    const tokens = trimmed.split(/\s+/);
    if (tokens.length > 3) return false;
    return tokens.every((token) => /^[a-zA-Z0-9#+.-]{1,16}$/.test(token));
  };

  const openImplicitFence = (lang: string, rest: string) => {
    out.push(lang ? `\`\`\`${lang}` : "```");
    if (rest) {
      pushArrowLines(rest);
      implicitFenceHasContent = rest.trim().length > 0;
    } else {
      implicitFenceHasContent = false;
    }
    implicitFence = true;
  };

  const closeImplicitFence = () => {
    if (!implicitFence) return;
    out.push("```");
    implicitFence = false;
    implicitFenceHasContent = false;
  };

  for (let i = 0; i < lines.length; i += 1) {
    const line = lines[i];
    const fenceMatch = line.match(/^\s*(```|~~~)/);
    if (fenceMatch) {
      if (pendingLangLine) {
        out.push(pendingLangLine);
        pendingLangLine = null;
      }
      closeImplicitFence();
      const token = fenceMatch[1] as "```" | "~~~";
      if (!inFence) {
        inFence = true;
        fenceToken = token;
      } else if (fenceToken === token) {
        inFence = false;
        fenceToken = null;
      }
      out.push(line);
      continue;
    }

    if (inFence) {
      out.push(line);
      continue;
    }

    if (!implicitFence && !pendingLangLine && isPotentialLangLine(line)) {
      const nextLine = lines[i + 1] ?? "";
      const nextLower = stripZeroWidth(nextLine).toLowerCase();
      if (/(copy\s*code|copia\s*codice)/.test(nextLower)) {
        pendingLangLine = line;
        continue;
      }
    }

    if (implicitFence) {
      if (!line.trim()) {
        if (implicitFenceHasContent) {
          closeImplicitFence();
          out.push(line);
        }
        continue;
      }
      pushArrowLines(line);
      implicitFenceHasContent = true;
      continue;
    }

    const lower = stripZeroWidth(line).toLowerCase();
    const tokenMatch = lower.match(/(copy\s*code|copia\s*codice)/);
    if (tokenMatch) {
      const tokenIndex = tokenMatch.index ?? -1;
      if (tokenIndex >= 0 && tokenIndex <= 80) {
        const before = line.slice(0, tokenIndex);
        const lang = pendingLangLine
          ? deriveLangFromLabel(pendingLangLine)
          : deriveLangFromLabel(before);
        const rest = line.slice(tokenIndex + tokenMatch[0].length).trim();
        pendingLangLine = null;
        openImplicitFence(lang, rest);
        continue;
      }

      if (line.trim().toLowerCase() === tokenMatch[0].trim()) {
        pendingLangLine = null;
        continue;
      }
    }

    if (pendingLangLine) {
      out.push(pendingLangLine);
      pendingLangLine = null;
    }
    out.push(line);
  }

  if (pendingLangLine) out.push(pendingLangLine);
  closeImplicitFence();
  return out.join("\n");
};

const fixLatexInlineTokens = (input: string) => {
  let line = input;
  line = line.replace(/\\text\{([^}\n]+)(?=\s|\\|$)/g, "\\text{$1}");
  line = line.replace(/(_\{[^}]*)(?=[\s\\]|$)/g, "$1}");
  line = line.replace(
    /(\\text\{[^}]+\})\s*\{\s*\\text\{([^}]+)\}\s*\}/g,
    "$1_{\\text{$2}}",
  );
  line = line.replace(
    /([A-Za-z0-9]+)\s*\{\s*\\text\{([^}]+)\}\s*\}/g,
    "$1_{\\text{$2}}",
  );
  line = line.replace(
    /\\text\{([^}]+)\}\s*\{\s*\\text\{([^}]+)\}\s*\}/g,
    "\\text{$1}_{\\text{$2}}",
  );
  // \text{logits} \text{mela} or \text{logits}\_\text{mela} -> \text{logits}_{\text{mela}}
  line = line.replace(
    /\\text\{([^}]+)\}\s*\\_?\s*\\text\{([^}]+)\}/g,
    "\\text{$1}_{\\text{$2}}",
  );
  if (
    /\\begin\{(bmatrix|pmatrix|matrix|aligned|align\*?)\}/.test(line) ||
    line.includes("&") ||
    /\\\\/.test(line) ||
    /\\end\{(bmatrix|pmatrix|matrix|aligned|align\*?)\}/.test(line)
  ) {
    line = line.replace(/(?<!\\)\\\s*(?=[0-9+\-\s\]])/g, "\\\\");
  }
  return line;
};

const normalizeLlmMarkdown = (input: string) => {
  const normalized = promoteInlineMathBlocks(liftUnbalancedInlineMath(input));
  const lines = normalized.split("\n");

  const out: string[] = [];
  let inFence = false;
  let fenceToken: "```" | "~~~" | null = null;
  let inDisplayMath = false;

  const latexEnvBegin =
    /\\begin\{(bmatrix|pmatrix|matrix|aligned|align\*?|cases|equation\*?|array)\}/;

  let envBlock:
    | { name: string; lines: string[]; alreadyWrapped: boolean }
    | null = null;

  const lastNonEmptyLine = () => {
    for (let i = out.length - 1; i >= 0; i -= 1) {
      const trimmed = out[i].trim();
      if (trimmed.length) return trimmed;
    }
    return null;
  };

  for (const originalLine of lines) {
    let line = originalLine;

    const fenceMatch = line.match(/^\s*(```|~~~)/);
    if (fenceMatch) {
      const token = fenceMatch[1] as "```" | "~~~";
      if (!inFence) {
        inFence = true;
        fenceToken = token;
      } else if (fenceToken === token) {
        inFence = false;
        fenceToken = null;
      }
      out.push(line);
      continue;
    }

    if (inFence) {
      out.push(line);
      continue;
    }

    // Convert common Unicode bullets (often produced by LLMs) to Markdown lists.
    line = line.replace(/^\s*[•‣∙◦]\s+/, "- ");

    // Gemini/LLM sometimes indents lists/equations => Markdown interprets them as code blocks.
    if (/^( {4,}|\t+)/.test(line)) {
      const withoutIndent = line.replace(/^(?: {4}|\t)+/, "");
      const trimmed = withoutIndent.trimStart();
      const looksAccidental =
        /^(\d+\.)\s/.test(trimmed) ||
        /^[-*+]\s/.test(trimmed) ||
        /^>\s/.test(trimmed) ||
        /^\\[a-zA-Z]+/.test(trimmed) ||
        /^\$\$/.test(trimmed) ||
        /^\\\[/.test(trimmed) ||
        /^\\\(/.test(trimmed);
      if (looksAccidental) line = withoutIndent;
    }

    line = fixLatexInlineTokens(line);

    const dollarSplit = splitDoubleDollarMath(line);
    const preExpanded = dollarSplit ?? [line];
    const expanded = preExpanded.flatMap((segment) =>
      splitDisplayDelimiters(segment),
    );

    for (const chunk of expanded) {
      const segment = chunk
        .replace(/\\\(/g, "$")
        .replace(/\\\)/g, "$")
        .trimEnd();

      if (segment.trim() === "$$") {
        inDisplayMath = !inDisplayMath;
        out.push(segment);
        continue;
      }

      if (envBlock) {
        envBlock.lines.push(segment);
        const envEnd = new RegExp(
          `\\\\end\\{${escapeRegExp(envBlock.name)}\\}`,
        );
        if (envEnd.test(segment)) {
          const blockText = envBlock.lines.join("\n").trim();
          if (envBlock.alreadyWrapped) {
            out.push(...envBlock.lines);
          } else {
            out.push("$$");
            out.push(blockText);
            out.push("$$");
          }
          envBlock = null;
        }
        continue;
      }

      if (inDisplayMath) {
        out.push(segment);
        continue;
      }

      const beginMatch = segment.match(latexEnvBegin);
      if (beginMatch) {
        const name = beginMatch[1];
        const segmentTrimmed = segment.trim();
        const alreadyWrapped =
          lastNonEmptyLine() === "$$" ||
          segment.includes("$$") ||
          (segmentTrimmed.startsWith("$") && segmentTrimmed.endsWith("$"));
        envBlock = { name, lines: [segment], alreadyWrapped };

        const envEnd = new RegExp(`\\\\end\\{${escapeRegExp(name)}\\}`);
        if (envEnd.test(segment)) {
          const blockText = segment.trim();
          if (alreadyWrapped) out.push(segment);
          else {
            out.push("$$");
            out.push(blockText);
            out.push("$$");
          }
          envBlock = null;
        }
        continue;
      }

      // Wrap bare LaTeX equations in list items: `1. \sin(...) = ...` -> `1. $...$`
      const listEquationMatch = segment.match(
        /^(\s*(?:\d+\.|[-*+])\s+)(.+?)\s*$/,
      );
      if (listEquationMatch) {
        const prefix = listEquationMatch[1];
        const body = listEquationMatch[2];
        if (looksLikeLatex(body) && !body.includes("$") && body.includes("=")) {
          out.push(`${prefix}$${body}$`);
          continue;
        }
      }

      // Wrap bare equations inside sentences: `... \omega_i(p+k) = ...:` -> `... $...$:`
      if (looksLikeLatex(segment) && !segment.includes("$") && segment.includes("=")) {
        const firstLatex = segment.search(/\\[a-zA-Z]+/);
        if (firstLatex >= 0) {
          const head = segment.slice(0, firstLatex);
          const tail = segment.slice(firstLatex);
          const endMatch = tail.match(/^(.*?)([:;,.!?])(\s*)$/);
          if (endMatch) {
            const math = endMatch[1];
            const punct = endMatch[2];
            const space = endMatch[3];
            out.push(`${head}$${math}$${punct}${space}`);
            continue;
          }
          out.push(`${head}$${tail}$`);
          continue;
        }
      }

      out.push(segment);
    }
  }

  if (envBlock) out.push(...envBlock.lines);

  const joined = out.join("\n");
  const promoted = promoteInlineMathBlocks(joined);
  const normalizedMath = normalizeMathNewlines(
    wrapAlignedMathBlocks(
      fixEscapedSubscriptsInMath(
        normalizeDisplayMathBlocks(
          liftMultilineInlineMath(
            liftInlineMathEnvironments(normalizeLatexEnvironments(promoted)),
          ),
        ),
      ),
    ),
  );
  return stripZeroWidth(normalizedMath).replace(/\n{3,}/g, "\n\n").trimEnd();
};

const normalizeLatexEnvironments = (input: string) => {
  const envRegex =
    /\\begin\{(bmatrix|pmatrix|matrix|aligned|align\*?|cases|equation\*?|array)\}([\s\S]*?)\\end\{\1\}/g;

  return input.replace(envRegex, (_match, env, body) => {
    const normalizedBody = String(body)
      .replace(/\r\n?/g, "\n")
      .replace(/\\\\\s*\n/g, "\\\\ ")
      .replace(/\n+/g, " \\\\ ")
      .replace(/(?<!\\)\\\s*(?=[0-9+\-\s\]])/g, "\\\\ ")
      .trim();
    return `\\begin{${env}}\n${normalizedBody}\n\\end{${env}}`;
  });
};

const liftInlineMathEnvironments = (input: string) => {
  const inlineEnvRegex =
    /(?<!\$)\$([\s\S]*?\\begin\{(bmatrix|pmatrix|matrix|aligned|align\*?|cases|equation\*?|array)\}[\s\S]*?\\end\{\2\}[\s\S]*?)\$(?!\$)/g;

  return input.replace(inlineEnvRegex, (_match, math) => {
    const trimmed = String(math).trim();
    return `\n\n$$\n${trimmed}\n$$\n\n`;
  });
};

const promoteInlineMathBlocks = (input: string) =>
  input.replace(/(?<!\$)\$([\s\S]*?)\$(?!\$)/g, (match, body) => {
    const content = String(body);
    const hasEnv =
      /\\begin\{[^}]+\}/.test(content) || /\\end\{[^}]+\}/.test(content);
    const hasRows = /\\\\/.test(content) || /\\\s*(?=[0-9+\-\s\]])/.test(content);
    if (!hasEnv && !hasRows) return match;
    const cleaned = content.replace(/\n{2,}/g, "\n").trim();
    if (!cleaned) return match;
    return `\n\n$$\n${cleaned}\n$$\n\n`;
  });

const normalizeDisplayMathBlocks = (input: string) =>
  input.replace(/(\$\$)([\s\S]*?)(\$\$)/g, (_m, open, body, close) => {
    const cleaned = String(body)
      .replace(/\r\n?/g, "\n")
      .split("\n")
      .filter((line) => line.trim() !== "")
      .join("\n")
      .trim();
    return `${open}\n${cleaned}\n${close}`;
  });

const normalizeMathNewlines = (input: string) => {
  const envRegex =
    /\\begin\{(bmatrix|pmatrix|matrix|array)\}[\s\S]*?\\end\{\1\}/g;
  const alignedRegex = /\\begin\{(aligned|align\*?|cases)\}/;
  const protectEnvs = (body: string) => {
    const placeholders: Array<{ token: string; text: string }> = [];
    let index = 0;
    const protectedBody = body.replace(envRegex, (match) => {
      const token = `__ENV_${index}__`;
      placeholders.push({ token, text: match });
      index += 1;
      return token;
    });
    return { protectedBody, placeholders };
  };

  const restoreEnvs = (body: string, placeholders: Array<{ token: string; text: string }>) =>
    placeholders.reduce(
      (acc, { token, text }) => acc.replace(token, text),
      body,
    );

  const flatten = (body: string) => {
    const { protectedBody, placeholders } = protectEnvs(body);
    const flattened = protectedBody
      .replace(/\r\n?/g, "\n")
      .split("\n")
      .map((line) => line.trim())
      .filter(Boolean)
      .join(" ");
    return restoreEnvs(flattened, placeholders);
  };

  let out = input.replace(/(\$\$)([\s\S]*?)(\$\$)/g, (_m, open, body, close) => {
    if (alignedRegex.test(String(body))) return `${open}${body}${close}`;
    return `${open}${flatten(String(body))}${close}`;
  });

  out = out.replace(/(?<!\$)\$([\s\S]*?)\$(?!\$)/g, (_m, body) => {
    if (alignedRegex.test(String(body))) return `$${body}$`;
    return `$${flatten(String(body))}$`;
  });

  return out;
};

const wrapAlignedMathBlocks = (input: string) =>
  input.replace(/(\$\$)([\s\S]*?)(\$\$)/g, (_m, open, body, close) => {
    const content = String(body);
    if (/\\begin\{(aligned|align\*?)\}/.test(content)) {
      return `${open}${content}${close}`;
    }

    const matrixRegex =
      /\\begin\{(bmatrix|pmatrix|matrix|array)\}[\s\S]*?\\end\{\1\}/g;
    const placeholders: Array<{ token: string; text: string }> = [];
    let index = 0;
    const protectedBody = content.replace(matrixRegex, (match) => {
      const token = `__ENV_${index}__`;
      placeholders.push({ token, text: match });
      index += 1;
      return token;
    });

    const lines = protectedBody
      .replace(/\r\n?/g, "\n")
      .split("\n")
      .map((line) => line.trim())
      .filter(Boolean);

    const merged: string[] = [];
    for (let i = 0; i < lines.length; i += 1) {
      const line = lines[i];
      if (line === "=" && i + 1 < lines.length) {
        merged.push(`= ${lines[i + 1]}`);
        i += 1;
        continue;
      }
      merged.push(line);
    }

    const needsAligned =
      merged.length > 1 &&
      merged.some((line) => line.includes("=") || line.trim().startsWith("="));
    if (!needsAligned) return `${open}${content}${close}`;

    const alignedLines = merged.map((line) => {
      let next = line;
      if (next.trim().startsWith("=")) {
        next = next.replace(/^\s*=\s*/, "&= ");
      } else if (next.includes("=") && !next.includes("&")) {
        next = next.replace(/=/, "&=");
      } else if (!next.includes("&")) {
        next = `& ${next}`;
      }
      return next;
    });

    const restored = alignedLines
      .join(" \\\\ ")
      .replace(/__ENV_(\d+)__/g, (_m, idx) => placeholders[Number(idx)]?.text ?? "");

    return `${open}\n\\begin{aligned}\n${restored}\n\\end{aligned}\n${close}`;
  });

const liftUnbalancedInlineMath = (input: string) => {
  const normalized = input.replace(/\r\n?/g, "\n");
  const lines = normalized.split("\n");
  const out: string[] = [];
  let inFence = false;
  let fenceToken: "```" | "~~~" | null = null;
  let pending: string[] | null = null;

  for (const line of lines) {
    const fenceMatch = line.match(/^\s*(```|~~~)/);
    if (fenceMatch) {
      if (pending) {
        out.push(`$${pending.join("\n")}`);
        pending = null;
      }
      const token = fenceMatch[1] as "```" | "~~~";
      if (!inFence) {
        inFence = true;
        fenceToken = token;
      } else if (fenceToken === token) {
        inFence = false;
        fenceToken = null;
      }
      out.push(line);
      continue;
    }

    if (inFence) {
      out.push(line);
      continue;
    }

    if (!pending) {
      if (line.includes("$$")) {
        out.push(line);
        continue;
      }
      const start = line.indexOf("$");
      if (start >= 0) {
        const end = line.indexOf("$", start + 1);
        if (end === -1) {
          const before = line.slice(0, start);
          const after = line.slice(start + 1);
          if (before.trim()) out.push(before.trimEnd());
          pending = [after];
          continue;
        }
      }
      out.push(line);
      continue;
    }

    const end = line.indexOf("$");
    if (end >= 0) {
      pending.push(line.slice(0, end));
      const content = pending.join("\n").trim();
      if (content) {
        out.push("$$");
        out.push(content);
        out.push("$$");
      }
      const rest = line.slice(end + 1);
      if (rest.trim()) out.push(rest.trimStart());
      pending = null;
      continue;
    }

    pending.push(line);
  }

  if (pending) out.push(`$${pending.join("\n")}`);
  return out.join("\n");
};

const liftMultilineInlineMath = (input: string) =>
  input.replace(/(?<!\$)\$([\s\S]*?)\$(?!\$)/g, (match, math) => {
    const body = String(math);
    const hasEnv = /\\begin\{/.test(body);
    if (!/\n/.test(body) && !hasEnv) return match;
    const trimmed = body.trim();
    if (!trimmed) return match;
    return `\n\n$$\n${trimmed}\n$$\n\n`;
  });

const fixEscapedSubscriptsInMath = (input: string) => {
  const zeroWidthRegex =
    /[\u200B\u200C\u200D\u2060\uFEFF\u00AD\u202A-\u202E\u2066-\u2069]+/g;

  const recoverSubscriptsFromZeroWidth = (body: string) => {
    const wrapBase = (value: string) =>
      /^[A-Za-z]{2,}$/.test(value) ? `\\text{${value}}` : value;
    const wrapSub = (value: string) =>
      /^[A-Za-z]+$/.test(value) ? `\\text{${value}}` : value;

    const splitSub = (value: string) => {
      const index = value.search(/[A-Z]/);
      if (index > 0) {
        return { sub: value.slice(0, index), rest: value.slice(index) };
      }
      return { sub: value, rest: "" };
    };

    const subRegex = new RegExp(
      `([A-Za-z0-9]+)${zeroWidthRegex.source}([A-Za-z0-9]+)`,
      "g",
    );

    let next = body;
    let prev = "";
    while (next !== prev) {
      prev = next;
      next = next.replace(subRegex, (_m, base, sub) => {
        const { sub: subPart, rest } = splitSub(String(sub));
        const baseWrapped = wrapBase(String(base));
        const subWrapped = wrapSub(String(subPart));
        return `${baseWrapped}_{${subWrapped}}${rest}`;
      });
    }

    return next.replace(zeroWidthRegex, "");
  };

  const repairTextSubscripts = (value: string) =>
    value
      .replace(
        /\\text\{([^}]+)\}\s*\{\s*\\text\{([^}]+)\}\s*\}/g,
        "\\text{$1}_{\\text{$2}}",
      )
      .replace(
        /([A-Za-z0-9]+)\s*\{\s*\\text\{([^}]+)\}\s*\}/g,
        "$1_{\\text{$2}}",
      )
      .replace(
        /\\text\{([^}]+)\}\s*\\_?\s*\\text\{([^}]+)\}/g,
        "\\text{$1}_{\\text{$2}}",
      );

  const fixBody = (body: string) => {
    const unescaped = body
      .replace(/\\_(\s*\{)/g, "_$1")
      .replace(/\\_([A-Za-z0-9])/g, "_$1");
    return repairTextSubscripts(recoverSubscriptsFromZeroWidth(unescaped));
  };

  let out = input.replace(/(\$\$)([\s\S]*?)(\$\$)/g, (_m, open, body, close) => {
    return `${open}${fixBody(body)}${close}`;
  });

  out = out.replace(/(?<!\$)\$([^$\n]+)\$(?!\$)/g, (_m, body) => {
    return `$${fixBody(body)}$`;
  });

  return out;
};
// ChatGPT: Tampermonkey already copies clean TeX; we only do small
// fixes and filter linear fallbacks (without touching normal text).
const normalizeChatgptMarkdown = (input: string) => {
  const normalized = promoteInlineMathBlocks(liftUnbalancedInlineMath(input));
  const lines = normalized.split("\n");

  const fixedLines = lines.map((raw) => {
    return fixLatexInlineTokens(raw);
  });

  // Filter linearized fallbacks (no spaces, alphanumeric-only, or just digits/brackets).
  const cleaned: string[] = [];
  let inFence = false;
  let fenceToken: "```" | "~~~" | null = null;
  let inMathBlock = false;
  for (const line of fixedLines) {
    const fenceMatch = line.match(/^\s*(```|~~~)/);
    if (fenceMatch) {
      const token = fenceMatch[1] as "```" | "~~~";
      if (!inFence) {
        inFence = true;
        fenceToken = token;
      } else if (fenceToken === token) {
        inFence = false;
        fenceToken = null;
      }
      cleaned.push(line);
      continue;
    }
    if (inFence) {
      cleaned.push(line);
      continue;
    }

    const dollarCount = line.match(/\$\$/g)?.length ?? 0;
    if (dollarCount > 0) {
      cleaned.push(line);
      if (dollarCount % 2 === 1) inMathBlock = !inMathBlock;
      continue;
    }

    if (inMathBlock) {
      cleaned.push(line);
      continue;
    }

    const t = stripZeroWidth(line).trim();
    const looksMathy = /(\$|\\begin\{|\\end\{|\\\\|&|\\[a-zA-Z]+)/.test(line);
    if (looksMathy) {
      cleaned.push(line);
      continue;
    }
    const denseDigits =
      /^([\[\]0-9\s]+)\$*$/.test(t) && t.replace(/[\[\]\s]/g, "").length >= 4;
    const denseAlphaNum =
      /^[A-Za-z0-9\\=+*^_\[\]\.]+$/.test(t) &&
      t.length >= 8 &&
      /[A-Za-z]/.test(t) &&
      /[0-9]/.test(t);
    const noSpaces = !/\s/.test(t) && t.length >= 8;
    const onlyNumsPunct =
      t.length >= 8 && t.replace(/[0-9\[\],.=]/g, "").length === 0;
    if (denseDigits || denseAlphaNum || noSpaces || onlyNumsPunct) continue;
    cleaned.push(line);
  }

  // Auto-add table separator if we detect a block of pipe-lines without '---'
  const finalLines = cleaned;
  const output: string[] = [];
  for (let i = 0; i < finalLines.length; i += 1) {
    const line = finalLines[i];
    const isPipeLine =
      line.includes("|") &&
      (line.match(/\|/g)?.length ?? 0) >= 2 &&
      !/^\s*```/.test(line) &&
      !line.includes("---");

    if (isPipeLine) {
      const block: string[] = [line];
      let j = i + 1;
      while (
        j < finalLines.length &&
        finalLines[j].includes("|") &&
        (finalLines[j].match(/\|/g)?.length ?? 0) >= 2 &&
        !finalLines[j].includes("---") &&
        !/^\s*```/.test(finalLines[j])
      ) {
        block.push(finalLines[j]);
        j += 1;
      }

      if (block.length >= 2) {
        const cols = block[0]
          .split("|")
          .filter((c) => c.trim().length > 0).length;
        const separator =
          cols > 0 ? "|" + Array(cols).fill(" --- ").join("|") + "|" : "| --- |";
        output.push(block[0]);
        output.push(separator);
        for (let k = 1; k < block.length; k += 1) output.push(block[k]);
        i = j - 1;
        continue;
      }
    }

    output.push(line);
  }

  const joined = output.join("\n");
  const promoted = promoteInlineMathBlocks(joined);
  const normalizedMath = normalizeMathNewlines(
    wrapAlignedMathBlocks(
      fixEscapedSubscriptsInMath(
        normalizeDisplayMathBlocks(
          liftMultilineInlineMath(
            liftInlineMathEnvironments(normalizeLatexEnvironments(promoted)),
          ),
        ),
      ),
    ),
  );
  return stripZeroWidth(normalizedMath).replace(/\n{3,}/g, "\n\n").trimEnd();
};

const markdownComponents: Components = {
  pre({ className, children }) {
    return (
      <pre
        className={clsx(
          "rounded-2xl border border-white/10 bg-slate-900/70 p-4 text-sm leading-7 text-indigo-50 shadow-inner shadow-black/40",
          className,
        )}
      >
        {children}
      </pre>
    );
  },
  code({ className, children }) {
    const content = String(children);
    const isProbablyBlock =
      Boolean(className && className.includes("language-")) ||
      content.includes("\n");

    return (
      <code
        className={clsx(
          isProbablyBlock
            ? "text-indigo-50"
            : "rounded-md bg-white/10 px-2 py-0.5 text-[0.95em] text-sky-100",
          className,
        )}
      >
        {children}
      </code>
    );
  },
};

const buttonStyles =
  "flex items-center gap-2 rounded-full border border-white/10 bg-white/10 px-4 py-2 text-sm font-medium text-slate-100 shadow-lg shadow-black/30 transition hover:-translate-y-0.5 hover:border-white/30 hover:bg-white/15 active:translate-y-0 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-sky-400";

export default function Home() {
  const previewRef = useRef<HTMLDivElement>(null);
  const pasteSequenceRef = useRef(0);
  const [value, setValue] = useState(starter);
  const [llmSource, setLlmSource] = useState<"gemini" | "chatgpt">("chatgpt");
  const [autoFix, setAutoFix] = useState(true);
  const [copyMode, setCopyMode] = useState(false);
  const [status, setStatus] = useState<string | null>(null);
  const [busy, setBusy] = useState<"docx" | "pdf" | null>(null);

  useEffect(() => {
    if (!status) return;
    const timer = setTimeout(() => setStatus(null), 2800);
    return () => clearTimeout(timer);
  }, [status]);

  const handlePaste = async () => {
    try {
      let plain = "";
      let html = "";
      let rtf = "";

      const clipboard = navigator.clipboard as unknown as
        | {
        read?: () => Promise<
          Array<{ types: string[]; getType: (t: string) => Promise<Blob> }>
        >;
        readText?: () => Promise<string>;
      }
        | undefined;

      const rtfTypes = [
        "text/rtf",
        "application/rtf",
        "text/x-rtf",
        "text/richtext",
      ];

      const pickBest = (current: string, candidate: string) =>
        candidate.length > current.length ? candidate : current;

      const tryRead = async (item: {
        types: string[];
        getType: (t: string) => Promise<Blob>;
      }) => {
        if (item.types.includes("text/plain")) {
          try {
            plain = pickBest(
              plain,
              await item.getType("text/plain").then((b) => b.text()),
            );
          } catch {
            // ignore per-type failures
          }
        }

        if (item.types.includes("text/html")) {
          try {
            html = pickBest(html, await item.getType("text/html").then((b) => b.text()));
          } catch {
            // ignore per-type failures
          }
        }

        for (const t of rtfTypes) {
          if (!item.types.includes(t)) continue;
          try {
            rtf = pickBest(rtf, await item.getType(t).then((b) => b.text()));
          } catch {
            // ignore per-type failures
          }
        }
      };

      if (clipboard && typeof clipboard.read === "function") {
        const items = await clipboard.read();
        for (const item of items) {
          await tryRead(item);
        }
      }

      if (!plain && clipboard && typeof clipboard.readText === "function") {
        plain = await clipboard.readText();
      }
      if (plain) {
        const smart = buildSmartPasteText(plain, html, rtf);
        setValue(smart);
        setStatus(
          smart !== plain
            ? "Pasted (math recovered)"
            : "Text pasted from clipboard",
        );
      }
    } catch {
      setStatus("Allow clipboard access to paste");
    }
  };

  const handleTextareaPaste = (event: ClipboardEvent<HTMLTextAreaElement>) => {
    const target = event.currentTarget;
    const start = target.selectionStart ?? 0;
    const end = target.selectionEnd ?? start;

    const plain = event.clipboardData.getData("text/plain");
    const html = event.clipboardData.getData("text/html");
    const rtf = event.clipboardData.getData("text/rtf");

    const types = Array.from(event.clipboardData.types ?? []);
    console.groupCollapsed(
      `[paste] types=${types.join(", ") || "(none)"} plain=${plain.length} html=${html.length} rtf=${rtf.length}`,
    );
    console.log("clipboardData.types", types);
    console.log("text/plain (preview)", JSON.stringify(plain.slice(0, 400)));
    console.log("text/html (preview)", JSON.stringify(html.slice(0, 400)));
    console.log("text/rtf  (preview)", JSON.stringify(rtf.slice(0, 400)));
    if (html) console.log("extracted math (html)", extractMathFromClipboardHtml(html));
    if (rtf) console.log("extracted math (rtf)", extractMathFromClipboardRtf(rtf));
    console.groupEnd();

    const insertText = (text: string) => {
      setValue((prev) => prev.slice(0, start) + text + prev.slice(end));
      requestAnimationFrame(() => {
        const next = start + text.length;
        target.selectionStart = next;
        target.selectionEnd = next;
      });
    };

    const smart = buildSmartPasteText(plain, html, rtf);

    if (smart !== plain) {
      // We recovered something synchronously (HTML/RTF → TeX or filled `$$ $$` placeholders).
      event.preventDefault();
      insertText(smart);
      setStatus("Pasted (math recovered)");
      return;
    }

    const normalizedPlain = plain.replace(/\r\n?/g, "\n");
    const plainHasMath = hasMathMarkers(plain);
    const hasEmptyDisplayMath = /\$\$\s*\$\$/.test(normalizedPlain);
    const looksLikeMissingMath =
      plain.trim().length > 0 &&
      (!plainHasMath || hasEmptyDisplayMath) &&
      /(matric|matrix|equaz|equation|formula|sistema|moltiplicazione)/i.test(
        plain,
      );

    if (!looksLikeMissingMath) return;

    // Async fallback: paste the plain text immediately, then try to recover math
    // from `navigator.clipboard.read()` (when available) and replace the just-pasted
    // segment if we find something better.
    event.preventDefault();
    insertText(plain);
    setStatus("Text pasted (trying to recover math...)");

    const pasteId = (pasteSequenceRef.current += 1);
    const insertedStart = start;
    const insertedEnd = start + plain.length;
    const insertedText = plain;

    void (async () => {
      try {
        let richPlain = "";
        let richHtml = "";
        let richRtf = "";

        const clipboard = navigator.clipboard as unknown as
          | {
              read?: () => Promise<
                Array<{ types: string[]; getType: (t: string) => Promise<Blob> }>
              >;
              readText?: () => Promise<string>;
            }
          | undefined;

        if (!clipboard || typeof clipboard.read !== "function") return;

        const rtfTypes = [
          "text/rtf",
          "application/rtf",
          "text/x-rtf",
          "text/richtext",
        ];
        const pickBest = (current: string, candidate: string) =>
          candidate.length > current.length ? candidate : current;

        const items = await clipboard.read();
        for (const item of items) {
          if (item.types.includes("text/plain")) {
            try {
              richPlain = pickBest(
                richPlain,
                await item.getType("text/plain").then((b) => b.text()),
              );
            } catch {
              // ignore
            }
          }

          if (item.types.includes("text/html")) {
            try {
              richHtml = pickBest(
                richHtml,
                await item.getType("text/html").then((b) => b.text()),
              );
            } catch {
              // ignore
            }
          }

          for (const t of rtfTypes) {
            if (!item.types.includes(t)) continue;
            try {
              richRtf = pickBest(
                richRtf,
                await item.getType(t).then((b) => b.text()),
              );
            } catch {
              // ignore
            }
          }
        }

        const base =
          richPlain && hasMathMarkers(richPlain) && !hasMathMarkers(insertedText)
            ? richPlain
            : insertedText;
        const upgraded = buildSmartPasteText(base, richHtml, richRtf);

        if (!upgraded || upgraded === insertedText) return;

        setValue((prev) => {
          if (pasteId !== pasteSequenceRef.current) return prev;
          if (prev.slice(insertedStart, insertedEnd) !== insertedText) return prev;
          return (
            prev.slice(0, insertedStart) + upgraded + prev.slice(insertedEnd)
          );
        });

        requestAnimationFrame(() => {
          const next = insertedStart + upgraded.length;
          target.selectionStart = next;
          target.selectionEnd = next;
        });
        setStatus("Pasted (math recovered)");
      } catch {
        // Best-effort: keep the plain text.
      }
    })();
  };

  const renderedMarkdown = useMemo(() => {
    const base = normalizeCopyCodeBlocks(value);
    if (!autoFix) return base;
    return llmSource === "gemini"
      ? normalizeLlmMarkdown(base)
      : normalizeChatgptMarkdown(base);
  }, [autoFix, llmSource, value]);

  const downloadBlob = (blob: Blob, filename: string) => {
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    link.click();
    URL.revokeObjectURL(url);
  };

  const exportMarkdown = () => {
    const content = renderedMarkdown.trim();
    const blob = new Blob([content], { type: "text/markdown;charset=utf-8" });
    downloadBlob(blob, "promptpress-export.md");
    setStatus("Markdown exported");
  };

  const exportDocx = async () => {
    if (!previewRef.current) return;
    setBusy("docx");
    setStatus("Generating DOCX...");
    try {
      const response = await fetch("/api/docx", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          html: previewRef.current.innerHTML,
        }),
      });
      if (!response.ok) {
        throw new Error("Docx API error");
      }
      const blob = await response.blob();
      downloadBlob(blob, "promptpress-export.docx");
      setStatus("DOCX exported");
    } catch (error) {
      console.error(error);
      setStatus("Error exporting DOCX");
    } finally {
      setBusy(null);
    }
  };

  const exportPdf = async () => {
    if (!previewRef.current) return;
    setBusy("pdf");
    setStatus("Generating PDF...");
    try {
      const response = await fetch("/api/pdf", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          html: previewRef.current.innerHTML,
        }),
      });
      if (!response.ok) {
        throw new Error("PDF API error");
      }
      const blob = await response.blob();
      downloadBlob(blob, "promptpress-export.pdf");
      setStatus("PDF exported");
    } catch (error) {
      console.error(error);
      setStatus("Error exporting PDF");
    } finally {
      setBusy(null);
    }
  };

  return (
    <div className="relative overflow-hidden">
      <div className="glow left-[-18rem] top-[-10rem] rounded-full bg-indigo-500/30" />
      <div className="glow right-[-12rem] top-24 rounded-full bg-cyan-400/30" />

      <main className="relative mx-auto flex min-h-screen max-w-6xl flex-col gap-8 px-6 py-10 md:px-10">
        <div className="flex flex-col gap-6 rounded-3xl border border-white/10 bg-gradient-to-br from-slate-900/70 via-slate-900/50 to-slate-900/20 p-6 shadow-2xl shadow-indigo-500/20 ring-1 ring-white/10 backdrop-blur">
          <div className="flex flex-col gap-3 md:flex-row md:items-start md:justify-between">
            <div className="space-y-2">
              <p className="inline-flex items-center rounded-full border border-white/10 bg-white/5 px-3 py-1 text-xs font-semibold uppercase tracking-[0.2em] text-indigo-100">
                PromptPress
              </p>
              <h1 className="text-3xl font-semibold text-white sm:text-4xl md:text-5xl">
                From LLM to MD/PDF in one click
              </h1>
              <p className="max-w-3xl text-lg text-slate-300">
                Paste raw output from Gemini or any LLM, preview the rendered
                LaTeX, and download versions ready for decks or shareable docs.
              </p>
              <div className="flex flex-wrap gap-2 text-sm text-slate-300">
                <span className="rounded-full bg-indigo-500/20 px-3 py-1">
                  Markdown + GFM
                </span>
                <span className="rounded-full bg-cyan-500/15 px-3 py-1">
                  LaTeX / KaTeX
                </span>
                <span className="rounded-full bg-emerald-500/15 px-3 py-1">
                  Docx &amp; PDF export
                </span>
              </div>
            </div>
            <div className="grid grid-cols-2 gap-3 rounded-2xl border border-white/10 bg-white/5 p-4 text-center text-sm font-medium text-slate-200 shadow-lg">
              <div className="flex flex-col gap-1 rounded-xl bg-black/30 px-3 py-2">
                <span className="text-xs uppercase tracking-[0.2em] text-slate-400">
                  Math
                </span>
                <span className="text-lg font-semibold text-cyan-200">
                  KaTeX ready
                </span>
              </div>
              <div className="flex flex-col gap-1 rounded-xl bg-black/30 px-3 py-2">
                <span className="text-xs uppercase tracking-[0.2em] text-slate-400">
                  Output
                </span>
                <span className="text-lg font-semibold text-indigo-100">
                  Clean &amp; export
                </span>
              </div>
            </div>
          </div>

          <div className="grid gap-6 lg:grid-cols-[1.05fr_1fr]">
            <section className="relative rounded-2xl border border-white/10 bg-white/5 p-4 shadow-xl shadow-black/30">
              <div className="flex items-center justify-between gap-3">
                <div className="flex items-center gap-2">
                  <span className="h-2 w-2 rounded-full bg-emerald-400" />
                  <h2 className="text-base font-semibold text-white">
                    LLM input
                  </h2>
                </div>
                <div className="flex gap-2">
                  <button
                    type="button"
                    onClick={() => {
                      setValue(starter);
                      setStatus("Sample reset");
                    }}
                    className={clsx(buttonStyles, "bg-slate-900/40")}
                  >
                    Load sample
                  </button>
                  <button
                    type="button"
                    aria-pressed={llmSource === "gemini"}
                    onClick={() => {
                      setLlmSource("gemini");
                      setCopyMode(false);
                      setStatus("Mode: Gemini");
                    }}
                    className={clsx(
                      buttonStyles,
                      "bg-slate-900/40",
                      llmSource === "gemini" &&
                        "border-emerald-400/60 bg-emerald-500/15 text-emerald-100 ring-2 ring-emerald-400/50",
                    )}
                  >
                    Gemini input
                  </button>
                  <button
                    type="button"
                    aria-pressed={llmSource === "chatgpt"}
                    onClick={() => {
                      setLlmSource("chatgpt");
                      setCopyMode(false);
                      setStatus("Mode: ChatGPT");
                    }}
                    className={clsx(
                      buttonStyles,
                      "bg-slate-900/40",
                      llmSource === "chatgpt" &&
                        "border-cyan-400/60 bg-cyan-500/15 text-cyan-100 ring-2 ring-cyan-400/50",
                    )}
                  >
                    ChatGPT input
                  </button>
                  <button
                    type="button"
                    aria-pressed={copyMode}
                    onClick={() => {
                      setAutoFix(true);
                      setLlmSource("chatgpt");
                      setCopyMode(true);
                      setStatus("Mode: Copy (Ctrl+C)");
                    }}
                    className={clsx(
                      buttonStyles,
                      "bg-indigo-500/25",
                      copyMode &&
                        "border-indigo-300/70 bg-indigo-500/35 text-white ring-2 ring-indigo-300/60",
                    )}
                  >
                    Copy mode (Ctrl+C)
                  </button>
                  <button
                    type="button"
                    onClick={handlePaste}
                    className={clsx(buttonStyles, "bg-indigo-500/30")}
                  >
                    Paste
                  </button>
                </div>
              </div>

              <div className="mt-3 rounded-xl border border-white/10 bg-slate-950/60 p-2 shadow-inner shadow-black/40">
                <textarea
                  value={value}
                  onChange={(event) => setValue(event.target.value)}
                  onPaste={handleTextareaPaste}
                  className="h-[380px] w-full resize-none rounded-lg border border-white/5 bg-transparent px-4 py-3 text-sm leading-6 text-slate-200 outline-none focus:border-indigo-400/60 focus:ring-2 focus:ring-indigo-500/30"
                  placeholder="Paste your model output here (Markdown and LaTeX supported)..."
                />
              </div>
              <p className="mt-2 text-xs text-slate-400">
                Supports Markdown (GFM), inline LaTeX $E=mc^2$ and blocks
                $$f(x)$$. No server dependency: everything happens in the
                browser.
              </p>
              <div className="mt-3 flex flex-col gap-3 rounded-xl border border-white/10 bg-black/20 px-4 py-3 text-xs text-slate-300 md:flex-row md:items-center md:justify-between">
                <div className="space-y-1">
                  <p className="font-semibold text-slate-200">
                    Auto-fix for LLM output
                  </p>
                  <p className="text-slate-400">
                    Cleans indentation (which becomes code) and converts `\( ... \)` / `\[ ... \]` into `$...$` / `$$...$$`.
                  </p>
                </div>
                <div className="flex flex-wrap items-center gap-2">
                  <button
                    type="button"
                    onClick={() => setAutoFix((prev) => !prev)}
                    className={clsx(
                      buttonStyles,
                      "bg-slate-900/40",
                      autoFix && "border-emerald-400/30 bg-emerald-500/10",
                    )}
                  >
                    {autoFix ? "On" : "Off"}
                  </button>
                </div>
              </div>
            </section>

            <section className="relative rounded-2xl border border-white/10 bg-slate-900/60 p-4 shadow-xl shadow-indigo-500/10">
              <div className="flex items-center justify-between gap-3">
                <div className="flex items-center gap-2">
                  <span className="h-2 w-2 rounded-full bg-indigo-400" />
                  <h2 className="text-base font-semibold text-white">
                    Preview ready
                  </h2>
                </div>
                <span className="rounded-full bg-indigo-500/20 px-3 py-1 text-xs font-semibold text-indigo-100">
                  Live render
                </span>
              </div>
              <div
                ref={previewRef}
                className="markdown prose prose-invert mt-4 max-h-[440px] overflow-y-auto rounded-xl border border-white/5 bg-gradient-to-b from-slate-900/50 to-black/60 p-4 shadow-inner shadow-black/40"
                data-preview-root="1"
              >
                <ReactMarkdown
                  remarkPlugins={[remarkGfm, remarkMath]}
                  rehypePlugins={[
                    [rehypeKatex, { throwOnError: false, strict: "ignore" }],
                  ]}
                  components={markdownComponents}
                >
                  {renderedMarkdown}
                </ReactMarkdown>
              </div>
            </section>
          </div>

          <div className="flex flex-col gap-3 rounded-2xl border border-white/10 bg-white/5 p-4 shadow-lg shadow-black/40 md:flex-row md:items-center md:justify-between">
            <div className="flex flex-wrap items-center gap-3">
              <button
                type="button"
                onClick={exportMarkdown}
                className={clsx(buttonStyles, "bg-emerald-500/20 px-5")}
              >
                Export Markdown
              </button>
              <button
                type="button"
                onClick={exportDocx}
                disabled={busy === "docx"}
                className={clsx(
                  buttonStyles,
                  "bg-indigo-500/30 px-5",
                  busy === "docx" && "cursor-not-allowed opacity-60",
                )}
              >
                {busy === "docx" ? "Exporting..." : "Export DOCX"}
              </button>
              <button
                type="button"
                onClick={exportPdf}
                disabled={busy === "pdf"}
                className={clsx(
                  buttonStyles,
                  "bg-cyan-500/30 px-5",
                  busy === "pdf" && "cursor-not-allowed opacity-60",
                )}
              >
                {busy === "pdf" ? "Exporting..." : "Export PDF"}
              </button>
            </div>
            <div className="flex items-center gap-2 text-sm text-slate-300">
              <span className="h-2 w-2 rounded-full bg-emerald-400" />
              {status ?? "Ready to export"}
            </div>
          </div>
        </div>
      </main>
    </div>
  );
}

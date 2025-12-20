import { NextResponse } from "next/server";
import {
  AlignmentType,
  BorderStyle,
  Document,
  HeadingLevel,
  LevelFormat,
  Packer,
  Paragraph,
  type ParagraphChild,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from "docx";
import { parse, type HTMLElement, type Node as HtmlNode } from "node-html-parser";
import katex from "katex";
import JSZip from "jszip";
import { mml2omml } from "mathml2omml";

export const runtime = "nodejs";

type DocxChild = Paragraph | Table;
type InlineChild = ParagraphChild;
type MathReplacement = { token: string; omml: string };
type MathContext = { nextId: number; replacements: MathReplacement[] };

type BlockContext = {
  list?: { ordered: boolean; level: number };
  cellHeader?: boolean;
};

const stripZeroWidth = (value: string) =>
  value.replace(
    /[\u200B\u200C\u200D\u2060\uFEFF\u00AD\u202A-\u202E\u2066-\u2069]/g,
    "",
  );

const stripUnpairedSurrogates = (value: string) => {
  let out = "";
  for (let i = 0; i < value.length; i += 1) {
    const code = value.charCodeAt(i);
    if (code >= 0xd800 && code <= 0xdbff) {
      const next = value.charCodeAt(i + 1);
      if (next >= 0xdc00 && next <= 0xdfff) {
        out += value[i] + value[i + 1];
        i += 1;
      }
      continue;
    }
    if (code >= 0xdc00 && code <= 0xdfff) continue;
    out += value[i];
  }
  return out;
};

const sanitizeXmlText = (value: string) =>
  stripUnpairedSurrogates(stripZeroWidth(value)).replace(
    /[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F-\u0084\u0086-\u009F]/g,
    "",
  );

const normalizeText = (value: string) =>
  sanitizeXmlText(value).replace(/\u00A0/g, " ").replace(/\s+/g, " ");

const normalizePreserveWhitespace = (value: string) =>
  sanitizeXmlText(value).replace(/\u00A0/g, " ");

const getTag = (el: HTMLElement) => el.tagName.toLowerCase();

const extractMathMlFromKatex = (el: HTMLElement) => {
  const math = el.querySelector("math");
  return math ? math.toString() : null;
};

const extractLatexFromKatex = (el: HTMLElement) => {
  const annotation = el.querySelector(
    'annotation[encoding="application/x-tex"], annotation[encoding="application/x-latex"], annotation[encoding="application/tex"]',
  );
  const texFromAnnotation = annotation?.textContent?.trim();
  if (texFromAnnotation) return texFromAnnotation;

  const fromTitle = el.getAttribute("title")?.trim();
  if (fromTitle && fromTitle.includes("\\") && fromTitle.length < 5000) {
    if (!/^ParseError:/i.test(fromTitle)) return fromTitle;
  }

  const text = el.textContent?.trim();
  return text && text.includes("\\") ? text : null;
};

const texToMathMl = (tex: string, displayMode: boolean) => {
  try {
    const rendered = katex.renderToString(tex, {
      output: "mathml",
      displayMode,
      throwOnError: true,
      strict: "ignore",
    });
    const match = rendered.match(/<math[\s\S]*?<\/math>/i);
    return match ? match[0] : null;
  } catch {
    return null;
  }
};

const cleanMathMl = (mathML: string) =>
  mathML.replace(/<annotation[\s\S]*?<\/annotation>/gi, "");

const buildOmmlFromElement = (el: HTMLElement): string | null => {
  const displayMode =
    el.classList.contains("katex-display") ||
    el.querySelector("math")?.getAttribute("display") === "block";

  const mathML = extractMathMlFromKatex(el);
  const sourceMathMl = mathML
    ? mathML
    : (() => {
        const tex = extractLatexFromKatex(el);
        if (!tex) return null;
        return texToMathMl(tex, displayMode);
      })();

  if (!sourceMathMl) return null;

  try {
    return mml2omml(cleanMathMl(sourceMathMl));
  } catch {
    return null;
  }
};

const createMathPlaceholder = (ctx: MathContext, el: HTMLElement) => {
  const omml = buildOmmlFromElement(el);
  if (!omml) return null;

  const token = `__MATH_${String(ctx.nextId).padStart(6, "0")}__`;
  ctx.nextId += 1;
  ctx.replacements.push({ token, omml });
  return new TextRun({ text: token });
};

const textRunFromNode = (
  text: string,
  opts?: { bold?: boolean; italics?: boolean; code?: boolean },
) =>
  new TextRun({
    text: sanitizeXmlText(text),
    bold: opts?.bold,
    italics: opts?.italics,
    font: opts?.code ? "Consolas" : undefined,
  });

const inlineFromNodes = (
  nodes: HtmlNode[],
  marks: { bold?: boolean; italics?: boolean; code?: boolean } = {},
  math: MathContext,
): InlineChild[] => {
  const out: InlineChild[] = [];

  for (const node of nodes) {
    if (node.nodeType === 3) {
      const text = normalizeText(node.rawText ?? "");
      if (!text) continue;
      out.push(textRunFromNode(text, marks));
      continue;
    }

    if (node.nodeType !== 1) continue;
    const el = node as HTMLElement;

    if (el.classList.contains("katex") || el.classList.contains("katex-display")) {
      const placeholder = createMathPlaceholder(math, el);
      if (placeholder) out.push(placeholder);
      else {
        const fallback = extractLatexFromKatex(el) ?? el.textContent ?? "";
        if (fallback.trim()) out.push(textRunFromNode(fallback, { ...marks, code: true }));
      }
      continue;
    }

    const tag = getTag(el);
    if (tag === "br") {
      out.push(new TextRun({ break: 1 }));
      continue;
    }

    const nextMarks = { ...marks };
    if (tag === "strong" || tag === "b") nextMarks.bold = true;
    if (tag === "em" || tag === "i") nextMarks.italics = true;
    if (tag === "code") nextMarks.code = true;

    if (el.classList.contains("katex-html")) continue;

    const children = inlineFromNodes(el.childNodes as HtmlNode[], nextMarks, math);
    out.push(...children);
  }

  return out;
};

const paragraphFromInline = (
  children: InlineChild[],
  opts?: {
    heading?: (typeof HeadingLevel)[keyof typeof HeadingLevel];
    bulletLevel?: number;
    numberLevel?: number;
  },
) =>
  new Paragraph({
    children,
    heading: opts?.heading,
    bullet: opts?.bulletLevel !== undefined ? { level: opts.bulletLevel } : undefined,
    numbering:
      opts?.numberLevel !== undefined
        ? { reference: "ordered", level: opts.numberLevel }
        : undefined,
    spacing: { after: 160 },
  });

const codeBlockParagraph = (code: string) => {
  const normalized = normalizePreserveWhitespace(code)
    .replace(/\r\n?/g, "\n")
    .replace(/\n$/, "");
  const lines = normalized.split("\n");
  const runs: TextRun[] = [];
  for (let i = 0; i < lines.length; i += 1) {
    runs.push(textRunFromNode(lines[i], { code: true }));
    if (i !== lines.length - 1) runs.push(new TextRun({ break: 1 }));
  }

  return new Paragraph({
    children: runs,
    spacing: { before: 80, after: 160 },
    shading: { type: ShadingType.CLEAR, color: "auto", fill: "0B1224" },
    border: {
      top: { style: BorderStyle.SINGLE, size: 6, color: "1E293B" },
      bottom: { style: BorderStyle.SINGLE, size: 6, color: "1E293B" },
      left: { style: BorderStyle.SINGLE, size: 6, color: "1E293B" },
      right: { style: BorderStyle.SINGLE, size: 6, color: "1E293B" },
    },
    indent: { left: 360, right: 360 },
  });
};

const blockFromNode = (
  node: HtmlNode,
  ctx: BlockContext | undefined,
  math: MathContext,
): DocxChild[] => {
  if (node.nodeType === 3) {
    const text = normalizeText(node.rawText ?? "");
    if (!text) return [];
    return [paragraphFromInline([textRunFromNode(text, ctx?.cellHeader ? { bold: true } : undefined)])];
  }

  if (node.nodeType !== 1) return [];
  const el = node as HTMLElement;
  const tag = getTag(el);

  if (el.classList.contains("katex") || el.classList.contains("katex-display")) {
    const placeholder = createMathPlaceholder(math, el);
    if (placeholder) return [paragraphFromInline([placeholder])];
    const fallback = extractLatexFromKatex(el) ?? el.textContent ?? "";
    if (!fallback.trim()) return [];
    return [paragraphFromInline([textRunFromNode(fallback, { code: true })])];
  }

  if (tag === "br") return [paragraphFromInline([])];

  if (tag === "pre") {
    const code = el.textContent ?? "";
    return [codeBlockParagraph(code)];
  }

  if (tag === "table") {
    const rows = Array.from(el.querySelectorAll("tr"));
    if (!rows.length) return [];

    const headerRow = el.querySelector("thead tr") ?? rows[0];
    const headerCells = Array.from(headerRow.querySelectorAll("th, td"));
    const colCount = headerCells.length || 1;

    const bodyRows =
      headerRow === rows[0]
        ? rows.slice(1)
        : rows.filter((row) => !row.closest("thead") && row !== headerRow);

    const headerShading = {
      type: ShadingType.CLEAR,
      color: "auto",
      fill: "1E293B",
    };

    const rowsOut: TableRow[] = [];
    const allRows = [headerRow, ...bodyRows];

    for (let r = 0; r < allRows.length; r += 1) {
      const row = allRows[r];
      const cells = Array.from(row.querySelectorAll("th, td"));
      const isHeaderRow = row === headerRow;
      const rowCells: TableCell[] = [];

      for (let c = 0; c < colCount; c += 1) {
        const cell = cells[c];
        const cellBlocks: Paragraph[] = [];

        if (!cell) {
          cellBlocks.push(
            paragraphFromInline([textRunFromNode("", isHeaderRow ? { bold: true } : undefined)]),
          );
        } else {
          for (const child of cell.childNodes as HtmlNode[]) {
            const blocks = blockFromNode(child, { cellHeader: isHeaderRow }, math);
            for (const block of blocks) {
              if (block instanceof Paragraph) cellBlocks.push(block);
            }
          }
          if (!cellBlocks.length) {
            cellBlocks.push(
              paragraphFromInline([textRunFromNode("", isHeaderRow ? { bold: true } : undefined)]),
            );
          }
        }

        rowCells.push(
          new TableCell({
            children: cellBlocks,
            shading: isHeaderRow ? headerShading : undefined,
          }),
        );
      }

      rowsOut.push(new TableRow({ children: rowCells }));
    }

    return [
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: {
          top: { style: BorderStyle.SINGLE, size: 4, color: "CBD5F5" },
          bottom: { style: BorderStyle.SINGLE, size: 4, color: "CBD5F5" },
          left: { style: BorderStyle.SINGLE, size: 4, color: "CBD5F5" },
          right: { style: BorderStyle.SINGLE, size: 4, color: "CBD5F5" },
          insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: "CBD5F5" },
          insideVertical: { style: BorderStyle.SINGLE, size: 4, color: "CBD5F5" },
        },
        rows: rowsOut,
      }),
    ];
  }

  if (tag === "ul" || tag === "ol") {
    const ordered = tag === "ol";
    const items = Array.from(el.children).filter(
      (child) => getTag(child) === "li",
    );
    const out: DocxChild[] = [];

    for (let i = 0; i < items.length; i += 1) {
      out.push(
        ...blockFromNode(
          items[i],
          { list: { ordered, level: (ctx?.list?.level ?? 0) + 1 } },
          math,
        ),
      );
    }
    return out;
  }

  if (tag === "li") {
    const list = ctx?.list;
    const children = inlineFromNodes(el.childNodes as HtmlNode[], {}, math);
    const paragraph = paragraphFromInline(children, {
      bulletLevel: list && !list.ordered ? list.level - 1 : undefined,
      numberLevel: list && list.ordered ? list.level - 1 : undefined,
    });
    return [paragraph];
  }

  if (/^h[1-6]$/.test(tag)) {
    const level = Number.parseInt(tag.slice(1), 10);
    const heading =
      level === 1
        ? HeadingLevel.HEADING_1
        : level === 2
          ? HeadingLevel.HEADING_2
          : level === 3
            ? HeadingLevel.HEADING_3
            : level === 4
              ? HeadingLevel.HEADING_4
              : level === 5
                ? HeadingLevel.HEADING_5
                : HeadingLevel.HEADING_6;
    const children = inlineFromNodes(el.childNodes as HtmlNode[], {}, math);
    return [paragraphFromInline(children, { heading })];
  }

  if (tag === "blockquote") {
    const children = inlineFromNodes(el.childNodes as HtmlNode[], { italics: true }, math);
    return [paragraphFromInline(children)];
  }

  if (tag === "p" || tag === "div") {
    const children = inlineFromNodes(el.childNodes as HtmlNode[], {}, math);
    if (!children.length) return [];
    return [paragraphFromInline(children, ctx?.cellHeader ? { } : undefined)];
  }

  const out: DocxChild[] = [];
  for (const child of el.childNodes as HtmlNode[]) {
    out.push(...blockFromNode(child, ctx, math));
  }
  return out;
};

const patchDocxWithMath = async (
  fileBuffer: Buffer,
  replacements: MathReplacement[],
) => {
  if (!replacements.length) return fileBuffer;

  const zip = await JSZip.loadAsync(fileBuffer);
  const documentXml = zip.file("word/document.xml");
  if (!documentXml) return fileBuffer;

  let xml = await documentXml.async("string");

  if (!xml.includes("xmlns:m=")) {
    xml = xml.replace(
      /<w:document\b([^>]*)>/,
      '<w:document$1 xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">',
    );
  }

  const cleanOmml = (value: string) =>
    value
      .replace(/\s+xmlns:m="[^"]+"/g, "")
      .replace(/\s+xmlns:w="[^"]+"/g, "");

  const sanitizeOmml = (value: string) => {
    const escaped = sanitizeXmlText(value).replace(
      /&(?!amp;|lt;|gt;|quot;|apos;|#\d+;|#x[0-9a-fA-F]+;)/g,
      "&amp;",
    );

    return escaped.replace(/<m:t([^>]*)>([\s\S]*?)<\/m:t>/g, (_m, attrs, text) => {
      const safeText = String(text)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");
      return `<m:t${attrs}>${safeText}</m:t>`;
    });
  };

  const wrapOmmlRun = (value: string) => {
    const trimmed = value.trim();
    if (/<w:r[\s>]/.test(trimmed)) return trimmed;
    return (
      `<w:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/>` +
      `</w:rPr>${trimmed}</w:r>`
    );
  };

  for (const { token, omml } of replacements) {
    const tokenIndex = xml.indexOf(token);
    if (tokenIndex === -1) continue;

    const runStart = xml.lastIndexOf("<w:r", tokenIndex);
    if (runStart === -1) continue;
    const runEnd = xml.indexOf("</w:r>", tokenIndex);
    if (runEnd === -1) continue;

    const runClose = runEnd + "</w:r>".length;
    const runSlice = xml.slice(runStart, runClose);
    if (!runSlice.includes(token)) continue;

    const safeOmml = wrapOmmlRun(sanitizeOmml(cleanOmml(omml)));
    xml = xml.slice(0, runStart) + safeOmml + xml.slice(runClose);
  }

  zip.file("word/document.xml", xml);
  return zip.generateAsync({ type: "nodebuffer" });
};

export async function POST(request: Request) {
  try {
    const { html } = await request.json();

    if (!html || typeof html !== "string") {
      return NextResponse.json({ error: "Missing HTML" }, { status: 400 });
    }

    const root = parse(`<div>${html}</div>`).querySelector("div");
    if (!root) {
      return NextResponse.json({ error: "Invalid HTML" }, { status: 400 });
    }

    const math: MathContext = { nextId: 1, replacements: [] };
    const children: DocxChild[] = [];
    for (const node of root.childNodes as HtmlNode[]) {
      children.push(...blockFromNode(node, undefined, math));
    }

    const doc = new Document({
      numbering: {
        config: [
          {
            reference: "ordered",
            levels: [
              {
                level: 0,
                format: LevelFormat.DECIMAL,
                text: "%1.",
                alignment: AlignmentType.START,
              },
              {
                level: 1,
                format: LevelFormat.LOWER_LETTER,
                text: "%2.",
                alignment: AlignmentType.START,
              },
              {
                level: 2,
                format: LevelFormat.LOWER_ROMAN,
                text: "%3.",
                alignment: AlignmentType.START,
              },
            ],
          },
        ],
      },
      sections: [{ children }],
    });

    const fileBuffer = await Packer.toBuffer(doc);
    const patched = await patchDocxWithMath(fileBuffer, math.replacements);

    const mime =
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    const body =
      patched instanceof Blob ? patched : new Blob([patched as BlobPart], { type: mime });

    return new NextResponse(body, {
      status: 200,
      headers: {
        "Content-Type": mime,
        "Content-Disposition": 'attachment; filename="promptpress-export.docx"',
      },
    });
  } catch (error) {
    console.error("DOCX export error", error);
    return NextResponse.json(
      { error: "Errore durante la generazione del DOCX" },
      { status: 500 },
    );
  }
}

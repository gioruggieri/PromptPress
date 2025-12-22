// app/api/export/docx/route.ts
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
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import type { Document as XmlDoc, Element as XmlEl, Node as XmlNode } from "@xmldom/xmldom";

export const runtime = "nodejs";

type DocxChild = Paragraph | Table;
type InlineChild = ParagraphChild;

type MathReplacement = { token: string; omml: string };
type MathContext = { nextId: number; replacements: MathReplacement[] };

type BlockContext = {
  list?: { ordered: boolean; level: number };
  cellHeader?: boolean;
};

const M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math";

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

/**
 * Normalizza ambienti LaTeX in modo che la semantica dei delimitatori sia esplicita.
 * Utile perché alcuni converter MathML->OMML degradano la "stretchiness".
 */
const normalizeDocxLatex = (tex: string) =>
  tex
    .replace(
      /\\begin\{bmatrix\}([\s\S]*?)\\end\{bmatrix\}/g,
      String.raw`\left[\begin{matrix}$1\end{matrix}\right]`,
    )
    .replace(
      /\\begin\{pmatrix\}([\s\S]*?)\\end\{pmatrix\}/g,
      String.raw`\left(\begin{matrix}$1\end{matrix}\right)`,
    )
    .replace(
      /\\begin\{Bmatrix\}([\s\S]*?)\\end\{Bmatrix\}/g,
      String.raw`\left\{\begin{matrix}$1\end{matrix}\right\}`,
    )
    .replace(
      /\\begin\{vmatrix\}([\s\S]*?)\\end\{vmatrix\}/g,
      String.raw`\left|\begin{matrix}$1\end{matrix}\right|`,
    )
    .replace(
      /\\begin\{Vmatrix\}([\s\S]*?)\\end\{Vmatrix\}/g,
      String.raw`\left\|\begin{matrix}$1\end{matrix}\right\|`,
    )
    .replace(
      /\\begin\{smallmatrix\}([\s\S]*?)\\end\{smallmatrix\}/g,
      String.raw`\left(\begin{smallmatrix}$1\end{smallmatrix}\right)`,
    );

const tagLocal = (name: string) => (name.includes(":") ? name.split(":")[1] : name);
const isM = (el: XmlEl, local: string) => tagLocal(el.tagName) === local;

const parseOmmlFragment = (omml: string): { doc: XmlDoc; root: XmlEl } | null => {
  // rimuovi xml decl ed eventuali namespace duplicati dentro il frammento
  const cleaned = omml
    .replace(/^\s*<\?xml[^>]*\?>\s*/i, "")
    .replace(/\s+xmlns:m="[^"]+"/g, "")
    .replace(/\s+xmlns:w="[^"]+"/g, "")
    .trim();

  // Wrappiamo in un root per parsare frammenti
  const wrapped = `<root xmlns:m="${M_NS}">${cleaned}</root>`;

  const doc = new DOMParser({
    errorHandler: { warning: () => {}, error: () => {}, fatalError: () => {} },
  }).parseFromString(wrapped, "application/xml") as unknown as XmlDoc;

  const root = doc.documentElement as unknown as XmlEl;
  if (!root || root.tagName !== "root") return null;

  // Se xmldom ha inserito <parsererror> (in alcuni ambienti), fallback
  const anyParserError = (root.getElementsByTagName("parsererror")?.length ?? 0) > 0;
  if (anyParserError) return null;

  return { doc, root };
};

const serializeOmmlChildren = (root: XmlEl) => {
  const ser = new XMLSerializer();
  let out = "";
  for (let n = root.firstChild; n; n = n.nextSibling) {
    // serializza solo nodi element/text rilevanti
    out += ser.serializeToString(n as unknown as XmlNode);
  }
  return out;
};

const runTextIfSingleChar = (run: XmlEl) => {
  if (!isM(run, "r")) return null;
  let tEl: XmlEl | null = null;

  for (let n = run.firstChild; n; n = n.nextSibling) {
    if (n.nodeType === 1) {
      const el = n as unknown as XmlEl;
      if (isM(el, "t")) {
        tEl = el;
        break;
      }
      // a volte <m:rPr> precede <m:t>, lo ignoriamo
    }
  }
  if (!tEl) return null;

  const txt = (tEl.textContent ?? "").trim();
  if (!txt) return null;

  // accettiamo solo singolo “glyph” (es. "[" , "|" , "‖")
  // NB: alcuni simboli sono surrogate pair; per semplicità controlliamo lunghezza in code points
  const codePoints = Array.from(txt);
  if (codePoints.length !== 1) return null;

  return codePoints[0];
};

const isTall = (el: XmlEl) => {
  const local = tagLocal(el.tagName);
  // Copertura pratica dei “tall constructs” che beneficiano di delimitatori stretchy
  return new Set([
    "m",       // matrix
    "eqArr",   // array/cases/aligned-like
    "f",       // fraction
    "rad",     // root
    "nary",    // sum/int/prod
    "stack",   // stacked
    "limLow",
    "limUpp",
    "box",
    "bar",
    "groupChr",
    "sSub",
    "sSup",
    "sSubSup",
  ]).has(local);
};

const createDelimiter = (doc: XmlDoc, beg: string, end: string, inner: XmlEl) => {
  const d = doc.createElementNS(M_NS, "m:d") as unknown as XmlEl;
  const dPr = doc.createElementNS(M_NS, "m:dPr") as unknown as XmlEl;

  const begChr = doc.createElementNS(M_NS, "m:begChr") as unknown as XmlEl;
  begChr.setAttribute("m:val", beg);

  const endChr = doc.createElementNS(M_NS, "m:endChr") as unknown as XmlEl;
  endChr.setAttribute("m:val", end);

  const grow = doc.createElementNS(M_NS, "m:grow") as unknown as XmlEl;
  // m:val opzionale; lo mettiamo esplicito per compatibilità
  grow.setAttribute("m:val", "1");

  dPr.appendChild(begChr as unknown as XmlNode);
  dPr.appendChild(endChr as unknown as XmlNode);
  dPr.appendChild(grow as unknown as XmlNode);

  const e = doc.createElementNS(M_NS, "m:e") as unknown as XmlEl;
  // sposta "inner" dentro <m:e>
  e.appendChild(inner as unknown as XmlNode);

  d.appendChild(dPr as unknown as XmlNode);
  d.appendChild(e as unknown as XmlNode);

  return d;
};

const promoteDelimitersInContainer = (container: XmlEl) => {
  // Scansiona triple (run open) (tall) (run close) tra i children diretti del container
  const pairs: Array<[string, string]> = [
    ["(", ")"],
    ["[", "]"],
    ["{", "}"],
    ["⟨", "⟩"],
    ["|", "|"],
    ["‖", "‖"],
  ];

  const nextElementIndex = (nodes: XmlNode[], start: number) => {
    for (let i = start; i < nodes.length; i += 1) {
      if (nodes[i]?.nodeType === 1) return i;
    }
    return -1;
  };

  let changed = true;
  while (changed) {
    changed = false;
    const nodes = Array.from(container.childNodes) as unknown as XmlNode[];

    for (let i = 0; i < nodes.length; i += 1) {
      if (nodes[i]?.nodeType !== 1) continue;
      const a = nodes[i] as unknown as XmlEl;

      const openChar = runTextIfSingleChar(a);
      if (!openChar) continue;

      const j = nextElementIndex(nodes, i + 1);
      if (j === -1) continue;

      const k = nextElementIndex(nodes, j + 1);
      if (k === -1) continue;

      const b = nodes[j] as unknown as XmlEl;
      const c = nodes[k] as unknown as XmlEl;

      if (!isTall(b)) continue;

      const closeChar = runTextIfSingleChar(c);
      if (!closeChar) continue;

      // verifica se (openChar, closeChar) è una coppia ammessa
      const ok = pairs.some(([beg, end]) => beg === openChar && end === closeChar);
      if (!ok) continue;

      // Costruisci <m:d> e sostituisci open/b/close
      const d = createDelimiter(container.ownerDocument as unknown as XmlDoc, openChar, closeChar, b);

      // inserisci prima dell'open run
      container.insertBefore(d as unknown as XmlNode, a as unknown as XmlNode);

      // rimuovi open run (a) e close run (c); b è già stato "spostato" dentro d
      container.removeChild(a as unknown as XmlNode);
      container.removeChild(c as unknown as XmlNode);

      changed = true;
      break;
    }
  }
};

const traverseElements = (root: XmlEl, fn: (el: XmlEl) => void) => {
  const stack: XmlEl[] = [root];
  while (stack.length) {
    const el = stack.pop()!;
    fn(el);

    for (let n = el.lastChild; n; n = n.previousSibling) {
      if (n.nodeType === 1) stack.push(n as unknown as XmlEl);
    }
  }
};

const repairInvalidScripts = (root: XmlEl) => {
  // Unwrap di script invalidi per evitare OMML non conforme (Word lo considera “corrotto”).
  // Strategia: se manca sub/sup, rimpiazza l'elemento con il contenuto di <m:e> (base).
  const unwrapToBase = (node: XmlEl) => {
    const parent = node.parentNode as unknown as XmlEl | null;
    if (!parent) return;

    let base: XmlEl | null = null;
    for (let n = node.firstChild; n; n = n.nextSibling) {
      if (n.nodeType === 1) {
        const el = n as unknown as XmlEl;
        if (isM(el, "e")) {
          base = el;
          break;
        }
      }
    }

    // inserisci i figli della base prima del nodo script
    if (base) {
      while (base.firstChild) {
        parent.insertBefore(base.firstChild as unknown as XmlNode, node as unknown as XmlNode);
      }
    }

    parent.removeChild(node as unknown as XmlNode);
  };

  // due passate per sicurezza (modifichiamo l'albero mentre iteriamo)
  for (let pass = 0; pass < 2; pass += 1) {
    const toCheck: XmlEl[] = [];
    traverseElements(root, (el) => {
      const local = tagLocal(el.tagName);
      if (local === "sSub" || local === "sSup" || local === "sSubSup") toCheck.push(el);
    });

    for (const el of toCheck) {
      const local = tagLocal(el.tagName);

      let hasE = false;
      let hasSub = false;
      let hasSup = false;

      for (let n = el.firstChild; n; n = n.nextSibling) {
        if (n.nodeType !== 1) continue;
        const c = n as unknown as XmlEl;
        if (isM(c, "e")) hasE = true;
        if (isM(c, "sub")) hasSub = true;
        if (isM(c, "sup")) hasSup = true;
      }

      if (!hasE) {
        // senza base, non ha senso: elimina
        unwrapToBase(el);
        continue;
      }

      if (local === "sSub" && !hasSub) {
        unwrapToBase(el);
      } else if (local === "sSup" && !hasSup) {
        unwrapToBase(el);
      } else if (local === "sSubSup" && (!hasSub || !hasSup)) {
        // se manca uno dei due, unwrap a base (soluzione conservativa)
        unwrapToBase(el);
      }
    }
  }
};

const postProcessOmml = (omml: string): string | null => {
  const parsed = parseOmmlFragment(omml);
  if (!parsed) return null;

  const { root } = parsed;

  // 1) Promuovi delimitatori (stretchy) in tutti i container
  traverseElements(root, (el) => {
    // Evita di promuovere dentro <m:t> ovviamente (anche se non dovrebbe contenere elementi)
    if (isM(el, "t")) return;
    promoteDelimitersInContainer(el);
  });

  // 2) Ripara script incompleti (m:sSub senza m:sub, ecc.)
  repairInvalidScripts(root);

  // 3) Serializza children del wrapper <root> come frammento OMML finale
  return serializeOmmlChildren(root).trim();
};

const buildOmmlFromElement = (el: HTMLElement): string | null => {
  const displayMode =
    el.classList.contains("katex-display") ||
    el.querySelector("math")?.getAttribute("display") === "block";

  const tex = extractLatexFromKatex(el);
  const hasMatrixLike =
    !!tex && /\\begin\{(?:bmatrix|pmatrix|Bmatrix|vmatrix|Vmatrix|matrix|smallmatrix)\}/.test(tex);

  // Preferisci l'HTML MathML di KaTeX quando non ci sono casi “critici”
  const mathML = !hasMatrixLike ? extractMathMlFromKatex(el) : null;

  const sourceMathMl =
    mathML ??
    (() => {
      if (!tex) return null;
      return texToMathMl(normalizeDocxLatex(tex), displayMode);
    })();

  if (!sourceMathMl) return null;

  try {
    const raw = mml2omml(cleanMathMl(sourceMathMl));
    const processed = postProcessOmml(raw);

    // se post-process fallisce, meglio tornare raw (di solito valido) che null
    return processed ?? raw;
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
    return [
      paragraphFromInline([
        textRunFromNode(text, ctx?.cellHeader ? { bold: true } : undefined),
      ]),
    ];
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
            paragraphFromInline([
              textRunFromNode("", isHeaderRow ? { bold: true } : undefined),
            ]),
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
              paragraphFromInline([
                textRunFromNode("", isHeaderRow ? { bold: true } : undefined),
              ]),
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
    const items = Array.from(el.children).filter((child) => getTag(child) === "li");
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
    return [paragraphFromInline(children, ctx?.cellHeader ? {} : undefined)];
  }

  const out: DocxChild[] = [];
  for (const child of el.childNodes as HtmlNode[]) {
    out.push(...blockFromNode(child, ctx, math));
  }
  return out;
};

const patchDocxWithMath = async (fileBuffer: Buffer, replacements: MathReplacement[]) => {
  if (!replacements.length) return fileBuffer;

  const zip = await JSZip.loadAsync(fileBuffer);
  const documentXml = zip.file("word/document.xml");
  if (!documentXml) return fileBuffer;

  let xml = await documentXml.async("string");

  // Ensure OMML namespace is declared
  if (!xml.includes("xmlns:m=")) {
    xml = xml.replace(
      /<w:document\b([^>]*)>/,
      '<w:document$1 xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">',
    );
  }

  const cleanOmml = (value: string) =>
    value
      .replace(/^\s*<\?xml[^>]*\?>\s*/i, "")
      .replace(/\s+xmlns:m="[^"]+"/g, "")
      .replace(/\s+xmlns:w="[^"]+"/g, "")
      .trim();

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

    // IMPORTANT: insert OMML directly (do NOT wrap in <w:r> and do NOT escape tags)
    const safeOmml = cleanOmml(omml);
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
              { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.START },
              { level: 1, format: LevelFormat.LOWER_LETTER, text: "%2.", alignment: AlignmentType.START },
              { level: 2, format: LevelFormat.LOWER_ROMAN, text: "%3.", alignment: AlignmentType.START },
            ],
          },
        ],
      },
      sections: [{ children }],
    });

    const fileBuffer = await Packer.toBuffer(doc);
    const patched = await patchDocxWithMath(fileBuffer, math.replacements);

    const mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    const body = patched instanceof Blob ? patched : new Blob([patched as BlobPart], { type: mime });

    return new NextResponse(body, {
      status: 200,
      headers: {
        "Content-Type": mime,
        "Content-Disposition": 'attachment; filename="promptpress-export.docx"',
      },
    });
  } catch (error) {
    console.error("DOCX export error", error);
    return NextResponse.json({ error: "Errore durante la generazione del DOCX" }, { status: 500 });
  }
}

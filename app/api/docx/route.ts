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

type MathReplacement = { token: string; omml: string; displayMode: boolean };
type MathContext = { nextId: number; replacements: MathReplacement[] };

type BlockContext = {
  list?: { ordered: boolean; level: number };
  cellHeader?: boolean;
};

const M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math";

/* --------------------------------- helpers -------------------------------- */

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

const stripFenceDelimiters = (mathML: string) =>
  mathML.replace(
    /<mo[^>]*fence=\"true\"[^>]*>\s*([\[\]\(\)\{\}])\s*<\/mo>/gi,
    "",
  );

const sanitizeXmlText = (value: string) =>
  stripUnpairedSurrogates(stripZeroWidth(value)).replace(
    /[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F-\u0084\u0086-\u009F]/g,
    "",
  );

const normalizeText = (value: string) =>
  sanitizeXmlText(value)
    .replace(/&nbsp;/g, " ")
    .replace(/\u00A0/g, " ")
    .replace(/\s+/g, " ");

const normalizePreserveWhitespace = (value: string) =>
  sanitizeXmlText(value).replace(/&nbsp;/g, " ").replace(/\u00A0/g, " ");

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
 * Pre-clean specifico per converter MathML -> OMML.
 * Obiettivo: eliminare token/spacing "invisibili" (KaTeX) che alcuni build di mathml2omml non digeriscono.
 */
const preCleanMathMlForOmml = (mathML: string) => {
  let out = cleanMathMl(mathML);

  // KaTeX: wrappers e annotation aggiuntive
  out = out.replace(/<\s*semantics[^>]*>/gi, "").replace(/<\/\s*semantics\s*>/gi, "");
  out = out.replace(/<annotation-xml[\s\S]*?<\/annotation-xml>/gi, "");

  // Remove MathML styling wrappers (often break mml2omml fraction handling)
  out = out.replace(/<mstyle[^>]*>/gi, "").replace(/<\/mstyle>/gi, "");

  // Caratteri "invisibili" (KaTeX li usa per function application / spacing)
  out = out.replace(/[\u2061\u2062\u2063\u2064]/g, "");

  // NBSP e spazi tipografici -> spazio normale
  out = out.replace(/&nbsp;|&#160;|&#xA0;/gi, " ");
  out = out.replace(/[\u00A0\u2005\u2009\u200A\u202F]/g, " ");

  // Convert/strip mtext to avoid mml2omml crashes.
  out = out.replace(/<mtext\b[^>]*>([\s\S]*?)<\/mtext>/gi, (_m, inner) => {
    const cleaned = String(inner)
      .replace(/<[^>]+>/g, "")
      .replace(/[\u200B-\u200D\uFEFF\u00AD\u2060-\u206F\u202A-\u202E]/g, "")
      .replace(/\s+/g, " ")
      .trim();
    return cleaned ? `<mi>${cleaned}</mi>` : "";
  });

  // Drop empty operators created by stripping invisible chars.
  out = out.replace(/<mo\b[^>]*>\s*<\/mo>/gi, "");

  // Rewrite norm msubsup into nested msub/msup to avoid converter crashes.
  out = out.replace(/<msubsup>\s*([\s\S]*?)\s*(<mn>2<\/mn>)\s*(<mn>2<\/mn>)\s*<\/msubsup>/gi, (_m, base, sub, sup) => {
    if (!String(base).includes(String.fromCharCode(0x2225))) return _m;
    return `<msup><msub>${base}${sub}</msub>${sup}</msup>`;
  });

  // Strip problematic fence/stretchy attributes on operators.
  out = out.replace(/<mo\b[^>]*>/gi, (match) => {
    return match.replace(/\s(?:fence|stretchy|minsize|maxsize|largeop|movablelimits|accent|separator|form|lspace|rspace)="[^"]*"/gi, "");
  });

  // Alcuni parser/converter sono namespace-agnostic: rimuovere xmlns MathML puÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â² aiutare
  out = out.replace(/\sxmlns="http:\/\/www\.w3\.org\/1998\/Math\/MathML"/i, "");

  return out;
};

const tryMathMlToOmml = (mathML: string) => {
  const cleanedMathML = preCleanMathMlForOmml(mathML);
  const hasFrac = /<mfrac\b/i.test(cleanedMathML);

  const run = (ml: string) => {
    const raw = mml2omml(ml);
    const processed = postProcessOmml(raw) ?? raw;
    return finalizeOmmlString(processed);
  };

  try {
    const first = run(cleanedMathML);
    if (!hasFrac || first.includes("<m:f")) return first;

    const stripped = stripFenceDelimiters(cleanedMathML);
    if (stripped == cleanedMathML) return first;

    const second = run(stripped);
    return second || first;
  } catch (e) {
    // Debug opt-in: DOCX_MATH_DEBUG=1
    if (process.env.DOCX_MATH_DEBUG === "1") {
      // eslint-disable-next-line no-console
      console.warn("mml2omml failed", {
        msg: e instanceof Error ? e.message : String(e),
        sample: mathML.slice(0, 400),
        cleanedSample: preCleanMathMlForOmml(mathML).slice(0, 400),
      });
    }
    return null;
  }
};

/**
 * Normalizza ambienti LaTeX per rendere esplicita la semantica dei delimitatori
 * (alcuni converter MathML->OMML degradano la "stretchiness").
 */
const stripLeftRight = (tex: string) =>
  tex.replace(/\\left/g, "").replace(/\\right/g, "");

const normalizeDocxLatex = (tex: string) =>
  tex
    .replace(/\\displaystyle\s*/g, "")
    .replace(/\\left\\\{([\\s\\S]*?)\\right\\\./g, "\\left\\{$1\\right\\}")
    .replace(/\\bigl/g, "\\left")
    .replace(/\\bigr/g, "\\right")
    .replace(/\\Bigl/g, "\\left")
    .replace(/\\Bigr/g, "\\right")
    .replace(/\\biggl/g, "\\left")
    .replace(/\\biggr/g, "\\right")
    .replace(/\\Biggl/g, "\\left")
    .replace(/\\Biggr/g, "\\right")
    .replace(/\\left\\langle/g, "\\langle")
    .replace(/\\right\\rangle/g, "\\rangle")
    .replace(/\\left\\lvert/g, "\\lvert")
    .replace(/\\right\\rvert/g, "\\rvert")
    .replace(/\\left\\lVert/g, "\\lVert")
    .replace(/\\right\\rVert/g, "\\rVert")
    .replace(/\\exp\\!/g, "\\exp")
    .replace(/\\!/g, "")
    .replace(/\\quad/g, " ")
    .replace(/\\qquad/g, " ")
    .replace(/\\,/g, " ")
    .replace(/\\;/g, " ")
    .replace(/\\begin\{aligned\}/g, "\\begin{array}{ll}")
    .replace(/\\end\{aligned\}/g, "\\end{array}")
    .replace(/\\begin\{alignedat\}\{[^}]*\}/g, "\\begin{array}{ll}")
    .replace(/\\end\{alignedat\}/g, "\\end{array}")
    .replace(
      /\\begin\{bmatrix\}([\s\S]*?)\\end\{bmatrix\}/g,
      String.raw`\\left[\\begin{matrix}$1\\end{matrix}\\right]`,
    )
    .replace(
      /\\begin\{pmatrix\}([\s\S]*?)\\end\{pmatrix\}/g,
      String.raw`\\left(\\begin{matrix}$1\\end{matrix}\\right)`,
    )
    .replace(
      /\\begin\{Bmatrix\}([\s\S]*?)\\end\{Bmatrix\}/g,
      String.raw`\\left\\{\\begin{matrix}$1\\end{matrix}\\right\\}`,
    )
    .replace(
      /\\begin\{vmatrix\}([\s\S]*?)\\end\{vmatrix\}/g,
      String.raw`\\left|\\begin{matrix}$1\\end{matrix}\\right|`,
    )
    .replace(
      /\\begin\{Vmatrix\}([\s\S]*?)\\end\{Vmatrix\}/g,
      String.raw`\\left\\|\\begin{matrix}$1\\end{matrix}\\right\\|`,
    )
    .replace(
      /\\begin\{cases\}([\s\S]*?)\\end\{cases\}/g,
      String.raw`\\left\\{\\begin{array}{ll}$1\\end{array}\\right.`,
    )
    .replace(
      /\\begin\{smallmatrix\}([\s\S]*?)\\end\{smallmatrix\}/g,
      String.raw`\\left(\\begin{smallmatrix}$1\\end{smallmatrix}\\right)`,
    );

/* -------------------------- OMML DOM postprocess -------------------------- */

const tagLocal = (name: string) => (name.includes(":") ? name.split(":")[1] : name);
const isM = (el: XmlEl, local: string) => tagLocal(el.tagName) === local;

const parseOmmlFragment = (omml: string): { doc: XmlDoc; root: XmlEl } | null => {
  const cleaned = omml
    .replace(/^\s*<\?xml[^>]*\?>\s*/i, "")
    .replace(/\s+xmlns:m="[^"]+"/g, "")
    .replace(/\s+xmlns:w="[^"]+"/g, "")
    .trim();

  const wrapped = `<root xmlns:m="${M_NS}">${cleaned}</root>`;

  const doc = new DOMParser({
    errorHandler: { warning: () => {}, error: () => {}, fatalError: () => {} },
  }).parseFromString(wrapped, "application/xml") as unknown as XmlDoc;

  const root = doc.documentElement as unknown as XmlEl;
  if (!root || root.tagName !== "root") return null;

  const anyParserError = (root.getElementsByTagName("parsererror")?.length ?? 0) > 0;
  if (anyParserError) return null;

  return { doc, root };
};

const serializeOmmlChildren = (root: XmlEl) => {
  const ser = new XMLSerializer();
  let out = "";
  for (let n = root.firstChild; n; n = n.nextSibling) {
    out += ser.serializeToString(n as unknown as XmlNode);
  }
  return out;
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

const removeIfChild = (parent: XmlEl, child: XmlNode | null | undefined) => {
  if (!child) return;
  try {
    if (child.parentNode === parent) parent.removeChild(child);
  } catch {
    // ignore
  }
};

const normalizeOmmlTextNodesDom = (root: XmlEl) => {
  traverseElements(root, (el) => {
    if (!isM(el, "t")) return;

    let t = el.textContent ?? "";

    t = t.replace(/&nbsp;|&#160;|&#xA0;/gi, " ");
    t = t.replace(/&thinsp;|&ensp;|&emsp;/gi, " ");
    t = t.replace(/\u00A0|\u202F/g, " ");
    t = t.replace(/[\u200B\u200C\u200D\uFEFF\u00AD\u2060-\u2064]/g, "");
    t = t.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, "");
    t = t.replace(/<\/?m:[^>]+>/g, "");

    if (/^\s|\s$| {2,}/.test(t)) {
      const has = el.getAttribute("xml:space");
      if (!has) el.setAttribute("xml:space", "preserve");
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (el as any).textContent = t;
  });
};

const removeNonMathElements = (root: XmlEl) => {
  const toRemove: XmlEl[] = [];
  traverseElements(root, (el) => {
    if (el.tagName === "root") return;

    if (/^(w|r|wp|w14|w15|w16|o|v|mc):/.test(el.tagName)) {
      toRemove.push(el);
      return;
    }

    if (el.tagName.startsWith("m:")) return;

    const ns = (el as unknown as { namespaceURI?: string | null }).namespaceURI;
    if (ns && ns !== M_NS) toRemove.push(el);
  });

  for (const el of toRemove) {
    const parent = el.parentNode as unknown as XmlEl | null;
    if (!parent) continue;
    parent.removeChild(el as unknown as XmlNode);
  }
};

const repairInvalidScripts = (root: XmlEl) => {
  const unwrapToBase = (node: XmlEl) => {
    const parent = node.parentNode as unknown as XmlEl | null;
    if (!parent) return;

    let base: XmlEl | null = null;
    for (let n = node.firstChild; n; n = n.nextSibling) {
      if (n.nodeType !== 1) continue;
      const el = n as unknown as XmlEl;
      if (isM(el, "e")) {
        base = el;
        break;
      }
    }

    if (base) {
      while (base.firstChild) {
        parent.insertBefore(base.firstChild as unknown as XmlNode, node as unknown as XmlNode);
      }
    }

    parent.removeChild(node as unknown as XmlNode);
  };

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
        unwrapToBase(el);
        continue;
      }

      if (local === "sSub" && !hasSub) unwrapToBase(el);
      else if (local === "sSup" && !hasSup) unwrapToBase(el);
      else if (local === "sSubSup" && (!hasSub || !hasSup)) unwrapToBase(el);
    }
  }
};

const runTextIfSingleChar = (run: XmlEl) => {
  if (!isM(run, "r")) return null;

  let tEl: XmlEl | null = null;
  for (let n = run.firstChild; n; n = n.nextSibling) {
    if (n.nodeType !== 1) continue;
    const el = n as unknown as XmlEl;
    if (isM(el, "t")) {
      tEl = el;
      break;
    }
  }
  if (!tEl) return null;

  const txt = (tEl.textContent ?? "").trim();
  if (!txt) return null;

  const cps = Array.from(txt);
  if (cps.length !== 1) return null;
  return cps[0];
};

const isTall = (el: XmlEl) => {
  const local = tagLocal(el.tagName);
  return new Set([
    "m",
    "eqArr",
    "f",
    "rad",
    "nary",
    "stack",
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

const promoteDelimitersInContainer = (container: XmlEl) => {
  const pairs: Array<[string, string]> = [
    ["(", ")"],
    ["[", "]"],
    ["{", "}"],
    ["ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â¦ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¸ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¨", "ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â¦ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¸ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â©"],
    ["|", "|"],
    ["ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã¢â‚¬Å“", "ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã¢â‚¬Å“"],
  ];

  const begOnly = new Set(["(", "[", "{", "ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â¦ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¸ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¨", "|", "ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã¢â‚¬Å“"]);

  const nextElementIndex = (nodes: XmlNode[], start: number) => {
    for (let i = start; i < nodes.length; i += 1) {
      if (nodes[i]?.nodeType === 1) return i;
    }
    return -1;
  };

  const createDelimiterWithOptionalEnd = (doc: XmlDoc, beg: string, end: string, inner: XmlEl) => {
    const d = doc.createElementNS(M_NS, "m:d") as unknown as XmlEl;
    const dPr = doc.createElementNS(M_NS, "m:dPr") as unknown as XmlEl;

    const begChr = doc.createElementNS(M_NS, "m:begChr") as unknown as XmlEl;
    begChr.setAttribute("m:val", beg);

    const endChr = doc.createElementNS(M_NS, "m:endChr") as unknown as XmlEl;
    endChr.setAttribute("m:val", end);

    const grow = doc.createElementNS(M_NS, "m:grow") as unknown as XmlEl;
    grow.setAttribute("m:val", "1");

    dPr.appendChild(begChr as unknown as XmlNode);
    dPr.appendChild(endChr as unknown as XmlNode);
    dPr.appendChild(grow as unknown as XmlNode);

    const e = doc.createElementNS(M_NS, "m:e") as unknown as XmlEl;
    e.appendChild(inner as unknown as XmlNode);

    d.appendChild(dPr as unknown as XmlNode);
    d.appendChild(e as unknown as XmlNode);

    return d;
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

      const b = nodes[j] as unknown as XmlEl;
      if (!isTall(b)) continue;

      const k = nextElementIndex(nodes, j + 1);
      const c = k !== -1 ? (nodes[k] as unknown as XmlEl) : null;
      const closeChar = c ? runTextIfSingleChar(c) : null;

      if (closeChar) {
        const okPair = pairs.some(([beg, end]) => beg === openChar && end === closeChar);
        if (okPair) {
          const d = createDelimiterWithOptionalEnd(
            container.ownerDocument as unknown as XmlDoc,
            openChar,
            closeChar,
            b,
          );

          container.insertBefore(d as unknown as XmlNode, a as unknown as XmlNode);
          removeIfChild(container, a as unknown as XmlNode);
          removeIfChild(container, c as unknown as XmlNode);

          changed = true;
          break;
        }
      }
    }
  }
};

const wrapParensInScripts = (root: XmlEl) => {
  const shouldWrap = new Set(["sSup", "sSub", "sSubSup"]);
  const pairs: Array<[string, string]> = [
    ["(", ")"],
    ["{", "}"],
  ];

  traverseElements(root, (el) => {
    const local = tagLocal(el.tagName);
    if (!shouldWrap.has(local)) return;

    let eNode: XmlEl | null = null;
    for (let n = el.firstChild; n; n = n.nextSibling) {
      if (n.nodeType !== 1) continue;
      const c = n as unknown as XmlEl;
      if (isM(c, "e")) {
        eNode = c;
        break;
      }
    }
    if (!eNode) return;

    const elementChildren = Array.from(eNode.childNodes).filter(
      (n) => n.nodeType === 1,
    ) as XmlEl[];
    if (elementChildren.length < 3) return;

    const first = elementChildren[0];
    const last = elementChildren[elementChildren.length - 1];

    const openChar = runTextIfSingleChar(first);
    const closeChar = runTextIfSingleChar(last);
    if (!openChar || !closeChar) return;

    const match = pairs.find(([beg, end]) => beg === openChar && end === closeChar);
    if (!match) return;

    const doc = eNode.ownerDocument as unknown as XmlDoc;
    const d = doc.createElementNS(M_NS, "m:d") as unknown as XmlEl;
    const dPr = doc.createElementNS(M_NS, "m:dPr") as unknown as XmlEl;
    const begChr = doc.createElementNS(M_NS, "m:begChr") as unknown as XmlEl;
    const endChr = doc.createElementNS(M_NS, "m:endChr") as unknown as XmlEl;
    const grow = doc.createElementNS(M_NS, "m:grow") as unknown as XmlEl;
    begChr.setAttribute("m:val", openChar);
    endChr.setAttribute("m:val", closeChar);
    grow.setAttribute("m:val", "1");
    dPr.appendChild(begChr as unknown as XmlNode);
    dPr.appendChild(endChr as unknown as XmlNode);
    dPr.appendChild(grow as unknown as XmlNode);

    const e = doc.createElementNS(M_NS, "m:e") as unknown as XmlEl;

    let started = false;
    const toMove: XmlNode[] = [];
    for (let n = eNode.firstChild; n; n = n.nextSibling) {
      if (n === first) {
        started = true;
        continue;
      }
      if (!started) continue;
      if (n === last) break;
      toMove.push(n as unknown as XmlNode);
    }

    for (const n of toMove) {
      removeIfChild(eNode, n);
      e.appendChild(n);
    }

    d.appendChild(dPr as unknown as XmlNode);
    d.appendChild(e as unknown as XmlNode);

    const anchor = first as unknown as XmlNode;
    eNode.insertBefore(d as unknown as XmlNode, anchor);
    removeIfChild(eNode, anchor);
    removeIfChild(eNode, last as unknown as XmlNode);
  });
};

const wrapBraceBeforeMatrix = (root: XmlEl) => {
  traverseElements(root, (el) => {
    if (isM(el, "t") || isM(el, "nary")) return;

    let changed = true;
    while (changed) {
      changed = false;
      const nodes = Array.from(el.childNodes) as unknown as XmlNode[];

      for (let i = 0; i < nodes.length - 1; i += 1) {
        const a = nodes[i];
        const b = nodes[i + 1];
        if (!a || !b || a.nodeType != 1 || b.nodeType != 1) continue;

        const openChar = runTextIfSingleChar(a as unknown as XmlEl);
        if (openChar != "{") continue;
        if (!isM(b as unknown as XmlEl, "m")) continue;

        const doc = el.ownerDocument as unknown as XmlDoc;
        const d = doc.createElementNS(M_NS, "m:d") as unknown as XmlEl;
        const dPr = doc.createElementNS(M_NS, "m:dPr") as unknown as XmlEl;
        const begChr = doc.createElementNS(M_NS, "m:begChr") as unknown as XmlEl;
        const endChr = doc.createElementNS(M_NS, "m:endChr") as unknown as XmlEl;
        const grow = doc.createElementNS(M_NS, "m:grow") as unknown as XmlEl;
        begChr.setAttribute("m:val", "{");
        endChr.setAttribute("m:val", "");
        grow.setAttribute("m:val", "1");
        dPr.appendChild(begChr as unknown as XmlNode);
        dPr.appendChild(endChr as unknown as XmlNode);
        dPr.appendChild(grow as unknown as XmlNode);

        const e = doc.createElementNS(M_NS, "m:e") as unknown as XmlEl;
        e.appendChild(b as unknown as XmlNode);
        d.appendChild(dPr as unknown as XmlNode);
        d.appendChild(e as unknown as XmlNode);

        el.insertBefore(d as unknown as XmlNode, a as unknown as XmlNode);
        removeIfChild(el, a);
        removeIfChild(el, b);
        changed = true;
        break;
      }
    }
  });
};
const wrapSquareBracketsAroundTall = (root: XmlEl) => {
  traverseElements(root, (el) => {
    if (isM(el, "t") || isM(el, "nary")) return;

    let changed = true;
    while (changed) {
      changed = false;
      const nodes = Array.from(el.childNodes) as unknown as XmlNode[];

      for (let i = 0; i < nodes.length - 1; i += 1) {
        const open = nodes[i];
        if (!open || open.nodeType !== 1) continue;
        const openChar = runTextIfSingleChar(open as unknown as XmlEl);
        if (openChar !== "[") continue;

        let closeIdx = -1;
        let hasTall = false;
        for (let j = i + 1; j < nodes.length; j += 1) {
          const candidate = nodes[j];
          if (!candidate || candidate.nodeType !== 1) continue;
          const closeChar = runTextIfSingleChar(candidate as unknown as XmlEl);
          if (closeChar === "]") {
            closeIdx = j;
            break;
          }
          if (isTall(candidate as unknown as XmlEl)) hasTall = true;
        }
        if (closeIdx === -1 || !hasTall) continue;

        const doc = el.ownerDocument as unknown as XmlDoc;
        const d = doc.createElementNS(M_NS, "m:d") as unknown as XmlEl;
        const dPr = doc.createElementNS(M_NS, "m:dPr") as unknown as XmlEl;
        const begChr = doc.createElementNS(M_NS, "m:begChr") as unknown as XmlEl;
        const endChr = doc.createElementNS(M_NS, "m:endChr") as unknown as XmlEl;
        const grow = doc.createElementNS(M_NS, "m:grow") as unknown as XmlEl;
        begChr.setAttribute("m:val", "[");
        endChr.setAttribute("m:val", "]");
        grow.setAttribute("m:val", "1");
        dPr.appendChild(begChr as unknown as XmlNode);
        dPr.appendChild(endChr as unknown as XmlNode);
        dPr.appendChild(grow as unknown as XmlNode);

        const e = doc.createElementNS(M_NS, "m:e") as unknown as XmlEl;
        for (let k = i + 1; k < closeIdx; k += 1) {
          const n = nodes[k];
          if (!n) continue;
          removeIfChild(el, n);
          e.appendChild(n as unknown as XmlNode);
        }

        d.appendChild(dPr as unknown as XmlNode);
        d.appendChild(e as unknown as XmlNode);

        el.insertBefore(d as unknown as XmlNode, open as unknown as XmlNode);
        removeIfChild(el, open);
        removeIfChild(el, nodes[closeIdx] as unknown as XmlNode);
        changed = true;
        break;
      }
    }
  });
};

const fixNormSupSubRuns = (root: XmlEl) => {
  const makeRun = (doc: XmlDoc, value: string) => {
    const r = doc.createElementNS(M_NS, "m:r") as unknown as XmlEl;
    const t = doc.createElementNS(M_NS, "m:t") as unknown as XmlEl;
    t.setAttribute("xml:space", "preserve");
    (t as any).textContent = value;
    r.appendChild(t as unknown as XmlNode);
    return r;
  };

  traverseElements(root, (el) => {
    if (isM(el, "t") || isM(el, "nary")) return;

    const nodes = Array.from(el.childNodes) as unknown as XmlNode[];
    for (let i = 0; i < nodes.length; i += 1) {
      const node = nodes[i];
      if (!node || node.nodeType !== 1) continue;
      const run = node as unknown as XmlEl;
      if (!isM(run, "r")) continue;

      let tEl: XmlEl | null = null;
      for (let n = run.firstChild; n; n = n.nextSibling) {
        if (n.nodeType !== 1) continue;
        const c = n as unknown as XmlEl;
        if (isM(c, "t")) {
          tEl = c;
          break;
        }
      }
      if (!tEl) continue;

      const text = tEl.textContent ?? "";
      const match = /(.*)\u2225([0-9]+)(.*)/.exec(text);
      if (!match) continue;

      const prefix = match[1] ?? "";
      const digits = match[2] ?? "";
      const suffix = match[3] ?? "";
      if (!digits) continue;

      const doc = el.ownerDocument as unknown as XmlDoc;
      const newNodes: XmlNode[] = [];

      if (prefix) newNodes.push(makeRun(doc, prefix) as unknown as XmlNode);

      if (digits.length >= 2) {
        const sub = digits[0];
        const sup = digits.slice(1);

        const sSubSup = doc.createElementNS(M_NS, "m:sSubSup") as unknown as XmlEl;
        const sSubSupPr = doc.createElementNS(M_NS, "m:sSubSupPr") as unknown as XmlEl;
        sSubSup.appendChild(sSubSupPr as unknown as XmlNode);

        const e = doc.createElementNS(M_NS, "m:e") as unknown as XmlEl;
        e.appendChild(makeRun(doc, "\u2225") as unknown as XmlNode);
        sSubSup.appendChild(e as unknown as XmlNode);

        const subEl = doc.createElementNS(M_NS, "m:sub") as unknown as XmlEl;
        subEl.appendChild(makeRun(doc, sub) as unknown as XmlNode);
        sSubSup.appendChild(subEl as unknown as XmlNode);

        const supEl = doc.createElementNS(M_NS, "m:sup") as unknown as XmlEl;
        supEl.appendChild(makeRun(doc, sup) as unknown as XmlNode);
        sSubSup.appendChild(supEl as unknown as XmlNode);

        newNodes.push(sSubSup as unknown as XmlNode);
      } else {
        const sSub = doc.createElementNS(M_NS, "m:sSub") as unknown as XmlEl;
        const sSubPr = doc.createElementNS(M_NS, "m:sSubPr") as unknown as XmlEl;
        sSub.appendChild(sSubPr as unknown as XmlNode);

        const e = doc.createElementNS(M_NS, "m:e") as unknown as XmlEl;
        e.appendChild(makeRun(doc, "\u2225") as unknown as XmlNode);
        sSub.appendChild(e as unknown as XmlNode);

        const subEl = doc.createElementNS(M_NS, "m:sub") as unknown as XmlEl;
        subEl.appendChild(makeRun(doc, digits) as unknown as XmlNode);
        sSub.appendChild(subEl as unknown as XmlNode);

        newNodes.push(sSub as unknown as XmlNode);
      }

      if (suffix) newNodes.push(makeRun(doc, suffix) as unknown as XmlNode);

      for (const n of newNodes) {
        el.insertBefore(n, run as unknown as XmlNode);
      }
      removeIfChild(el, run as unknown as XmlNode);
    }
  });
};


const fixRunIntegrals = (root: XmlEl) => {
  const integral = String.fromCharCode(0x222B);
  const infinity = String.fromCharCode(0x221E);

  const makeRun = (doc: XmlDoc, value: string) => {
    const r = doc.createElementNS(M_NS, "m:r") as unknown as XmlEl;
    const t = doc.createElementNS(M_NS, "m:t") as unknown as XmlEl;
    t.setAttribute("xml:space", "preserve");
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (t as any).textContent = value;
    r.appendChild(t as unknown as XmlNode);
    return r;
  };
  traverseElements(root, (el) => {
    if (isM(el, "nary") || isM(el, "t")) return;

    const nodes = Array.from(el.childNodes) as unknown as XmlNode[];
    let i = 0;
    while (i < nodes.length) {
      const node = nodes[i];
      if (!node || node.nodeType !== 1) {
        i += 1;
        continue;
      }
      const run = node as unknown as XmlEl;
      if (!isM(run, "r")) {
        i += 1;
        continue;
      }

      const text = (run.textContent ?? "").trim();
      let sub: string | null = null;
      let sup: string | null = null;
      let consume = 1;

      const integralOnly = text === integral;
      const integralWithSub = text.startsWith(integral) && text.length > 1;
      if (integralOnly || integralWithSub) {
        if (integralWithSub) {
          const rest = text.slice(integral.length).trim();
          if (/^\d+$/.test(rest)) sub = rest;
        }

        const next = nodes[i + 1] as unknown as XmlEl | undefined;
        const nextText = next && next.nodeType === 1 ? (next.textContent ?? "").trim() : "";
        if (nextText === infinity) {
          sup = infinity;
          consume = 2;
        } else if (integralOnly && /^\d+$/.test(nextText)) {
          sub = nextText;
          const next2 = nodes[i + 2] as unknown as XmlEl | undefined;
          const next2Text = next2 && next2.nodeType === 1 ? (next2.textContent ?? "").trim() : "";
          if (next2Text === infinity) {
            sup = infinity;
            consume = 3;
          } else {
            consume = 2;
          }
        }
      }

      if (!integralOnly && !integralWithSub) {
        i += 1;
        continue;
      }

      const doc = el.ownerDocument as unknown as XmlDoc;
      const nary = doc.createElementNS(M_NS, "m:nary") as unknown as XmlEl;
      const naryPr = doc.createElementNS(M_NS, "m:naryPr") as unknown as XmlEl;
      const chr = doc.createElementNS(M_NS, "m:chr") as unknown as XmlEl;
      chr.setAttribute("m:val", integral);
      const limLoc = doc.createElementNS(M_NS, "m:limLoc") as unknown as XmlEl;
      limLoc.setAttribute("m:val", "undOvr");
      const grow = doc.createElementNS(M_NS, "m:grow") as unknown as XmlEl;
      grow.setAttribute("m:val", "1");
      const subHide = doc.createElementNS(M_NS, "m:subHide") as unknown as XmlEl;
      subHide.setAttribute("m:val", "off");
      const supHide = doc.createElementNS(M_NS, "m:supHide") as unknown as XmlEl;
      supHide.setAttribute("m:val", "off");
      naryPr.appendChild(chr as unknown as XmlNode);
      naryPr.appendChild(limLoc as unknown as XmlNode);
      naryPr.appendChild(grow as unknown as XmlNode);
      naryPr.appendChild(subHide as unknown as XmlNode);
      naryPr.appendChild(supHide as unknown as XmlNode);
      nary.appendChild(naryPr as unknown as XmlNode);

      const subEl = doc.createElementNS(M_NS, "m:sub") as unknown as XmlEl;
      if (sub) subEl.appendChild(makeRun(doc, sub) as unknown as XmlNode);
      nary.appendChild(subEl as unknown as XmlNode);

      const supEl = doc.createElementNS(M_NS, "m:sup") as unknown as XmlEl;
      if (sup) supEl.appendChild(makeRun(doc, sup) as unknown as XmlNode);
      nary.appendChild(supEl as unknown as XmlNode);

      const e = doc.createElementNS(M_NS, "m:e") as unknown as XmlEl;
      nary.appendChild(e as unknown as XmlNode);

      el.insertBefore(nary as unknown as XmlNode, run as unknown as XmlNode);
      for (let c = 0; c < consume; c += 1) {
        const toRemove = nodes[i + c] as unknown as XmlNode | undefined;
        if (toRemove && toRemove.parentNode === el) {
          el.removeChild(toRemove as unknown as XmlNode);
        }
      }

      nodes.splice(i, consume, nary as unknown as XmlNode);
      i += 1;
    }
  });
};

const fixEmptyNaryIntegrand = (root: XmlEl) => {
  traverseElements(root, (el) => {
    if (!isM(el, "nary")) return;

    let eNode: XmlEl | null = null;
    for (let n = el.firstChild; n; n = n.nextSibling) {
      if (n.nodeType !== 1) continue;
      const c = n as unknown as XmlEl;
      if (isM(c, "e")) {
        eNode = c;
        break;
      }
    }
    if (!eNode) return;

    const hasChildEl = Array.from(eNode.childNodes).some((n) => n.nodeType === 1);
    if (hasChildEl) return;

    const parent = el.parentNode as unknown as XmlEl | null;
    if (!parent) return;

    let next = el.nextSibling as unknown as XmlNode | null;
    while (next && next.nodeType !== 1) next = next.nextSibling as unknown as XmlNode | null;
    if (!next) return;

    const nextEl = next as unknown as XmlEl;

    if (isM(nextEl, "r")) {
      const t = (nextEl.textContent ?? "").replace(/\s+/g, "");
      if (/^d.*x/i.test(t)) return;
    }

    parent.removeChild(nextEl as unknown as XmlNode);
    eNode.appendChild(nextEl as unknown as XmlNode);
  });
};

const forceIntegralLimitsUnderOver = (root: XmlEl) => {
  const integrals = new Set([
    String.fromCharCode(0x222B),
    String.fromCharCode(0x222E),
    String.fromCharCode(0x222F),
    String.fromCharCode(0x2230),
  ]);

  traverseElements(root, (el) => {
    if (!isM(el, "nary")) return;

    let naryPr: XmlEl | null = null;
    for (let n = el.firstChild; n; n = n.nextSibling) {
      if (n.nodeType !== 1) continue;
      const c = n as unknown as XmlEl;
      if (isM(c, "naryPr")) {
        naryPr = c;
        break;
      }
    }
    if (!naryPr) return;

    let chrEl: XmlEl | null = null;
    for (let n = naryPr.firstChild; n; n = n.nextSibling) {
      if (n.nodeType !== 1) continue;
      const c = n as unknown as XmlEl;
      if (isM(c, "chr")) {
        chrEl = c;
        break;
      }
    }

    const chrVal = chrEl?.getAttribute("m:val") ?? chrEl?.getAttribute("val") ?? "";
    if (!integrals.has(chrVal)) return;

    let limLoc: XmlEl | null = null;
    for (let n = naryPr.firstChild; n; n = n.nextSibling) {
      if (n.nodeType !== 1) continue;
      const c = n as unknown as XmlEl;
      if (isM(c, "limLoc")) {
        limLoc = c;
        break;
      }
    }

    if (!limLoc) {
      const doc = naryPr.ownerDocument as unknown as XmlDoc;
      limLoc = doc.createElementNS(M_NS, "m:limLoc") as unknown as XmlEl;
      naryPr.appendChild(limLoc as unknown as XmlNode);
    }

    limLoc.setAttribute("m:val", "undOvr");
  });
};
const getRunText = (run: XmlEl) => {
  if (!isM(run, "r")) return "";
  return (run.textContent ?? "").trim();
};

const makeRunWithText = (doc: XmlDoc, value: string) => {
  const r = doc.createElementNS(M_NS, "m:r") as unknown as XmlEl;
  const t = doc.createElementNS(M_NS, "m:t") as unknown as XmlEl;
  t.setAttribute("xml:space", "preserve");
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (t as any).textContent = value;
  r.appendChild(t as unknown as XmlNode);
  return r;
};

const isNaryWithChr = (node: XmlEl, chr: string) => {
  if (!isM(node, "nary")) return false;
  let naryPr: XmlEl | null = null;
  for (let n = node.firstChild; n; n = n.nextSibling) {
    if (n.nodeType !== 1) continue;
    const c = n as unknown as XmlEl;
    if (isM(c, "naryPr")) { naryPr = c; break; }
  }
  if (!naryPr) return false;
  for (let n = naryPr.firstChild; n; n = n.nextSibling) {
    if (n.nodeType !== 1) continue;
    const c = n as unknown as XmlEl;
    if (isM(c, "chr")) {
      const val = c.getAttribute("m:val") ?? c.getAttribute("val") ?? "";
      return val === chr;
    }
  }
  return false;
};

const fixBracketedFraction = (root: XmlEl) => {
  const integral = String.fromCharCode(0x222B);
  const summation = String.fromCharCode(0x2211);

  traverseElements(root, (el) => {
    if (isM(el, "t")) return;

    const nodes = Array.from(el.childNodes) as unknown as XmlNode[];
    for (let i = 0; i < nodes.length; i += 1) {
      const n0 = nodes[i] as unknown as XmlEl | undefined;
      const n1 = nodes[i + 1] as unknown as XmlEl | undefined;
      const n2 = nodes[i + 2] as unknown as XmlEl | undefined;
      const n3 = nodes[i + 3] as unknown as XmlEl | undefined;
      const n4 = nodes[i + 4] as unknown as XmlEl | undefined;
      if (!n0 || !n1 || !n2) continue;
      if (n0.nodeType !== 1 || n1.nodeType !== 1 || n2.nodeType !== 1) continue;

      if (getRunText(n0) !== "[") continue;
      if (!isNaryWithChr(n1, integral)) continue;

      const run2Text = getRunText(n2);
      if (!run2Text) continue;

      let denomRun: XmlEl | null = null;
      let sumNode: XmlEl | null = null;
      let closeNode: XmlEl | null = null;
      let numeratorTail = run2Text;
      let denomHead = "";
      let consume = 3;

      if (run2Text.includes("1+")) {
        const parts = run2Text.split("1+");
        numeratorTail = parts[0];
        denomHead = "1+" + parts.slice(1).join("1+");
        if (n3 && n3.nodeType === 1 && isNaryWithChr(n3 as unknown as XmlEl, summation)) {
          sumNode = n3 as unknown as XmlEl;
          closeNode = n4 && n4.nodeType === 1 ? (n4 as unknown as XmlEl) : null;
          consume = 4 + (closeNode && getRunText(closeNode) === "]" ? 1 : 0);
        }
      } else {
        if (n3 && n3.nodeType === 1 && getRunText(n3 as unknown as XmlEl) === "1+") {
          denomRun = n3 as unknown as XmlEl;
          if (n4 && n4.nodeType === 1 && isNaryWithChr(n4 as unknown as XmlEl, summation)) {
            sumNode = n4 as unknown as XmlEl;
            const n5 = nodes[i + 5] as unknown as XmlEl | undefined;
            closeNode = n5 && n5.nodeType === 1 ? (n5 as unknown as XmlEl) : null;
            consume = 5 + (closeNode && getRunText(closeNode) === "]" ? 1 : 0);
          }
        }
      }

      if (!sumNode || !closeNode || getRunText(closeNode) !== "]") continue;

      const doc = el.ownerDocument as unknown as XmlDoc;
      const d = doc.createElementNS(M_NS, "m:d") as unknown as XmlEl;
      const dPr = doc.createElementNS(M_NS, "m:dPr") as unknown as XmlEl;
      const begChr = doc.createElementNS(M_NS, "m:begChr") as unknown as XmlEl;
      begChr.setAttribute("m:val", "[");
      const endChr = doc.createElementNS(M_NS, "m:endChr") as unknown as XmlEl;
      endChr.setAttribute("m:val", "]");
      const grow = doc.createElementNS(M_NS, "m:grow") as unknown as XmlEl;
      grow.setAttribute("m:val", "1");
      dPr.appendChild(begChr as unknown as XmlNode);
      dPr.appendChild(endChr as unknown as XmlNode);
      dPr.appendChild(grow as unknown as XmlNode);
      d.appendChild(dPr as unknown as XmlNode);
      const e = doc.createElementNS(M_NS, "m:e") as unknown as XmlEl;
      const f = doc.createElementNS(M_NS, "m:f") as unknown as XmlEl;
      const fPr = doc.createElementNS(M_NS, "m:fPr") as unknown as XmlEl;
      const fType = doc.createElementNS(M_NS, "m:type") as unknown as XmlEl;
      fType.setAttribute("m:val", "bar");
      fPr.appendChild(fType as unknown as XmlNode);
      f.appendChild(fPr as unknown as XmlNode);
      const num = doc.createElementNS(M_NS, "m:num") as unknown as XmlEl;
      num.appendChild(n1 as unknown as XmlNode);
      if (numeratorTail.trim()) {
        num.appendChild(makeRunWithText(doc, numeratorTail) as unknown as XmlNode);
      }
      const den = doc.createElementNS(M_NS, "m:den") as unknown as XmlEl;
      if (denomHead) {
        den.appendChild(makeRunWithText(doc, denomHead) as unknown as XmlNode);
      } else if (denomRun) {
        den.appendChild(denomRun as unknown as XmlNode);
      }
      den.appendChild(sumNode as unknown as XmlNode);
      f.appendChild(num as unknown as XmlNode);
      f.appendChild(den as unknown as XmlNode);
      e.appendChild(f as unknown as XmlNode);
      d.appendChild(e as unknown as XmlNode);

      el.insertBefore(d as unknown as XmlNode, n0 as unknown as XmlNode);
      for (let c = 0; c < consume; c += 1) {
        const toRemove = nodes[i + c] as unknown as XmlNode | undefined;
        if (toRemove && toRemove.parentNode === el) {
          el.removeChild(toRemove as unknown as XmlNode);
        }
      }
      break;
    }
  });
};


const removeUndefinedStyle = (root: XmlEl) => {
  const toRemove: XmlEl[] = [];
  traverseElements(root, (el) => {
    if (!isM(el, "sty")) return;
    const v = el.getAttribute("m:val") ?? el.getAttribute("val") ?? "";
    if (v === "undefined") toRemove.push(el);
  });
  for (const el of toRemove) {
    const p = el.parentNode as unknown as XmlEl | null;
    if (!p) continue;
    p.removeChild(el as unknown as XmlNode);
  }
};

const postProcessOmmlWithoutDelimiters = (omml: string): string | null => {
  const parsed = parseOmmlFragment(omml);
  if (!parsed) return null;

  const { root } = parsed;

  removeNonMathElements(root);
  fixRunIntegrals(root);
  fixNormSupSubRuns(root);
  fixEmptyNaryIntegrand(root);
  forceIntegralLimitsUnderOver(root);
  wrapParensInScripts(root);
  wrapSquareBracketsAroundTall(root);
  wrapBraceBeforeMatrix(root);
  fixBracketedFraction(root);
  repairInvalidScripts(root);
  removeUndefinedStyle(root);
  normalizeOmmlTextNodesDom(root);

  return serializeOmmlChildren(root).trim();
};

const postProcessOmml = (omml: string): string | null => {
  const parsed = parseOmmlFragment(omml);
  if (!parsed) return null;

  const { root } = parsed;

  removeNonMathElements(root);
  fixRunIntegrals(root);
  fixNormSupSubRuns(root);
  fixEmptyNaryIntegrand(root);
  forceIntegralLimitsUnderOver(root);
  wrapParensInScripts(root);
  wrapSquareBracketsAroundTall(root);
  wrapBraceBeforeMatrix(root);
  fixBracketedFraction(root);

  const hasFraction = omml.includes("<m:f");
  if (!hasFraction) {
    traverseElements(root, (el) => {
      if (isM(el, "t")) return;
      promoteDelimitersInContainer(el);
    });
  }

  repairInvalidScripts(root);
  removeUndefinedStyle(root);
  normalizeOmmlTextNodesDom(root);

  const full = serializeOmmlChildren(root).trim();
  if (omml.includes("<m:f") && !full.includes("<m:f")) {
    return postProcessOmmlWithoutDelimiters(omml) ?? full;
  }

  return full;
};

const finalizeOmmlString = (omml: string) =>
  omml
    .replace(/^\s*<\?xml[^>]*\?>\s*/i, "")
    .replace(/\s+xmlns:m="[^"]+"/g, "")
    .replace(/\s+xmlns:w="[^"]+"/g, "")
    .trim();

const buildOmmlFromElement = (el: HTMLElement): string | null => {
  const displayMode =
    el.classList.contains("katex-display") ||
    el.querySelector("math")?.getAttribute("display") === "block";

  const tex = extractLatexFromKatex(el);

  const hasMatrixLike =
    !!tex && /\\begin\{(?:bmatrix|pmatrix|Bmatrix|vmatrix|Vmatrix|matrix|smallmatrix)\}/.test(tex);

  const mathMLFromDom = !hasMatrixLike ? extractMathMlFromKatex(el) : null;

  if (mathMLFromDom) {
    const omml = tryMathMlToOmml(mathMLFromDom);
    if (omml) return omml;
  }

  if (tex) {
    const normalized = normalizeDocxLatex(tex);
    const hasFrac = /\\frac\s*\{/.test(tex);

    const regenerated = texToMathMl(normalized, displayMode);
    if (regenerated) {
      const omml = tryMathMlToOmml(regenerated);
      if (omml && (!hasFrac || omml.includes("<m:f"))) return omml;
    }

    if (hasFrac) {
      const stripped = stripLeftRight(normalized);
      if (stripped != normalized) {
        const regeneratedAlt = texToMathMl(stripped, displayMode);
        if (regeneratedAlt) {
          const ommlAlt = tryMathMlToOmml(regeneratedAlt);
          if (ommlAlt) return ommlAlt;
        }
      }
    }
  }

  const mathMLAny = extractMathMlFromKatex(el);
  if (mathMLAny && mathMLAny !== mathMLFromDom) {
    const omml = tryMathMlToOmml(mathMLAny);
    if (omml) return omml;
  }

  return null;
};

/* ------------------------------ DOC generation ----------------------------- */

const createMathPlaceholder = (ctx: MathContext, el: HTMLElement) => {
  const displayMode =
    el.classList.contains("katex-display") ||
    el.querySelector("math")?.getAttribute("display") === "block";

  const omml = buildOmmlFromElement(el);
  if (!omml) return null;

  const token = `__MATH_${String(ctx.nextId).padStart(6, "0")}__`;
  ctx.nextId += 1;
  ctx.replacements.push({ token, omml, displayMode });
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

    out.push(...inlineFromNodes(el.childNodes as HtmlNode[], nextMarks, math));
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

    const headerShading = { type: ShadingType.CLEAR, color: "auto", fill: "1E293B" };

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

    for (const item of items) {
      out.push(
        ...blockFromNode(
          item,
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
    return [
      paragraphFromInline(children, {
        bulletLevel: list && !list.ordered ? list.level - 1 : undefined,
        numberLevel: list && list.ordered ? list.level - 1 : undefined,
      }),
    ];
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
    return [paragraphFromInline(children)];
  }

  const out: DocxChild[] = [];
  for (const child of el.childNodes as HtmlNode[]) {
    out.push(...blockFromNode(child, ctx, math));
  }
  return out;
};

/* ---------------------------- DOCX patch (OMML) ---------------------------- */

const patchDocxWithMath = async (fileBuffer: Buffer, replacements: MathReplacement[]) => {
  if (!replacements.length) return fileBuffer;

  const zip = await JSZip.loadAsync(fileBuffer);
  const documentXml = zip.file("word/document.xml");
  if (!documentXml) return fileBuffer;

  let xml = await documentXml.async("string");

  if (!xml.includes('xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"')) {
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

  for (const { token, omml, displayMode } of replacements) {
    const tokenIndex = xml.indexOf(token);
    if (tokenIndex === -1) continue;

    const runStart = xml.lastIndexOf("<w:r", tokenIndex);
    if (runStart === -1) continue;

    const runEnd = xml.indexOf("</w:r>", tokenIndex);
    if (runEnd === -1) continue;

    const runClose = runEnd + "</w:r>".length;
    const runSlice = xml.slice(runStart, runClose);
    if (!runSlice.includes(token)) continue;

    let safeOmml = cleanOmml(omml);
    if (displayMode && !safeOmml.includes("<m:oMathPara")) {
      safeOmml = `<m:oMathPara>${safeOmml}</m:oMathPara>`;
    }

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
    // eslint-disable-next-line no-console
    console.error("DOCX export error", error);
    return NextResponse.json({ error: "Errore durante la generazione del DOCX" }, { status: 500 });
  }
}









/**
 * A PERCENTAGE line-height in a Word or LibreOffice clipboard, rewritten as the
 * CSS ratio it actually draws as in those applications.
 *
 * Both write proportional line spacing as a percentage — Word's "1.5 lines" and
 * LibreOffice's "Proportional 150%" both export `line-height:150%` — and both
 * mean it as a multiple of the FONT'S NATURAL LINE. A browser reads the same
 * declaration as a multiple of the FONT SIZE, which on Calibri is 17% tighter.
 * So a report typed at 1.5 lines opened here at what Word calls 1.23, and the
 * difference grew with every line of a table. See lib/line-gap.ts.
 *
 * Only percentages, and only from those two applications:
 *
 * - A percentage is what Word and LibreOffice write for proportional spacing,
 *   and a BARE ratio is what this editor's own serializer writes. Converting
 *   bare ratios too would multiply a report a second time every time somebody
 *   copied a table out of this editor and pasted it back.
 * - `normal` is left alone because it already agrees: it is the font's natural
 *   line in CSS, and that is exactly what "Single" means in both applications.
 * - An absolute leading (`line-height:12.0pt`, which Word writes for "Exactly"
 *   spacing beside `mso-line-height-rule:exactly`) is left alone because it is
 *   absolute in both models — there is no multiplier in it to correct.
 */
import { UNIT_TO_PT } from "@/lib/font-size";
import { DEFAULT_PASTED_LINE_GAP, naturalLineHeight } from "@/lib/line-gap";
import { matchRtfParagraphs, readRtfParagraphs, type RtfLineSpacing, type RtfParagraph } from "@/lib/rtf-paragraphs";

/** LibreOffice and OpenOffice name themselves in a generator `<meta>`. */
export function isLibreOfficeClipboard(html: string): boolean {
  return /content=["'][^"']*LibreOffice/i.test(html) || /content=["'][^"']*OpenOffice/i.test(html);
}

/**
 * Is this clipboard Word's or LibreOffice's?
 *
 * Word announces itself in the Office namespaces and in the `mso-` properties
 * it writes on nearly every element; LibreOffice in its generator meta tag.
 * Anything else — a web page, another report page, this editor itself — states
 * CSS and is meant as CSS.
 */
export function isOfficeClipboard(html: string): boolean {
  return (
    /urn:schemas-microsoft-com:office/i.test(html) ||
    /\bmso-[a-z-]+\s*:/i.test(html) ||
    /class=["']?Mso/i.test(html) ||
    isLibreOfficeClipboard(html)
  );
}

/** The font in effect on an element: its own, else the nearest ancestor's. */
function effectiveFontFamily(el: HTMLElement): string | undefined {
  for (let node: HTMLElement | null = el; node; node = node.parentElement) {
    const stated = node.style?.fontFamily;
    if (stated) return stated;
  }

  // Nothing above says anything, so ask the text itself — Word states the font
  // on the run when the paragraph carries no style of its own.
  const run = el.querySelector<HTMLElement>("[style*='font-family']");
  return run?.style.fontFamily || undefined;
}

/** The blocks a line gap can sit on — anything that is a line of text. */
const TEXT_BLOCKS = "p, h1, h2, h3, h4, h5, h6, li";

/**
 * What tells us a cell already has a block of its own to put a gap on — or
 * something that is not a line of text at all and must not be wrapped in one.
 */
const CELL_BLOCKS = `${TEXT_BLOCKS}, div, table, ul, ol, blockquote, pre, img`;

/**
 * Give a cell's bare text a paragraph to live in.
 *
 * A gap can only reach this editor on a BLOCK — the only node in a table that
 * carries one — and a spreadsheet writes none: Excel's and Calc's clipboards
 * put the text straight into the `<td>`, so a pasted spreadsheet had nowhere to
 * put a line gap and every row came out at the cell's own `normal`, a gap of 1,
 * while the paragraphs around the table took the document's spacing.
 *
 * It costs nothing in structure: Plate wraps a cell's loose text in a paragraph
 * regardless, so this is the same node tree either way — with the difference
 * that this one can be told how far apart its lines go.
 */
function wrapLooseCellText(doc: Document): boolean {
  let wrapped = false;

  doc.body.querySelectorAll<HTMLElement>("td, th").forEach((cell) => {
    if (cell.querySelector(CELL_BLOCKS)) return;
    if (!/\S/.test(cell.textContent ?? "")) return;

    const p = doc.createElement("p");
    while (cell.firstChild) p.append(cell.firstChild);
    cell.append(p);
    wrapped = true;
  });

  return wrapped;
}

/**
 * The line gap in force on an element: its own, else the nearest ancestor's.
 *
 * `line-height` INHERITS, and a table does not break that chain the way it
 * breaks `text-align`'s — so a gap stated on a `<td>`, a `<tr>`, the `<table>`
 * or a `<div>` wrapping a cell's paragraph is the gap the document draws that
 * paragraph at, exactly as if it were stated on the paragraph itself. A
 * `<style>` rule Juice inlines onto the cell is the same thing again.
 *
 * `normal` does not count. It is a real declaration, but it says "ask the font"
 * rather than naming a gap, and a document that only says that is a document
 * whose gap was not collected — which is the case the default is for. Reported
 * 2026-09-16: "doc file any line gap table copy and after paste show our editor
 * still line gap 1".
 */
function statedLineHeight(el: HTMLElement): string | null {
  for (let node: HTMLElement | null = el; node; node = node.parentElement) {
    const stated = node.style?.lineHeight?.trim();
    if (stated && stated.toLowerCase() !== "normal") return stated;
  }
  return null;
}

/** A percentage line-height as a multiple, or null when it is anything else. */
function statedPercentage(value: string): number | null {
  const match = value.trim().match(/^([\d.]+)\s*%$/);
  if (!match) return null;

  const n = Number.parseFloat(match[1]);
  if (!Number.isFinite(n) || n <= 0) return null;

  const gap = n / 100;
  // Beyond this is not spacing anybody set, it is a broken document.
  return gap > 0 && gap <= 10 ? gap : null;
}

/**
 * A gap, as the CSS ratio that draws it against this element's own font — or,
 * `fromRun`, against the font its TEXT is set in, which is the one a line gap is
 * a multiple of. LibreOffice's `<p>` often says its style's family while its runs
 * say another; see `runStyle`.
 */
function asRatio(gap: number, el: HTMLElement, fromRun = false): string {
  return ratioString(gap * naturalLineHeight(fromRun ? runStyle(el, "fontFamily") : effectiveFontFamily(el)));
}

function ratioString(ratio: number): string {
  return String(Math.round(ratio * 10000) / 10000);
}

/**
 * A style property as it is in force on a block's TEXT: the innermost value
 * stated around its first character, else the block's own or an ancestor's.
 *
 * The text is what a line gap is a multiple of. A LibreOffice cell's `<p>` says
 * its style's family (`p.western` — Liberation Serif) while its runs say the
 * document's (`<font face="Calibri">`), and Plate keeps only the run's.
 */
function runStyle(block: HTMLElement, property: "fontFamily" | "fontSize"): string | undefined {
  const walker = block.ownerDocument.createTreeWalker(block, NodeFilter.SHOW_TEXT);
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    if (!/\S/.test(node.textContent ?? "")) continue;
    for (let el = node.parentElement; el && el !== block; el = el.parentElement) {
      if (el.style?.[property]) return el.style[property];
    }
    break;
  }
  for (let el: HTMLElement | null = block; el; el = el.parentElement) {
    if (el.style?.[property]) return el.style[property];
  }
  return undefined;
}

function lengthInPt(value: string | undefined): number | null {
  const match = /^(\d*\.?\d+)(pt|px|pc|in|cm|mm)$/i.exec(String(value ?? "").trim());
  const pt = match ? Number.parseFloat(match[1]) * (UNIT_TO_PT[match[2].toLowerCase()] ?? 0) : 0;
  return pt > 0 ? pt : null;
}

/**
 * An RTF line spacing as the `line-height` that draws it here, or null when it
 * cannot be worked out. "Exactly" is absolute in both models, so it stays in pt.
 */
function lineHeightFromRtf(spacing: RtfLineSpacing, block: HTMLElement): string | null {
  if (spacing.rule === "exact") return spacing.pt > 0 ? `${Math.round(spacing.pt * 100) / 100}pt` : null;

  const natural = naturalLineHeight(runStyle(block, "fontFamily"));
  if (spacing.rule === "multiple") {
    // Beyond this is not spacing anybody set, it is a broken document.
    return spacing.lines > 0 && spacing.lines <= 10 ? ratioString(spacing.lines * natural) : null;
  }

  // "At least" against the size the text is set in: never tighter than single.
  const sizePt = lengthInPt(runStyle(block, "fontSize"));
  return sizePt ? ratioString(Math.max(natural, spacing.pt / sizePt)) : null;
}

/**
 * A LibreOffice table's line gap, read from the clipboard's RTF.
 *
 * LibreOffice's HTML writer leaves the line spacing off every paragraph inside a
 * table cell — a double-spaced table and a single-spaced one copy as the same
 * HTML — and what Juice inlines there instead is the document's default
 * `p { line-height: 115% }`, so every pasted row was drawn at the same gap
 * whatever the document said. The RTF it copies alongside states each one
 * (`\intbl\sl480\slmult1`), so each cell paragraph is matched to its RTF
 * paragraph by text and given that gap, on the font its text is set in.
 *
 * Only a paragraph INSIDE a cell, and only one the RTF can be matched to: the
 * HTML already states the gap of everything else, and whatever does not match
 * is left as it was. Returns the input untouched — the same string — when
 * nothing changes.
 */
export function inlineLibreOfficeCellLineGap(html: string, rtf: string | null | undefined): string {
  const paragraphs = readRtfParagraphs(rtf);
  if (!html || !paragraphs.length) return html;

  const doc = new DOMParser().parseFromString(html, "text/html");
  if (!doc.body) return html;

  // Every block, not just the cells', so the walk through the RTF stays in step.
  const blocks = Array.from(doc.body.querySelectorAll<HTMLElement>(TEXT_BLOCKS));
  const rtfOf = matchRtfParagraphs(blocks, paragraphs);
  matchEmptyLinesByPosition(blocks, paragraphs, rtfOf);
  const cellBlocks = blocks.filter((block) => block.closest("td, th"));
  let changed = false;

  for (const block of cellBlocks) {
    const paragraph = rtfOf.get(block) ?? nearestMatched(block, cellBlocks, rtfOf);
    const lineHeight = paragraph ? lineHeightFromRtf(paragraph.lineSpacing, block) : null;
    if (lineHeight) {
      block.style.lineHeight = lineHeight;
      changed = true;
      continue;
    }
    // Nothing in its table could be placed: it keeps what it was drawn at before,
    // spelled as the bare ratio the percentage drew as, so that the conversion
    // `inlineWordLineGap` makes next leaves it alone.
    const percent = /^([\d.]+)%$/.exec(block.style.lineHeight.trim());
    if (!percent) continue;
    block.style.lineHeight = String(Number.parseFloat(percent[1]) / 100);
    changed = true;
  }

  return changed ? doc.body.innerHTML : html;
}

/**
 * An EMPTY line has no text to be matched by, so it takes the RTF paragraph in
 * the same place between the two matched lines around it — only where both
 * sides hold the same number of lines there, and only an empty one for an empty
 * one. Anything less certain stays unmatched.
 */
function matchEmptyLinesByPosition(
  blocks: HTMLElement[],
  paragraphs: RtfParagraph[],
  matched: Map<HTMLElement, RtfParagraph>
): void {
  const indexOf = new Map(paragraphs.map((paragraph, i) => [paragraph, i]));
  const anchors: Array<[number, number]> = [[-1, -1]];
  blocks.forEach((block, i) => {
    const paragraph = matched.get(block);
    if (paragraph) anchors.push([i, indexOf.get(paragraph) ?? -1]);
  });
  anchors.push([blocks.length, paragraphs.length]);

  for (let a = 1; a < anchors.length; a++) {
    const [fromBlock, fromParagraph] = anchors[a - 1];
    const [toBlock, toParagraph] = anchors[a];
    const between = blocks.slice(fromBlock + 1, toBlock);
    const betweenRtf = paragraphs.slice(fromParagraph + 1, toParagraph);
    if (!between.length || between.length !== betweenRtf.length) continue;
    between.forEach((block, k) => {
      if (/\S/.test((block.textContent ?? "").replace(/\u00a0/g, " ")) || /\S/.test(betweenRtf[k].text)) return;
      matched.set(block, betweenRtf[k]);
    });
  }
}

/**
 * The RTF paragraph of the matched line nearest an unmatched one: in the same
 * cell first, then anywhere in the same table. A table's lines share their
 * spacing far more often than not, and the document's default `p` rule Juice
 * put there instead is never a cell's.
 */
function nearestMatched(
  block: HTMLElement,
  cellBlocks: HTMLElement[],
  matched: Map<HTMLElement, RtfParagraph>
): RtfParagraph | undefined {
  const at = cellBlocks.indexOf(block);
  for (const scope of [block.closest("td, th"), block.closest("table")]) {
    for (let distance = 1; distance < cellBlocks.length; distance++) {
      for (const i of [at - distance, at + distance]) {
        const other = cellBlocks[i];
        if (other && scope?.contains(other) && matched.has(other)) return matched.get(other);
      }
    }
  }
  return undefined;
}

/**
 * Give every block in an Office clipboard the gap the document draws it at,
 * written as the CSS ratio that draws the same line here. Any other HTML is
 * returned untouched, byte for byte.
 *
 * Four passes, in this order because each needs the one before it: a cell's
 * loose text is given a paragraph to carry a gap at all, an inherited gap is
 * brought down onto that paragraph, every proportional gap is converted, and
 * whatever still has none takes the default.
 *
 * `libreOffice` says the RAW clipboard was LibreOffice's. It has to be told:
 * LibreOffice always sends RTF, the RTF makes the docx cleaner run, and the
 * cleaner removes the generator `<meta>` — so by now this HTML carries no mark
 * of it, and a paragraph's `line-height: 100%` was left to mean CSS's 1 x the
 * font size: 0.73 of a line in the editor's face, which "line spacing 1" then
 * visibly opened up. Its gaps are measured against the font the text is set in.
 */
export function inlineWordLineGap(html: string, { libreOffice = false }: { libreOffice?: boolean } = {}): string {
  if (!html || (!libreOffice && !isOfficeClipboard(html))) return html;

  const doc = new DOMParser().parseFromString(html, "text/html");
  if (!doc.body) return html;

  let changed = wrapLooseCellText(doc);

  // THEN bring an inherited gap DOWN onto the block that has to carry it.
  //
  // Only a block reaches this editor as a node with a line-height on it: a
  // `<td>`'s style is read for its padding and its borders and for nothing
  // else, and a wrapper `<div>` becomes no node at all. So a table whose gap
  // is stated on the cell — which is what a `<style>` rule like
  // `td {line-height:150%}` becomes once Juice has inlined it — arrived with
  // the gap on an element that drops it, and the paragraph inside was drawn at
  // the cell's own `normal`: a gap of 1, whatever the document said. It was
  // doubly lost, because an ancestor stating a gap also stopped the default
  // below from filling one in.
  //
  // Done before the conversion, not after, so that a percentage is turned into
  // a ratio against the font THIS paragraph is set in rather than the first one
  // found under the table.
  doc.body.querySelectorAll<HTMLElement>(TEXT_BLOCKS).forEach((el) => {
    const own = el.style?.lineHeight?.trim().toLowerCase();
    if (own && own !== "normal") return;

    const inherited = el.parentElement ? statedLineHeight(el.parentElement) : null;
    if (inherited === null) return;

    el.style.lineHeight = inherited;
    changed = true;
  });

  doc.body.querySelectorAll<HTMLElement>("[style*='line-height']").forEach((el) => {
    const gap = statedPercentage(el.style.lineHeight || "");
    if (gap === null) return;

    el.style.lineHeight = asRatio(gap, el, libreOffice);
    changed = true;
  });

  // Whatever is left without a usable gap takes the default, so a pasted report
  // has ONE spacing rather than falling back to whatever the editor happens to
  // draw an unstated block at. `normal` counts as "no gap collected": it names
  // no multiple, and a document that only says that used to open at a gap of 1.
  doc.body.querySelectorAll<HTMLElement>(TEXT_BLOCKS).forEach((el) => {
    if (statedLineHeight(el) !== null) return;
    el.style.lineHeight = asRatio(DEFAULT_PASTED_LINE_GAP, el, libreOffice);
    changed = true;
  });

  return changed ? doc.body.innerHTML : html;
}

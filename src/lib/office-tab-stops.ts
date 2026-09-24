import { UNIT_TO_PT } from "./font-size";
import { matchRtfParagraphs, readRtfParagraphs, type RtfParagraph } from "./rtf-paragraphs";

/**
 * A paste from Word or LibreOffice, laid out on the document's own tab stops.
 *
 * The editor's tab stops are every half inch — Word's and LibreOffice's default
 * (RichTextEdit's `[tab-size:0.5in]`) — so a line that only uses default
 * stops already lands where the document put it. What a document adds on top,
 * CSS cannot express: a ruler stop part-way along (`Status<TAB><TAB>…Alive`
 * relies on one at 1548 twips) and the stop a hanging indent makes at its own
 * edge (`Placenta<TAB>Fundal` jumps straight to 2in). So for every run of tabs
 * this works out where the document puts the text after it, and writes however
 * many editor tabs reach that point on the half-inch grid. Measured, not
 * estimated: each block is laid out offscreen in the editor's own fonts.
 *
 * Where the stops come from: Word's HTML states them (`tab-stops:77.4pt`);
 * LibreOffice's leaves them out and only its RTF has them (`\tx1548`), so an
 * RTF paragraph is matched to its HTML block by text. See lib/rtf-paragraphs.ts.
 *
 * Two more things that stood between the paste and the document:
 *  - A heading here is a Word/LibreOffice paragraph STYLE, and a template often
 *    sets its body lines in "Heading 3". The editor draws a heading its own way
 *    — tighter letter-spacing, weight 500, its own margins — so those lines came
 *    out narrower than the document and their tabs fell short. They become
 *    paragraphs, keeping the document's own font, size and weight on the runs.
 *  - A hanging indent (`margin-left:2in; text-indent:-2in`) was read as a plain
 *    left indent, moving the whole line 2in right. The block now starts where
 *    the document starts its first line.
 *
 * Right, centre and decimal stops go to the editor stop NEAREST where the
 * document starts the text; a left-aligned grid cannot line those up exactly.
 */

/** Word's default tab stop, when the clipboard does not state its own. */
export const OFFICE_DEFAULT_TAB_STOP_PT = 36;

/** The editable's `tab-size`, for when it cannot be asked. */
const EDITOR_TAB_SIZE = "0.5in";

const PX_PER_PT = 96 / 72;

/** Sub-pixel slack when a measured position is compared with a target. */
const TOLERANCE_PX = 0.5;

/** A run that needs more than this is not lining anything up; it is left alone. */
const MAX_TABS = 64;

export type TabAlignment = "left" | "center" | "right" | "decimal";

export interface OfficeTabStop {
  /** From the left edge of the text column, in pt. */
  positionPt: number;
  alignment: TabAlignment;
}

const ALIGNMENTS: ReadonlySet<string> = new Set(["left", "center", "right", "decimal"]);

const LENGTH = /^(-?\d*\.?\d+)(pt|px|pc|in|cm|mm)$/i;

function lengthToPt(value: string | null | undefined): number {
  const match = LENGTH.exec(String(value ?? "").trim());
  if (!match) return 0;
  const pt = Number.parseFloat(match[1]) * (UNIT_TO_PT[match[2].toLowerCase()] ?? 0);
  return Number.isFinite(pt) ? pt : 0;
}

/**
 * The stops a Word paragraph sets, from its `tab-stops` property.
 *
 * `tab-stops:108.0pt right dotted 450.0pt` — an optional alignment and leader
 * before each position. Leaders, `list` and anything unrecognised are skipped;
 * a `bar` stop draws a rule and is not somewhere a tab goes.
 */
export function parseTabStops(value: string | null | undefined): OfficeTabStop[] {
  const stops: OfficeTabStop[] = [];
  let alignment: TabAlignment = "left";
  let isBar = false;

  for (const token of String(value ?? "").trim().toLowerCase().split(/\s+/)) {
    if (ALIGNMENTS.has(token)) {
      alignment = token as TabAlignment;
      continue;
    }
    if (token === "bar") {
      isBar = true;
      continue;
    }
    if (!LENGTH.test(token)) continue;

    const positionPt = lengthToPt(token);
    if (!isBar && positionPt > 0) stops.push({ positionPt, alignment });
    alignment = "left";
    isBar = false;
  }

  return stops.sort((a, b) => a.positionPt - b.positionPt);
}

/**
 * The stop a tab at `fromPt` takes the text to in Word.
 *
 * The paragraph's own stops first. Past the last of them the default stops take
 * over — only past it, because a set stop clears every default stop before it.
 */
export function nextOfficeTabStop(
  fromPt: number,
  stops: readonly OfficeTabStop[],
  defaultStopPt: number = OFFICE_DEFAULT_TAB_STOP_PT
): OfficeTabStop {
  const epsilon = 0.01;
  const custom = stops.find((stop) => stop.positionPt > fromPt + epsilon);
  if (custom) return custom;

  const lastCustomPt = stops.length ? stops[stops.length - 1].positionPt : 0;
  const from = Math.max(fromPt, lastCustomPt);
  const step = defaultStopPt > 0 ? defaultStopPt : OFFICE_DEFAULT_TAB_STOP_PT;
  return { positionPt: (Math.floor((from + epsilon) / step) + 1) * step, alignment: "left" };
}

/**
 * Clipboard HTML from an office suite. Only those are laid out here: a tab from
 * anywhere else was not typed against a document's stops.
 */
const OFFICE_MARKERS: readonly RegExp[] = [
  // Word: its namespace, its paragraph classes and its `mso-` properties.
  /urn:schemas-microsoft-com:office/i,
  /class=["']?Mso/,
  /mso-[a-z-]+\s*:/i,
  // LibreOffice and OpenOffice name themselves in a generator <meta>.
  /<meta[^>]*content=["']?(?:LibreOffice|OpenOffice)/i,
];

export function isOfficeClipboard(html: string | null | undefined): boolean {
  const source = String(html ?? "");
  return OFFICE_MARKERS.some((marker) => marker.test(source));
}

/** Blocks that are a line — or lines — of the document. */
const LINE_BLOCKS = "p, h1, h2, h3, h4, h5, h6, li";

const HEADINGS = "h1, h2, h3, h4, h5, h6";

/** What a tab's position is measured within, in the editor and on the page. */
const BLOCK_SELECTOR = "p, h1, h2, h3, h4, h5, h6, li, td, th, blockquote, pre, div, body";

/** Font properties that decide how wide a run of text is. */
const FONT_PROPERTIES = [
  "font-family",
  "font-size",
  "font-weight",
  "font-style",
  "letter-spacing",
  "word-spacing",
] as const;

function textNodesOf(root: Node): Text[] {
  const walker = (root.ownerDocument ?? (root as Document)).createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const nodes: Text[] = [];
  for (let node = walker.nextNode(); node; node = walker.nextNode()) nodes.push(node as Text);
  return nodes;
}

/** The nearest inline value of `property` on `element` or an ancestor of it. */
function inheritedInlineStyle(element: Element | null, property: string): string {
  for (let el = element; el; el = el.parentElement) {
    const value = (el as HTMLElement).style?.getPropertyValue(property);
    if (value) return value;
  }
  return "";
}

/** A property the browser's CSSOM would throw away — Word's `tab-stops`. */
function rawStyleProperty(element: Element, property: string): string {
  const match = new RegExp(`(?:^|;)\\s*${property}\\s*:\\s*([^;]+)`, "i").exec(
    element.getAttribute("style") ?? ""
  );
  return match ? match[1].trim() : "";
}

function charRect(text: Text, index: number): DOMRect {
  const range = text.ownerDocument.createRange();
  range.setStart(text, index);
  range.setEnd(text, index + 1);
  return range.getBoundingClientRect();
}

function sameLine(a: DOMRect, b: DOMRect): boolean {
  return a.top < b.bottom - 1 && b.top < a.bottom - 1;
}

/** An offscreen box in the editable's font and with its `tab-size`. */
function createHost(editable: HTMLElement | null): HTMLElement {
  const host = document.createElement("div");
  host.setAttribute("aria-hidden", "true");
  host.style.cssText =
    "position:absolute;left:-99999px;top:0;visibility:hidden;width:max-content;" +
    "margin:0;padding:0;border:0";

  if (editable) {
    const computed = getComputedStyle(editable);
    for (const property of [...FONT_PROPERTIES, "line-height", "tab-size"]) {
      host.style.setProperty(property, computed.getPropertyValue(property));
    }
  }
  if (!host.style.getPropertyValue("tab-size")) host.style.setProperty("tab-size", EDITOR_TAB_SIZE);

  document.body.appendChild(host);
  return host;
}

/** A browser that lays nothing out (jsdom) has no stops to count against. */
function hasLayout(host: HTMLElement): boolean {
  const probe = document.createElement("span");
  probe.textContent = "M";
  host.appendChild(probe);
  const width = probe.getBoundingClientRect().width;
  probe.remove();
  return width > 0;
}

interface Segment {
  /** The text after a run, up to the next tab or the end of the line. */
  widthPx: number;
  /** Up to its first `.`, for a decimal stop. */
  toDecimalPx: number | null;
}

/** How wide the text that follows a tab run is — what a right stop aligns. */
function measureSegment(texts: Text[], index: number, offset: number, line: DOMRect): Segment {
  const range = texts[index].ownerDocument.createRange();
  range.setStart(texts[index], offset);

  let endNode = texts[texts.length - 1];
  let endOffset = endNode.data.length;
  let decimal: { node: Text; offset: number } | null = null;

  scan: for (let i = index; i < texts.length; i++) {
    const data = texts[i].data;
    for (let j = i === index ? offset : 0; j < data.length; j++) {
      if (data[j] === "\t" || data[j] === "\n") {
        endNode = texts[i];
        endOffset = j;
        break scan;
      }
      if (data[j] === "." && !decimal) decimal = { node: texts[i], offset: j };
    }
  }

  const widthOf = (node: Text, end: number): number => {
    range.setEnd(node, end);
    const rects = Array.from(range.getClientRects()).filter((rect) => sameLine(rect, line));
    if (!rects.length) return 0;
    return Math.max(...rects.map((r) => r.right)) - Math.min(...rects.map((r) => r.left));
  };

  return {
    widthPx: widthOf(endNode, endOffset),
    toDecimalPx: decimal ? widthOf(decimal.node, decimal.offset) : null,
  };
}

/**
 * The number of editor tabs that takes the text to `targetPx`: the first stop at
 * or past it, or for a right/centre/decimal stop the nearest one.
 */
function fitTabs(
  text: Text,
  start: number,
  count: number,
  targetPx: number,
  nearest: boolean,
  boxLeft: number
): { count: number; endPx: number } {
  const original = text.data;
  const head = original.slice(0, start);
  const tail = original.slice(start + count);
  const endWith = (tabs: number): number => {
    text.data = head + "\t".repeat(tabs) + tail;
    return charRect(text, start + tabs - 1).right - boxLeft;
  };

  let previousEnd = Number.NEGATIVE_INFINITY;
  for (let tabs = 1; tabs <= MAX_TABS; tabs++) {
    const end = endWith(tabs);
    if (end < targetPx - TOLERANCE_PX) {
      previousEnd = end;
      continue;
    }
    if (nearest && tabs > 1 && targetPx - previousEnd < end - targetPx) {
      return { count: tabs - 1, endPx: endWith(tabs - 1) };
    }
    return { count: tabs, endPx: end };
  }

  text.data = original;
  return { count, endPx: targetPx };
}

/** Where a block's text column starts, and how it indents, in the document. */
interface BlockGeometry {
  stops: OfficeTabStop[];
  /** The left indent. */
  marginPt: number;
  /** The first line's indent relative to it; negative for a hanging indent. */
  firstIndentPt: number;
}

function geometryOf(block: HTMLElement, rtf: RtfParagraph | undefined): BlockGeometry {
  const cssStops = rawStyleProperty(block, "tab-stops");
  const stops = cssStops ? parseTabStops(cssStops) : [...(rtf?.stops ?? [])];
  const hasCssIndent = !!(block.style.marginLeft || block.style.textIndent);
  const marginPt = hasCssIndent ? lengthToPt(block.style.marginLeft) : (rtf?.leftPt ?? 0);
  const firstIndentPt = hasCssIndent ? lengthToPt(block.style.textIndent) : (rtf?.firstPt ?? 0);

  // A hanging indent is a stop of its own, on its first line: Word takes the
  // first tab there, past every default stop before it.
  if (firstIndentPt < 0 && marginPt > 0 && !stops.some((stop) => Math.abs(stop.positionPt - marginPt) < 0.5)) {
    stops.push({ positionPt: marginPt, alignment: "left" });
    stops.sort((a, b) => a.positionPt - b.positionPt);
  }

  return { stops, marginPt, firstIndentPt };
}

/**
 * Re-count one block's tab runs on a live copy of it in `host`.
 *
 * `originals` are the block's text nodes in document order. Returns the new text
 * of each one that changed, by its index there — the copy is a deep clone, node
 * for node, so the indexes line up.
 */
function realignBlock(
  block: HTMLElement,
  originals: Text[],
  host: HTMLElement,
  geometry: BlockGeometry,
  defaultStopPt: number
): Map<number, string> {
  const changed = new Map<number, string>();

  const box = document.createElement("div");
  box.style.fontSize = inheritedInlineStyle(block, "font-size");
  // As Plate will read it: the docx cleaner's `pre-wrap` wrapper, or the
  // `pre` spans a plain paste's tabs sit in.
  box.style.whiteSpace = inheritedInlineStyle(block, "white-space") || "normal";
  const inline = document.createElement("span");
  for (const property of FONT_PROPERTIES) {
    const value = inheritedInlineStyle(block, property);
    if (value) inline.style.setProperty(property, value);
  }
  inline.append(...Array.from(block.childNodes, (node) => document.importNode(node, true)));
  box.append(inline);
  host.append(box);

  try {
    const texts = textNodesOf(inline);
    const boxLeft = box.getBoundingClientRect().left;
    const first = texts.find((text) => /\S/.test(text.data));
    const firstLine = first ? charRect(first, first.data.search(/\S/)) : null;
    const { stops, marginPt, firstIndentPt } = geometry;

    let line: DOMRect | null = null;
    /** The document's x minus the editor's, on the current line, in px. */
    let drift = 0;

    for (let i = 0; i < texts.length; i++) {
      if (originals[i].parentElement?.closest(BLOCK_SELECTOR) !== block) continue;
      const text = texts[i];

      for (let from = 0; ; ) {
        const tabRun = /\t+/g;
        tabRun.lastIndex = from;
        const match = tabRun.exec(text.data);
        if (!match) break;

        const start = match.index;
        const count = match[0].length;
        from = start + count;
        // Against a newline: the source file's indentation, not the report's.
        if (/[\n\r]/.test(text.data[start - 1] ?? "") || /[\n\r]/.test(text.data[from] ?? "")) continue;

        const startRect = charRect(text, start);
        if (!line || !sameLine(startRect, line)) drift = 0;
        line = startRect;

        // The editor starts every line at the block's edge; the document starts
        // its first line at the first-line indent, the rest at the left indent.
        const onFirstLine = !!firstLine && sameLine(startRect, firstLine);
        const lineStartPt = marginPt + (onFirstLine ? firstIndentPt : 0);
        const lineStops = onFirstLine ? stops : stops.filter((stop) => stop.positionPt !== marginPt || firstIndentPt >= 0);
        let positionPt = lineStartPt + (startRect.left - boxLeft + drift) / PX_PER_PT;
        let beforeLastPt = positionPt;
        let stop: OfficeTabStop = { positionPt, alignment: "left" };
        for (let k = 0; k < count; k++) {
          beforeLastPt = positionPt;
          stop = nextOfficeTabStop(positionPt, lineStops, defaultStopPt);
          positionPt = stop.positionPt;
        }

        if (stop.alignment !== "left") {
          const segment = measureSegment(texts, i, from, startRect);
          const leadPx =
            stop.alignment === "center"
              ? segment.widthPx / 2
              : stop.alignment === "decimal"
                ? (segment.toDecimalPx ?? segment.widthPx)
                : segment.widthPx;
          positionPt = Math.max(beforeLastPt, positionPt - leadPx / PX_PER_PT);
        }

        const targetPx = (positionPt - lineStartPt) * PX_PER_PT;
        const fitted = fitTabs(text, start, count, targetPx, stop.alignment !== "left", boxLeft);
        drift = targetPx - fitted.endPx;
        from = start + fitted.count;
      }

      if (text.data !== originals[i].data) changed.set(i, text.data);
    }
  } finally {
    box.remove();
  }

  return changed;
}

/**
 * A hanging-indent block, started where the document starts its first line.
 *
 * Its left indent was read as an indent for the WHOLE block, so `Placenta`
 * printed 2in in from where the document prints it. The block keeps whatever
 * the first line's own offset is (`margin + text-indent`, never below zero).
 */
function startAtFirstLine(block: HTMLElement): boolean {
  const firstIndentPt = lengthToPt(block.style.textIndent);
  if (firstIndentPt >= 0) return false;
  const firstLinePt = Math.max(0, lengthToPt(block.style.marginLeft) + firstIndentPt);
  block.style.marginLeft = firstLinePt ? `${firstLinePt}pt` : "";
  block.style.textIndent = "";
  return true;
}

/** A heading element as the paragraph it is in the document. */
function asParagraph(heading: HTMLElement): HTMLElement {
  const paragraph = heading.ownerDocument.createElement("p");
  for (const attribute of Array.from(heading.attributes)) {
    paragraph.setAttribute(attribute.name, attribute.value);
  }
  paragraph.append(...Array.from(heading.childNodes));
  heading.replaceWith(paragraph);
  return paragraph;
}

export interface OfficeLayoutOptions {
  /** The editor's contenteditable, whose font and `tab-size` are measured in. */
  editable?: HTMLElement | null;
  /** The clipboard's `text/rtf`, for the stops LibreOffice's HTML leaves out. */
  rtf?: string | null;
}

/**
 * The clipboard's HTML laid out on the document's tab stops. See the top of
 * this file. Returns the input untouched — the same string — when there is
 * nothing to change; the tab counts are only re-worked where text is laid out.
 */
export function layOutOfficePaste(html: string, { editable = null, rtf = null }: OfficeLayoutOptions = {}): string {
  if (!html || typeof document === "undefined" || !document.body) return html;

  const doc = new DOMParser().parseFromString(html, "text/html");
  let changed = false;

  // Stops and indents are read BEFORE either is rewritten below.
  const lineBlocks = Array.from(doc.body.querySelectorAll<HTMLElement>(LINE_BLOCKS));
  const rtfParagraphs = readRtfParagraphs(rtf);
  const rtfOf = matchRtfParagraphs(lineBlocks, rtfParagraphs);
  const defaultStopPt = defaultTabOf(rtf);
  const geometries = new Map<HTMLElement, BlockGeometry>(
    lineBlocks.map((block) => [block, geometryOf(block, rtfOf.get(block))])
  );

  const tabbed = new Set<HTMLElement>();
  for (const text of textNodesOf(doc.body)) {
    if (!text.data.includes("\t")) continue;
    const block = text.parentElement?.closest(BLOCK_SELECTOR) as HTMLElement | null;
    if (block && block.tagName !== "PRE") tabbed.add(block);
  }

  if (tabbed.size) {
    const host = createHost(editable);
    try {
      if (hasLayout(host)) {
        for (const block of tabbed) {
          const geometry = geometries.get(block) ?? geometryOf(block, undefined);
          const originals = textNodesOf(block);
          for (const [index, data] of realignBlock(block, originals, host, geometry, defaultStopPt)) {
            originals[index].data = data;
            changed = true;
          }
        }
      }
    } finally {
      host.remove();
    }
  }

  for (const block of lineBlocks) {
    if (startAtFirstLine(block)) changed = true;
  }
  for (const heading of Array.from(doc.body.querySelectorAll<HTMLElement>(HEADINGS))) {
    asParagraph(heading);
    changed = true;
  }

  return changed ? doc.documentElement.outerHTML : html;
}

/** The document's default tab stop, from its RTF (`\deftab720`), in pt. */
function defaultTabOf(rtf: string | null | undefined): number {
  const match = /\\deftab(\d+)/.exec(String(rtf ?? "").slice(0, 20000));
  const pt = match ? Number(match[1]) / 20 : 0;
  return pt > 0 ? pt : OFFICE_DEFAULT_TAB_STOP_PT;
}

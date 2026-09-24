import { isDocxContent } from "@platejs/docx";

/**
 * Whitespace that a report LAYS OUT with.
 *
 * A lab report pasted as plain text does its whole layout with horizontal
 * whitespace — tabs to line the values up in a column, a run of spaces to
 * centre a heading, a blank line to open a gap:
 *
 *     Test Name<TAB><TAB>:  RT-PCR FOR COVID-19
 *     Specimen<TAB><TAB>:  Nasopharyngeal Swab
 *     <SPACE>
 *     <77 SPACES>RESULT
 *
 * HTML collapses every one of those to a single space (and drops it entirely at
 * the edges of a block), so the report came out of the serializer as
 * `<p>Test Name\t\t: RT-PCR</p>` — right in the editor, which is
 * `white-space: pre-wrap`, and flattened everywhere the report is actually
 * read: the print page, the PDF, the patient portal. The columns closed up and
 * the blank line disappeared, because `<p> </p>` is an empty block.
 *
 * The reference for what should happen is not a guess. The editor this replaces
 * did nothing at all with a plain-text paste — `summernote_table_resize.blade.php`
 * bails out of its paste handler when the clipboard carries no HTML ("plain-text
 * paste -> let the browser handle it") — so what the old editor produced IS
 * Chrome's own contenteditable paste, and that is measurable. It writes:
 *
 *     <div>Test Name<span style="white-space:pre">\t\t</span>:&nbsp; RT-PCR</div>
 *     <div>&nbsp;</div>
 *     <div>&nbsp; &nbsp; &nbsp;RESULT</div>
 *
 * Tabs preserved by a `white-space: pre` span, and every space run rebalanced
 * so that no two collapsible spaces ever sit next to each other. That is what
 * this module writes on the way out, and it is measured against the real thing
 * in tests/browser/plain-text-paste.mjs.
 *
 * The way IN needs the mirror of it. Plate's HTML deserializer implements CSS
 * whitespace collapsing itself (`collapseString` in @platejs/core), and its
 * regex is JavaScript's `\s`, which — unlike CSS — counts U+00A0 as collapsible.
 * So a non-breaking space arrives at the editor as a plain space, and a run of
 * them as ONE, and a leading run as nothing at all. Every legacy Summernote
 * report indented with `&nbsp;`, every Word paste whose tabs are `mso-tab-count`
 * spans full of them, and anything this serializer had just written, all lost
 * their alignment the moment they were opened. `protectWhitespace` puts those
 * runs inside a `white-space: pre` span before Plate sees them, which is the
 * one rule its collapser does honour.
 */

/** U+00A0 — a space that is not whitespace as far as layout is concerned. */
export const NBSP = "\u00A0";

/** What keeps a run of whitespace intact, in the editor and in the report. */
const PRE = "white-space: pre";

/**
 * Whitespace HTML collapses, minus the newlines.
 *
 * Newlines are deliberately not part of a run here: in the source of a document
 * they are how it is INDENTED, not how it is laid out, and a Word file is one
 * long tree of `\n\t\t<p>`. Protecting those would paste the source's own
 * indentation into the report. See `protectWhitespace`.
 */
const HORIZONTAL_RUN = /[ \t\u00A0]+/g;

/** Elements whose text is not the report's text. */
const NOT_CONTENT = new Set(["SCRIPT", "STYLE", "TITLE", "TEXTAREA", "NOSCRIPT"]);

/** CSS `white-space` values that already keep a run as written. */
const PRESERVING = new Set(["pre", "pre-wrap", "pre-line", "break-spaces"]);

/**
 * The text of one leaf as HTML that keeps its spacing.
 *
 * `escaped` must already be HTML-escaped — this only ever adds markup of its
 * own, and the characters it looks at (space, tab, U+00A0) survive escaping
 * untouched, so the two steps cannot interfere.
 *
 * A space is written as `&nbsp;` exactly when HTML would otherwise lose it:
 * when the previous character is a space that was kept as one, or when it sits
 * against the start or the end of the block, where a collapsible space is
 * dropped outright. Every other space stays a space — that is not tidiness, it
 * is the only place a line is allowed to WRAP, and a paragraph of prose whose
 * spaces were all made non-breaking would run off the side of the page.
 *
 * `atBlockStart` / `atBlockEnd` say whether this leaf is the first / last child
 * of its block. A leaf that follows another one starts mid-line, where a space
 * is safe: `<strong>Result</strong> : Negative` must not gain an `&nbsp;`.
 *
 * `afterSpace` says the leaf before this one ended in a space written as a
 * plain space — so this one's first space would be swallowed by it. A run of
 * spaces split across leaves (`Fundal ` + a bold ` ` + ` Ant`) was three plain
 * spaces in a row on the page, and printed as one.
 */
export function serializeWhitespace(
  escaped: string,
  {
    atBlockStart = true,
    atBlockEnd = true,
    afterSpace = false,
  }: { atBlockStart?: boolean; atBlockEnd?: boolean; afterSpace?: boolean } = {}
): string {
  if (!/[ \t\u00A0]/.test(escaped)) return escaped;

  let out = "";
  /** Would a space written here be swallowed by the one before it? */
  let afterKeptSpace = afterSpace;

  for (let i = 0; i < escaped.length; i++) {
    const char = escaped[i];

    if (char === "\t") {
      // The whole run in one span, the way Chrome writes it: a tab's width is
      // the distance to the next tab stop, so two of them inside one preserved
      // run is not the same as two preserved runs of one.
      let end = i;
      while (end < escaped.length && escaped[end] === "\t") end++;
      out += `<span style="${PRE}">${escaped.slice(i, end)}</span>`;
      i = end - 1;
      afterKeptSpace = false;
      continue;
    }

    if (char === NBSP) {
      out += "&nbsp;";
      afterKeptSpace = false;
      continue;
    }

    if (char === " ") {
      const atEdge = (atBlockStart && i === 0) || (atBlockEnd && i === escaped.length - 1);
      if (afterKeptSpace || atEdge) {
        out += "&nbsp;";
        afterKeptSpace = false;
      } else {
        out += " ";
        afterKeptSpace = true;
      }
      continue;
    }

    out += char;
    afterKeptSpace = false;
  }

  return out;
}

/** Is this node inside something that already keeps its whitespace as written? */
function inPreservingContext(node: Node): boolean {
  let element = node.parentElement;
  while (element) {
    if (NOT_CONTENT.has(element.tagName)) return true;
    if (element.tagName === "PRE") return true;
    if (PRESERVING.has(element.style?.whiteSpace ?? "")) return true;
    element = element.parentElement;
  }
  return false;
}

/**
 * Elements that hold cells and rows, never text.
 *
 * A browser renders nothing for the whitespace between a `<tr>`'s cells, and
 * neither does Word. Plate's deserializer is not so lucky.
 */
const TABLE_STRUCTURE: ReadonlySet<string> = new Set([
  "TABLE",
  "THEAD",
  "TBODY",
  "TFOOT",
  "TR",
  "COLGROUP",
]);

/**
 * The same HTML with the markup whitespace BETWEEN cells and rows removed.
 *
 * A table whose tags are written one per line is fine: a run of whitespace
 * touching a newline is the source file's own indentation and every pass here
 * already skips it. A table written on ONE line is not — Word does that to
 * tables it has re-flowed — and those single spaces between `</td>` and `<td>`
 * become text nodes Plate turns into PARAGRAPHS inside the table and inside the
 * rows.
 *
 * That is not cosmetic. The stray blocks shift every index the table is read
 * by: the first row is no longer `children[0]` and the first cell no longer
 * column 0, so Plate's border emulation draws neither the table's top edge nor
 * its left one, and the column count comes out short. Measured on a patient
 * header table written on one line: one `<col>` where there are two columns,
 * and a first cell with `0/1/1/0` borders where the document has all four. The
 * same table with newlines between its tags: two columns, `1/1/1/1`.
 *
 * Returns the input untouched — the same string, not a re-serialized copy —
 * when there is nothing of the sort to remove.
 */
export function dropTableMarkupWhitespace(html: string): string {
  if (!html || !/<table[\s>]/i.test(html)) return html;

  const doc = new DOMParser().parseFromString(html, "text/html");
  const walker = doc.createTreeWalker(doc.body, NodeFilter.SHOW_TEXT);

  const stray: Text[] = [];
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const text = node as Text;
    const parent = text.parentElement;
    if (!parent || !TABLE_STRUCTURE.has(parent.tagName)) continue;
    if (/\S/.test(text.data)) continue;
    stray.push(text);
  }

  if (stray.length === 0) return html;

  for (const text of stray) text.remove();

  return doc.documentElement.outerHTML;
}

/**
 * What `protectWhitespace` needs to know on the paste path.
 *
 * Plate's docx cleaner runs whenever the clipboard carries RTF — LibreOffice's
 * always does — or the HTML is Word's, and it deletes every element whose
 * `innerHTML.trim()` is empty. A `white-space: pre` span holding only tabs is
 * such an element, so `BPD<TAB><TAB><TAB><TAB>89 mm` pasted out of LibreOffice
 * as `BPD89 mm`: the one thing protecting the tabs was what removed them. When
 * that cleaner is going to run, a tab run is written the way Word writes one
 * instead — an `mso-tab-count` span, which the same cleaner turns back into
 * tabs inside the `pre-wrap` wrapper it puts round everything.
 */
export interface ClipboardWhitespace {
  /** The clipboard carries `text/rtf`, which alone makes the docx cleaner run. */
  hasRtf: boolean;
  /**
   * LibreOffice wrote the HTML. A word processor keeps every space as typed, and
   * LibreOffice's HTML writer writes them as typed too — a run of spaces, not
   * `&nbsp;` — so a line indented with spaces (`<53 spaces>No retro-placental…`)
   * collapsed to nothing. And it wraps long source lines by turning a SPACE into
   * a newline, and puts a newline of its own after a block's opening tag. See
   * `normalizeLibreOfficeNewlines`.
   */
  libreOffice?: boolean;
}

/** Blocks whose inline content is a line of the document. */
const TEXT_BLOCKS = "p, h1, h2, h3, h4, h5, h6, li, dt, dd, td, th";

function textBlockOf(node: Node): Element | null {
  return node.parentElement?.closest(TEXT_BLOCKS) ?? null;
}

/** Elements that hold blocks rather than being a line of text themselves. */
const BLOCK_LEVEL: ReadonlySet<string> = new Set([
  "P", "DIV", "TABLE", "THEAD", "TBODY", "TFOOT", "TR", "TD", "TH", "COLGROUP", "COL", "CAPTION",
  "UL", "OL", "LI", "DL", "DT", "DD", "H1", "H2", "H3", "H4", "H5", "H6",
  "BLOCKQUOTE", "PRE", "CENTER", "HR", "SECTION", "ARTICLE", "HEADER", "FOOTER",
]);

/** A cell or list item that holds paragraphs: its own whitespace is markup. */
function holdsBlocks(element: Element): boolean {
  return Array.from(element.children).some((child) => BLOCK_LEVEL.has(child.tagName));
}

/**
 * How many tabs LibreOffice indents this block's source lines by.
 *
 * Its HTML writer indents by nesting: nothing for a top-level paragraph,
 * `\t\t\t` for one inside a table cell. A wrapped line inside the block
 * carries that indentation after its newline — `Inv. Date:\n\t\t\t13-10-2025`
 * is `Inv. Date: 13-10-2025` — and it is not text. A cell states its own depth
 * on the line it closes on (`</p>\n\t\t</td>`), one less than its content's;
 * failing that, the block's first line says it (`<p>\n\t\t\t<font>`).
 */
function libreOfficeIndent(block: Element, texts: Text[]): number {
  for (let ancestor = block.parentElement; ancestor && ancestor.tagName !== "BODY"; ancestor = ancestor.parentElement) {
    const closing = ancestor.lastChild;
    const match = closing instanceof Text ? /\r?\n(\t*)$/.exec(closing.data) : null;
    if (match) return match[1].length + 1;
  }
  const first = texts.find((text) => text.data);
  const match = first ? /^\r?\n(\t*)/.exec(first.data) : null;
  return match ? match[1].length : 0;
}

/**
 * LibreOffice's newlines, put back to what they stand for.
 *
 * Its HTML writer adds a newline (and the block's indentation) of its own right
 * after a block's opening tag and before its closing one — formatting, not text
 * — and breaks every long source line by REPLACING a space with a newline plus
 * that same indentation: `Number \t\t\tSingle` at the top level is written
 * `Number\n\t\t\tSingle` (the tabs are the document's), while in a cell
 * `Inv. Date: 13-10-2025` is written `Inv. Date:\n\t\t\t13-10-2025` (the tabs
 * are indentation). So the formatting goes, each wrap becomes the space it
 * replaced, and only tabs past the block's indentation stay — after which a
 * block has no newline left for Plate's cleaners to collapse spaces around.
 *
 * A cell or list item that holds paragraphs is left alone: the whitespace
 * between its paragraphs is markup, and Plate already drops it.
 */
function normalizeLibreOfficeNewlines(body: HTMLElement): void {
  for (const block of Array.from(body.querySelectorAll(TEXT_BLOCKS))) {
    if (holdsBlocks(block)) continue;
    const texts = textNodesIn(block).filter((text) => textBlockOf(text) === block && !inPreservingContext(text));
    if (!texts.some((text) => /[\r\n]/.test(text.data))) continue;

    const indent = libreOfficeIndent(block, texts);
    const leadingBreak = new RegExp(`^(?:\\r?\\n\\t{0,${indent}})+`);
    const wrap = new RegExp(`\\r?\\n\\t{0,${indent}}`, "g");

    for (const text of texts) {
      text.data = text.data.replace(leadingBreak, "");
      if (text.data) break;
    }
    for (const text of [...texts].reverse()) {
      text.data = text.data.replace(/(?:\r?\n\t*)+$/, "");
      if (text.data) break;
    }
    for (const text of texts) {
      if (/[\r\n]/.test(text.data)) text.data = text.data.replace(wrap, " ");
    }

    // A line indented with spaces is written with the spaces OUTSIDE the run's
    // `<font face>` — `<h3>     …<font face="Calibri">No retro-placental` — so
    // they took the paragraph's style-block font (Times New Roman, a wider
    // space) and pushed the text ~20px past where the document has it. They
    // belong to the run they indent.
    const leading = block.firstChild;
    const run = leading?.nextSibling;
    if (leading instanceof Text && !leading.data.trim() && leading.data && run instanceof HTMLElement) {
      const firstText = textNodesIn(run).find((text) => text.data);
      if (firstText) {
        firstText.data = leading.data + firstText.data;
        leading.remove();
      }
    }
  }
}

function textNodesIn(root: Node): Text[] {
  const walker = (root.ownerDocument ?? (root as Document)).createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const texts: Text[] = [];
  for (let node = walker.nextNode(); node; node = walker.nextNode()) texts.push(node as Text);
  return texts;
}

/**
 * Whether a run of plain spaces in a LibreOffice block is the document's own
 * spacing: two or more of them, a line's leading indent, or the only content
 * of an inline element (which the docx cleaner would delete as empty).
 */
function isLibreOfficeSpacing(text: Text, start: number, end: number): boolean {
  const block = textBlockOf(text);
  if (!block || holdsBlocks(block)) return false;
  const texts = textNodesIn(block).filter((node) => textBlockOf(node) === block);
  const before = texts.slice(0, texts.indexOf(text)).map((node) => node.data).join("") + text.data.slice(0, start);
  const after = text.data.slice(end) + texts.slice(texts.indexOf(text) + 1).map((node) => node.data).join("");
  if (!after.trim()) return false; // trailing: nothing to push along
  if (!before.trim()) return true; // leading: the line's indent
  if (end - start >= 2) return true;
  return !text.data.trim() && text.parentElement !== block;
}

function preSpan(doc: Document, run: string): HTMLElement {
  const span = doc.createElement("span");
  span.setAttribute("style", PRE);
  span.textContent = run;
  return span;
}

/**
 * A run as Word would write it: each group of tabs an `mso-tab-count` span, the
 * spaces between left as text for the cleaner's `pre-wrap` wrapper to keep.
 */
function asDocxTabs(doc: Document, run: string, soleContent: boolean): Node[] {
  return (run.match(/\t+| +/g) ?? []).map((part) => {
    if (part[0] !== "\t") {
      // The cleaner's `pre-wrap` wrapper keeps spaces as they are — unless they
      // are all an element holds, when it deletes the element. Then they
      // alternate with `&nbsp;` (it also deletes a node of nothing but those).
      return doc.createTextNode(soleContent ? part.replace(/ /g, (_, i: number) => (i % 2 ? " " : NBSP)) : part);
    }
    const span = doc.createElement("span");
    span.setAttribute("style", `mso-tab-count:${part.length}`);
    span.textContent = NBSP;
    return span;
  });
}

/**
 * The same HTML with every run of whitespace the document LAID OUT with wrapped
 * in a `white-space: pre` span, so Plate's deserializer leaves it alone.
 *
 * Only runs that contain a tab or a non-breaking space are protected. A run of
 * ordinary spaces means one space in HTML — that is what a browser showed the
 * author, and inventing a gap where the document only had source formatting
 * would be worse than losing one. A run touching a newline is skipped for the
 * same reason: it is how the file is indented, not what the report says.
 *
 * The whole document comes back, `<head>` and all, for the reason
 * `inlineLegacyAlignment` returns it: Word ships its formatting as a `<style>`
 * block that JuicePlugin inlines later, and returning the body alone would drop
 * every font, size and column width with it.
 *
 * `clipboard` is for the paste path only, where two more things are true — see
 * `ClipboardWhitespace`. The load path and the Word-upload button leave it out
 * and get exactly the behaviour above.
 */
export function protectWhitespace(html: string, clipboard?: ClipboardWhitespace): string {
  // Markup whitespace inside a table goes first: it is not the report's, it is
  // not protectable, and left alone it breaks the table. See
  // `dropTableMarkupWhitespace` — a no-op, and byte-identical, for everything
  // that does not have it.
  const source = dropTableMarkupWhitespace(html);

  // The entity forms count as well as the character: `&nbsp;` is how legacy
  // Summernote reports — the ones this is here for — write every one of theirs.
  // A LibreOffice paste has its space runs and newlines to see to as well.
  const libreOffice = !!clipboard?.libreOffice;
  if (!source || (!libreOffice && !/[\t\u00A0]|&nbsp;|&#0*160;|&#x0*a0;/i.test(source))) return source;

  const doc = new DOMParser().parseFromString(source, "text/html");
  if (libreOffice) normalizeLibreOfficeNewlines(doc.body);
  const walker = doc.createTreeWalker(doc.body, NodeFilter.SHOW_TEXT);
  const docxCleanerRuns = !!clipboard && (clipboard.hasRtf || isDocxContent(doc.body));


  const texts: Text[] = [];
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    texts.push(node as Text);
  }

  for (const text of texts) {
    const value = text.data;
    if (!libreOffice && !/[\t\u00A0]/.test(value)) continue;
    if (inPreservingContext(text)) continue;

    const pieces: Node[] = [];
    let cursor = 0;
    HORIZONTAL_RUN.lastIndex = 0;

    for (let match = HORIZONTAL_RUN.exec(value); match; match = HORIZONTAL_RUN.exec(value)) {
      const run = match[0];
      const start = match.index;
      const end = start + run.length;
      const spacing = libreOffice && !/[\t\u00A0]/.test(run) && isLibreOfficeSpacing(text, start, end);
      if (!/[\t\u00A0]/.test(run) && !spacing) continue;
      // Against a newline: source indentation, not the report's own spacing.
      // (A LibreOffice block has none left by now; see normalizeLibreOfficeNewlines.)
      if (/[\n\r]/.test(value[start - 1] ?? "") || /[\n\r]/.test(value[end] ?? "")) continue;

      if (start > cursor) pieces.push(doc.createTextNode(value.slice(cursor, start)));
      const soleContent = !(text.parentElement?.textContent ?? "").replace(/[ \t\u00A0]/g, "");
      pieces.push(...(docxCleanerRuns && !run.includes(NBSP) ? asDocxTabs(doc, run, soleContent) : [preSpan(doc, run)]));
      cursor = end;
    }

    if (!pieces.length) continue;
    if (cursor < value.length) pieces.push(doc.createTextNode(value.slice(cursor)));
    text.replaceWith(...pieces);
  }

  return doc.documentElement.outerHTML;
}

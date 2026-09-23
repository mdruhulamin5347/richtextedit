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
 */
export function serializeWhitespace(
  escaped: string,
  { atBlockStart = true, atBlockEnd = true }: { atBlockStart?: boolean; atBlockEnd?: boolean } = {}
): string {
  if (!/[ \t\u00A0]/.test(escaped)) return escaped;

  let out = "";
  /** Would a space written here be swallowed by the one before it? */
  let afterKeptSpace = false;

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
 */
export function protectWhitespace(html: string): string {
  // Markup whitespace inside a table goes first: it is not the report's, it is
  // not protectable, and left alone it breaks the table. See
  // `dropTableMarkupWhitespace` — a no-op, and byte-identical, for everything
  // that does not have it.
  const source = dropTableMarkupWhitespace(html);

  // The entity forms count as well as the character: `&nbsp;` is how legacy
  // Summernote reports — the ones this is here for — write every one of theirs.
  if (!source || !/[\t\u00A0]|&nbsp;|&#0*160;|&#x0*a0;/i.test(source)) return source;

  const doc = new DOMParser().parseFromString(source, "text/html");
  const walker = doc.createTreeWalker(doc.body, NodeFilter.SHOW_TEXT);

  const texts: Text[] = [];
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    texts.push(node as Text);
  }

  for (const text of texts) {
    const value = text.data;
    if (!/[\t\u00A0]/.test(value)) continue;
    if (inPreservingContext(text)) continue;

    const pieces: Node[] = [];
    let cursor = 0;
    HORIZONTAL_RUN.lastIndex = 0;

    for (let match = HORIZONTAL_RUN.exec(value); match; match = HORIZONTAL_RUN.exec(value)) {
      const run = match[0];
      const start = match.index;
      const end = start + run.length;
      if (!/[\t\u00A0]/.test(run)) continue;
      // Against a newline: source indentation, not the report's own spacing.
      if (/[\n\r]/.test(value[start - 1] ?? "") || /[\n\r]/.test(value[end] ?? "")) continue;

      if (start > cursor) pieces.push(doc.createTextNode(value.slice(cursor, start)));
      const span = doc.createElement("span");
      span.setAttribute("style", PRE);
      span.textContent = run;
      pieces.push(span);
      cursor = end;
    }

    if (!pieces.length) continue;
    if (cursor < value.length) pieces.push(doc.createTextNode(value.slice(cursor)));
    text.replaceWith(...pieces);
  }

  return doc.documentElement.outerHTML;
}

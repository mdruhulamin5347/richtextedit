/**
 * The one walk behind every "restate an inherited style on the runs" pass.
 *
 * `font-size` and `color` both INHERIT in CSS, so `<td style="color:#c00000">
 * <p>Positive</p>` is red in every browser and in the printed report. A Plate
 * node cannot say that: only a text LEAF carries either mark, and the
 * deserializer drops the declaration outright from any element that maps to a
 * Plate element — a `<p>`, `<td>`, `<li>`, a heading. So the run arrives
 * carrying nothing.
 *
 * And it is worse on the way out than on screen: the serializer writes neither
 * a block font-size nor a block colour, so opening a stored report that stated
 * one and saving it wrote the value out of the report FOR GOOD.
 *
 * Same shape as `lib/legacy-alignment.ts`, and for the same reason: a value
 * that inherits has to be pushed onto the nodes that can actually hold it,
 * before Plate sees the HTML. Only the property differs between the two passes,
 * so only the property is a parameter — see `lib/inherited-font-size.ts` and
 * `lib/inherited-color.ts` for the ends of the walk that do differ.
 */

/** The properties this walk knows how to restate. */
export type InheritedStyleKey = "fontSize" | "color";

/**
 * Elements whose declaration Plate turns into a mark on the text inside them.
 *
 * These need no help: a value on one of them already reaches the leaf. Anything
 * else — a block, a cell, a table, a wrapper — is where it is lost.
 *
 * `<a>` is deliberately absent. It deserializes to an ELEMENT, not a mark, so a
 * value on a link never reaches the text: that text is wrapped instead.
 */
const MARK_CARRIERS: ReadonlySet<string> = new Set([
  "SPAN",
  "B",
  "STRONG",
  "I",
  "EM",
  "U",
  "S",
  "STRIKE",
  "DEL",
  "INS",
  "SUB",
  "SUP",
  "CODE",
  "MARK",
  "SMALL",
  "BIG",
  "FONT",
]);

/**
 * Elements that can hold no text of their own.
 *
 * Whatever sits between a `<tr>`'s cells is markup whitespace, never the
 * report — and wrapping it in a `<span>` does not merely add noise, it BREAKS
 * THE TABLE. A span is not legal there, so the next time the HTML is parsed the
 * browser foster-parents it out of the table, taking the row structure with it:
 * the table arrives with its rows and borders in pieces.
 *
 * It only bites on some documents, which is what made it look arbitrary: Word
 * writes most tables one tag per line, and a run of whitespace containing a
 * NEWLINE was already skipped as the source file's own indentation. A table
 * written on ONE line has the same whitespace with no newline in it.
 */
const NO_TEXT_CHILDREN: ReadonlySet<string> = new Set([
  "TABLE",
  "THEAD",
  "TBODY",
  "TFOOT",
  "TR",
  "COLGROUP",
]);

/** Text in these is not the report's text. */
const NOT_CONTENT: ReadonlySet<string> = new Set([
  "STYLE",
  "SCRIPT",
  "TITLE",
  "HEAD",
  "NOSCRIPT",
  "TEMPLATE",
]);

/**
 * What a stated value should be restated AS, or undefined to restate nothing.
 *
 * Undefined STOPS the walk rather than continuing up: the nearest declaration
 * still wins even when it is one worth leaving behind, exactly as it does in a
 * browser. Anything else would give a run the colour of some ancestor the
 * document had already overridden.
 */
export type AcceptStated = (stated: string) => string | undefined;

/** The value in effect on a text node, and the element that should carry it. */
interface Restatement {
  text: Text;
  value: string;
  /** The innermost element between the text and the block, if one can hold it. */
  carrier: HTMLElement | null;
}

function restatementFor(
  text: Text,
  styleKey: InheritedStyleKey,
  accept: AcceptStated
): Restatement | null {
  let carrier: HTMLElement | null = null;

  for (let el: HTMLElement | null = text.parentElement; el; el = el.parentElement) {
    if (NOT_CONTENT.has(el.tagName)) return null;

    const stated = el.style?.[styleKey];
    if (stated) {
      // The nearest thing that states a value is already something Plate reads
      // onto the leaf — this text is fine as it stands.
      if (MARK_CARRIERS.has(el.tagName)) return null;

      const value = accept(stated);
      return value ? { text, value, carrier } : null;
    }

    // Remember the innermost carrier on the way up: giving it the value keeps
    // the markup as it was rather than adding a wrapper.
    if (!carrier && MARK_CARRIERS.has(el.tagName)) carrier = el;
  }

  return null;
}

/**
 * Is this text node the source file's own indentation rather than the report's?
 *
 * Word HTML is a pretty-printed tree of `\n\t\t<p>`. A run of spaces WITHOUT a
 * newline is the report's own spacing and does get restated — the width of a
 * space is the size of the font it is set in, so leaving it behind would move
 * the columns a report lays out with. See lib/whitespace.ts.
 */
function isSourceIndentation(text: string): boolean {
  return /^\s*$/.test(text) && text.includes("\n");
}

/**
 * State `styleKey` on every run that only inherits it, in place. Returns how
 * many runs were restated, so a caller can tell an untouched document from a
 * rewritten one and hand back the original string rather than a re-serialized
 * copy of it.
 *
 * Collected first and applied after: wrapping a text node changes the tree the
 * walker is walking.
 */
export function restateInheritedStyle(
  doc: Document,
  styleKey: InheritedStyleKey,
  accept: AcceptStated = (stated) => stated
): number {
  const walker = doc.createTreeWalker(doc.documentElement, NodeFilter.SHOW_TEXT);

  const restatements: Restatement[] = [];
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const text = node as Text;
    if (!text.data || isSourceIndentation(text.data)) continue;
    if (text.parentElement && NO_TEXT_CHILDREN.has(text.parentElement.tagName)) continue;

    const restatement = restatementFor(text, styleKey, accept);
    if (restatement) restatements.push(restatement);
  }

  for (const { text, value, carrier } of restatements) {
    if (carrier) {
      carrier.style[styleKey] = value;
      continue;
    }
    const span = doc.createElement("span");
    span.style[styleKey] = value;
    text.replaceWith(span);
    span.append(text);
  }

  return restatements.length;
}

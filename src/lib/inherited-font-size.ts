/**
 * A font size the document states on a BLOCK, restated on the runs inside it.
 *
 * `font-size` inherits in CSS, so `<td style="font-size:12px"><p>Detected</p>`
 * renders at 12px in every browser and in the printed report. A Plate node
 * cannot say that: only a text LEAF carries a size, and the deserializer drops
 * a block's outright. So the run arrived carrying nothing, fell back to the
 * editable's 18px base, and a report's body text came out BIGGER on screen than
 * in the document it was copied from — while its header, which states its size
 * on the span, came out right. That is what makes the mismatch look arbitrary.
 *
 * Worse on the way out: the serializer writes no block font-size either, so
 * opening a stored report that stated one and saving it wrote the size out of
 * the report for good.
 *
 * Only Word's own clipboard escaped, and only partly: `cleanDocx` copies a
 * paragraph's marks onto its span children. Nothing covers a size on the
 * `<table>`, and nothing at all covers HTML that is not Word's — Word Online,
 * Google Docs, another report page, or the LibreOffice conversion behind the
 * Word-import button.
 *
 * The walk itself lives in lib/inherited-style.ts, which `lib/inherited-color.ts`
 * shares; what is left here is the part that is about SIZES — including the two
 * spellings of a size that are not CSS at all: `<font size=N>` and Word's
 * `mso-ansi-font-size`.
 */
import { restateInheritedStyle } from "@/lib/inherited-style";

/**
 * The px a browser draws `<font size=N>` at — HTML's absolute size table, on
 * the 16px base a browser starts from.
 *
 * Pre-CSS, and still what a legacy report can carry. Plate reads a size off a
 * `<font>`'s STYLE like any other inline element, and off its `size` attribute
 * not at all, so the run rendered at the editor's base instead.
 */
const FONT_ATTRIBUTE_PX = ["10px", "13px", "16px", "18px", "24px", "32px", "48px"];

/**
 * Word's OWN spelling of a run's size, which no browser reads.
 *
 * `mso-ansi-font-size` is the size Word draws Latin text at (and
 * `mso-bidi-font-size` the size it draws right-to-left text at). Word normally
 * writes the CSS `font-size` beside it, so this is a fallback rather than a
 * translation — but when it writes only the Word property, the size reached
 * nothing at all and the run fell back to the editor's base.
 *
 * Read off the STYLE ATTRIBUTE, not off `element.style`: CSSOM drops a property
 * it does not know, so by the time it is an `HTMLElement` this declaration is
 * already gone.
 */
const MSO_FONT_SIZE = /(?:^|;)\s*mso-(?:ansi|bidi)-font-size\s*:\s*([^;]+)/i;

/**
 * Does anything from here up state a CSS `font-size`?
 *
 * The Word property is applied ONLY where the answer is no — where the
 * alternative is not a different size but no size, and so the editor's base.
 * That keeps it from ever overruling a size the document states in the spelling
 * a browser actually draws: on a document where the two disagree, CSS wins here
 * exactly as it does everywhere else.
 */
function statesFontSize(element: HTMLElement | null): boolean {
  for (let el = element; el; el = el.parentElement) {
    if (el.style?.fontSize) return true;
  }
  return false;
}

/** `1`-`7`, or `+n` / `-n` relative to the 3 a browser treats as normal. */
function fontAttributeSize(raw: string | null): string | undefined {
  const value = (raw ?? "").trim();
  const match = value.match(/^([+-]?)(\d+)$/);
  if (!match) return undefined;

  const n = Number.parseInt(match[2], 10);
  const index = (match[1] === "+" ? 3 + n : match[1] === "-" ? 3 - n : n) - 1;
  return FONT_ATTRIBUTE_PX[Math.min(Math.max(index, 0), FONT_ATTRIBUTE_PX.length - 1)];
}

/**
 * The same HTML with every inherited font size stated where a Plate leaf can
 * carry it.
 *
 * The document's own unit is kept — `11.0pt` stays `11.0pt` — so the paste
 * normalizer converts it onto the editor's scale exactly as it does a size the
 * document stated on a span, and `faithful` mode stays faithful.
 */
export function inlineInheritedFontSize(html: string): string {
  if (!html || !/font-size|<font[\s>]/i.test(html)) return html;

  const doc = new DOMParser().parseFromString(html, "text/html");

  // `<font size>` first, so the walk below has one size to inherit rather than
  // two spellings of it. A CSS size on the same element still wins, as it does
  // in a browser.
  doc.querySelectorAll("font[size]").forEach((element) => {
    const font = element as HTMLElement;
    if (font.style.fontSize) return;
    const size = fontAttributeSize(font.getAttribute("size"));
    if (size) font.style.fontSize = size;
  });

  // Word's own property next, and only where nothing states a CSS size — see
  // MSO_FONT_SIZE. In document order, so an ancestor is settled before the
  // descendants that inherit from it are asked.
  doc.querySelectorAll("[style]").forEach((element) => {
    const el = element as HTMLElement;
    if (statesFontSize(el)) return;

    const stated = MSO_FONT_SIZE.exec(el.getAttribute("style") ?? "")?.[1]?.trim();
    if (stated) el.style.fontSize = stated;
  });

  restateInheritedStyle(doc, "fontSize");

  // Not the input string: a `<font size>` above may already have been restated.
  return doc.documentElement.outerHTML;
}

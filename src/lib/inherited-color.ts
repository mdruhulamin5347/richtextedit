/**
 * A text colour the document states on a BLOCK — or in the pre-CSS spelling —
 * restated on the runs inside it.
 *
 * The same hole `lib/inherited-font-size.ts` fills, one property over, and it
 * cost the report more: a colour is the whole point of the run that carries it.
 * `color` inherits, so `<td style="color:#c00000"><p>POSITIVE</p>` is red
 * everywhere HTML is drawn — but Plate reads a colour onto a LEAF only, and
 * drops the declaration from every element that maps to a Plate element. So a
 * pasted `<p style="color:red">`, a red `<td>`, a coloured `<ul>` and a
 * `<div>`-wrapped abnormal-result line all arrived BLACK, and the first save
 * wrote the colour out of the stored report for good.
 *
 * Word's own clipboard was the one thing that mostly worked, because `cleanDocx`
 * copies a paragraph's marks onto its span children. Everything else lost it:
 * Google Docs, Word Online, another report page, a template pasted from a
 * browser, and the LibreOffice conversion behind the Word-import button — which
 * states colour the pre-CSS way, `<font color=...>`, on nearly every run it
 * writes.
 */
import { restateInheritedStyle } from "@/lib/inherited-style";

/** Keywords that are the ABSENCE of a stated colour rather than a colour. */
const NOT_A_COLOUR = /^(inherit|initial|unset|revert|revert-layer|currentcolor|auto|none)$/;

/**
 * Black in every spelling a report can carry.
 *
 * Worth singling out because it is the colour of nearly all of a report, and
 * because Plate's own default for the mark is the literal `black`: a leaf whose
 * colour equals that default is dropped on deserialization, which is exactly
 * what should happen. Restating `#000000` instead would defeat that comparison
 * and hang a `<span style="color: rgb(0, 0, 0)">` on every ordinary run of
 * every LibreOffice document, which writes `<font color="#000000">` throughout.
 *
 * `windowtext` is Word's spelling of the same thing.
 */
function isBlack(value: string): boolean {
  return (
    value === "black" ||
    value === "windowtext" ||
    /^#(0{3,4}|0{6}|0{8})$/.test(value) ||
    /^rgba?\(0,0,0(,(1|1\.0+|100%))?\)$/.test(value)
  );
}

/**
 * What a stated colour should be restated as, or undefined for one that says
 * nothing — in which case nothing is restated and, per `AcceptStated`, no
 * ancestor's colour is inherited past it either.
 *
 * A real colour keeps the document's own spelling, exactly as the size pass
 * keeps its unit.
 */
export function normalizeColor(stated: string): string | undefined {
  const value = stated.trim();
  if (!value) return undefined;

  const compact = value.replace(/\s+/g, "").toLowerCase();
  if (NOT_A_COLOUR.test(compact)) return undefined;
  if (isBlack(compact)) return "black";

  return value;
}

/**
 * The same HTML with every inherited text colour stated where a Plate leaf can
 * carry it, and `<font color>` restated as CSS.
 */
export function inlineInheritedColor(html: string): string {
  // Deliberately loose — `background-color`, `bgcolor` and the word "color" in
  // a report's own text all get past it. What keeps this pass INERT on a
  // document it has nothing to say about is not the guard but the return below:
  // HTML that gained nothing is handed back as the very string it came in as,
  // never a re-serialized copy. A report body reopens and re-saves byte for
  // byte, which is the property the whole load path is written around.
  if (!html || !/color|<font[\s>]/i.test(html)) return html;

  const doc = new DOMParser().parseFromString(html, "text/html");
  let changed = 0;

  // The `color` ATTRIBUTE first, so the walk below has one spelling to inherit
  // rather than two. A CSS colour on the same element still wins, as it does in
  // a browser.
  //
  // Read off any element and not just `<font>`, because by the time a Word
  // paste reaches here it may not be a `<font>` any more: Plate's own docx
  // cleaner renames the tag to `<span>` and leaves the attribute sitting on it,
  // where it means nothing to a browser and nothing to Plate. That is a colour
  // the document plainly states, and the only thing still carrying it.
  doc.querySelectorAll("[color]").forEach((element) => {
    const el = element as HTMLElement;
    if (!el.style || el.style.color) return;
    const color = normalizeColor(el.getAttribute("color") ?? "");
    if (!color) return;
    el.style.color = color;
    changed += 1;
  });

  changed += restateInheritedStyle(doc, "color", normalizeColor);

  return changed > 0 ? doc.documentElement.outerHTML : html;
}

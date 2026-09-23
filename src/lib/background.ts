/**
 * What counts as a FILL, in every spelling a report can state one.
 *
 * One definition, because three places ask the same question of three different
 * elements — a table cell (`extractCellBackground`), a paragraph or heading
 * (`components/block-background-plugin.ts`), and a row or table standing behind
 * either — and a fill that counted in one place and not another is how a shaded
 * report comes back half shaded.
 */

/** `transparent` and friends are the ABSENCE of a fill, not a fill. */
const NO_FILL = /^(transparent|none|initial|inherit|unset|revert)$/i;

/**
 * A colour, in the spellings CSS writes one.
 *
 * The `background` shorthand is only read when the longhand says nothing, and
 * what is left in that case is a shorthand with no colour in it: `background:
 * url(qr.png) no-repeat`. Passed through, that would be written back out as
 * `background-color: url(qr.png) no-repeat` — not a fill, not even valid CSS.
 */
const COLOUR = /^(#[0-9a-f]{3,8}|(rgba?|hsla?)\([^()]*\)|[a-z]+)$/i;

/**
 * The fill this element states on itself, or undefined for none.
 *
 * Three spellings, all of them live: the CSS longhand, the `background`
 * shorthand Word and LibreOffice both write, and the pre-CSS `bgcolor`
 * attribute — which is not a museum piece here, it is what LibreOffice puts on
 * a shaded cell (`<td bgcolor="#d9e2f3" style="background: #d9e2f3">`), what
 * Word 97 HTML writes, and what every legacy report body in the database
 * carries.
 */
export function statedBackground(el: Element): string | undefined {
  const style = (el as HTMLElement).style;

  const stated = (
    style?.backgroundColor ||
    style?.background ||
    el.getAttribute?.("bgcolor") ||
    ""
  ).trim();

  if (!stated || NO_FILL.test(stated) || !COLOUR.test(stated)) return undefined;

  return stated;
}

/**
 * The scale this editor measures type on: POINTS, the same scale Word and
 * LibreOffice measure it on.
 *
 * Reported by name: "currently main doc has font size 10 but after paste our
 * editor show 13". Both numbers described the same text. Word states type in
 * points and the editor used to convert every pasted size into whole CSS
 * pixels, so 10pt arrived as `13px` — the control then named it 13, and 13px is
 * not even 10pt: 10pt is 13.333px, so the text was also drawn 2.5% small.
 *
 * Two things follow from that, and they are the whole of this module.
 *
 * **A document's own absolute unit is kept.** `10.0pt` stays `10pt`, so nothing
 * is rounded on the way in and the browser resolves it exactly as Word does —
 * 72 points to the inch, 96 CSS pixels to the same inch. Points are CSS, so the
 * node, the saved HTML and the printed report all carry them with no conversion
 * anywhere between; this is the same shape as the line-gap rule in
 * lib/line-gap.ts, where everything in the middle speaks CSS and only the edges
 * convert. A legacy report stating `15px` is likewise left in px — it renders
 * as it always did, and open + save does not rewrite it.
 *
 * **Everything the control says is in points.** A size is NAMED in points
 * whatever unit it is stated in, so one report cannot show two scales: a
 * pasted `10pt` reads 10 and a legacy `15px` reads 11.25, which is what 15px is.
 * Setting a size writes points.
 *
 * A relative size (`em`, `rem`, `%`) has lost the context it was a multiple of
 * by the time a run reaches this editor, so it is resolved against the size the
 * run will actually be drawn at and becomes points like any other foreign unit.
 */

/**
 * CSS absolute length units, in points. 72pt to the inch and 96 CSS px to the
 * same inch — the 96dpi every browser assumes, and what makes a stated point
 * size render at the size Word prints it.
 */
export const UNIT_TO_PT: Record<string, number> = {
  pt: 1,
  px: 72 / 96,
  pc: 12,
  in: 72,
  cm: 72 / 2.54,
  mm: 72 / 25.4,
};

/**
 * CSS's ABSOLUTE size keywords, in px.
 *
 * These are a defined table, not an opinion: the values a browser draws them at
 * off the 16px `medium` it starts from. A document that states one used to lose
 * its size outright and fall back to the editor's base, which is the one thing
 * worse than either number — a converter that turns `<font size>` into CSS
 * writes these, and so does hand-written report HTML.
 *
 * They are kept in PX for the same reason `FONT_ATTRIBUTE_PX` is
 * (lib/inherited-font-size.ts): the table itself is defined in pixels, so those
 * pixels are the faithful reading of the keyword rather than a conversion of it.
 */
const KEYWORD_PX: Record<string, number> = {
  "xx-small": 9,
  "x-small": 10,
  small: 13,
  medium: 16,
  large: 18,
  "x-large": 24,
  "xx-large": 32,
  "xxx-large": 48,
};

/**
 * CSS's RELATIVE size keywords — one step of 1.2 off the size inherited from
 * the parent, which is the factor CSS names.
 *
 * Relative to a context that is gone by the time a run reaches this editor, so
 * they resolve against the base exactly as `em` and `%` do, and become points
 * for the same reason.
 */
const RELATIVE_KEYWORD: Record<string, number> = {
  smaller: 1 / 1.2,
  larger: 1.2,
};

/**
 * The editor's own text size, in points — `text-[18px]` on the `report`
 * variant in components/ui/editor.tsx, which is 13.5pt EXACTLY (13.5 x 96/72 =
 * 18). Both statements describe the same rendering; this one is on the scale
 * the control reads in, so an unsized run is named honestly instead of being
 * called 18.
 */
export const BASE_FONT_PT = 13.5;

/**
 * Sizes outside this range are a broken document, not a design decision.
 *
 * Also the range the control accepts, deliberately the same one: a size it let
 * somebody type but this rejected on reopen would be silently lost on the next
 * save.
 */
export const MIN_FONT_PT = 1;
export const MAX_FONT_PT = 150;

/** A bare number means px, which is how CSS reads one in most places. */
const SIZE_PATTERN = /^([\d.]+)\s*(px|pt|pc|in|cm|mm|em|rem|%)?$/;

/**
 * A number with no trailing zeros: 11.0 -> "11", 11.25 -> "11.25".
 *
 * Two places rounding to two decimals matters. A stated size survives a save
 * and a reopen unchanged instead of drifting a fraction on every pass, and a
 * converted one is a number a reader can act on — 11.25, not 11.249999999999998.
 */
export function formatFontSizeNumber(n: number): string {
  return String(Math.round(n * 100) / 100);
}

/**
 * A stated size split into its number and its unit, or null if it is not one.
 *
 * A KEYWORD is resolved here too, so everything downstream sees one shape: an
 * absolute keyword becomes its defined pixels, a relative one becomes a
 * multiple of the base in `em` — which is what it is.
 */
function parseFontSize(raw: unknown): { value: number; unit: string } | null {
  const value = String(raw ?? "")
    .trim()
    .toLowerCase();

  if (value in KEYWORD_PX) return { value: KEYWORD_PX[value], unit: "px" };
  if (value in RELATIVE_KEYWORD) return { value: RELATIVE_KEYWORD[value], unit: "em" };

  const match = value.match(SIZE_PATTERN);
  if (!match) return null;

  const n = Number.parseFloat(match[1]);
  if (!Number.isFinite(n) || n <= 0) return null;

  return { value: n, unit: match[2] || "px" };
}

/**
 * A stated font size in POINTS — the scale to compare two sizes on and the
 * scale to name one on.
 *
 * Null for anything that names no size at all (`inherit`, `initial`, an empty
 * string) and for a length no report would use. CSS's size KEYWORDS do name
 * one and are resolved, not rejected.
 */
export function fontSizeToPt(raw: unknown): number | null {
  const parsed = parseFontSize(raw);
  if (!parsed) return null;

  const { value, unit } = parsed;
  const pt =
    unit === "%"
      ? (value / 100) * BASE_FONT_PT
      : unit === "em" || unit === "rem"
        ? value * BASE_FONT_PT
        : value * (UNIT_TO_PT[unit] ?? 1);

  if (!Number.isFinite(pt) || pt < MIN_FONT_PT || pt > MAX_FONT_PT) return null;
  return Math.round(pt * 100) / 100;
}

/**
 * What a stated font size should be STORED as — on the node, and so in the
 * saved HTML the printed report is drawn from.
 *
 * Points and pixels are both kept in their own unit, and that is the point of
 * this function rather than an oversight:
 *
 * - **Points are what Word and LibreOffice state**, and keeping them is what
 *   makes a paste exact. Converting 10pt to whole px cost a quarter of a pixel
 *   on every run and cost the control the only number the document ever named.
 * - **Pixels are what the reports already in the database state**, from the
 *   years this editor worked in them. Rewriting those to points would change
 *   every stored body the first time it was opened and saved — the same size on
 *   screen, but 15px named 11.25 rather than 15, on ~70 live installations. A
 *   document that states px is left in px; only a size somebody SETS becomes
 *   points, and only the paragraph they set it on.
 *
 * Every other unit is resolved to points, because none of them is a scale
 * anything here reads in and all of them are exactly convertible.
 */
export function toEditorFontSize(raw: unknown): string | null {
  const parsed = parseFontSize(raw);
  if (!parsed) return null;

  const pt = fontSizeToPt(raw);
  if (pt === null) return null;

  return parsed.unit === "pt" || parsed.unit === "px"
    ? `${formatFontSizeNumber(parsed.value)}${parsed.unit}`
    : `${formatFontSizeNumber(pt)}pt`;
}

/** A size as the font-size control states it: a point value. */
export function pointsToFontSize(pt: number): string {
  return `${formatFontSizeNumber(pt)}pt`;
}

/**
 * The number the font-size control shows for a stated size — its size in
 * points, whatever unit it is stated in.
 *
 * Plate's own `toUnitLess` cannot do this and not only because of the unit: it
 * matches the FIRST RUN OF DIGITS (`/\d+/`), so it reads `13.5pt` back as "13"
 * and `11.25pt` as "11". A control that cannot show a fraction cannot show the
 * editor's own base size, which is 13.5pt.
 */
export function fontSizeLabel(raw: unknown, fallbackPt: number = BASE_FONT_PT): string {
  return formatFontSizeNumber(fontSizeToPt(raw) ?? fallbackPt);
}

/**
 * A value two stated sizes can be compared on, so `10pt` and `13.33px` count as
 * the same size rather than as two paragraphs disagreeing.
 *
 * Worth having for one case in particular: text pasted from Word (points) next
 * to text a legacy report states in pixels, in the same paragraph. Comparing
 * the strings would find no single size there and drop the block's strut. See
 * `agreedFontSize` in lib/block-font-size.ts.
 */
export function comparableFontSize(raw: unknown): string {
  const pt = fontSizeToPt(raw);
  return pt === null ? String(raw ?? "") : formatFontSizeNumber(pt);
}

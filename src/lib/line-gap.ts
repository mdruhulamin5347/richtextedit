/**
 * A LINE GAP, the way Word and LibreOffice mean it.
 *
 * The two models do not agree, and that is the whole of this file. Word's
 * "Multiple N" and LibreOffice's "Proportional N%" multiply the FONT'S NATURAL
 * LINE HEIGHT — what the font's own metrics say a line of it is. CSS
 * `line-height: N` multiplies the FONT SIZE. They coincide only for a font
 * whose natural line is exactly 1em, which no text font is.
 *
 * Measured in Chrome at 15px: Calibri's natural line is 1.2, so Word draws a
 * gap of 1 at 18px, 1.2 at 21.6px and 1.5 at 27px, where this editor drew
 * 15 / 18 / 22.5px — 17% tight at every step, and more on a font with a taller
 * natural line (the editor's own UI stack is 1.333, so 25%).
 *
 *     css ratio = gap x natural(font)
 *
 * WHY THE NATURAL LINE IS MEASURED AND NOT TABULATED. It is a property of the
 * font FILE, so it is only right if the font is actually installed: on a Linux
 * box without Word's fonts, Chrome substitutes and reports 1.15 for Arial,
 * Times, Verdana, Georgia and eighteen others alike — one substitute's metric
 * wearing every name. A table built from that would be fiction, and a table of
 * hand-copied metrics would be worse: unverifiable, and wrong the day a font is
 * updated. The editor and the print page both run in the same browser on the
 * same machine as Word does, so asking that browser is asking the same font
 * Word asked.
 *
 * The answer is baked into the saved report as a plain ratio, so the printed
 * page needs no measuring and no script — and a ratio resolves against the
 * element's own font SIZE, never its family, so it cannot drift afterwards.
 */

/**
 * What a line is worth where nothing can be measured.
 *
 * jsdom does no layout, so every rect is zero there and the unit tests would
 * otherwise divide by it. 1.2 is Calibri's, the font these reports are written
 * in. A browser never reaches this.
 */
export const FALLBACK_NATURAL_LINE_HEIGHT = 1.2;

/**
 * The gap a pasted block is given when the document states none at all.
 *
 * Asked for on 2026-09-16: "jodi kono karone line gap na paoya jai tahole
 * default hisebe tumi 1.2 set kore rakho". It applies only where there is
 * nothing to read — a document that states `normal` HAS stated its spacing
 * (that is Word's "Single", a gap of 1) and keeps it.
 */
export const DEFAULT_PASTED_LINE_GAP = 1.2;

/** Measured once per font stack — the metric cannot change under us. */
const cache = new Map<string, number>();

/**
 * Big enough that the line box's rounding to whole pixels cannot move the
 * ratio: at 15px Calibri measures 1.2 and at 1000px it measures its true
 * 1.2207, and it is the true one that has to be multiplied.
 */
const PROBE_PX = 1000;

let probe: HTMLElement | null = null;

function probeElement(): HTMLElement | null {
  if (typeof document === "undefined" || !document.body) return null;
  if (probe && probe.isConnected) return probe;

  probe = document.createElement("div");
  // Out of flow, out of sight, and out of the way of anything that walks the
  // document — including this editor's own serializer.
  probe.setAttribute("aria-hidden", "true");
  probe.style.cssText =
    "position:absolute;left:-99999px;top:0;visibility:hidden;white-space:nowrap;" +
    "padding:0;border:0;margin:0;line-height:normal";
  probe.textContent = "Hg";
  document.body.appendChild(probe);
  return probe;
}

/**
 * The natural line height of a font, as a multiple of its size — `line-height:
 * normal`, which is what "Single" spacing means in both Word and LibreOffice.
 */
export function naturalLineHeight(fontFamily?: string | null): number {
  const stack = String(fontFamily ?? "").trim() || "inherit";

  const cached = cache.get(stack);
  if (cached !== undefined) return cached;

  const el = probeElement();
  let ratio = FALLBACK_NATURAL_LINE_HEIGHT;

  if (el) {
    el.style.fontFamily = stack === "inherit" ? "" : stack;
    el.style.fontSize = `${PROBE_PX}px`;
    const measured = el.getBoundingClientRect().height / PROBE_PX;
    // A browser that laid nothing out (jsdom, a detached document) answers 0.
    // A ratio outside this range is not a font, it is a broken measurement.
    if (Number.isFinite(measured) && measured >= 0.5 && measured <= 4) {
      ratio = measured;
    }
  }

  cache.set(stack, ratio);
  return ratio;
}

/**
 * The natural line of the font an ELEMENT is actually drawn in.
 *
 * Asking the DOM rather than reconstructing a font stack from the node: the
 * font a paragraph is set in may be stated on the paragraph, on the runs inside
 * it, on the editable around it or on nothing at all, and the browser has
 * already resolved all of that — including which families are installed, which
 * is the part no node can know. Falls back to the editor's own where there is
 * no element to ask.
 */
export function naturalLineHeightOfElement(el: Element | null | undefined): number {
  if (!el || typeof window === "undefined" || !window.getComputedStyle) {
    return naturalLineHeight(undefined);
  }
  return naturalLineHeight(window.getComputedStyle(el).fontFamily);
}

/** Ratios are compared and stored to this many places, so they round-trip. */
function round(value: number): number {
  return Math.round(value * 10000) / 10000;
}

/**
 * The CSS `line-height` that draws a Word/LibreOffice line gap of `gap` against
 * a font whose natural line is `natural`. This is what gets stored on the node
 * and written into the report.
 *
 * Takes the measured ratio rather than a font name so that the arithmetic can
 * be tested without a browser, and so the caller decides how the font was
 * resolved — from an element, which is the only reliable way, or by name.
 */
export function lineGapToCssRatio(gap: number, natural: number): number {
  return round(gap * (natural > 0 ? natural : FALLBACK_NATURAL_LINE_HEIGHT));
}

/**
 * The line gap a stored CSS ratio represents — what the control shows as
 * selected. The inverse of `lineGapToCssRatio`.
 */
export function cssRatioToLineGap(ratio: number, natural: number): number {
  return round(ratio / (natural > 0 ? natural : FALLBACK_NATURAL_LINE_HEIGHT));
}

/** Test seam: forget what was measured, so a changed font stack is re-asked. */
export function resetNaturalLineHeightCache(): void {
  cache.clear();
  if (probe?.isConnected) probe.remove();
  probe = null;
}

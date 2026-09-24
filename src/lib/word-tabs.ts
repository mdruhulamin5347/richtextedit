/**
 * Every tab a Word paste carries, kept as a real tab.
 *
 * Word writes a tab as a span full of `&nbsp;` marked `mso-tab-count:N`, and
 * Plate's docx cleaner turns that span into N tabs — but only when its `style`
 * STARTS with `mso-tab-count:`. Word also writes `font-size:12.0pt;mso-tab-count:1`,
 * and that one pasted as a fixed run of `&nbsp;`: not a tab at all, just spaces
 * as wide as Word happened to need.
 *
 * Each becomes as many editor tabs as Word wrote. Where the text after them
 * lands — Word's half-inch stops, a ruler stop, a hanging indent's — is worked
 * out afterwards, in lib/office-tab-stops.ts.
 */

const MSO_TAB_COUNT = /(^|;)\s*mso-tab-count\s*:\s*(\d+)[^;]*;?/i;

/**
 * The same HTML with every `mso-tab-count` span still in it turned into that
 * many tabs. The span itself is kept, minus the property, so the tab stays in
 * the run's font.
 *
 * Returns the input untouched — the same string — when there is none.
 */
export function expandWordTabSpans(html: string): string {
  if (!html || !/mso-tab-count/i.test(html)) return html;

  const doc = new DOMParser().parseFromString(html, "text/html");
  let changed = false;

  for (const span of Array.from(doc.querySelectorAll("span[style]"))) {
    const style = span.getAttribute("style") ?? "";
    const match = MSO_TAB_COUNT.exec(style);
    if (!match) continue;

    span.textContent = "\t".repeat(Number(match[2]));
    const rest = style.replace(MSO_TAB_COUNT, "$1").replace(/^[\s;]+|[\s;]+$/g, "");
    if (rest) span.setAttribute("style", rest);
    else span.removeAttribute("style");
    changed = true;
  }

  return changed ? doc.documentElement.outerHTML : html;
}

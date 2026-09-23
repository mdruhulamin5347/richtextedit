/**
 * A box Word draws with Borders & Shading rather than with a text box.
 *
 * It is the other half of lib/word-textbox.ts, and from the reader's side the
 * same thing: a heading in a box at the top of a report. Word writes it as a
 * DIV around the paragraph, marked as its own:
 *
 *   <div style='border:solid #2E74B5 1.0pt; padding:1.0pt 4.0pt;
 *               background:#DEEAF6; mso-element:para-border-div'>
 *     <p class=MsoNormal style='border:none'>CYP2C19 GENOTYPING</p></div>
 *
 * Plate drops the div and everything on it, so the box vanished and — worse —
 * its fill did not: a background on a div reaches the RUNS inside it, so the
 * heading came back as highlighted TEXT sitting where a box used to be.
 *
 * LibreOffice states the same box on the paragraph itself (`<p style='border:
 * 1pt solid #2e74b5'>`), which is what an uploaded .doc or .docx brings, so
 * both spellings are read here.
 */
import { borderedBoxHtml, boxStyleOf, drawnBorderOf } from "@/lib/bordered-box";

/** Word's marker for "this div is a paragraph's border". */
const PARA_BORDER_DIV = /mso-element:\s*para-border-div/i;

/** A box already inside a table is a cell's border, not a box to rebuild. */
function isInsideTable(element: Element): boolean {
  return element.closest("table") !== null;
}

/**
 * The elements that are a box: Word's own border div, and a paragraph that
 * states a border of its own.
 *
 * Deliberately narrow. Any bordered `<div>` would sweep up the scaffolding of
 * every web page ever pasted; these two are what a document means by a box.
 */
function boxes(doc: Document): HTMLElement[] {
  const found: HTMLElement[] = [];

  for (const element of Array.from(doc.body.querySelectorAll("div, p"))) {
    const el = element as HTMLElement;
    if (isInsideTable(el)) continue;

    if (el.tagName === "DIV") {
      if (PARA_BORDER_DIV.test(el.getAttribute("style") || "")) found.push(el);
      continue;
    }

    // A paragraph is a box only when it draws a line of its own. Word writes
    // `border:none` on the paragraph inside its border div, which says the
    // opposite.
    if (drawnBorderOf(el)) found.push(el);
  }

  // Innermost first, so a paragraph inside a border div is left to the div.
  return found.filter((el) => !found.some((other) => other !== el && other.contains(el)));
}

/**
 * The same HTML with every such box rebuilt as the one-cell table that can
 * carry a line, or the original string when the document draws none.
 */
export function inlineWordBorderBoxes(html: string): string {
  if (!html || !/border/i.test(html)) return html;

  const doc = new DOMParser().parseFromString(html, "text/html");
  const found = boxes(doc);
  if (found.length === 0) return html;

  let changed = false;
  for (const element of found) {
    const box = boxStyleOf(element);
    if (!box) continue;

    // The paragraph inside Word's border div carries `border:none; padding:0`
    // to say the div owns both; a bordered paragraph is its own content.
    const content =
      element.tagName === "DIV" ? element.innerHTML : `<p>${element.innerHTML}</p>`;

    const replacement = doc.createElement("div");
    replacement.innerHTML = borderedBoxHtml(content, box);
    element.replaceWith(...Array.from(replacement.childNodes));
    changed = true;
  }

  return changed ? doc.documentElement.outerHTML : html;
}

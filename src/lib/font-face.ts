import { matchRtfParagraphs, readRtfParagraphs } from "./rtf-paragraphs";

/**
 * A run's font, stated the pre-CSS way, restated as CSS.
 *
 * LibreOffice writes a run's font as `<font face="Calibri, serif">` and the
 * paragraph's in its `<style>` block (`h3.western { font-family: "Times New
 * Roman" }`). Plate reads a family off `style` only, and its docx cleaner —
 * which runs on every paste that carries RTF, and LibreOffice's always does —
 * renames `<font>` to `<span>`, keeping an attribute CSS ignores. So the run's
 * Calibri was lost and the style block's Times New Roman drew the whole paste:
 * it looked nothing like the document, and every label came out a different
 * width, which moved the tab columns along with it.
 *
 * The run's own family wins over the paragraph's, as it does in the document:
 * the docx cleaner wraps a block's font AROUND its runs, never over them.
 */

/**
 * The same HTML with each `<font face>` also stating its family in `style`.
 * A `<font>` that already states one in CSS keeps it.
 *
 * Returns the input untouched — the same string — when there is nothing to do.
 */
export function inlineFontFaces(html: string): string {
  if (!html || !/<font[^>]*\sface\s*=/i.test(html)) return html;

  const doc = new DOMParser().parseFromString(html, "text/html");
  let changed = false;

  doc.querySelectorAll<HTMLElement>("font[face]").forEach((font) => {
    const face = font.getAttribute("face")?.trim();
    if (!face || font.style.fontFamily) return;
    font.style.fontFamily = face;
    changed = true;
  });

  return changed ? doc.documentElement.outerHTML : html;
}

/** Blocks that are a line of the document. */
const TEXT_BLOCKS = "p, h1, h2, h3, h4, h5, h6, li";

/** What makes a block hold other blocks rather than be a line of text itself. */
const NESTED_BLOCKS = "p, div, table, ul, ol, li, h1, h2, h3, h4, h5, h6, blockquote, pre";

/** Is some family already stated around the block's text — its own, a run's, or an ancestor's? */
function statesAFont(block: HTMLElement): boolean {
  const walker = block.ownerDocument.createTreeWalker(block, NodeFilter.SHOW_TEXT);
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    if (!/\S/.test(node.textContent ?? "")) continue;
    for (let el = node.parentElement; el; el = el.parentElement) {
      if (el.style?.fontFamily) return true;
    }
    return false;
  }
  return true; // no text: nothing to draw, nothing to change
}

/** A family name as CSS writes it: quoted when it is more than one word. */
function cssFamily(name: string): string {
  return /^[a-z-]+$/i.test(name) ? name : `"${name.replace(/"/g, "")}"`;
}

/**
 * The document's font on every line of a LibreOffice paste whose HTML names none.
 *
 * LibreOffice writes a run's font only where it differs from the paragraph
 * style's, and the style's own font only when the paragraph carries its class
 * (`p.western { font-family: … }`). A plain paragraph in the document's default
 * font arrives naming no font at all — so the editor drew it in its own UI
 * face, a line 1.7px taller and every word a different width than the
 * document's Calibri, which moved its tab columns too. The RTF it copies
 * alongside names each paragraph's font (`\f4` → Calibri), and each line is
 * matched to its RTF paragraph by text.
 *
 * Wrapped round the line's runs, where Plate keeps a font — never over one: a
 * family already stated on the line, a run or anything above wins, as it does
 * in the document. Returns the input untouched — the same string — when
 * nothing changes.
 */
export function inlineRtfParagraphFonts(html: string, rtf: string | null | undefined): string {
  const paragraphs = readRtfParagraphs(rtf);
  if (!html || !paragraphs.some((paragraph) => paragraph.font)) return html;

  const doc = new DOMParser().parseFromString(html, "text/html");
  if (!doc.body) return html;

  const blocks = Array.from(doc.body.querySelectorAll<HTMLElement>(TEXT_BLOCKS));
  const rtfOf = matchRtfParagraphs(blocks, paragraphs);
  let changed = false;

  for (const block of blocks) {
    const font = rtfOf.get(block)?.font;
    if (!font || block.querySelector(NESTED_BLOCKS) || statesAFont(block)) continue;
    const run = doc.createElement("span");
    run.style.fontFamily = cssFamily(font);
    run.append(...Array.from(block.childNodes));
    block.append(run);
    changed = true;
  }

  return changed ? doc.body.innerHTML : html;
}

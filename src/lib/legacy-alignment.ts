/**
 * Alignment that predates CSS.
 *
 * A paragraph can say it is centred in two ways, and Word uses both:
 *
 *   <p align=center style='text-align:center'>   the modern pair
 *   <p align=center>                             the attribute alone
 *
 * Only the CSS half is ever read — TextAlignPlugin is configured with
 * `styleKey: "textAlign"`, so it looks at the style and nothing else. A
 * document that states alignment the old way therefore pastes flush left, which
 * is what happened to the centred "TEST RESULT" heading and the rule of dashes
 * above it. LibreOffice, which converts the legacy `.doc` uploads on the server,
 * writes the attribute form constantly; so does anything old enough to emit
 * `<center>`.
 *
 * Rewriting those into an inline `text-align` before the HTML is deserialized
 * means one representation reaches the editor, and everything downstream —
 * which already handles the CSS correctly — needs no special case.
 */

const ALIGNMENTS = new Set(["left", "center", "right", "justify"]);

/**
 * Elements where `align` means something else entirely.
 *
 * On a table it positions the TABLE on the page, and on an image it floats the
 * image or sets its vertical alignment. Turning either into `text-align` would
 * centre the contents of something the document only asked to place.
 */
const NOT_TEXT_ALIGNMENT = new Set(["TABLE", "IMG", "COL", "COLGROUP", "HR"]);

/**
 * Blocks a wrapper's alignment can be pushed down onto.
 *
 * No TABLE: a `<div align=center>` around a table asks for the TABLE to sit in
 * the middle of the page, not for every cell's text to be centred. Excluding
 * the tag is only half of that — see `crossesTable`, which is what stops the
 * selector from reaching straight through the table to the paragraphs inside
 * its cells.
 */
const ALIGNABLE_BLOCKS = "p, h1, h2, h3, h4, h5, h6, li, div";

/**
 * Would the alignment have to cross a TABLE to reach this block?
 *
 * A browser stops `text-align` at a table — its own stylesheet says
 * `table { text-align: start }` — so `<div align=center>` (or `<center>`, or a
 * CSS `text-align:center`) around a table centres the TABLE on the page and
 * leaves every cell's text exactly where the document left it. Measured in a
 * bare iframe, which is the same reference the rest of the paste work uses: a
 * centred three-column signature block reads `start` in its cells there.
 *
 * Pushing the wrapper's alignment through anyway is what centred every one of
 * those cells here, on a table the author had only asked to place.
 */
function crossesTable(source: Element, block: Element): boolean {
  for (let node = block.parentElement; node && node !== source; node = node.parentElement) {
    if (node.tagName === "TABLE") return true;
  }
  return false;
}

/**
 * Table parts that can state an alignment for everything beneath them.
 *
 * `text-align` INHERITS, so a `<tr align=center>` centres every cell in that
 * row and Word shows it centred. None of these survives as a node carrying an
 * alignment, though — only the cells do — so the value has to be moved onto
 * them or it is simply lost, which is how a centred header row came back left
 * even after the paragraph and cell forms were both handled.
 */
const ALIGNMENT_CONTAINERS = "table, thead, tbody, tfoot, tr";

/**
 * The same HTML with every legacy alignment stated as CSS.
 *
 * The whole document is returned, `<head>` included: Word ships its formatting
 * as a `<style>` block that JuicePlugin later inlines, and returning only the
 * body would throw away the fonts, sizes and column widths along with it.
 */
export function inlineLegacyAlignment(html: string): string {
  // `text-align` counts too, not just the legacy notations: a `<tr>` or
  // `<tbody>` can state it in CSS, and that still has to be moved onto the
  // cells because a row is not a node that carries an alignment.
  if (!html || !/<center[\s>]|align\s*=|text-align/i.test(html)) return html;

  const doc = new DOMParser().parseFromString(html, "text/html");

  // <center> is a block that centres its contents — a div that says so in CSS.
  doc.querySelectorAll("center").forEach((element) => {
    const replacement = doc.createElement("div");
    replacement.setAttribute("style", "text-align: center");
    while (element.firstChild) {
      replacement.append(element.firstChild);
    }
    element.replaceWith(replacement);
  });

  doc.querySelectorAll("[align]").forEach((element) => {
    if (NOT_TEXT_ALIGNMENT.has(element.tagName)) return;

    const align = (element.getAttribute("align") ?? "").trim().toLowerCase();
    if (!ALIGNMENTS.has(align)) return;

    // The CSS wins where a document states both, because that is what a browser
    // showed the author.
    const styled = element as HTMLElement;
    if (!styled.style?.textAlign) {
      styled.style.textAlign = align;
    }
  });

  // A <div> is not a node in this editor: it is unwrapped, and whatever it said
  // about alignment is unwrapped with it. So the alignment has to be moved
  // somewhere that survives — down onto the blocks it contains, or onto the
  // <div> itself once it has become a paragraph.
  doc.querySelectorAll("div").forEach((element) => {
    const wrapper = element as HTMLElement;
    const align = wrapper.style?.textAlign;
    if (!align || !ALIGNMENTS.has(align)) return;

    const blocks = Array.from(wrapper.querySelectorAll(ALIGNABLE_BLOCKS)).filter(
      (block) => !crossesTable(wrapper, block)
    );
    if (blocks.length > 0) {
      blocks.forEach((block) => {
        const styled = block as HTMLElement;
        // Only where the block is silent: a paragraph that states its own
        // alignment is doing so deliberately, against its container.
        if (!styled.style?.textAlign) {
          styled.style.textAlign = align;
        }
      });
      return;
    }

    // A wrapper holding nothing but a table is PLACING that table. Nothing in
    // it can carry the alignment, and the branch below would wrap the table in
    // a paragraph, so leave it as it stands.
    if (wrapper.querySelector("table")) return;

    // Nothing inside to carry it: this div IS the paragraph.
    const paragraph = doc.createElement("p");
    for (const attribute of Array.from(wrapper.attributes)) {
      paragraph.setAttribute(attribute.name, attribute.value);
    }
    while (wrapper.firstChild) {
      paragraph.append(wrapper.firstChild);
    }
    wrapper.replaceWith(paragraph);
  });

  // Same problem one level up in a table: the row states it, only the cell can
  // carry it.
  doc.querySelectorAll(ALIGNMENT_CONTAINERS).forEach((container) => {
    const align = (container as HTMLElement).style?.textAlign;
    if (!align || !ALIGNMENTS.has(align)) return;

    container.querySelectorAll("td, th").forEach((cell) => {
      const styled = cell as HTMLElement;
      // A cell of a table NESTED in this one is behind another table boundary,
      // and a browser resets the alignment there too.
      if (crossesTable(container, cell)) return;
      // A cell that states its own — in either notation — outranks the row.
      if (!styled.style?.textAlign && !styled.getAttribute("align")) {
        styled.style.textAlign = align;
      }
    });
  });

  return doc.documentElement.outerHTML;
}

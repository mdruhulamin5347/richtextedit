/**
 * A Word TEXT BOX, recovered from the clipboard.
 *
 * Insert → Text Box is how a report gets its title bar: a full-width box, a
 * fill behind it, coloured text inside. Copy that and the editor showed a
 * BROKEN IMAGE — no text, no colour, nothing to edit — and the reason is not
 * the editor's:
 *
 *   <p><!--[if gte vml 1]><v:shape … fillcolor="#a8cdea"><v:textbox>
 *        <div><p><span style='color:red'>MOLECULAR BIOLOGY REPORT</span></p></div>
 *      </v:textbox></v:shape><![endif]-->
 *      <![if !vml]><img src="file:///C:/…/clip_image001.png"><![endif]></p>
 *
 * The text and the fill are inside an HTML COMMENT — Word's conditional
 * comment for IE's VML renderer — so no browser parses them and no pass over
 * the DOM can see them. What every other browser is left with is the second
 * branch: a PICTURE of the box, at a `file:///` path on the machine it was
 * copied from, which nothing on a web page is allowed to load. Hence the empty
 * frame. Summernote had exactly the same hole.
 *
 * So the comment is opened and read. The box becomes what it looks like — its
 * paragraphs, each laid on the shape's fill, which the editor draws as a
 * full-width band (components/block-background-plugin.ts) — and the picture
 * fallback is dropped so the report does not carry both.
 *
 * A box Word DREW A LINE AROUND becomes a one-cell table instead, because a
 * paragraph cannot carry an outline and a table cell carries all three — the
 * line, the fill and the inset — with machinery this editor already has. Which
 * of the two a box gets is decided by the box itself: only an outline it
 * actually states earns the table, so a fill-only box stays the plain band it
 * looks like.
 */

import { borderedBoxHtml } from "@/lib/bordered-box";
import { normalizeColor } from "@/lib/inherited-color";

/** VML shapes that can hold a text box. */
const SHAPE_TAGS = new Set(["v:shape", "v:rect", "v:roundrect", "v:oval"]);

/** Blocks a fill is laid on, so the band is the width of the report. */
const BLOCK_TAGS = new Set(["P", "H1", "H2", "H3", "H4", "H5", "H6", "LI"]);

function tag(el: Element): string {
  return el.tagName.toLowerCase();
}

function descendants(root: ParentNode): Element[] {
  return Array.from(root.querySelectorAll("*"));
}

/**
 * A colour off a VML attribute, in CSS.
 *
 * Word does not write a plain colour on a shape. It writes the colour AND the
 * theme slot it came from:
 *
 *   fillcolor="#deeaf6 [660]"   strokecolor="#2e74b5 [2404]"   fillcolor="white [3201]"
 *
 * The bracket is meaningless to CSS and invalidates the whole declaration, so
 * `background: #deeaf6 [660]` set nothing at all and `border: 1pt solid
 * #2e74b5 [2404]` fell back to a default black line. A box with its outline
 * turned off then had NOTHING left — it came through as bare text, which is
 * what a Word text box did on every paste until this was read properly.
 *
 * Everything else about a colour — `windowtext`, a name, a hex — is what
 * normalizeColor already knows.
 */
function vmlColor(raw: string | null | undefined): string | undefined {
  return normalizeColor((raw ?? "").split("[")[0]);
}

/**
 * The shape's fill, or undefined when it is drawn without one.
 *
 * `fillcolor` on the shape and `<v:fill color>` inside it are the same value in
 * two places; Word writes whichever suits it, and `filled="f"` / `on="f"` is
 * how it says "no fill at all".
 */
function shapeFill(shape: Element): string | undefined {
  if ((shape.getAttribute("filled") || "").toLowerCase() === "f") return undefined;

  const fill = Array.from(shape.children).find((el) => tag(el) === "v:fill");
  if ((fill?.getAttribute("on") || "").toLowerCase() === "f") return undefined;

  return vmlColor(shape.getAttribute("fillcolor") || fill?.getAttribute("color"));
}

/**
 * The line Word drew around the box, or undefined for one drawn without.
 *
 * VML's default is a thin black line, but a default is not a decision: a box
 * that states nothing about its stroke is left as a band rather than given an
 * outline the author may never have seen. `stroked="f"` — what Word writes when
 * the outline is turned off — says so outright.
 */
function shapeStroke(shape: Element): { color: string; weight: string } | undefined {
  const stroked = (shape.getAttribute("stroked") || "").toLowerCase();
  const stroke = Array.from(shape.children).find((el) => tag(el) === "v:stroke");
  if (stroked === "f" || (stroke?.getAttribute("on") || "").toLowerCase() === "f") return undefined;

  const color = (shape.getAttribute("strokecolor") || stroke?.getAttribute("color") || "").trim();
  const weight = (shape.getAttribute("strokeweight") || stroke?.getAttribute("weight") || "").trim();
  if (!color && !weight && stroked !== "t") return undefined;

  return {
    // `windowtext` is Word's black, and a theme colour drags its slot number
    // along behind it — see vmlColor.
    color: vmlColor(color) ?? "black",
    weight: /^[\d.]+\s*(pt|px|in|cm|mm)$/i.test(weight) ? weight : "0.75pt",
  };
}

/**
 * The gap the box keeps between its line and its text.
 *
 * VML states it as `left,top,right,bottom`; CSS wants `top right bottom left`.
 * Word's own default, when it states nothing, is 0.1in by 0.05in — which is
 * what a text box visibly has.
 */
function textboxPadding(box: Element): string {
  const inset = (box.getAttribute("inset") || "").split(",").map((part) => part.trim());
  const [left, top, right, bottom] = inset;

  return inset.length === 4 && left && top
    ? `${top} ${right || left} ${bottom || top} ${left}`
    : "0.05in 0.1in";
}

/** The width the box was drawn at, when its style states one. */
function shapeWidth(shape: Element): string | undefined {
  const width = (shape as HTMLElement).style?.width?.trim();

  return width && !/^0(px|pt)?$/.test(width) ? width : undefined;
}

/**
 * Word wraps a text box's content in a one-cell table for renderers that do not
 * speak VML (`<![if !mso]><table…>`), and that wrapper parses into real markup
 * the moment the comment is opened. It is scaffolding, not the report: kept, it
 * would put a stray borderless table around every recovered box.
 */
function unwrapSingleCellTable(root: HTMLElement): void {
  const children = Array.from(root.children);
  if (children.length !== 1 || tag(children[0]) !== "table") return;

  const cells = children[0].querySelectorAll("td, th");
  if (cells.length !== 1) return;

  root.innerHTML = cells[0].innerHTML;
}

/**
 * Give the box's content the blocks a band can be the width of.
 *
 * Word wraps it in a `<div>`, which is not one: Plate reads a fill off a div
 * onto the RUNS inside it, and a highlight as wide as the text is not the band
 * the box was. So a div that holds paragraphs is unwrapped, and one that holds
 * only runs becomes the paragraph it was standing in for.
 */
function blocksFor(root: HTMLElement): HTMLElement[] {
  for (const child of Array.from(root.children)) {
    if (child.tagName !== "DIV") continue;

    if (descendants(child).some((el) => BLOCK_TAGS.has(el.tagName))) {
      child.replaceWith(...Array.from(child.childNodes));
      continue;
    }

    const p = root.ownerDocument.createElement("p");
    p.append(...Array.from(child.childNodes));
    child.replaceWith(p);
  }

  const blocks = descendants(root).filter((el) => BLOCK_TAGS.has(el.tagName)) as HTMLElement[];
  if (blocks.length > 0) return blocks;

  // Bare runs, with nothing around them at all.
  const p = root.ownerDocument.createElement("p");
  p.append(...Array.from(root.childNodes));
  root.append(p);

  return [p];
}

/** Lay the shape's fill on those blocks, leaving one that states its own. */
function applyFill(root: HTMLElement, fill: string): void {
  for (const block of blocksFor(root)) {
    if (block.style?.backgroundColor || block.style?.background) continue;
    block.style.backgroundColor = fill;
  }
}

interface Recovery {
  /** The text boxes, as ordinary HTML. */
  html: string;
  /** The first text box's shape, which knows whether it was floating. */
  shape?: Element;
  /**
   * What is left of the comment once the text boxes are out of it, or null
   * when nothing worth keeping remains.
   *
   * A PICTURE is worth keeping, and this is the whole reason for it: Word
   * writes a pasted image as a shape too, and Plate matches that shape to the
   * image bytes in the clipboard's RTF flavour by its `o:spid`. One comment can
   * hold a text box and a picture together, so tearing the comment out whole to
   * get the box would take the picture's only identity with it and the image
   * would arrive as the `file:///` placeholder it can never load.
   */
  remaining: string | null;
}

/** The text boxes inside one conditional comment, and what is left behind. */
function recoverShapes(data: string): Recovery {
  // `<!--[if gte vml 1]>…<![endif]-->` — the comment's own conditional wrapper,
  // which has to go back on whatever is put back.
  const wrapper = data.match(/^(\s*\[if[^\]]*\]>)([\s\S]*)(<!\[endif\]\s*)$/i);
  const doc = new DOMParser().parseFromString(wrapper ? wrapper[2] : data, "text/html");

  const parts: string[] = [];
  let first: Element | undefined;
  for (const shape of descendants(doc.body).filter((el) => SHAPE_TAGS.has(tag(el)))) {
    const box = descendants(shape).find((el) => tag(el) === "v:textbox");
    if (!box) continue;

    const content = doc.createElement("div");
    content.innerHTML = box.innerHTML;
    unwrapSingleCellTable(content);

    const fill = shapeFill(shape);
    const stroke = shapeStroke(shape);

    // With a line around it the CELL carries the fill; laying it on the
    // paragraphs as well would paint a band inside the box's own padding.
    if (fill && !stroke) applyFill(content, fill);
    else blocksFor(content);

    const inner = content.innerHTML.trim();
    const html =
      inner && stroke
        ? borderedBoxHtml(inner, {
            border: `${stroke.weight} solid ${stroke.color}`,
            background: fill,
            padding: textboxPadding(box),
            width: shapeWidth(shape),
          })
        : inner;

    if (html) {
      parts.push(html);
      first = first ?? shape;
    }
    shape.remove();
  }

  if (parts.length === 0) return { html: "", remaining: data };

  const keeps = descendants(doc.body).some((el) => SHAPE_TAGS.has(tag(el)));
  const leftover = doc.body.innerHTML.trim();

  return {
    html: parts.join(""),
    shape: first,
    remaining: keeps ? (wrapper ? wrapper[1] + leftover + wrapper[3] : leftover) : null,
  };
}

/** Is this the opening marker of the picture fallback Word writes for a shape? */
function isFallbackStart(node: Node): boolean {
  return node.nodeType === Node.COMMENT_NODE && /^\s*\[if\s*!vml\]/i.test((node as Comment).data);
}

function isFallbackEnd(node: Node): boolean {
  return node.nodeType === Node.COMMENT_NODE && /^\s*\[endif\]/i.test((node as Comment).data);
}

/**
 * Drop `<![if !vml]>…<![endif]>` — the picture of the shape, at a `file:///`
 * path the browser cannot load. Now that the box itself has been recovered,
 * keeping it would leave a broken image beside the real thing.
 */
function removeFallback(from: Node): void {
  const parent = from.parentNode;
  if (!parent) return;

  let node: ChildNode | null = from as ChildNode;
  let dropping = false;
  while (node) {
    const next: ChildNode | null = node.nextSibling;
    if (isFallbackStart(node)) dropping = true;
    if (dropping) {
      const done = isFallbackEnd(node) && node !== from;
      parent.removeChild(node);
      if (done) return;
    }
    node = next;
  }
}

/**
 * Is this shape FLOATING — drawn over the page rather than sitting in the text?
 *
 * It decides what happens to the blank lines under it. A floating box hovers
 * above the paragraphs beneath it, and in Word those paragraphs are the room it
 * is hovering IN: nothing of them shows. Paste the box inline and it takes its
 * own space instead, so the blank lines it used to cover become a gap of their
 * own beneath it — the box's height and the space it floated over, one after
 * the other. That is the three or four blank lines that appear between a boxed
 * heading and the table under it, where the document has almost none.
 */
function isFloating(shape: Element): boolean {
  const style = (shape.getAttribute("style") || "").toLowerCase();

  return /position:\s*absolute/.test(style) || /mso-position-/.test(style);
}

/** A blank line in Word: a paragraph holding a non-breaking space and nothing else. */
function isBlankParagraph(node: ChildNode | null): node is HTMLElement {
  if (!node || node.nodeType !== Node.ELEMENT_NODE) return false;

  const el = node as HTMLElement;

  return el.tagName === "P" && isSpent(el);
}

/**
 * How many of those blank lines the box was actually covering.
 *
 * From the box's own height, so a short box takes back a short gap and a
 * deliberate space further down survives. A box that states no height covered
 * whatever followed it, which is the only honest reading of no information.
 */
function coveredLines(shape: Element): number {
  const height = ((shape as HTMLElement).style?.height || "").match(/^([\d.]+)pt$/i);
  if (!height) return Number.POSITIVE_INFINITY;

  // A blank line in a report is about 14pt: an 11pt run plus its leading.
  return Math.max(1, Math.round(Number.parseFloat(height[1]) / 14));
}

/** Is there nothing left in this element that a reader would see? */
function isSpent(el: Element): boolean {
  if (el.querySelector("img, table, hr, br")) return false;
  return !(el.textContent || "").replace(/\u00a0/g, " ").trim();
}

/**
 * The same HTML with every Word text box replaced by the content it holds.
 *
 * A no-op — the original string, byte for byte — for HTML that has no text box
 * in it, which is all HTML that did not come from Word.
 */
export function inlineWordTextboxes(html: string): string {
  if (!html || !/v:textbox/i.test(html)) return html;

  const doc = new DOMParser().parseFromString(html, "text/html");
  const walker = doc.createTreeWalker(doc.documentElement, NodeFilter.SHOW_COMMENT);

  const comments: Comment[] = [];
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    if (/v:textbox/i.test((node as Comment).data)) comments.push(node as Comment);
  }
  if (comments.length === 0) return html;

  let changed = false;

  for (const comment of comments) {
    const { html: recovered, remaining, shape } = recoverShapes(comment.data);
    if (!recovered) continue;

    const floating = shape ? isFloating(shape) : false;

    // The block the shape is anchored in — Word writes an empty paragraph of
    // its own for a text box, and that paragraph is what the recovered content
    // takes the place of.
    const host = comment.parentElement;
    const holder = doc.createElement("div");
    holder.innerHTML = recovered;

    const fragment = doc.createDocumentFragment();
    fragment.append(...Array.from(holder.childNodes));

    const anchor = host && host.tagName === "P" ? host : comment;
    anchor.parentNode?.insertBefore(fragment, anchor);
    changed = true;

    if (remaining !== null) {
      // A picture is still in there. Its VML stays, and so does the fallback
      // image the RTF flavour rewrites — see `Recovery.remaining`.
      comment.data = remaining;
      continue;
    }

    removeFallback(comment);

    // What follows the anchor, before the anchor goes: the blank lines the box
    // was floating over.
    let next: ChildNode | null = anchor.nextSibling;

    comment.remove();
    // Word's anchor paragraph holds the shape and nothing else. Left behind, it
    // would open a blank line above every recovered box.
    if (host && host.tagName === "P" && isSpent(host)) host.remove();

    if (!floating || !shape) continue;

    let budget = coveredLines(shape);
    while (budget > 0) {
      // Whitespace between Word's tags is not a line.
      while (next && next.nodeType === Node.TEXT_NODE && !(next.textContent || "").trim()) {
        next = next.nextSibling;
      }
      if (!isBlankParagraph(next)) break;

      const blank = next as HTMLElement;
      next = blank.nextSibling;
      blank.remove();
      budget -= 1;
    }
  }

  return changed ? doc.documentElement.outerHTML : html;
}

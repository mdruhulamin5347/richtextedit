/**
 * The font size a BLOCK is set in — which is what decides how tall its line is.
 *
 * Every line box carries a STRUT: an invisible inline box in the block's own
 * font. So a paragraph's height is not set by the text in it but by the size
 * the BLOCK is, and here that was always the editable's base — `text-[18px]`
 * on the `report` variant. A row of 10px text and a row of 15px text came out
 * the same height, both taller than the document, because in each of them the
 * strut was the tallest thing on the line. Measured against a browser rendering
 * the same markup: 24px / 18px / 15px / 10px text gave rows of 44 / 38 / 34 /
 * 28px there and 49 / 40 / 40 / 40px here.
 *
 * So a block is given the size its own runs agree on, and an empty one the size
 * in effect around it. Written onto the node, drawn by FontSizePlugin's
 * injection, and emitted by `blockStyle` — the printed report states
 * `body p {font-size: 15px}`, which an inline size outranks, so the page and
 * the screen stay the same shape.
 *
 * The blank-line half of this came first and is the same idea: an empty block
 * has no run to measure at all.
 *
 * A blank line takes the size IN EFFECT on it: the one it carries itself, else
 * the size of the line above, else of the line below. It carries one of its own
 * more often than it looks — Enter splits the run you were typing in, so the
 * new empty leaf keeps its size — and the line above covers the rest: a legacy
 * report's `<p><br></p>` states nothing at all.
 *
 * A block whose runs DISAGREE, or any one of which states no size, keeps the
 * editor's base: there is no single size to speak of, and the base is the
 * honest answer.
 *
 * Widening this from blank lines to every paragraph was asked for by name on
 * 2026-09-02, after the row-height question ("small text hole margin ta akoi
 * thakbe na text er opor dynamic hobe?"). It had been built once before across
 * the whole document and reverted; this time it was chosen deliberately, over
 * the narrower table-only option. See
 * ai/memory/radiolens-editor-paste-scope.md.
 */
import type { TElement, TText } from "platejs";

import { comparableFontSize } from "./font-size";

/** Blocks a blank line will take a size from — anything that is a line of text. */
const TEXT_BLOCKS: ReadonlySet<string> = new Set([
  "p",
  "h1",
  "h2",
  "h3",
  "h4",
  "h5",
  "h6",
  "blockquote",
]);

/**
 * A line with NOTHING on it.
 *
 * Newlines only. A line of tabs or of spaces is a line the report lays out
 * with (see lib/whitespace.ts) and `tab-size` is counted in character widths,
 * so resizing one would move the columns it aligns. The two forms that do count
 * are `""` — a line made with Enter — and `"\n"`, which is how a saved
 * `<p><br/></p>` reads back.
 */
const BLANK_TEXT = /^\n*$/;

function isElementNode(node: unknown): node is TElement {
  return !!node && typeof node === "object" && Array.isArray((node as TElement).children);
}

/** Is this an empty paragraph — a blank line in the report? */
export function isBlankLine(node: unknown): boolean {
  if (!isElementNode(node) || node.type !== "p") return false;
  return (node.children as (TElement | TText)[]).every(
    (child) => "text" in child && BLANK_TEXT.test(String(child.text))
  );
}

/**
 * The font size every run in a block agrees on.
 *
 * Undefined when they disagree, or when any one of them states none — that run
 * is at the editor's base, so the block has no single size to speak of.
 *
 * `skipEmpty` is the difference between the two questions asked here: what size
 * is this line's TEXT (empty leaves are Slate's padding around inlines and say
 * nothing), and what size is this BLANK line set in (empty leaves are all it
 * has).
 *
 * A run of pure WHITESPACE that states no size is skipped along with the empty
 * ones, and that is not a detail. Word writes a line's indentation as a run of
 * its own — `<span style='mso-spacerun:yes'>&nbsp;&nbsp; </span>` — and puts no
 * size on it, so an indented line looked like a paragraph whose runs disagreed
 * and fell back to the editor's base while its unindented neighbours did not:
 * in a three-signature footer the first column's lines stood 24px against the
 * other two columns' 20px, and its indentation was drawn in 18px spaces instead
 * of the document's. A space shows nothing and says nothing; it takes the size
 * of the line it sits on, exactly as it does in a browser.
 */
function agreedFontSize(node: TElement, skipEmpty: boolean): string | undefined {
  let agreed: string | undefined;
  let agreedPoints: string | undefined;

  const visit = (n: TElement | TText): boolean => {
    if ("text" in n) {
      const size = (n as { fontSize?: string }).fontSize;
      if (skipEmpty && !size && /^\s*$/.test(String(n.text))) return true;
      if (!size) return false;

      // Compared in POINTS, not as strings: a paragraph can hold text pasted
      // from Word (`10pt`) beside text a legacy report states in pixels
      // (`13.33px`), and those are one size, not two runs disagreeing. See
      // lib/font-size.ts.
      const points = comparableFontSize(size);
      if (agreedPoints !== undefined && points !== agreedPoints) return false;

      agreed ??= size;
      agreedPoints = points;
      return true;
    }
    return ((n as TElement).children as (TElement | TText)[]).every(visit);
  };

  return (node.children as (TElement | TText)[]).every(visit) ? agreed : undefined;
}

/**
 * The size a blank line states for itself.
 *
 * A line made with Enter holds one empty run carrying the size that was being
 * typed in; a saved one reopens as a run of `"\n"` carrying it. So ask the runs
 * that have something in them first, and fall back to the empty ones, which on
 * a blank line are the only ones there are.
 */
export function ownBlankLineFontSize(node: TElement): string | undefined {
  return agreedFontSize(node, true) ?? agreedFontSize(node, false);
}

/**
 * The size a neighbouring line is set in.
 *
 * A blank neighbour answers with the size it was itself given, so a run of
 * several empty lines between two paragraphs is one height rather than
 * alternating with the base.
 */
function lineFontSize(node: unknown): string | undefined {
  if (!isElementNode(node) || !TEXT_BLOCKS.has(String(node.type))) return undefined;
  return (node as { fontSize?: string }).fontSize ?? agreedFontSize(node, true);
}

/**
 * What size this block should be drawn and saved at, or undefined for the
 * editor's base — which is also the answer for every line that has text on it.
 */
export function blankLineFontSize(
  node: unknown,
  previous?: unknown,
  next?: unknown
): string | undefined {
  if (!isBlankLine(node)) return undefined;
  return ownBlankLineFontSize(node as TElement) ?? lineFontSize(previous) ?? lineFontSize(next);
}

/**
 * The size a paragraph should be DRAWN and SAVED at, so its line is as tall as
 * its text — or undefined for the editor's base, which is the answer whenever
 * the block has no one size of its own.
 *
 * A blank line asks its neighbours; a line with text asks its own runs and
 * nobody else. Two paragraphs set differently must not lean on each other, or
 * one line's edit would silently resize the next.
 */
export function blockFontSize(
  node: unknown,
  previous?: unknown,
  next?: unknown
): string | undefined {
  if (!isElementNode(node) || node.type !== "p") return undefined;
  return isBlankLine(node) ? blankLineFontSize(node, previous, next) : agreedFontSize(node, true);
}

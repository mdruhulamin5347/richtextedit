"use client";

import { createSlatePlugin, type NodeEntry, type SlateEditor, type TNode, KEYS } from "platejs";

import { BASE_FONT_PT, fontSizeToPt, toEditorFontSize, UNIT_TO_PT } from "@/lib/font-size";

/**
 * Properties to strip from block-level (element) nodes.
 *
 * These have no counterpart in the saved HTML — the serializer writes only
 * alignment, indent and line-height for a block — so keeping them would style
 * the editor with spacing the printed report never gets.
 *
 * The same reasoning is why `padding` is EXEMPT on a table cell: there the
 * serializer does write it, so it is not foreign spacing but the document's
 * own. See CELL_EXEMPT_PROPS.
 */
const ELEMENT_PROPS_TO_STRIP: readonly string[] = [
  "margin",
  "marginTop",
  "marginBottom",
  "marginLeft",
  "marginRight",
  "padding",
  "paddingTop",
  "paddingBottom",
  "paddingLeft",
  "paddingRight",
  "textIndent",
  "wordSpacing",
  "letterSpacing",
  "widows",
  "orphans",
];

/**
 * Properties a TABLE CELL keeps that any other block loses.
 *
 * A `<td>`'s padding is the only one of these the serializer can write back
 * out (`serializeCell` emits `padding: ${node.padding}`), so on a cell it is
 * not spacing the printed report will never get — it is the gap the report
 * prints WITH. The old Summernote editor kept it by doing nothing at all: it
 * dropped the clipboard HTML into a contenteditable and the browser obeyed
 * `padding:0cm 5.4pt` as written.
 *
 * Stripping it here is what undid `extractCellPadding` on the way in, so every
 * pasted cell fell back to a 12px/8px box and each row of a Word table gained
 * ~16px of height it never had in the document.
 *
 * Deliberately narrow: MARGIN is still stripped from a cell, exactly as the old
 * editor's `cleanWordSpacing` stripped it from everything.
 */
const CELL_EXEMPT_PROPS: ReadonlySet<string> = new Set([
  "padding",
  "paddingTop",
  "paddingBottom",
  "paddingLeft",
  "paddingRight",
]);

/**
 * Properties a block INSIDE a table cell keeps that the same block anywhere
 * else loses.
 *
 * The space above and below a paragraph is most of the gap between one table
 * row's text and the next, and the serializer writes it back out for a block in
 * a cell — so there it is not foreign spacing, it is the row's own height. See
 * `extractBlockSpacing`, and the exact same reasoning one comment up: stripping
 * it here is what would silently undo the extraction, as it once did the cell
 * padding's.
 *
 * Vertical only. `marginLeft` stays stripped because indent already owns it.
 */
const IN_CELL_EXEMPT_PROPS: ReadonlySet<string> = new Set(["marginTop", "marginBottom"]);

/**
 * Properties to strip from inline (text / leaf) nodes.
 *
 * `lineHeight` belongs to the paragraph, not to a run inside it: the serializer
 * only writes it on a block, and Word puts it on both.
 */
const TEXT_PROPS_TO_STRIP: readonly string[] = ["letterSpacing", "wordSpacing", "lineHeight"];

/**
 * A pasted font size as this editor stores one, or null when the value is not a
 * length at all (`smaller`, `inherit`, an empty string).
 *
 * The document's own unit is KEPT — see lib/font-size.ts, which owns the rule
 * and the reasoning. Word states type almost exclusively in POINTS, and those
 * used to be converted to whole px here: 10pt became `13px`, so the font-size
 * control named a 10pt run 13 and drew it a quarter of a pixel small. Before
 * that they were dropped outright and a 16pt heading, 11pt body and 9pt
 * footnote all rendered at the editor's flat base, which is why a pasted report
 * lost its hierarchy.
 */
export function toPtFontSize(raw: unknown): string | null {
  return toEditorFontSize(raw);
}

/**
 * A pasted line-height as a plain ratio.
 *
 * Three notations reach here. A bare number and a percentage are ratios
 * already. The third is ABSOLUTE LEADING — `line-height:12.0pt`, which is what
 * Word writes whenever a paragraph's spacing is set to "Exactly", beside
 * `mso-line-height-rule:exactly`. That one is a ratio only against the size the
 * text is set in, so it needs `fontSizePt` and is null without it. It used to
 * be dropped outright, so a tightly led document opened at whatever the editor
 * drew an unstated block at.
 *
 * The leading and the font size are both reduced to POINTS before they are
 * divided. A ratio has no unit, so any one unit on both sides gives the same
 * answer — points is simply the one this editor states type in.
 *
 * Absolute leading is NOT multiplied by the font's natural line the way a
 * percentage is (see lib/word-line-gap.ts): "Exactly 12pt" means twelve points
 * in Word and in CSS alike — there is no multiplier in it to correct.
 *
 * Still null for `normal`, which is not a ratio but an instruction to ask the
 * font, and which CSS already answers the same way Word's "Single" does.
 */
export function toLineHeightRatio(raw: unknown, fontSizePt?: number): number | null {
  const value = String(raw ?? "")
    .trim()
    .toLowerCase();
  const match = value.match(/^([\d.]+)\s*(%|px|pt|pc|in|cm|mm|em|rem)?$/);
  if (!match) return null;

  const n = Number.parseFloat(match[1]);
  if (!Number.isFinite(n) || n <= 0) return null;

  const unit = match[2];
  let ratio: number;

  if (!unit) {
    ratio = n;
  } else if (unit === "%") {
    ratio = n / 100;
  } else if (unit === "em" || unit === "rem") {
    // Already relative to the element's own size, which is what a ratio is.
    ratio = n;
  } else {
    if (!fontSizePt || !Number.isFinite(fontSizePt) || fontSizePt <= 0) return null;
    ratio = (n * (UNIT_TO_PT[unit] ?? 1)) / fontSizePt;
  }

  return ratio > 0 && ratio <= 10 ? ratio : null;
}

/**
 * The size, in points, the text of this block is set in — what an absolute
 * leading has to be measured against.
 *
 * The block's own `fontSize` when it has one, else the size its runs agree on,
 * else the editor's base. Read through `fontSizeToPt` because the value arrives
 * in whatever unit the document stated it in, and READ ONLY: writing a block's
 * font size here is what once deadlocked normalization, as the note in
 * `elementPatch` records.
 */
function blockFontSizePt(node: TNode): number {
  const own = (node as { fontSize?: string }).fontSize;
  let agreed: string | undefined = own;

  if (!agreed) {
    const visit = (n: any): void => {
      if (agreed) return;
      if ("text" in n) {
        if (n.fontSize && /\S/.test(String(n.text))) agreed = n.fontSize;
        return;
      }
      (n.children ?? []).forEach(visit);
    };
    ((node as any).children ?? []).forEach(visit);
  }

  return fontSizeToPt(agreed) ?? BASE_FONT_PT;
}

/**
 * What a pasted `lineHeight` should become, or undefined to drop one that means
 * nothing here.
 *
 * The document's OWN ratio, kept as stated. It used to be snapped to the
 * nearest of the five gaps the control offers, which cannot survive alongside
 * an exact line gap: a gap of 1.5 on Calibri IS the ratio 1.8306 (see
 * lib/line-gap.ts), and the nearest offered value to that is 2. Every exactly
 * spaced paragraph would have been rounded to a third more spacing than it
 * asked for, on the way in and again on every reopen.
 *
 * Rounded to four places for the same reason the conversion is: so a value
 * survives a save and reopen instead of drifting on each pass.
 */
function normalizedLineHeight(raw: unknown, fontSizePt?: number): number | undefined {
  const ratio = toLineHeightRatio(raw, fontSizePt);
  return ratio === null ? undefined : Math.round(ratio * 10000) / 10000;
}

/** Is this node a table cell, whose padding belongs to the document? */
function isTableCell(editor: SlateEditor, node: TNode | undefined): boolean {
  const type = (node as { type?: string } | undefined)?.type;
  if (!type) return false;
  return type === editor.getType(KEYS.td) || type === editor.getType(KEYS.th);
}

/** The changes one element node needs; empty when it is already conformant. */
function elementPatch(
  node: TNode,
  { isCell, inCell }: { isCell: boolean; inCell: boolean }
): Record<string, unknown> {
  const patch: Record<string, unknown> = {};

  for (const prop of ELEMENT_PROPS_TO_STRIP) {
    if (isCell && CELL_EXEMPT_PROPS.has(prop)) continue;
    if (inCell && IN_CELL_EXEMPT_PROPS.has(prop)) continue;
    // A paragraph spaced by the document keeps its space above and below for
    // the same reason, and the serializer writes it back out with its mark.
    if ((node as { documentSpacing?: boolean }).documentSpacing && IN_CELL_EXEMPT_PROPS.has(prop)) continue;
    if (prop in node) patch[prop] = undefined;
  }

  if ("lineHeight" in node) {
    const lineHeight = normalizedLineHeight((node as any).lineHeight, blockFontSizePt(node));
    if (lineHeight !== (node as any).lineHeight) patch.lineHeight = lineHeight;
  }

  // A block's `fontSize` is deliberately NOT touched here.
  //
  // It is not a pasted value: Plate's deserializer drops a font size stated on
  // an element outright, so nothing arrives carrying one. The only thing that
  // puts one on a block is BlockFontSizePlugin, which MIRRORS the size its runs
  // already carry — and those are normalized below, by `textPatch`.
  //
  // Converting it here meant the two normalizers fought over the same property
  // forever: this one rewrote the block's `11pt` to `15px`, the other read the
  // run again and wrote `11pt` back. Slate gave up after 42 passes per dirty
  // path ("Could not completely normalize the editor") and the editor rendered
  // nothing at all — a whole report gone, on any document whose sizes are in
  // points and are not converted before the block is derived from them.

  return patch;
}

/** The changes one text node needs; empty when it is already conformant. */
function textPatch(node: TNode): Record<string, unknown> {
  const patch: Record<string, unknown> = {};

  for (const prop of TEXT_PROPS_TO_STRIP) {
    if (prop in node) patch[prop] = undefined;
  }

  if ("fontSize" in node) {
    const fontSize = toPtFontSize((node as any).fontSize) ?? undefined;
    if (fontSize !== (node as any).fontSize) patch.fontSize = fontSize;
  }

  return patch;
}

/**
 * Conform pasted content to the editor's own options.
 *
 * Foreign spacing the saved report cannot carry is dropped, while the two
 * things a reader actually sees — how big the text is and how far apart the
 * lines sit — are KEPT instead. Both were being discarded, so every Word paste
 * came out at one flat size and one flat spacing, no matter what the document
 * said. A size keeps the document's own unit (lib/font-size.ts) and a
 * proportional line gap is converted to the CSS ratio that draws the same line
 * (lib/word-line-gap.ts).
 *
 * Runs through Slate's normalizeNode, after the HtmlPlugin's deserialization
 * rather than inside it. Every patch it writes is already conformant, so the
 * re-normalization it triggers produces an empty patch and stops.
 */
export const PasteNormalizationPlugin = createSlatePlugin({
  key: "pasteNormalization",
  extendEditor: ({ editor }) => {
    const originalNormalizeNode = editor.normalizeNode as (entry: NodeEntry) => void;

    editor.normalizeNode = (entry: NodeEntry) => {
      const [node, path] = entry;

      const isElement = "children" in node;
      const isText = "text" in node && !isElement;
      const patch = isElement
        ? elementPatch(node, {
            isCell: isTableCell(editor as SlateEditor, node),
            inCell: isTableCell(
              editor as SlateEditor,
              path.length > 0
                ? (editor as SlateEditor).api.node(path.slice(0, -1))?.[0]
                : undefined
            ),
          })
        : isText
          ? textPatch(node)
          : {};

      if (Object.keys(patch).length > 0) {
        (editor as SlateEditor).tf.setNodes(patch as any, { at: path });
        return;
      }

      originalNormalizeNode(entry);
    };

    return editor;
  },
});

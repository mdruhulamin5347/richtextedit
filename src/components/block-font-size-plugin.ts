"use client";

import { createSlatePlugin, type NodeEntry, type SlateEditor, type TElement } from "platejs";

import { blockFontSize } from "@/lib/block-font-size";

/**
 * Give every paragraph the size its own text is set in, and take it away again
 * from any that has no single size to speak of.
 *
 * The size goes on the BLOCK because that is what a line box is measured
 * against: every line carries a STRUT in the block's own font, so a paragraph
 * of 10px text sitting on the editable's 18px base is 18px tall however small
 * its text is. FontSizePlugin injects `fontSize` into paragraphs (see
 * components/font-kit.tsx), so writing it here draws it, and `blockStyle`
 * writes the same value into the saved HTML — where it outranks the print
 * page's own `body p {font-size: 15px}`, keeping the page and the screen the
 * same shape.
 *
 * `editor.children` is re-read each turn because `setNodes` replaces the node,
 * and consecutive blank lines take their size from the one above.
 */
function sizeBlocks(editor: SlateEditor): boolean {
  let changed = false;

  const visit = (children: readonly TElement[], path: number[]) => {
    for (let i = 0; i < children.length; i++) {
      const node = children[i];
      if (!node || !Array.isArray(node.children)) continue;

      if (node.type === "p") {
        const fontSize = blockFontSize(node, children[i - 1], children[i + 1]);
        if (fontSize !== (node as { fontSize?: string }).fontSize) {
          editor.tf.setNodes({ fontSize } as Partial<TElement>, { at: [...path, i] });
          changed = true;
        }
        continue;
      }

      // Into tables and list items too: a table ROW is exactly where this was
      // asked for, and a cell's paragraph is measured the same way as any other.
      visit(node.children as TElement[], [...path, i]);
    }
  };

  visit(editor.children as TElement[], []);

  return changed;
}

/**
 * Runs on the EDITOR node rather than on each paragraph: a blank line is sized
 * from the lines around it, a paragraph cannot see its neighbours, and editing
 * one line does not mark the next dirty. Slate walks the root on every change,
 * which is exactly the pass this needs.
 *
 * `normalizeInitialValue` covers opening a report: Plate only force-normalizes
 * at init when asked to, and this editor does not ask.
 */
export const BlockFontSizePlugin = createSlatePlugin({
  key: "blockFontSize",
  normalizeInitialValue: ({ editor }) => {
    sizeBlocks(editor);
  },
  extendEditor: ({ editor }) => {
    const originalNormalizeNode = editor.normalizeNode as (entry: NodeEntry) => void;

    editor.normalizeNode = (entry: NodeEntry) => {
      const [, path] = entry;

      // Re-normalization carries on from here: every setNodes marks the root
      // dirty again, so the pass repeats until nothing changes.
      if (path.length === 0 && sizeBlocks(editor as SlateEditor)) return;

      originalNormalizeNode(entry);
    };

    return editor;
  },
});

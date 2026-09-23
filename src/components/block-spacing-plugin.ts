"use client";

import { createSlatePlugin, KEYS, type TElement } from "platejs";

import { extractBlockSpacing } from "@/lib/table-widths";

/**
 * Carry a pasted paragraph's space-before / space-after onto the node — inside
 * a table CELL only.
 *
 * Most of the air between one table row's text and the next is not the cell's
 * padding, it is the paragraph's own margins: `margin-top:6.0pt` and its pair
 * are ordinary on a Word report's table. Plate's deserializer drops a block's
 * margins outright — in `faithful` mode too, so this is the parser's doing and
 * not the paste normalizer's — so a row standing 35px in the document arrived
 * here at 24px, and the saved report lost the same gap again.
 *
 * `extractBlockSpacing` returns nothing for a block outside a cell, which is
 * what keeps the editor's own paragraph spacing — settled ground — exactly
 * where it is. See the web app's editor paste-scope notes in ai/memory/.
 *
 * INJECTED into the paragraph plugin rather than configured on it: DocxPlugin
 * `override`s `p`'s deserializer parse outright with its own (the one that
 * reads Word's list indent), and an override replaces, so anything configured
 * there is simply discarded — which is exactly what happened on the first
 * attempt, silently. An injected parse is merged into the node whichever plugin
 * won the element, and it is how `align` and `lineHeight` reach a paragraph
 * through the same docx path.
 */
export const BlockSpacingPlugin = createSlatePlugin({
  key: "blockSpacing",
  inject: {
    plugins: {
      [KEYS.p]: {
        parsers: {
          html: {
            deserializer: {
              parse: ({ element }: { element: HTMLElement }) =>
                (extractBlockSpacing(element) ?? {}) as Partial<TElement>,
            },
          },
        },
      },
    },
  },
});

import { createSlatePlugin, KEYS } from "platejs";

import { inlineWordBorderBoxes } from "@/lib/word-border-box";

/**
 * Rebuild the boxes a document draws with Borders & Shading, on the way in.
 *
 * Registered beside the text-box pass and for the same reason: both run on the
 * rawest clipboard, before the docx cleaner and Juice rewrite the document
 * around them. A box is a box whichever of the two ways Word drew it.
 */
export const WordBorderBoxPlugin = createSlatePlugin({
  key: "wordBorderBox",
  inject: {
    plugins: {
      [KEYS.html]: {
        parser: {
          transformData: ({ data }: { data: string }) => inlineWordBorderBoxes(data),
        },
      },
    },
  },
});

import { createSlatePlugin, KEYS } from "platejs";

import { inlineWordTextboxes } from "@/lib/word-textbox";

/**
 * Recover Word's text boxes on the way in from the clipboard.
 *
 * LAST in the plugin list, which makes it FIRST to run — see the note on
 * ordering in plugins.ts. It has to see the RAWEST clipboard there is: the box
 * lives in a conditional comment, and both the docx cleaner and Juice rewrite
 * the document around it. Once the box is ordinary HTML, every other pass —
 * alignment, whitespace, colour, size, shading — treats it as what it now is.
 */
export const WordTextboxPlugin = createSlatePlugin({
  key: "wordTextbox",
  inject: {
    plugins: {
      [KEYS.html]: {
        parser: {
          transformData: ({ data }: { data: string }) => inlineWordTextboxes(data),
        },
      },
    },
  },
});

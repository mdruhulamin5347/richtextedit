import { createSlatePlugin, KEYS } from "platejs";

import { inlineWordLineGap } from "@/lib/word-line-gap";

/**
 * Correct Word's and LibreOffice's proportional line spacing on the way in.
 *
 * Registered at the FRONT of the plugin list, which makes it LAST to run:
 * `transformData` is piped in reverse registration order, and this pass has to
 * see the HTML after Juice, because a Word document states both the spacing
 * (`p.MsoNormal {line-height:115%}`) and the font it is measured against in a
 * `<style>` block rather than inline. The same reason InheritedFontSizePlugin
 * sits beside it. See lib/word-line-gap.ts.
 */
export const WordLineGapPlugin = createSlatePlugin({
  key: "wordLineGap",
  inject: {
    plugins: {
      [KEYS.html]: {
        parser: {
          transformData: ({ data }: { data: string }) => inlineWordLineGap(data),
        },
      },
    },
  },
});

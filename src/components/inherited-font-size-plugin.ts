import { createSlatePlugin, KEYS } from "platejs";

import { inlineInheritedFontSize } from "@/lib/inherited-font-size";

/**
 * Restate an inherited font size on the runs, on the way in from the clipboard.
 *
 * Hooks the HTML plugin's `transformData`, the same seam DocxPlugin, JuicePlugin
 * and LegacyAlignmentPlugin use. Registered after those three so it works on the
 * HTML they have already normalised — which matters most for Juice, since a
 * Word paste states its body size in a `<style>` block (`p.MsoNormal
 * {font-size:11.0pt}`) that only becomes an inline style once Juice has run.
 */
export const InheritedFontSizePlugin = createSlatePlugin({
  key: "inheritedFontSize",
  inject: {
    plugins: {
      [KEYS.html]: {
        parser: {
          transformData: ({ data }: { data: string }) => inlineInheritedFontSize(data),
        },
      },
    },
  },
});

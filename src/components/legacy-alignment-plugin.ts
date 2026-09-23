import { createSlatePlugin, KEYS } from "platejs";

import { inlineLegacyAlignment } from "@/lib/legacy-alignment";

/**
 * Restate pre-CSS alignment as CSS, on the way in from the clipboard.
 *
 * Hooks the HTML plugin's `transformData`, the same seam DocxPlugin and
 * JuicePlugin use to clean up what Word puts on the clipboard. Registered after
 * both so it runs on the HTML they have already normalised — and it hands back
 * the whole document, so the `<style>` block Juice depends on survives whatever
 * the order turns out to be.
 */
export const LegacyAlignmentPlugin = createSlatePlugin({
  key: "legacyAlignment",
  inject: {
    plugins: {
      [KEYS.html]: {
        parser: {
          transformData: ({ data }: { data: string }) => inlineLegacyAlignment(data),
        },
      },
    },
  },
});

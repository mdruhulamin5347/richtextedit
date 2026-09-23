import { createSlatePlugin, KEYS } from "platejs";

import { inlineInheritedColor } from "@/lib/inherited-color";

/**
 * Restate an inherited text colour on the runs, on the way in from the
 * clipboard.
 *
 * The colour twin of InheritedFontSizePlugin, and registered beside it for the
 * same reason: a Word document states its body colour in a `<style>` block
 * (`p.MsoNormal {color:#C00000}`), which is not an inline style — the only
 * thing a Plate node can be read from — until JuicePlugin has run.
 */
export const InheritedColorPlugin = createSlatePlugin({
  key: "inheritedColor",
  inject: {
    plugins: {
      [KEYS.html]: {
        parser: {
          transformData: ({ data }: { data: string }) => inlineInheritedColor(data),
        },
      },
    },
  },
});

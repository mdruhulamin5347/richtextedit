import { createSlatePlugin, KEYS } from "platejs";

import { protectWhitespace } from "@/lib/whitespace";

/**
 * Keep the whitespace a pasted document lays itself out with.
 *
 * Hooks the HTML plugin's `transformData`, the same seam DocxPlugin, JuicePlugin
 * and LegacyAlignmentPlugin use. Registered last of the four so it sees the HTML
 * they have already normalised — Juice in particular can only inline Word's
 * `<style>` block while the document is still intact, and this hands the whole
 * document back for the same reason.
 *
 * Why it is needed at all: Plate collapses whitespace itself when deserializing,
 * and its rule counts U+00A0 as collapsible where CSS does not. Word writes its
 * tabs as spans full of `&nbsp;`, and so did every report the old Summernote
 * editor saved — all of which arrived here with the gaps closed up. See
 * lib/whitespace.ts.
 */
export const WhitespacePlugin = createSlatePlugin({
  key: "whitespacePreservation",
  inject: {
    plugins: {
      [KEYS.html]: {
        parser: {
          transformData: ({ data }: { data: string }) => protectWhitespace(data),
        },
      },
    },
  },
});

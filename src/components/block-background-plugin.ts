"use client";

import { createSlatePlugin, KEYS, type TElement } from "platejs";

import { statedBackground } from "@/lib/background";

/**
 * Carry a pasted paragraph's or heading's own SHADING onto the node.
 *
 * Word calls it paragraph shading and reports are full of it — a section
 * heading laid on a grey band, a caution line on pale yellow. It is a fill on
 * the BLOCK, and Plate reads a fill onto a text leaf only: the deserializer
 * drops the declaration from anything that maps to a Plate element, so
 * `<p style="background:#D9E2F3">` arrived plain white and, since the
 * serializer wrote no block fill either, the first save took the band out of
 * the report for good. Exactly the story of a block's colour and size — see
 * lib/inherited-color.ts — except that a band cannot be restated on the runs
 * without changing what it IS: paragraph shading runs the full width of the
 * block, a highlight only as wide as the text.
 *
 * Which is why this is a block property rather than a mark. Plate renders it
 * already: FontKit injects `backgroundColor` into the paragraph and heading
 * plugins (`components/font-kit.tsx`), so a node that carries one is drawn with
 * it, and `blockStyle` in lib/html-serializer.ts writes it back out.
 *
 * Paragraphs and headings only, deliberately — those are the two the injection
 * covers, and a fill the editor could not draw but the print page could would
 * put the two out of step.
 *
 * INJECTED rather than configured, for the reason spelled out in
 * components/block-spacing-plugin.ts: DocxPlugin `override`s the paragraph
 * deserializer outright, and an override replaces.
 */
const blockParsers = {
  parsers: {
    html: {
      deserializer: {
        parse: ({ element }: { element: HTMLElement }) => {
          const background = statedBackground(element);
          return (background ? { backgroundColor: background } : {}) as Partial<TElement>;
        },
      },
    },
  },
};

export const BlockBackgroundPlugin = createSlatePlugin({
  key: "blockBackground",
  inject: {
    plugins: Object.fromEntries(
      [KEYS.p, ...KEYS.heading].map((key) => [key, blockParsers])
    ),
  },
});

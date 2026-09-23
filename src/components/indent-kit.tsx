"use client";

import { IndentPlugin } from "@platejs/indent/react";
import { KEYS } from "platejs";

export const IndentKit = [
  IndentPlugin.configure({
    inject: {
      targetPlugins: [
        ...KEYS.heading,
        KEYS.p,
        KEYS.blockquote,
        KEYS.codeBlock,
        KEYS.toggle,
        KEYS.img,
      ],
    },
    options: {
      /**
       * One indent level, in px — and it has to be the SERIALIZER's number.
       *
       * `blockStyle` writes `margin-left: indent * 40px` (lib/html-serializer.ts)
       * and the deserializer reads a level back at the same 40px, so a report
       * round-trips unchanged. This is only what the editor DRAWS, and at 24 it
       * drew every indent a little under two-thirds of the width the same
       * document printed at. Measured: levels 1/2/3 rendered 24/48/72px against
       * the stored — and printed — 40/80/120px.
       */
      offset: 40,
    },
  }),
];

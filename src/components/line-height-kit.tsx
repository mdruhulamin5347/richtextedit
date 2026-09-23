"use client";

import { LineHeightPlugin } from "@platejs/basic-styles/react";
import { KEYS } from "platejs";

/**
 * The line gaps the control OFFERS, in Word's and LibreOffice's sense: a
 * multiple of the font's natural line. See lib/line-gap.ts.
 */
export const LINE_GAPS = [1, 1.2, 1.5, 2, 3] as const;

/** What the control shows for a paragraph that states no gap of its own. */
export const DEFAULT_LINE_GAP = 1.5;

export const LineHeightKit = [
  LineHeightPlugin.configure({
    inject: {
      nodeProps: {
        // Both gates are off, and neither is an oversight.
        //
        // A node now carries the CSS ratio a line gap WORKS OUT TO — `1.5` on
        // Calibri is stored as `1.8` — so a list of five permitted values would
        // reject nearly every one of them. Plate's `pluginInjectNodeProps`
        // refuses to inject a value that is not in `validNodeValues`, so the
        // paragraph would simply be drawn with no `line-height` at all.
        //
        // `defaultNodeValue` is the same trap one step along: Plate ALSO skips
        // the injection when a value equals it, so declaring 1.5 there meant a
        // paragraph set to exactly 1.5 was drawn with nothing. Outside a table
        // that hid, because the editable's own `leading-normal` is 1.5 too;
        // inside a cell `leading-[normal]` is CSS `normal`, so it did not.
        // `configure` MERGES, so leaving the field out keeps
        // BaseLineHeightPlugin's own 1.5 — it has to be overridden with a value
        // no ratio can equal.
        defaultNodeValue: null as unknown as number,
        validNodeValues: undefined,
      },
      targetPlugins: [...KEYS.heading, KEYS.p],
    },
  }),
];

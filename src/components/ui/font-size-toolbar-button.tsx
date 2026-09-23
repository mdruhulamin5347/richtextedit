"use client";

import * as React from "react";

import type { TElement } from "platejs";

import { FontSizePlugin } from "@platejs/basic-styles/react";
import { Minus, Plus } from "lucide-react";
import { KEYS } from "platejs";
import { useEditorPlugin, useEditorSelector } from "platejs/react";

import { isBlankLine } from "@/lib/block-font-size";
import {
  BASE_FONT_PT,
  fontSizeLabel,
  formatFontSizeNumber,
  MAX_FONT_PT,
  MIN_FONT_PT,
  pointsToFontSize,
} from "@/lib/font-size";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { cn } from "@/lib/utils";

import { ToolbarButton } from "./toolbar";

/**
 * This control reads and writes POINTS — the scale Word and LibreOffice state
 * type on, and the scale a report's author names a size in.
 *
 * It used to work in whole CSS pixels, so a paragraph the document set in 10pt
 * showed as 13 (10pt is 13.333px, rounded). Reported as "main doc has font size
 * 10 but after paste our editor show 13". The number is now the size in points
 * whatever unit the run states it in, so a legacy report stating `15px` reads
 * 11.25 — which is what 15px is — and a pasted `10pt` reads 10.
 *
 * lib/font-size.ts owns the conversion and the reasoning.
 */

/** The editor's own text size, as this control names it: `text-[18px]` = 13.5pt. */
const DEFAULT_FONT_SIZE = formatFontSizeNumber(BASE_FONT_PT);

/**
 * What a heading is drawn at when no run in it states a size, in points — the
 * 36 / 24 / 20px this editor draws h1 / h2 / h3 at.
 */
const HEADING_FONT_SIZE = {
  h1: formatFontSizeNumber(27),
  h2: formatFontSizeNumber(18),
  h3: formatFontSizeNumber(15),
} as const;

/**
 * Word's own font-size list, which is the point of it: a report is written in
 * one application and finished in this one, so the sizes offered here are the
 * sizes offered there.
 */
const FONT_SIZES = [
  "8",
  "9",
  "10",
  "11",
  "12",
  "14",
  "16",
  "18",
  "20",
  "22",
  "24",
  "26",
  "28",
  "36",
  "48",
  "72",
] as const;

export function FontSizeToolbarButton() {
  const [inputValue, setInputValue] = React.useState(DEFAULT_FONT_SIZE);
  const [isFocused, setIsFocused] = React.useState(false);
  const { editor, tf } = useEditorPlugin(FontSizePlugin);

  const cursorFontSize = useEditorSelector((editor) => {
    const fontSize = editor.api.marks()?.[KEYS.fontSize];

    if (fontSize) {
      return fontSizeLabel(fontSize);
    }

    const [block] = editor.api.block<TElement>() || [];

    if (!block?.type) return DEFAULT_FONT_SIZE;

    return block.type in HEADING_FONT_SIZE
      ? HEADING_FONT_SIZE[block.type as keyof typeof HEADING_FONT_SIZE]
      : DEFAULT_FONT_SIZE;
  }, []);

  /**
   * A caret on a BLANK line has no run to hold the mark, so Slate keeps it
   * pending — it applies to whatever is typed next and the empty line itself
   * never changes height. Written onto the empty run instead, which is what
   * BlankLineSizePlugin measures the line with, and what the saved HTML keeps.
   */
  const applyFontSize = (points: number) => {
    const size = pointsToFontSize(points);
    tf.fontSize.addMark(size);

    const [block, path] = editor.api.block<TElement>() ?? [];
    if (block && path && isBlankLine(block)) {
      editor.tf.setNodes({ fontSize: size }, { at: path, match: (n) => "text" in n });
    }
  };

  const handleInputChange = () => {
    const points = Number.parseFloat(inputValue);

    // The same range lib/font-size.ts accepts, deliberately: a size this let
    // somebody type but that was rejected on reopen would be lost on the next
    // save.
    if (!Number.isFinite(points) || points < MIN_FONT_PT || points > MAX_FONT_PT) {
      editor.tf.focus();

      return;
    }
    if (formatFontSizeNumber(points) !== cursorFontSize) {
      applyFontSize(points);
    }

    editor.tf.focus();
  };

  /** Steps by a whole POINT, which is what the number shown is measured in. */
  const handleFontSizeChange = (delta: number) => {
    const points = Number.parseFloat(displayValue) + delta;
    if (points < MIN_FONT_PT || points > MAX_FONT_PT) return;

    applyFontSize(points);
    editor.tf.focus();
  };

  const displayValue = isFocused ? inputValue : cursorFontSize;

  return (
    <div className="bg-muted/60 flex h-7 items-center gap-1 rounded-md p-0">
      <ToolbarButton onClick={() => handleFontSizeChange(-1)}>
        <Minus />
      </ToolbarButton>

      <Popover open={isFocused} modal={false}>
        <PopoverTrigger asChild>
          <input
            className={cn(
              "hover:bg-muted h-full w-10 shrink-0 bg-transparent px-1 text-center text-sm"
            )}
            value={displayValue}
            onBlur={() => {
              setIsFocused(false);
              handleInputChange();
            }}
            onChange={(e) => setInputValue(e.target.value)}
            onFocus={() => {
              setIsFocused(true);
              setInputValue(cursorFontSize);
            }}
            onKeyDown={(e) => {
              if (e.key === "Enter") {
                e.preventDefault();
                handleInputChange();
              }
            }}
            data-plate-focus="true"
            type="text"
          />
        </PopoverTrigger>
        <PopoverContent
          className="max-h-72 w-10 overflow-y-auto px-px py-1"
          onOpenAutoFocus={(e) => e.preventDefault()}
        >
          {FONT_SIZES.map((size) => (
            <button
              key={size}
              className={cn(
                "hover:bg-accent data-[highlighted=true]:bg-accent flex h-8 w-full items-center justify-center text-sm"
              )}
              onClick={() => {
                applyFontSize(Number.parseFloat(size));
                setIsFocused(false);
              }}
              data-highlighted={size === displayValue}
              type="button"
            >
              {size}
            </button>
          ))}
        </PopoverContent>
      </Popover>

      <ToolbarButton onClick={() => handleFontSizeChange(1)}>
        <Plus />
      </ToolbarButton>
    </div>
  );
}

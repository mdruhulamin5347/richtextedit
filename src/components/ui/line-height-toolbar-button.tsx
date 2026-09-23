"use client";

import * as React from "react";

import type { DropdownMenuProps } from "@radix-ui/react-dropdown-menu";

import { LineHeightPlugin } from "@platejs/basic-styles/react";
import { DropdownMenuItemIndicator } from "@radix-ui/react-dropdown-menu";
import { CheckIcon, WrapText } from "lucide-react";
import { useEditorRef, useSelectionFragmentProp } from "platejs/react";

import { DEFAULT_LINE_GAP, LINE_GAPS } from "@/components/line-height-kit";
import {
  cssRatioToLineGap,
  lineGapToCssRatio,
  naturalLineHeightOfElement,
} from "@/lib/line-gap";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuRadioGroup,
  DropdownMenuRadioItem,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";

import { ToolbarButton } from "./toolbar";

/**
 * The line-gap control, which speaks Word's language on both sides.
 *
 * What the reader picks is a Word/LibreOffice line gap — a multiple of the
 * font's natural line — and what the document stores is the CSS ratio that
 * works out to. The conversion lives here and nowhere else: past this button
 * the node, the serializer and the printed page all speak plain CSS, so a
 * report is a ratio and needs nothing at render time to be drawn right.
 * See lib/line-gap.ts.
 */
export function LineHeightToolbarButton(props: DropdownMenuProps) {
  const editor = useEditorRef();
  const stored = useSelectionFragmentProp({ getProp: (node) => node.lineHeight });

  const [open, setOpen] = React.useState(false);

  /**
   * The natural line of the font the caret is sitting in, read from the DOM.
   *
   * Not from the node: a pasted paragraph states its font on its RUNS and not
   * on itself, so asking the block answered "the editor's own" and every gap
   * came out measured against the UI stack (1.362) instead of the document's
   * Calibri (1.221) — 11% wrong at every step, which is most of the error this
   * whole conversion exists to remove.
   *
   * Read when the menu OPENS: the selection cannot move while it is up, and a
   * `getComputedStyle` on every render would be for nothing.
   */
  const [natural, setNatural] = React.useState(() => naturalLineHeightOfElement(null));

  /**
   * The gap the caret's line is actually DRAWN at.
   *
   * The control used to fall back to a flat 1.5 whenever the selection had no
   * single stored value — which is every paragraph that states no gap of its
   * own — so it read 1.5 whatever was on screen. A line always has a height,
   * though, and the browser has already worked it out: the ratio it is drawn
   * at, over the font's natural line, IS the gap in Word's sense.
   */
  const [effective, setEffective] = React.useState<number | undefined>(undefined);

  const readNatural = React.useCallback(() => {
    // Slate's own node, not `window.getSelection()`: opening the menu moves
    // focus into it, so by the time this runs the DOM selection is the menu's
    // and the font read off it is the toolbar's.
    let element: Element | null = null;
    try {
      const block = (editor as any).api?.block?.();
      element = block ? ((editor as any).api?.toDOMNode?.(block[0]) ?? null) : null;
      if (!element) element = (editor as any).api?.toDOMNode?.(editor) ?? null;
    } catch {
      element = null;
    }

    // The font is read off the RUN, not the block. Plate drops a `font-family`
    // stated on a `<p>` at parse — only a text leaf keeps one — so a pasted
    // paragraph's block computes to the editor's own UI stack (natural 1.362)
    // while its text is the document's Calibri (1.221). Measuring the block
    // made every gap 11% too loose, which is most of the error this conversion
    // exists to remove.
    const run = element?.querySelector("[data-slate-string]") ?? null;
    const measuredNatural = naturalLineHeightOfElement(run ?? element);
    setNatural(measuredNatural);

    // What the line is drawn at. `line-height: normal` computes to the word
    // `normal` in Chrome rather than to a length, and that IS the natural
    // line — a gap of exactly 1.
    if (element && typeof window !== "undefined") {
      const style = window.getComputedStyle(element);
      const size = Number.parseFloat(style.fontSize);
      const drawn =
        style.lineHeight === "normal"
          ? measuredNatural * size
          : Number.parseFloat(style.lineHeight);
      setEffective(
        Number.isFinite(drawn) && Number.isFinite(size) && size > 0
          ? drawn / size / measuredNatural
          : undefined
      );
    } else {
      setEffective(undefined);
    }
  }, [editor]);

  /**
   * The gap to show as selected: the one the document states, else the one the
   * line is drawn at, snapped to the nearest the control offers so that a tick
   * always lands on an item. Never a flat default.
   */
  const measured = stored === undefined ? effective : cssRatioToLineGap(Number(stored), natural);

  const selected =
    measured === undefined || !Number.isFinite(measured)
      ? DEFAULT_LINE_GAP
      : LINE_GAPS.reduce((best, gap) =>
          Math.abs(gap - measured) < Math.abs(best - measured) ? gap : best
        );

  return (
    <DropdownMenu
      open={open}
      onOpenChange={(next) => {
        if (next) readNatural();
        setOpen(next);
      }}
      modal={false}
      {...props}
    >
      <DropdownMenuTrigger asChild>
        <ToolbarButton pressed={open} tooltip="Line height" isDropdown>
          <WrapText />
        </ToolbarButton>
      </DropdownMenuTrigger>

      <DropdownMenuContent className="min-w-0" align="start">
        <DropdownMenuRadioGroup
          value={String(selected)}
          onValueChange={(newValue) => {
            editor
              .getTransforms(LineHeightPlugin)
              .lineHeight.setNodes(lineGapToCssRatio(Number(newValue), natural));
            editor.tf.focus();
          }}
        >
          {LINE_GAPS.map((gap) => (
            <DropdownMenuRadioItem
              key={gap}
              className="min-w-[180px] pl-2 *:first:[span]:hidden"
              value={String(gap)}
            >
              <span className="pointer-events-none absolute right-2 flex size-3 items-center justify-center">
                <DropdownMenuItemIndicator>
                  <CheckIcon />
                </DropdownMenuItemIndicator>
              </span>
              {gap}
            </DropdownMenuRadioItem>
          ))}
        </DropdownMenuRadioGroup>
      </DropdownMenuContent>
    </DropdownMenu>
  );
}

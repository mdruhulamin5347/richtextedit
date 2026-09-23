"use client";

import * as React from "react";

import type { PlateElementProps } from "platejs/react";

import { PlateElement } from "platejs/react";

import { cn } from "@/lib/utils";

export function ParagraphElement(props: PlateElementProps) {
  /**
   * The space the DOCUMENT puts above and below this paragraph.
   *
   * Only a paragraph inside a table cell ever carries it — `extractBlockSpacing`
   * reads it nowhere else — so the editor's own paragraph spacing outside
   * tables is untouched by construction, and `m-0` below still decides it. An
   * inline style outranks the class, which is what lets a row be as deep as the
   * document set it.
   */
  const { marginTop, marginBottom, indent } = props.element as {
    marginTop?: string;
    marginBottom?: string;
    indent?: number;
  };

  const spacing = marginTop || marginBottom;

  return (
    <PlateElement
      {...props}
      // `m-0` is dropped rather than overridden when the document states a gap:
      // build/scope-editor-css.mjs marks a declaration `!important` wherever
      // Bootstrap uses that class name too, and Bootstrap's spacing utilities
      // are all important — so `.m-0` beat the inline style and the row stayed
      // flat with the right margins written on it. The scoped preflight still
      // zeroes every margin, so nothing is left to a browser default.
      //
      // An INDENT is the same trap on the other axis, and it was missed. Plate's
      // IndentPlugin states the level as an inline `margin-left`, so `.m-0` beat
      // that too and an indented paragraph drew FLAT — while the serializer wrote
      // `margin-left: indent * 40px` all along, so the report saved and printed
      // with an indentation the author was never shown. Measured before the fix:
      // three paragraphs at levels 1/2/3 all computed `margin-left: 0px`.
      //
      // `my-0 me-0` rather than dropping the class outright: this editor loads
      // Bootstrap, whose `p { margin-bottom: 1rem }` is UNLAYERED and so beats the
      // scoped preflight at any specificity — leaving an indented paragraph with a
      // 16px gap under it that the printed report does not have. Both of those are
      // Bootstrap utility names as well, so both come out `!important` and win,
      // and neither says anything about `margin-left`.
      className={cn("px-0 py-1", !spacing && (indent ? "my-0 me-0" : "m-0"))}
      style={spacing ? { marginTop, marginBottom } : undefined}
    >
      {props.children}
    </PlateElement>
  );
}

/**
 * A box drawn around text, as the one thing this editor can draw one with.
 *
 * A paragraph carries no border. A table CELL carries the line, the fill and
 * the inset together, and every part of it is read straight back off the HTML
 * by `extractCellBorders`, `extractCellBackground` and `extractCellPadding` in
 * lib/table-widths.ts. So a box of any origin — Word's floating text box, its
 * Borders & Shading on a paragraph, LibreOffice's bordered paragraph — is
 * handed to the paste pipeline as a one-cell table and needs no new node type,
 * no new serializer branch and no new print rule.
 */

export interface BoxStyle {
  /** The line, as CSS shorthand: `1pt solid #2e74b5`. */
  border: string;
  background?: string;
  padding?: string;
  width?: string;
}

/** Word's own default text-box inset: 0.1in across, 0.05in down. */
export const DEFAULT_BOX_PADDING = "0.05in 0.1in";

/**
 * The line an element draws, rebuilt from the LONGHANDS.
 *
 * Not from the `border` shorthand, which cannot be trusted to say what it was
 * given: `border: none` comes back out of the CSSOM as `medium` — the width,
 * with the `none` that was the whole point of it dropped — so reading the
 * shorthand drew a box around every paragraph Word had explicitly told not to
 * draw one. That is the paragraph inside Word's own border div, on every boxed
 * heading there is.
 */
export function drawnBorderOf(element: HTMLElement): string | undefined {
  const style = element.style;
  if (!style) return undefined;

  const lineStyle = (style.borderTopStyle || style.borderStyle || "").trim();
  if (!lineStyle || /^(none|hidden)$/i.test(lineStyle)) return undefined;

  const width = (style.borderTopWidth || style.borderWidth || "").trim();
  if (/^0(\.0+)?(px|pt|in|cm|mm|em)?$/i.test(width)) return undefined;

  const color = (style.borderTopColor || style.borderColor || "").trim();

  return `${width || "1px"} ${lineStyle} ${color || "black"}`;
}

/** The box an element's own style describes, or undefined when it draws none. */
export function boxStyleOf(element: HTMLElement): BoxStyle | undefined {
  const border = drawnBorderOf(element);
  if (!border) return undefined;

  const style = element.style;

  const background = (style.backgroundColor || style.background || "").trim();

  return {
    border,
    background: background && !/^(transparent|none)$/i.test(background) ? background : undefined,
    padding: (style.padding || "").trim() || undefined,
    width: (style.width || "").trim() || undefined,
  };
}

/** The box as markup: one row, one cell, the content inside it. */
export function borderedBoxHtml(content: string, box: BoxStyle): string {
  const cell = [
    `border: ${box.border}`,
    box.background ? `background: ${box.background}` : "",
    `padding: ${box.padding || DEFAULT_BOX_PADDING}`,
  ]
    .filter(Boolean)
    .join("; ");

  return (
    `<table border="1" style="border-collapse: collapse; width: ${box.width || "100%"}">` +
    `<tr><td style="${cell}">${content}</td></tr></table>`
  );
}

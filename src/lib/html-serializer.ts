/**
 * HTML <-> Plate conversion for the Laravel integration.
 *
 * WHY NOT `serializeHtml` FROM `platejs/static`:
 * that serializer emits class-based markup that only renders correctly with the
 * Tailwind stylesheet linked alongside it. Everything downstream here — the
 * print blades, dompdf, the patient portal — embeds `report_body` raw and has
 * no such stylesheet, and dompdf supports only a small CSS subset. So HTML is
 * produced the way Summernote produced it: self-contained, inline-styled, plain
 * tags.
 *
 * This is a hardened version of the serializer in RadioLens's
 * `rich-text-editor.tsx`. The additions are all round-trip fixes — that
 * serializer silently dropped table column widths, cell borders/shading, nested
 * lists, hard breaks and every non-image media node.
 */
import type { TElement, TText, Value } from "platejs";
import { ownBlankLineFontSize } from "./block-font-size";
import { serializeWhitespace } from "./whitespace";
import {
  borderShorthand,
  DEFAULT_CELL_PADDING,
  hasVisibleBorder,
  nodeCellSpan,
  type CellBorders,
} from "./table-widths";

/**
 * Applied only to cells that say nothing about their own borders — which is
 * what a table created inside the editor looks like, and matches the 1px line
 * Plate draws for it. A cell that arrived with explicit borders (including
 * explicitly NONE, as Word writes for the borderless tables it uses to align
 * label/value columns) is serialized from those instead, so the printed report
 * never shows a grid the source document did not have.
 */
const DEFAULT_CELL_BORDER = borderShorthand({ size: 1 });

const CELL_SIDES = ["top", "right", "bottom", "left"] as const;

/** No rule on any edge — width zero AND style none, so neither is inherited. */
const NO_BORDER_STYLE = "border: 0";

/**
 * Border declarations for one cell, faithful to what it actually carries.
 *
 * Written per side from the same `borderShorthand` the editor draws with, so
 * the width, style and colour on screen are the ones the report prints.
 *
 * The edges that carry NO rule are stated too, as a leading `border: 0`. That
 * is not noise: the print pages include `_partials/print_richtext_fixes`, whose
 * `.report_body table td { border-style: solid }` supplies the style Word
 * leaves off its own markup — and it lands on every side this serializer stays
 * silent about, where a missing width then resolves to CSS's `medium`. So an
 * edge left unsaid does not print as nothing, it prints as a 3px black line,
 * and a table with a frame but no inner grid came out of the printer boxed in
 * one. An inline declaration outranks that stylesheet; the longhands after it
 * outrank the shorthand.
 */
function cellBorderStyles(borders: CellBorders | undefined): string[] {
  if (!borders) return [`border: ${DEFAULT_CELL_BORDER}`];

  const drawn = CELL_SIDES.filter((side) => (borders[side]?.size ?? 0) > 0);
  if (drawn.length === 0) return [NO_BORDER_STYLE];

  return [NO_BORDER_STYLE, ...drawn.map((side) => `border-${side}: ${borderShorthand(borders[side])}`)];
}

/** Does any cell anywhere in this table carry a visible border? */
function tableHasBorders(table: any): boolean {
  for (const row of table?.children ?? []) {
    for (const cell of row?.children ?? []) {
      if (hasVisibleBorder(cell?.borders)) return true;
    }
  }
  return false;
}

export function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

export function isPlateJson(str: string): boolean {
  if (!str) return false;
  const t = str.trim();
  return t.startsWith("[") && t.endsWith("]");
}

/**
 * Is this the "blank document" — nothing, or a single empty paragraph?
 * Conservative on purpose: anything else counts as content.
 */
export function isPlateValueEmpty(value: Value | null | undefined): boolean {
  if (!value || value.length === 0) return true;
  if (value.length > 1) return false;
  const node = value[0] as TElement;
  if (!node || node.type !== "p") return false;
  const children = (node.children ?? []) as (TElement | TText)[];
  if (children.length === 0) return true;
  return children.every((c) => "text" in c && String((c as TText).text).trim() === "");
}

/** Inline style string for an element's block-level properties. */
function blockStyle(node: any): string {
  const parts: string[] = [];
  if (node.align) parts.push(`text-align: ${node.align}`);
  if (node.indent) parts.push(`margin-left: ${Number(node.indent) * 40}px`);
  if (node.lineHeight) parts.push(`line-height: ${node.lineHeight}`);
  // The space above and below a paragraph in a table CELL, which is most of a
  // row's height — see `extractBlockSpacing`. No block outside a cell carries
  // one, so this writes nothing anywhere else.
  if (node.marginTop) parts.push(`margin-top: ${node.marginTop}`);
  if (node.marginBottom) parts.push(`margin-bottom: ${node.marginBottom}`);
  // The size the block itself is set in — BlockFontSizePlugin puts it there,
  // from the runs the paragraph holds or, on a blank line, from the lines
  // around it. It has to be on the block: a line's height is its STRUT, and a
  // strut is the size of the element the line sits in, not of anything inside
  // it. The print page states `body p {font-size: 15px}`, which this outranks.
  // See lib/block-font-size.ts.
  if (node.fontSize) parts.push(`font-size: ${node.fontSize}`);
  // Paragraph shading — the full-width band Word lays a heading or a caution
  // line on. On the block because that is what it is: a highlight on the runs
  // would stop where the text stops. See components/block-background-plugin.ts.
  if (node.backgroundColor) parts.push(`background-color: ${node.backgroundColor}`);
  return parts.length ? ` style="${parts.join("; ")}"` : "";
}

/**
 * The `<br/>` that gives an empty `<p>` a line box.
 *
 * Wrapped in the size the line was set in when it states one of its own, so a
 * blank line the author sized deliberately reopens as itself rather than being
 * re-derived from its neighbours: Plate reads a font size off a `<span>`, never
 * off the `<p>`.
 */
function blankLineBreak(node: any): string {
  const fontSize = ownBlankLineFontSize(node);
  return fontSize ? `<span style="font-size: ${fontSize}"><br/></span>` : "<br/>";
}

/**
 * Where a leaf sits inside its block, which decides whether a space at either
 * end of it is one HTML would drop. See serializeWhitespace.
 */
export type LeafEdges = {
  atBlockStart?: boolean;
  atBlockEnd?: boolean;
  /**
   * Is this node being written INSIDE a paragraph?
   *
   * An image deserialized from `<p><img></p>` — which is how Word, LibreOffice
   * and this serializer's own output all state one — becomes an image node
   * nested in a paragraph node. Both wrote a `<p>` of their own around it, so
   * the report was saved as `<p><p><img/></p></p>`, which no browser accepts:
   * the print page reparsed it into three paragraphs and drew a blank line
   * above and below every picture in the report.
   */
  insideParagraph?: boolean;
};

/**
 * Text, with hard breaks and layout whitespace preserved.
 *
 * The old serializer wrote "\n" straight into the output, which HTML collapses
 * to a space — so every shift+enter in a pasted Word document quietly
 * disappeared. Emitted as <br/> now.
 *
 * A <br/> ends a line, so each side of one is its own start-and-end for the
 * purpose of a space that would otherwise be dropped, which is why the segments
 * are spaced individually rather than the whole string at once.
 */
function textWithBreaks(text: string, edges: LeafEdges): string {
  const lines = escapeHtml(text).split("\n");
  return lines
    .map((line, i) =>
      serializeWhitespace(line, {
        atBlockStart: i > 0 || (edges.atBlockStart ?? true),
        atBlockEnd: i < lines.length - 1 || (edges.atBlockEnd ?? true),
      })
    )
    .join("<br/>");
}

function serializeText(node: TText, edges: LeafEdges = {}): string {
  const n = node as any;
  let text = textWithBreaks(n.text ?? "", edges);
  if (!text) return "";

  if (n.bold) text = `<strong>${text}</strong>`;
  if (n.italic) text = `<em>${text}</em>`;
  if (n.underline) text = `<u>${text}</u>`;
  if (n.strikethrough) text = `<s>${text}</s>`;
  if (n.code) text = `<code>${text}</code>`;
  if (n.subscript) text = `<sub>${text}</sub>`;
  if (n.superscript) text = `<sup>${text}</sup>`;
  if (n.highlight) text = `<mark>${text}</mark>`;

  const styles: string[] = [];
  if (n.color) styles.push(`color: ${n.color}`);
  if (n.backgroundColor) styles.push(`background-color: ${n.backgroundColor}`);
  if (n.fontSize) styles.push(`font-size: ${n.fontSize}`);
  if (n.fontFamily) {
    // Double quotes inside the value would terminate the style="" attribute.
    styles.push(`font-family: ${String(n.fontFamily).replace(/"/g, "'")}`);
  }
  if (styles.length) text = `<span style="${styles.join("; ")}">${text}</span>`;

  return text;
}

function serializeChildren(node: TElement, insideParagraph = false): string {
  const children = (node.children || []) as any[];
  return children
    .map((c, i) =>
      serializeNode(c, {
        atBlockStart: i === 0,
        atBlockEnd: i === children.length - 1,
        insideParagraph,
      })
    )
    .join("");
}

/** <colgroup> carrying the column widths the user dragged. */
function serializeColgroup(node: any): string {
  const sizes: number[] | undefined = node.colSizes;
  if (!Array.isArray(sizes) || sizes.length === 0) return "";
  const cols = sizes
    .map((w) => (w ? `<col style="width: ${Math.round(w)}px" />` : "<col />"))
    .join("");
  return `<colgroup>${cols}</colgroup>`;
}

function serializeCell(node: any, tag: "td" | "th"): string {
  const attrs: string[] = [];
  // Spans live in two places depending on how the cell got here — see nodeCellSpan.
  const colSpan = nodeCellSpan(node, "colSpan");
  const rowSpan = nodeCellSpan(node, "rowSpan");
  if (colSpan > 1) attrs.push(`colspan="${colSpan}"`);
  if (rowSpan > 1) attrs.push(`rowspan="${rowSpan}"`);

  const styles: string[] = [
    ...cellBorderStyles(node.borders),
    // The document's own padding when it stated one — `extractCellPadding`
    // reads Word's `padding:0cm 5.4pt` and a legacy report's `cellpadding`
    // alike — otherwise the shared default, which is also what the editor
    // draws, so a table looks the same on screen as it prints.
    `padding: ${node.padding || DEFAULT_CELL_PADDING}`,
  ];
  // Cell shading set in the editor, and whatever Word's tables brought with them.
  if (node.background) styles.push(`background-color: ${node.background}`);
  const width = node.width ?? node.attributes?.width;
  if (width) styles.push(`width: ${String(width).match(/\D/) ? width : `${width}px`}`);
  if (node.verticalAlign) styles.push(`vertical-align: ${node.verticalAlign}`);
  // Alignment the source stated on the cell rather than on its paragraphs.
  if (node.align) styles.push(`text-align: ${node.align}`);
  // No invented header shading: a header cell is only tinted if the document
  // actually said so.

  const attrStr = attrs.length ? ` ${attrs.join(" ")}` : "";
  return `<${tag}${attrStr} style="${styles.join("; ")}">${serializeChildren(node)}</${tag}>`;
}

/**
 * A list item. Its content lives in a `lic` child; any nested `ul`/`ol` are
 * siblings of that content and must be emitted INSIDE the <li>, which is what
 * the previous serializer lost.
 */
function serializeListItem(node: TElement): string {
  const children = (node.children || []) as any[];
  const content: string[] = [];
  const nested: string[] = [];
  for (const child of children) {
    if (child && (child.type === "ul" || child.type === "ol")) {
      nested.push(serializeNode(child));
    } else {
      content.push(serializeNode(child));
    }
  }
  const checked = (node as any).checked;
  const marker =
    typeof checked === "boolean"
      ? `<input type="checkbox" ${checked ? "checked" : ""} disabled /> `
      : "";
  return `<li>${marker}${content.join("")}${nested.join("")}</li>`;
}

function serializeCaption(caption: any): string {
  if (!Array.isArray(caption)) return "";
  return caption.map((c: any) => serializeNode(c)).join("");
}

export function serializeNode(node: TElement | TText, edges: LeafEdges = {}): string {
  if ("text" in node) return serializeText(node as TText, edges);

  const el = node as TElement;
  const n = el as any;
  const style = blockStyle(n);
  // `lic` is a list item's content and writes no wrapper of its own, so what
  // sits in one is in a paragraph in every sense that matters here.
  const children = serializeChildren(el, el.type === "p" || el.type === "lic");

  switch (el.type) {
    case "p":
      return `<p${style}>${children || blankLineBreak(n)}</p>`;
    case "h1":
    case "h2":
    case "h3":
    case "h4":
    case "h5":
    case "h6":
      return `<${el.type}${style}>${children}</${el.type}>`;
    case "blockquote":
      return `<blockquote${style}>${children}</blockquote>`;
    case "ul":
    case "ol":
      return `<${el.type}${style}>${children}</${el.type}>`;
    case "li":
      return serializeListItem(el);
    case "lic":
      // List item content: no wrapper, it is the text of the <li>.
      return children;
    case "table": {
      const width = n.width ? `width: ${n.width}px` : "width: 100%";
      // `border="1"` only when the table really is bordered — it is what the
      // print stylesheets and dompdf key a visible grid off. And `border="0"`
      // when it is not, rather than no attribute at all: that is the hook
      // `print_richtext_fixes` uses to keep a borderless table borderless
      // (`table[border="0"] td { border-width: 0 }`), and it is what Word
      // itself writes for one.
      const attr = tableHasBorders(n) ? ' border="1"' : ' border="0"';
      return `<table${attr} style="border-collapse: collapse; ${width};">${serializeColgroup(n)}${children}</table>`;
    }
    case "tr": {
      // `size` is the row height @platejs/table stores; emitting it keeps a
      // pasted table's row heights across a save and reopen.
      const height = n.size ? ` style="height: ${Math.round(n.size)}px"` : "";
      return `<tr${height}>${children}</tr>`;
    }
    case "td":
      return serializeCell(n, "td");
    case "th":
      return serializeCell(n, "th");
    case "img": {
      const url = n.url || "";
      const w = n.width ? ` width="${n.width}"` : "";
      const cap = n.caption ? serializeCaption(n.caption) : "";
      const img = `<img src="${url}"${w} alt="${escapeHtml(n.alt || "")}" style="max-width: 100%; height: auto;" />`;
      // Already inside a paragraph: that paragraph is the picture's block, and
      // it is the one carrying the alignment the document centred it with.
      if (edges.insideParagraph) return img;
      return cap
        ? `<figure${style}>${img}<figcaption>${cap}</figcaption></figure>`
        : `<p${style}>${img}</p>`;
    }
    case "video":
    case "audio":
    case "file": {
      // Not renderable in a PDF, but a dropped node is worse than a link.
      const link = `<a href="${n.url || "#"}">${escapeHtml(n.name || n.url || el.type)}</a>`;
      return edges.insideParagraph ? link : `<p${style}>${link}</p>`;
    }
    case "a":
      return `<a href="${n.url || "#"}">${children}</a>`;
    case "hr":
      return "<hr />";
    case "code_block":
      return `<pre><code>${children}</code></pre>`;
    case "code_line":
      return `${children}\n`;
    default:
      // Unknown block: keep the text rather than dropping the node.
      return children ? `<p${style}>${children}</p>` : "";
  }
}

/** Plate value -> self-contained, inline-styled HTML for storage and printing. */
export function plateValueToHtml(value: Value): string {
  if (!value || value.length === 0) return "";
  return value.map((n) => serializeNode(n as TElement)).join("");
}

export function plateValueToJson(value: Value): string {
  if (!value || value.length === 0) {
    return JSON.stringify([{ type: "p", children: [{ text: "" }] }]);
  }
  return JSON.stringify(value);
}

export const EMPTY_VALUE: Value = [{ type: "p", children: [{ text: "" }] }] as Value;

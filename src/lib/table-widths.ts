/**
 * Column widths for pasted tables.
 *
 * Plate's table deserializer keeps rows, cells and spans, but throws away every
 * width: neither `<td width="113">` nor `style="width:85.0pt"` nor a
 * `<colgroup>` survives. With no `colSizes` on the table node, TableCellElement
 * falls back to `minWidth: 120 / maxWidth: 240` for every column, so a pasted
 * table comes out with equal columns.
 *
 * That is very visible on the label/colon/value tables these reports are full
 * of — a middle column holding just ":" is ~20pt in Word and rendered a third
 * of the table wide here.
 *
 * The widths are read back off the HTML and written to `colSizes`, which is
 * what @platejs/table sizes columns from. The serializer emits them again as a
 * <colgroup>, so a saved report reopens with its columns intact.
 */

import { statedBackground } from "@/lib/background";

/** CSS absolute units to px, at the 96dpi the browser assumes. */
const UNIT_TO_PX: Record<string, number> = {
  px: 1,
  pt: 96 / 72,
  pc: 16,
  in: 96,
  cm: 96 / 2.54,
  mm: 96 / 25.4,
};

/**
 * Reference width for percentage columns — roughly the printable width of A4
 * at 96dpi. Only the ratios between columns matter; this just turns them into
 * the absolute numbers `colSizes` is defined in.
 */
const PERCENT_REFERENCE_PX = 700;

/**
 * Only a guard against a zero/negative width, which would break the layout
 * maths. A width the source actually states is otherwise reproduced exactly —
 * including very narrow columns, which Word uses for the ":" separator in
 * label/value tables.
 */
const MIN_COLUMN_PX = 1;

/** A single width, in px, from an attribute or a CSS length. Null if absent. */
function parseWidth(raw: string | null | undefined): number | null {
  if (!raw) return null;
  const value = raw.trim();
  if (!value) return null;

  const match = value.match(/^(-?[\d.]+)\s*(px|pt|pc|in|cm|mm|%)?$/i);
  if (!match) return null;

  const n = Number.parseFloat(match[1]);
  if (!Number.isFinite(n) || n <= 0) return null;

  const unit = (match[2] || "px").toLowerCase();
  if (unit === "%") return (n / 100) * PERCENT_REFERENCE_PX;

  return n * (UNIT_TO_PX[unit] ?? 1);
}

/** A CSS length from an element's inline style, e.g. `width` or `height`. */
function styleLength(el: Element, prop: "width" | "height"): number | null {
  return parseWidth((el as HTMLElement).style?.[prop]);
}

/** Width of one <col> / <td> / <th>, preferring the CSS length Word writes. */
function cellWidth(el: Element): number | null {
  return styleLength(el, "width") ?? parseWidth(el.getAttribute("width"));
}

function spanOf(el: Element): number {
  const n = Number.parseInt(el.getAttribute("colspan") ?? "1", 10);
  return Number.isFinite(n) && n > 1 ? n : 1;
}

/**
 * Append one source element's contribution to the column list.
 *
 * `share` distinguishes the two spanning rules, which are NOT the same:
 *  - `<td colspan="2" style="width:200px">` — the 200px covers both columns, so
 *    it is divided between them.
 *  - `<col span="2" width="60">` — per the HTML spec the width applies to EACH
 *    column described, so it is repeated.
 */
function pushSpanned(
  into: (number | null)[],
  width: number | null,
  span: number,
  { share }: { share: boolean }
) {
  const each = width === null || !share ? width : width / span;
  for (let i = 0; i < span; i++) into.push(each);
}

/**
 * How many columns the grid is WIDE, measured over every row.
 *
 * The first row is not always the widest. A report table is routinely ragged —
 * ten `label : value` rows with two of them broken into five cells for a
 * measurement pair — and the columns only those two rows reach are columns all
 * the same. Counting the first row alone leaves them out of `colSizes`, and a
 * column with no entry there has no width to resize: see
 * `components/table-colsizes-plugin.ts`.
 */
function rowGridWidth(element: HTMLElement): number {
  return tableRows(element).reduce(
    (widest, row) =>
      Math.max(
        widest,
        rowCells(row).reduce((total, cell) => total + spanOf(cell), 0)
      ),
    0
  );
}

/**
 * Column widths in px, or undefined when the markup carries none — in which
 * case Plate's own defaults should stand rather than a guess.
 */
export function extractColSizes(element: HTMLElement): number[] | undefined {
  if (!element || element.tagName?.toLowerCase() !== "table") return undefined;

  const widths: (number | null)[] = [];

  // A <colgroup> is authoritative when present: it describes the columns
  // directly rather than being inferred from one row.
  const cols = element.querySelectorAll(":scope > colgroup > col, :scope > col");
  if (cols.length > 0) {
    cols.forEach((col) => {
      const span = Number.parseInt(col.getAttribute("span") ?? "1", 10);
      pushSpanned(widths, cellWidth(col), Number.isFinite(span) && span > 1 ? span : 1, {
        share: false,
      });
    });
  } else {
    // Otherwise the rows define the grid. The FIRST row stays the authority for
    // every column it reaches — that is where Word puts the per-column widths,
    // repeated on every row — and a WIDER row later on contributes only the
    // columns the first row never got to. Purely additive: a table whose rows
    // all have the same shape comes out of here exactly as it always did.
    const rows = tableRows(element);
    if (rows.length === 0) return undefined;

    rows.forEach((row) => {
      const fromRow: (number | null)[] = [];
      rowCells(row).forEach((cell) =>
        pushSpanned(fromRow, cellWidth(cell), spanOf(cell), { share: true })
      );
      for (let i = widths.length; i < fromRow.length; i++) widths.push(fromRow[i]);
    });
  }

  if (widths.length === 0 || widths.every((w) => w === null)) return undefined;

  // A table with only some widths stated: give the rest the average of the
  // known ones, so the stated columns still hold their proportions. A column
  // the source did state is passed through exactly.
  const known = widths.filter((w): w is number => w !== null);
  const fallback = known.reduce((a, b) => a + b, 0) / known.length;

  return widths.map((w) => Math.max(MIN_COLUMN_PX, Math.round(w ?? fallback)));
}

/**
 * Row height in px, or undefined when the markup states none.
 *
 * Becomes `size` on the row node, which @platejs/table turns into the cells'
 * `minHeight`. Word states it on the <tr>, but sometimes only on the row's
 * first cell, so both are checked.
 */
export function extractRowSize(element: HTMLElement): number | undefined {
  if (!element || element.tagName?.toLowerCase() !== "tr") return undefined;

  let height = styleLength(element, "height") ?? parseWidth(element.getAttribute("height"));

  if (height === null) {
    const firstCell = Array.from(element.children).find((c) => /^t[dh]$/i.test(c.tagName));
    if (firstCell) {
      height = styleLength(firstCell, "height") ?? parseWidth(firstCell.getAttribute("height"));
    }
  }

  if (height === null || height <= 0) return undefined;
  return Math.round(height);
}

/**
 * A cell's span, read from BOTH places it can live.
 *
 * A cell merged inside the editor carries numeric `colSpan` / `rowSpan` node
 * props. A cell that arrived by paste or from legacy HTML instead carries the
 * raw attributes Plate's deserializer preserved — `attributes.colspan`, as a
 * string. Reading only the first is how a merged cell silently flattened.
 */
export function nodeCellSpan(node: any, key: "colSpan" | "rowSpan"): number {
  const raw = node?.[key] ?? node?.attributes?.[key.toLowerCase()];
  const n = Number.parseInt(String(raw ?? "1"), 10);
  return Number.isFinite(n) && n > 1 ? n : 1;
}

/* -------------------------------------------------------------------------
 * Cell borders
 * ---------------------------------------------------------------------- */

export interface CellBorder {
  size: number;
  style?: string;
  color?: string;
}

export type CellBorders = Partial<Record<"top" | "right" | "bottom" | "left", CellBorder>>;

const SIDES = ["top", "right", "bottom", "left"] as const;

type Side = (typeof SIDES)[number];

const OPPOSITE: Record<Side, Side> = {
  top: "bottom",
  right: "left",
  bottom: "top",
  left: "right",
};

/**
 * `thin` / `medium` / `thick`, the keyword widths, at the px a browser uses.
 * A border that names a style but no width at all is left at 1px — see
 * `declaredBorder`.
 */
const KEYWORD_WIDTH: Record<string, number> = { thin: 1, medium: 3, thick: 5 };

/**
 * How a border style ranks when two edges of equal width meet — CSS 2.1
 * §17.6.2.1, "border conflict resolution". A double rule beats a solid one of
 * the same width, a solid beats a dashed, and so on down.
 *
 * `hidden` and `none` are NOT on this scale: `hidden` beats everything and is
 * handled first, and `none` is the absence of a declaration rather than a
 * losing one.
 */
const STYLE_RANK: Record<string, number> = {
  double: 8,
  solid: 7,
  dashed: 6,
  dotted: 5,
  ridge: 4,
  outset: 3,
  groove: 2,
  inset: 1,
};

/**
 * Where a declaration came from, and how much that is worth when width and
 * style tie. Same order as the spec: the cell wins over its row, the row over
 * the table, and the table over the `border` attribute's presentational hint.
 */
const SOURCE_RANK = { cell: 4, row: 3, table: 2, attribute: 1 } as const;

type BorderSource = keyof typeof SOURCE_RANK;

interface BorderCandidate extends CellBorder {
  hidden?: boolean;
  source: BorderSource;
}

/** A colour worth carrying: not Word's `windowtext`, not a CSS-wide keyword. */
function borderColor(color: string): string | undefined {
  const value = (color || "").trim();
  if (!value) return undefined;
  return /^(windowtext|currentcolor|initial|inherit|unset)$/i.test(value) ? undefined : value;
}

/**
 * What one element declares for one of its edges, or null for "nothing".
 *
 * `none` is not a border that lost, it is the absence of one — which matters
 * because Word writes the interior edges of every grid table as
 * `border-left:none` and lets the neighbour's `border-right` draw the line.
 *
 * A WIDTH with no style is a border here, even though a bare browser draws
 * nothing for one. That is not a browser quirk being papered over, it is this
 * application's own rule: `_partials/print_richtext_fixes` supplies
 * `border-style: solid` for exactly this case, because some of Word's exports
 * write a cell's width and colour and omit the style. Those rules DO print, so
 * they have to show on screen too.
 */
function declaredBorder(
  el: Element | null | undefined,
  side: Side,
  source: BorderSource
): BorderCandidate | null {
  const style = (el as HTMLElement | null)?.style;
  if (!style) return null;

  const stated = style.getPropertyValue(`border-${side}-style`).trim().toLowerCase();
  if (stated === "none") return null;
  if (stated === "hidden") return { size: 0, hidden: true, source };

  const rawWidth = style.getPropertyValue(`border-${side}-width`).trim().toLowerCase();
  if (!stated && !rawWidth) return null;

  const lineStyle = stated || "solid";
  const px = KEYWORD_WIDTH[rawWidth] ?? parseWidth(rawWidth);
  // A style with no width of its own keeps the 1px this editor has always
  // drawn for it, rather than the browser's 3px `medium` — Word states a width
  // on every border it writes, so this only ever catches hand-written markup,
  // where a sudden 3px rule in a report is the more surprising answer.
  // Whole pixels: Word's 1.0pt is 1.333px, and a fractional border prints
  // unevenly. Anything visible is at least 1px.
  const size = px === null ? 1 : Math.max(1, Math.round(px));
  const color = borderColor(style.getPropertyValue(`border-${side}-color`));

  return {
    size,
    ...(lineStyle !== "solid" ? { style: lineStyle } : {}),
    ...(color ? { color } : {}),
    source,
  };
}

/** The `border` attribute as a number; null when absent or unparseable. */
function tableBorderAttribute(table: Element | null): number | null {
  const raw = table?.getAttribute("border");
  if (raw === null || raw === undefined) return null;
  const n = Number.parseInt(raw.trim(), 10);
  return Number.isFinite(n) ? n : null;
}

/** Every `<tr>` of THIS table, in order — not those of a nested one. */
function tableRows(table: Element): Element[] {
  return Array.from(
    table.querySelectorAll(":scope > tr, :scope > thead > tr, :scope > tbody > tr, :scope > tfoot > tr")
  );
}

function rowCells(row: Element): Element[] {
  return Array.from(row.children).filter((c) => /^t[dh]$/i.test(c.tagName));
}

/** The grid column a cell starts at, counting the colspans before it. */
function cellColumnIndex(cell: Element): number {
  let index = 0;
  for (const sibling of rowCells(cell.parentElement as Element)) {
    if (sibling === cell) return index;
    index += spanOf(sibling);
  }
  return index;
}

/** The cell covering a grid column in a row, ignoring rows spanned into it. */
function cellAtColumn(row: Element | undefined, column: number): Element | null {
  if (!row) return null;
  let index = 0;
  for (const cell of rowCells(row)) {
    const span = spanOf(cell);
    if (column >= index && column < index + span) return cell;
    index += span;
  }
  return null;
}

/**
 * The cell sharing this edge, or null at the table's rim.
 *
 * Rows are looked up by GRID column rather than by sibling position, so a
 * `colspan` above a pair of cells still resolves to the cell they actually
 * touch. A `rowspan` reaching down past the previous row is not tracked: the
 * lookup then finds no neighbour, and the cell's own declaration stands, which
 * is the same answer in every table these reports contain.
 */
function adjacentCell(cell: Element, side: Side): Element | null {
  const row = cell.parentElement;
  if (!row) return null;

  if (side === "left" || side === "right") {
    const cells = rowCells(row);
    const at = cells.indexOf(cell);
    return (side === "left" ? cells[at - 1] : cells[at + 1]) ?? null;
  }

  const table = cell.closest("table");
  if (!table) return null;
  const rows = tableRows(table);
  const at = rows.indexOf(row);
  if (at === -1) return null;

  return cellAtColumn(rows[side === "top" ? at - 1 : at + 1], cellColumnIndex(cell));
}

/** Is this edge on the outside of the table, where the table's own frame sits? */
function isOuterEdge(cell: Element, side: Side): boolean {
  return adjacentCell(cell, side) === null;
}

/** The winner of a border conflict, per CSS 2.1 §17.6.2.1. */
function strongest(candidates: (BorderCandidate | null)[]): CellBorder {
  let best: BorderCandidate | null = null;

  for (const candidate of candidates) {
    if (!candidate) continue;
    // `hidden` suppresses the edge outright, whatever else meets it there.
    if (candidate.hidden) return { size: 0 };
    if (
      !best ||
      candidate.size > best.size ||
      (candidate.size === best.size &&
        (STYLE_RANK[candidate.style ?? "solid"] ?? 0) > (STYLE_RANK[best.style ?? "solid"] ?? 0)) ||
      (candidate.size === best.size &&
        candidate.style === best.style &&
        SOURCE_RANK[candidate.source] > SOURCE_RANK[best.source])
    ) {
      best = candidate;
    }
  }

  if (!best) return { size: 0 };
  const { hidden: _hidden, source: _source, ...border } = best;
  return border;
}

/**
 * Does this table say ANYTHING about borders?
 *
 * The answer decides between reading the document and leaving Plate's own 1px
 * default alone, so it is deliberately generous: a `border` attribute of any
 * value counts (`border=0` is a statement that there are none), and so does a
 * single `border-*-style` or `border-*-width` anywhere in the table —
 * including `none`, which is how Word writes the borderless tables it uses to
 * align label / value columns.
 *
 * Memoised per table element: every cell asks, and the answer walks the whole
 * table.
 */
const TABLE_DECLARES_BORDERS = new WeakMap<Element, boolean>();

function tableDeclaresBorders(table: Element): boolean {
  const cached = TABLE_DECLARES_BORDERS.get(table);
  if (cached !== undefined) return cached;

  let declared = tableBorderAttribute(table) !== null;
  if (!declared) {
    const elements = [table, ...Array.from(table.querySelectorAll("tr, thead, tbody, tfoot, td, th"))];
    declared = elements.some((el) => {
      const style = (el as HTMLElement).style;
      return (
        !!style &&
        SIDES.some(
          (side) =>
            !!style.getPropertyValue(`border-${side}-style`) ||
            !!style.getPropertyValue(`border-${side}-width`)
        )
      );
    });
  }

  TABLE_DECLARES_BORDERS.set(table, declared);
  return declared;
}

/**
 * Per-side borders for a pasted cell, in Plate's own `borders` shape.
 *
 * This is a table-wide question, not a per-cell one, and reading only the
 * cell's own inline style is how a table's border DESIGN — a double rule under
 * the header, a heavy outer frame, a hairline grid inside it — arrived as one
 * flat 1px box on every cell. Under `border-collapse: collapse` (which is what
 * Word writes, and what this editor and every print blade render) each edge is
 * shared, and the browser picks a single winner for it out of everything that
 * describes it: the two cells that meet there, their rows, and the table
 * itself. So that is what is resolved here — width, style and colour together,
 * because a 2.25pt double rule is not a 1px solid one.
 *
 * Undefined only when the table says nothing about borders at all, which is
 * the one case where Plate's own default (a 1px line, the same one a table
 * created in this editor gets) is the right answer rather than a guess.
 */
export function extractCellBorders(el: Element): CellBorders | undefined {
  const table = el.closest("table");
  if (!table || !tableDeclaresBorders(table)) return undefined;

  const row = el.parentElement;
  const attribute = tableBorderAttribute(table);
  // `<table border=1>`: the HTML rendering rules give every cell a 1px rule,
  // which is what a legacy Summernote report relies on for its grid. It is the
  // weakest declaration there is, so anything stated in CSS overrides it.
  const attributeBorder: BorderCandidate | null =
    attribute !== null && attribute > 0 ? { size: 1, source: "attribute" } : null;

  const borders: CellBorders = {};
  for (const side of SIDES) {
    const neighbour = adjacentCell(el, side);
    borders[side] = strongest([
      declaredBorder(el, side, "cell"),
      declaredBorder(neighbour, OPPOSITE[side], "cell"),
      declaredBorder(row, side, "row"),
      declaredBorder(neighbour?.parentElement, OPPOSITE[side], "row"),
      // The table's frame reaches only the edges that ARE the table's edges.
      isOuterEdge(el, side) ? declaredBorder(table, side, "table") : null,
      attributeBorder,
    ]);
  }

  return borders;
}

/* -------------------------------------------------------------------------
 * Cell padding
 * ---------------------------------------------------------------------- */

/**
 * The gap a cell keeps around its text when the document states none.
 *
 * The same value the serializer has always written, so a table authored in
 * this editor looks on screen exactly like the report it prints — those two
 * disagreed, the editor drawing a 12px/8px gap and the print a 5px/2px one.
 */
export const DEFAULT_CELL_PADDING = "2px 5px";

const PADDING_SIDES = ["top", "right", "bottom", "left"] as const;

/**
 * A CSS length in px, **zero included**.
 *
 * `parseWidth` rejects 0 because a zero-width column is a broken table. A zero
 * PADDING is the opposite: it is what Word writes on every report cell
 * (`padding:0cm 5.4pt`) and dropping it is exactly how a pasted table came out
 * taller here than in the document it was copied from.
 */
function parseLength(raw: string | null | undefined): number | null {
  if (raw === null || raw === undefined) return null;
  const value = String(raw).trim();
  if (!value) return null;

  const match = value.match(/^(-?[\d.]+)\s*(px|pt|pc|in|cm|mm)?$/i);
  if (!match) return null;

  const n = Number.parseFloat(match[1]);
  if (!Number.isFinite(n)) return null;

  const px = n * (UNIT_TO_PX[(match[2] || "px").toLowerCase()] ?? 1);
  // A negative padding is not a thing; treat it as none rather than as a value
  // that would pull the text out of its own cell.
  return px < 0 ? 0 : px;
}

/** A px number as short as it can be written without changing it. */
function pxValue(value: number): string {
  return `${Number(value.toFixed(2))}px`;
}

/** Four side values as the shortest equivalent CSS shorthand. */
function paddingShorthand([top, right, bottom, left]: number[]): string {
  if (top === bottom && right === left) {
    return top === right ? pxValue(top) : `${pxValue(top)} ${pxValue(right)}`;
  }
  return `${pxValue(top)} ${pxValue(right)} ${pxValue(bottom)} ${pxValue(left)}`;
}

/**
 * The padding a pasted cell states, as a CSS shorthand, or undefined for none.
 *
 * This is the one piece of the old Summernote editor's table handling that had
 * no counterpart here. That editor never parsed a table at all — it dropped the
 * clipboard HTML into a contenteditable and let the BROWSER render it, so a
 * Word cell's `padding:0cm 5.4pt` was simply obeyed: no gap above or below the
 * text, 7.2px either side. This editor instead rebuilt every cell with a fixed
 * `px-3 py-2`, which is 8px of air on each of a row's two sides — ~16px taller
 * per row, and a table of ten rows visibly deeper than the document it came
 * from. That, and not the borders or the column widths, is what made a pasted
 * table "not the same".
 *
 * Read from the two places the padding can live, in the order the browser
 * itself resolves them:
 *  1. the cell's own `padding` — which the inline style declaration has already
 *     expanded into longhands for us, whether the source wrote the shorthand or
 *     the sides separately;
 *  2. the table's `cellpadding` attribute, the pre-CSS spelling of the same
 *     thing, and what the legacy Summernote reports in the database carry.
 *
 * Undefined when neither says anything, so `DEFAULT_CELL_PADDING` still stands
 * for tables created in the editor.
 */
export function extractCellPadding(el: Element): string | undefined {
  const style = (el as HTMLElement).style;

  if (style) {
    const sides = PADDING_SIDES.map((side) =>
      parseLength(style.getPropertyValue(`padding-${side}`))
    );
    // A partially stated padding leaves the rest at the browser's own 0, which
    // is what the source rendered as — not at this editor's default.
    if (sides.some((value) => value !== null)) {
      return paddingShorthand(sides.map((value) => value ?? 0));
    }
  }

  // `cellpadding` is a bare pixel count and applies to every side of every cell
  // in the table. Word writes `cellpadding=0` alongside per-cell padding, so it
  // is only ever reached when the cell itself said nothing.
  const attr = el.closest("table")?.getAttribute("cellpadding");
  const uniform = parseLength(attr);
  return uniform === null ? undefined : pxValue(uniform);
}

/* -------------------------------------------------------------------------
 * Block spacing inside a cell
 * ---------------------------------------------------------------------- */

/** The gap a block keeps above and below itself, in the shape a node carries. */
export interface BlockSpacing {
  marginTop?: string;
  marginBottom?: string;
}

/**
 * Is this the last block in its cell — the one whose space-after becomes the
 * gap under the whole row?
 *
 * Walked up to the cell rather than asking `cell.lastElementChild`, because a
 * paragraph is not always a direct child of the `<td>`: Word nests one inside a
 * `<div>` often enough that the simple check would call a nested block "last"
 * and leave the row's gap in place.
 */
function isLastBlockInCell(el: Element, cell: Element): boolean {
  let node: Element | null = el;
  while (node && node !== cell) {
    if (node.nextElementSibling) return false;
    node = node.parentElement;
  }
  return node === cell;
}

/**
 * The space a paragraph inside a table cell keeps above itself, and between
 * itself and the next paragraph in the same cell.
 *
 * A row's height is not the cell's padding alone. Word sets its paragraphs with
 * space before and after — `margin-top:6.0pt;margin-bottom:6.0pt` is ordinary
 * on a report's table — and that is most of the air between one row's text and
 * the next. Nothing carried it: Plate's deserializer drops a block's margins
 * outright (in `faithful` mode too, so this is the parser's doing and not the
 * paste normalizer's), so a row that stands 35px in the document arrived here
 * at 24px and the printed report lost the same 11px again.
 *
 * THE LAST BLOCK'S SPACE-AFTER IS DROPPED, and that is the whole of this note.
 * Word's `Normal` style carries `margin-bottom:10.0pt` and nothing on top —
 * not something an author sets, just what every unmodified document has on
 * every paragraph — so the last paragraph in a cell handed every pasted row
 * 13px of empty space under its text and none above it. Measured: a row of
 * 20px text stood 33px, 0 above and 13 below, and the serializer wrote
 * `margin-bottom: 13.33px` so the printed report kept it too. Asked for by name
 * on 2026-09-16: "when table data paste then generate row bottom size some
 * extra padding or margin which is generate big problem".
 *
 * Only the LAST one, so a cell holding several paragraphs keeps the gaps
 * BETWEEN them — those are the spacing inside the cell, not the gap under the
 * row, and collapsing them would run a three-line finding together.
 *
 * Scoped to the inside of a CELL on purpose. The editor's own paragraph spacing
 * outside tables is settled ground and must not move; a block that is not in a
 * cell returns undefined here and carries no margin at all, so the rule cannot
 * reach it.
 *
 * Only the vertical sides. `margin-left` is how this editor states INDENT (the
 * serializer writes `indent * 40px` into it), and Word's own left margins on a
 * paragraph are the same thing said differently — not spacing between rows.
 */
export function extractBlockSpacing(el: Element): BlockSpacing | undefined {
  const cell = el.closest("td, th");
  if (!cell) return undefined;

  const style = (el as HTMLElement).style;
  if (!style) return undefined;

  const top = parseLength(style.getPropertyValue("margin-top"));
  const bottom = isLastBlockInCell(el, cell)
    ? null
    : parseLength(style.getPropertyValue("margin-bottom"));

  // A stated zero is the same as no statement — it is what a block gets here
  // anyway — so it is not worth carrying onto the node or into the report.
  const spacing: BlockSpacing = {
    ...(top ? { marginTop: pxValue(top) } : {}),
    ...(bottom ? { marginBottom: pxValue(bottom) } : {}),
  };

  return spacing.marginTop || spacing.marginBottom ? spacing : undefined;
}

/* -------------------------------------------------------------------------
 * Cell alignment
 * ---------------------------------------------------------------------- */

const ALIGNMENTS = new Set(["left", "center", "right", "justify"]);

/**
 * How a pasted cell aligns its contents, or undefined when it says nothing.
 *
 * Word states this in EITHER place, and which one it picks is not the author's
 * doing: a centred heading row often arrives as `<td align=center>` with a
 * plain paragraph inside. The paragraph carries no alignment of its own, so
 * reading only the paragraph — which is all the align plugin looks at — landed
 * every one of those cells back on the left.
 */
export function extractCellAlign(el: Element): string | undefined {
  const stated =
    (el as HTMLElement).style?.textAlign || el.getAttribute("align") || "";
  const align = stated.trim().toLowerCase();
  return ALIGNMENTS.has(align) ? align : undefined;
}

/* -------------------------------------------------------------------------
 * Cell shading and vertical alignment
 * ---------------------------------------------------------------------- */

const VERTICAL_ALIGNMENTS = new Set(["top", "middle", "bottom", "baseline"]);

/**
 * The fill a pasted cell states, or undefined for none.
 *
 * Word shades header rows and section bands — `background:#D9E2F3` — and the
 * serializer has always been able to write `node.background` back out. Nothing
 * was reading it, so a banded table arrived plain white.
 *
 * `bgcolor` is the same fill in the pre-CSS spelling, and it is not a museum
 * piece here: it is what LibreOffice writes for a shaded cell, what Word 97
 * HTML writes, and what every legacy report body in the database carries. Read
 * off no style, it left those tables white too. See lib/background.ts.
 *
 * A fill the ROW or the TABLE states counts as this cell's, because that is
 * what the reader sees: a cell paints nothing of its own, so the row's colour
 * shows straight through it. HTML shades a banded table either way round —
 * `<tr bgcolor>` on a web page, `<td>` in Word — and a Plate cell is the only
 * node in a table that can hold a fill at all, so a row's has to come to rest
 * here or nowhere. The climb stops AT the table: whatever is behind that is the
 * page, not the table's own shading.
 */
export function extractCellBackground(el: Element): string | undefined {
  for (let node: Element | null = el; node; node = node.parentElement) {
    const stated = statedBackground(node);
    if (stated) return stated;
    if (node.tagName === "TABLE") break;
  }

  return undefined;
}

/**
 * Where a pasted cell sits its content vertically.
 *
 * Word writes `valign=top` on practically every report cell, because that is
 * what a Word table does. An HTML cell defaults to `middle`, so dropping this
 * moved the text of every multi-line row down off the line it shares with its
 * neighbours — visible the moment one column wraps and the others do not.
 */
export function extractCellVerticalAlign(el: Element): string | undefined {
  const stated =
    (el as HTMLElement).style?.verticalAlign || el.getAttribute("valign") || "";
  const align = stated.trim().toLowerCase();
  return VERTICAL_ALIGNMENTS.has(align) ? align : undefined;
}

/** Does any cell in this value carry a visible border? */
export function hasVisibleBorder(borders: CellBorders | undefined): boolean {
  if (!borders) return true; // Plate's default is a 1px line.
  return SIDES.some((side) => (borders[side]?.size ?? 0) > 0);
}

/**
 * The colour a border with none of its own is drawn in.
 *
 * Word's `windowtext` is dropped during extraction because it is a keyword for
 * "the default text colour" rather than a colour, and a report is printed in
 * black.
 */
export const DEFAULT_BORDER_COLOR = "#000";

/** No line at all, as a `border` shorthand. */
export const NO_BORDER = "0 none";

/**
 * One edge as a complete CSS `border` shorthand.
 *
 * The single definition the editor and the serializer both draw from, so the
 * line on screen and the line in the printed report cannot drift apart — the
 * same reason `DEFAULT_CELL_PADDING` lives here. Width AND style AND colour,
 * because a 3px double rule is not a 1px solid one and rendering it as one is
 * what flattened every pasted table's border design.
 */
export function borderShorthand(border: Partial<CellBorder> | undefined): string {
  // `size` is optional in Plate's own `TTableCellBorder`, and an absent one is
  // the same as a zero: nothing to draw.
  if (!border?.size || border.size <= 0) return NO_BORDER;
  return `${border.size}px ${border.style ?? "solid"} ${border.color ?? DEFAULT_BORDER_COLOR}`;
}

/* -------------------------------------------------------------------------
 * Fitting a pasted table to the editor
 * ---------------------------------------------------------------------- */

/**
 * Slack left between the fitted table and the editable's content box.
 *
 * The table block sits flush inside that box — its `-ml-2` and the 8px control
 * cell each row carries cancel out — so this covers only what hangs off the
 * right edge: the last column's resize handle, which is drawn 3px past it and
 * would otherwise sit inside the block's `overflow-x-auto`, out of reach.
 */
const FIT_SAFETY_PX = 6;

/** Below this the measurement is not a laid-out editor and is ignored. */
const MIN_FIT_WIDTH_PX = 80;

/**
 * Width a table has to work with inside the MOUNTED editor, in px.
 *
 * Word states a table at the width of a PAGE — the 6.5in text column of a
 * Letter page is 624px — while this editor is as wide as the admin panel's
 * card, comfortably past 1200px on a desktop. So reproducing the stated widths
 * exactly drew a pasted report across the left half of the editor with the
 * right half blank.
 *
 * Undefined when there is nothing to measure: a headless editor in a test, or
 * the initial value, which is deserialized before the editable exists. Both
 * then keep the widths the source stated.
 */
export function editorContentWidth(editor: any): number | undefined {
  if (typeof window === "undefined") return undefined;

  let dom: HTMLElement | undefined;
  try {
    dom = editor?.api?.toDOMNode?.(editor);
  } catch {
    return undefined;
  }
  if (!dom) return undefined;

  const style = window.getComputedStyle(dom);
  const width =
    dom.clientWidth -
    (Number.parseFloat(style.paddingLeft) || 0) -
    (Number.parseFloat(style.paddingRight) || 0) -
    FIT_SAFETY_PX;

  return Number.isFinite(width) && width >= MIN_FIT_WIDTH_PX ? width : undefined;
}

/**
 * How many grid columns a source table describes.
 *
 * Word writes plenty of tables with no width on anything at all. They have to
 * be filled too, or they render at Plate's own fallback — a fraction of this
 * editor — which is the very thing this fitting exists to stop.
 *
 * A `<colgroup>` is taken at its word; otherwise the grid is as wide as its
 * WIDEST row, not its first. See `rowGridWidth`.
 */
function sourceColumnCount(element: HTMLElement): number {
  const cols = element.querySelectorAll(":scope > colgroup > col, :scope > col");
  if (cols.length > 0) {
    const declared = Array.from(cols).reduce((total, col) => {
      const span = Number.parseInt(col.getAttribute("span") ?? "1", 10);
      return total + (Number.isFinite(span) && span > 1 ? span : 1);
    }, 0);
    // A colgroup narrower than the rows it sits over is Word's, not the
    // document's: the rows are what actually exist.
    return Math.max(declared, rowGridWidth(element));
  }

  return rowGridWidth(element);
}

/**
 * The `colSizes` a table deserialized from HTML should carry: the widths the
 * source stated, or equal columns when it stated none, scaled to fill
 * `availableWidth`.
 *
 * With no width to fit — a headless deserialization, or the initial value,
 * which is read before the editable exists — this is exactly what
 * `extractColSizes` returns and nothing is invented.
 */
export function tableColSizes(
  element: HTMLElement,
  availableWidth: number | undefined
): number[] | undefined {
  const stated = extractColSizes(element);
  if (!availableWidth) return stated;

  // A neutral ratio per column; `fitColSizes` turns it into whole pixels that
  // add up to the width available.
  const count = sourceColumnCount(element);
  const base = stated ?? (count > 0 ? new Array<number>(count).fill(1) : undefined);

  return fitColSizes(base, availableWidth);
}

/**
 * Scale column widths so the table spans the width it is given.
 *
 * Only the RATIOS between columns ever mattered: the serializer emits every
 * table as `width: 100%`, so the print blades, dompdf and the patient portal
 * already stretch it across the page whatever the numbers say. Scaling here is
 * what makes the editor show the same thing they do.
 *
 * The input is returned untouched when there is nothing to scale to, which is
 * how a headless deserialization keeps the source document's exact widths.
 */
export function fitColSizes(
  sizes: number[] | undefined,
  availableWidth: number | undefined
): number[] | undefined {
  if (!sizes || sizes.length === 0) return sizes;
  if (!availableWidth || !Number.isFinite(availableWidth) || availableWidth <= 0) return sizes;

  const total = sizes.reduce((a, b) => a + b, 0);
  if (total <= 0) return sizes;

  const target = Math.round(availableWidth);
  if (Math.abs(total - target) <= 1) return sizes;

  const scale = target / total;
  const scaled = sizes.map((w) => Math.max(MIN_COLUMN_PX, Math.round(w * scale)));

  // Rounding drift is handed to the widest column — the one where a pixel is
  // least visible — so the row ends exactly on the available width.
  const drift = target - scaled.reduce((a, b) => a + b, 0);
  if (drift !== 0) {
    let widest = 0;
    for (let i = 1; i < scaled.length; i++) {
      if (scaled[i] > scaled[widest]) widest = i;
    }
    scaled[widest] = Math.max(MIN_COLUMN_PX, scaled[widest] + drift);
  }

  return scaled;
}

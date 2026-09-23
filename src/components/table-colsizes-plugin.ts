import { createSlatePlugin, type NodeEntry, type SlateEditor, KEYS } from "platejs";
import { nodeCellSpan } from "@/lib/table-widths";

/**
 * Guarantee every table node carries `colSizes`.
 *
 * Column resizing is entirely gated on it. `getTableCellSize` computes a cell's
 * width as `(colSizes ?? []).slice(col, col + colSpan).reduce(...)`, so with no
 * colSizes the width is 0 and TableCellElement falls back to its constant
 * `minWidth: 120 / maxWidth: 240`. Dragging a column border then writes an
 * override that nothing reads — the handle moves, the column does not. The
 * insert/delete-column transforms are guarded by `if (colSizes)` for the same
 * reason.
 *
 * Plate only populates colSizes when a table is created with `initialTableWidth`
 * set. Anything that arrives another way — pasted from Word, deserialized from
 * a legacy Summernote report — has none, which is every table in this editor.
 *
 * So any table without usable colSizes gets evenly distributed columns here.
 * That matches what it already rendered as, and makes the border draggable.
 */

/** Roughly the printable width of A4 at 96dpi. */
const DEFAULT_TABLE_WIDTH = 700;
const MIN_COLUMN_PX = 24;

/**
 * Columns in the grid — over EVERY row, with both spans taken into account.
 *
 * Counting the first row alone is how a ragged table lost its extra columns. A
 * report table is often ten `label : value` rows with two of them broken into
 * five cells for a measurement pair; the first row says three, colSizes came
 * out three long, and the two columns that exist only in those rows had no
 * entry — so `getTableCellSize` summed them to 0, the cells fell back to the
 * constant min/max width, and dragging their border moved the handle and
 * nothing else. The three columns every row shared resized normally, which is
 * what made it look like a resize bug rather than a missing column.
 *
 * A rowSpan pushes the row below it along, so the grid is walked with an
 * occupancy map rather than by cell index — the same reading the print page's
 * `fitReportTablesToPage` needs, and for the same reason.
 */
function columnCount(table: any): number {
  const rows: any[] = table?.children ?? [];
  const taken = new Set<string>();
  let columns = 0;

  rows.forEach((row: any, y: number) => {
    let x = 0;
    (row?.children ?? []).forEach((cell: any) => {
      while (taken.has(`${y}:${x}`)) x++;
      const span = nodeCellSpan(cell, "colSpan");
      const down = nodeCellSpan(cell, "rowSpan");
      for (let dy = 0; dy < down; dy++) {
        for (let dx = 0; dx < span; dx++) taken.add(`${y + dy}:${x + dx}`);
      }
      x += span;
      if (x > columns) columns = x;
    });
  });

  return columns;
}

function needsColSizes(table: any, count: number): boolean {
  const sizes = table?.colSizes;
  if (!Array.isArray(sizes)) return true;
  if (sizes.length !== count) return true;
  // A 0 entry is as unresizable as a missing one — it sums to a 0 width.
  return sizes.some((size: unknown) => !Number.isFinite(size as number) || (size as number) <= 0);
}

/**
 * Keep whatever widths are already stated and fill only the gaps, so a table
 * pasted with real widths is never flattened to equal columns by this.
 */
function fillColSizes(table: any, count: number): number[] {
  const existing: unknown[] = Array.isArray(table?.colSizes) ? table.colSizes : [];
  const stated = existing.filter(
    (size): size is number => Number.isFinite(size as number) && (size as number) > 0
  );
  const fallback = stated.length
    ? stated.reduce((a, b) => a + b, 0) / stated.length
    : DEFAULT_TABLE_WIDTH / count;

  return Array.from({ length: count }, (_, i) => {
    const size = existing[i];
    const usable = Number.isFinite(size as number) && (size as number) > 0 ? (size as number) : fallback;
    return Math.max(MIN_COLUMN_PX, Math.round(usable));
  });
}

export const TableColSizesPlugin = createSlatePlugin({
  key: "tableColSizes",
  extendEditor: ({ editor }) => {
    const originalNormalizeNode = editor.normalizeNode as (entry: NodeEntry) => void;

    editor.normalizeNode = (entry: NodeEntry) => {
      const [node, path] = entry;

      if ((node as any).type === editor.getType(KEYS.table)) {
        const count = columnCount(node);
        if (count > 0 && needsColSizes(node, count)) {
          // Idempotent: the value written always satisfies needsColSizes(), so
          // the re-normalization this triggers stops here.
          (editor as SlateEditor).tf.setNodes({ colSizes: fillColSizes(node, count) } as any, {
            at: path,
          });
          return;
        }
      }

      originalNormalizeNode(entry);
    };

    return editor;
  },
});

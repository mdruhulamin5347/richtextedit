"use client";

import {
  TableCellHeaderPlugin,
  TableCellPlugin,
  TablePlugin,
  TableRowPlugin,
} from "@platejs/table/react";

import {
  TableCellElement,
  TableCellHeaderElement,
  TableElement,
  TableRowElement,
} from "@/components/ui/table-node";
import {
  editorContentWidth,
  extractCellAlign,
  extractCellBackground,
  extractCellBorders,
  extractCellPadding,
  extractCellVerticalAlign,
  extractRowSize,
  tableColSizes,
} from "@/lib/table-widths";

/**
 * Carry pasted column widths onto the table node.
 *
 * Plate's own table deserializer drops every width, so a pasted table renders
 * with equal columns regardless of what the source said. `colSizes` is what
 * @platejs/table sizes columns from.
 *
 * The widths are then SCALED to the width of the editor. Word states a table at
 * the width of a page, which is little more than half of this editor, so
 * reproducing the numbers literally left a pasted report in the left half of
 * the screen with the right half blank. Only the ratios between the columns
 * were ever meaningful — the serializer emits the table as `width: 100%`, so
 * every print already stretches it across the page.
 *
 * Supplying `parse` replaces the default one, which is why it returns `type`
 * itself — the same pattern list-classic-kit uses.
 */
const tableWidthParsers = {
  html: {
    deserializer: {
      parse: ({ editor, element, type }: { editor: any; element: HTMLElement; type: string }) => {
        const colSizes = tableColSizes(element, editorContentWidth(editor));
        return { type, ...(colSizes ? { colSizes } : {}) };
      },
    },
  },
};

/**
 * Carry a pasted row's height onto the row node. @platejs/table reads
 * `size` off the <tr> and turns it into the cells' minHeight.
 */
const tableRowHeightParsers = {
  html: {
    deserializer: {
      parse: ({ element, type }: { element: HTMLElement; type: string }) => {
        const size = extractRowSize(element);
        return { type, ...(size ? { size } : {}) };
      },
    },
  },
};

/**
 * Carry a pasted cell's borders onto the cell node, in Plate's own `borders`
 * shape.
 *
 * Word writes the borderless tables it uses purely to align label / value
 * columns as cells with `border:none`. Without this they arrive with no border
 * information at all, Plate applies its default 1px line, and the editor draws
 * — and the report prints — a grid that was never in the source document.
 *
 * The alignment comes along for the same reason: Word puts it on the CELL as
 * often as on the paragraph inside it (`<td align=center>` around a plain
 * `<p>`), and the align plugin only ever looks at the paragraph — so a centred
 * header row came back left-aligned.
 *
 * Shading and vertical alignment likewise: the serializer could always write
 * `background` and `verticalAlign`, but nothing ever put them on the node, so
 * a banded table arrived plain white and every cell sat at HTML's default
 * `middle` where Word had put it at `top`.
 *
 * And the PADDING, which is what the old Summernote editor got right for free:
 * it pasted the clipboard HTML straight into a contenteditable, so `padding:0cm
 * 5.4pt` on a Word cell was simply what the browser drew. Here every cell was
 * rebuilt with a fixed 12px/8px box instead, adding ~16px to the height of
 * every single row — the reason a pasted table stood taller than the document
 * it was copied from.
 */
const tableCellParsers = {
  html: {
    deserializer: {
      parse: ({ element, type }: { element: HTMLElement; type: string }) => {
        const borders = extractCellBorders(element);
        const align = extractCellAlign(element);
        const background = extractCellBackground(element);
        const verticalAlign = extractCellVerticalAlign(element);
        const padding = extractCellPadding(element);
        return {
          type,
          ...(borders ? { borders } : {}),
          ...(align ? { align } : {}),
          ...(background ? { background } : {}),
          ...(verticalAlign ? { verticalAlign } : {}),
          ...(padding ? { padding } : {}),
        };
      },
    },
  },
};

export const TableKit = [
  TablePlugin.withComponent(TableElement).configure({
    parsers: tableWidthParsers,
  }),
  TableRowPlugin.withComponent(TableRowElement).configure({ parsers: tableRowHeightParsers }),
  TableCellPlugin.withComponent(TableCellElement).configure({ parsers: tableCellParsers }),
  TableCellHeaderPlugin.withComponent(TableCellHeaderElement).configure({
    parsers: tableCellParsers,
  }),
];

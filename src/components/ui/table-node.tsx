"use client";

import * as React from "react";

import type * as DropdownMenuPrimitive from "@radix-ui/react-dropdown-menu";

import { useDraggable, useDropLine } from "@platejs/dnd";
import { BlockSelectionPlugin, useBlockSelected } from "@platejs/selection/react";
import { setCellBackground } from "@platejs/table";
import {
  TablePlugin,
  TableProvider,
  useTableBordersDropdownMenuContentState,
  useTableCellElement,
  useTableCellElementResizable,
  useTableElement,
  useTableMergeState,
} from "@platejs/table/react";
import { PopoverAnchor } from "@radix-ui/react-popover";
import { cva } from "class-variance-authority";
import {
  ArrowDown,
  ArrowLeft,
  ArrowRight,
  ArrowUp,
  CombineIcon,
  EraserIcon,
  Grid2X2Icon,
  GripVertical,
  PaintBucketIcon,
  SquareSplitHorizontalIcon,
  Trash2Icon,
  XIcon,
} from "lucide-react";
import {
  type TElement,
  type TTableCellElement,
  type TTableElement,
  type TTableRowElement,
  KEYS,
  PathApi,
} from "platejs";
import {
  type PlateElementProps,
  PlateElement,
  useComposedRef,
  useEditorPlugin,
  useEditorRef,
  useEditorSelector,
  useElement,
  useElementSelector,
  useFocusedLast,
  usePluginOption,
  useReadOnly,
  useRemoveNodeButton,
  useSelected,
  withHOC,
} from "platejs/react";

import { Button } from "@/components/ui/button";
import {
  DropdownMenu,
  DropdownMenuCheckboxItem,
  DropdownMenuContent,
  DropdownMenuGroup,
  DropdownMenuItem,
  DropdownMenuPortal,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import { Popover, PopoverContent } from "@/components/ui/popover";
import {
  borderShorthand as cellBorderShorthand,
  type CellBorder,
  DEFAULT_CELL_PADDING,
} from "@/lib/table-widths";
import { cn } from "@/lib/utils";

import { blockSelectionVariants } from "./block-selection";
import { ColorDropdownMenuItems, DEFAULT_COLORS } from "./font-color-toolbar-button";
import { ResizeHandle } from "./resize-handle";
import {
  BorderAllIcon,
  BorderBottomIcon,
  BorderLeftIcon,
  BorderNoneIcon,
  BorderRightIcon,
  BorderTopIcon,
} from "./table-icons";
import { Toolbar, ToolbarButton, ToolbarGroup, ToolbarMenuGroup } from "./toolbar";
/**
 * Room for the last column's resize handle, which is drawn a few pixels past
 * the table's right edge.
 *
 * It lives as PADDING on the scroll box, not as a `calc()` on the table's
 * max-width: the table's parent is `w-fit`, so a percentage max-width there
 * resolves against the table's own preferred width and shrank every table by
 * 6px whether or not anything was constraining it.
 */
const HANDLE_OVERHANG = "pr-1.5";

/**
 * The row-control column TableRowElement puts in front of every row — `w-2`,
 * so 8px — while the table is editable.
 *
 * It is a real cell in a real column, so the table is that much wider than its
 * own columns add up to. Both the table's width and the percentage each column
 * is sized with have to allow for it, or the control column silently takes its
 * 8px OUT of the data columns and every one of them renders a little narrower
 * than the document said.
 */
const CONTROL_COLUMN_PX = 8;

export const TableElement = withHOC(
  TableProvider,
  function TableElement({ children, ...props }: PlateElementProps<TTableElement>) {
    const readOnly = useReadOnly();
    const isSelectionAreaVisible = usePluginOption(BlockSelectionPlugin, "isSelectionAreaVisible");
    const hasControls = !readOnly && !isSelectionAreaVisible;
    const { marginLeft, props: tableProps } = useTableElement();

    // What the columns add up to. `colSizes` is the only place a table's widths
    // live, and every table that arrives by paste or import has them.
    const colSizes = (props.element.colSizes as number[] | undefined) ?? [];
    const naturalWidth = colSizes.reduce(
      (total, size) => total + (Number.isFinite(size) ? size : 0),
      0
    );
    // The same total TableCellElement sizes a cell against, so a <col> below and
    // the cell inside it ask for exactly the same width.
    const trackWidth = naturalWidth + (hasControls ? CONTROL_COLUMN_PX : 0);

    const isSelectingTable = useBlockSelected(props.element.id as string);

    const content = (
      <PlateElement
        {...props}
        className={cn(
          // Only what the drag handles need to stay inside the scroll box: the
          // column handle is drawn 8px above the table, the row handle 4px
          // below. `py-5` was 20px each way — a gap under the table that Word
          // does not have, and the first thing you see under a pasted one.
          // It cannot go to zero: `overflow-x-auto` forces overflow-y to auto,
          // so anything outside the box is clipped.
          "overflow-x-auto pt-2 pb-1",
          HANDLE_OVERHANG,
          hasControls && "-ml-2 *:data-[slot=block-selection]:left-2"
        )}
        style={{ paddingLeft: marginLeft }}
      >
        {/*
          `max-w-full` so the wrapper can never be wider than the block it sits
          in, and the table below carries its natural width as a CAP rather than
          a floor. See the width/maxWidth pair on the table itself.
        */}
        <div className="group/table relative w-fit max-w-full">
          <table
            className={cn(
              // Deviation from RadioLens, and a deliberate one: the `table`
              // utility (display: table) is REDUNDANT on a <table>, and its
              // class name is Bootstrap's table component. Carrying it here let
              // Bootstrap apply `width: 100%` — which makes a fixed-layout table
              // fill its container, divide columns equally and ignore the
              // per-cell widths that column resizing drives — and
              // `.table > :not(caption) > * > * { border-bottom-width: 1px }`,
              // which drew a line under every row of even a borderless table.
              // Dropping the class changes nothing about the rendering and
              // detaches the editor from that component entirely.
              "mr-0 ml-px h-px table-fixed border-collapse"
              // NOT `selection:bg-transparent`, which is what upstream puts
              // here. That class hides the browser's own selection highlight
              // because the block/cell-selection overlay paints its own instead
              // — and BlockSelectionPlugin is not in this editor's plugin set,
              // so nothing painted anything. Text inside a table selected and
              // copied perfectly well; it just never LOOKED selected, which is
              // indistinguishable from being unselectable. The container's
              // `selection:bg-brand/25` now reaches into the table like it does
              // everywhere else.
            )}
            {...tableProps}
            style={{
              ...(tableProps as { style?: React.CSSProperties }).style,
              // The table's own width is its NATURAL one, capped at the block.
              // On a narrow window the cap wins, and because every column is
              // sized as a percentage (see TableCellElement) they all shrink
              // together instead of the block growing a horizontal scrollbar.
              //
              // Both halves are needed. A px width alone does not shrink: under
              // `table-layout: fixed` a px column width beats `max-width`, and
              // the browser widens the table past the cap anyway — measured.
              // Room for the resize handle comes from the scroll box's own
              // padding (HANDLE_OVERHANG), so `100%` here is already short of
              // the block by that much.
              ...(naturalWidth
                ? {
                    width: naturalWidth + (hasControls ? CONTROL_COLUMN_PX : 0),
                    maxWidth: "100%",
                  }
                : null),
            }}
          >
            {/*
              The columns, stated once for the whole table.

              Under `table-layout: fixed` a browser takes its column widths from
              the <colgroup> if there is one and from THE FIRST ROW if there is
              not — and a report table is routinely ragged, ten `label : value`
              rows with two of them broken into five cells for a measurement
              pair. With no colgroup the two columns that exist only in those
              rows were never sized: dragging their border updated `colSizes`
              correctly and the render ignored it, so the handle moved and the
              column did not, while the three columns the first row did cover
              resized normally.

              Percentages, not pixels, and against the same `trackWidth`
              TableCellElement uses — see the note on the table's own width for
              why a px column width would stop the table shrinking to a narrow
              window.
            */}
            {naturalWidth > 0 && (
              <colgroup contentEditable={false}>
                {hasControls && (
                  <col style={{ width: `${(CONTROL_COLUMN_PX / trackWidth) * 100}%` }} />
                )}
                {colSizes.map((size, index) => (
                  <col
                    key={index}
                    style={{
                      width: `${((Number.isFinite(size) ? size : 0) / trackWidth) * 100}%`,
                    }}
                  />
                ))}
              </colgroup>
            )}
            <tbody className="min-w-full">{children}</tbody>
          </table>

          {isSelectingTable && <div className={blockSelectionVariants()} contentEditable={false} />}
        </div>
      </PlateElement>
    );

    if (readOnly) {
      return content;
    }

    return <TableFloatingToolbar>{content}</TableFloatingToolbar>;
  }
);

function TableFloatingToolbar({ children, ...props }: React.ComponentProps<typeof PopoverContent>) {
  const { tf } = useEditorPlugin(TablePlugin);
  const selected = useSelected();
  const element = useElement<TTableElement>();
  const { props: buttonProps } = useRemoveNodeButton({ element });
  const collapsedInside = useEditorSelector(
    (editor) => selected && editor.api.isCollapsed(),
    [selected]
  );
  const isFocusedLast = useFocusedLast();

  const { canMerge, canSplit } = useTableMergeState();

  return (
    <Popover open={isFocusedLast && (canMerge || canSplit || collapsedInside)} modal={false}>
      <PopoverAnchor asChild>{children}</PopoverAnchor>
      <PopoverContent
        asChild
        onOpenAutoFocus={(e) => e.preventDefault()}
        contentEditable={false}
        {...props}
      >
        <Toolbar
          className="scrollbar-hide bg-popover flex w-auto max-w-[80vw] flex-row overflow-x-auto rounded-md border p-1 shadow-md print:hidden"
          contentEditable={false}
        >
          <ToolbarGroup>
            <ColorDropdownMenu tooltip="Background color">
              <PaintBucketIcon />
            </ColorDropdownMenu>
            {canMerge && (
              <ToolbarButton
                onClick={() => tf.table.merge()}
                onMouseDown={(e) => e.preventDefault()}
                tooltip="Merge cells"
              >
                <CombineIcon />
              </ToolbarButton>
            )}
            {canSplit && (
              <ToolbarButton
                onClick={() => tf.table.split()}
                onMouseDown={(e) => e.preventDefault()}
                tooltip="Split cell"
              >
                <SquareSplitHorizontalIcon />
              </ToolbarButton>
            )}

            <DropdownMenu modal={false}>
              <DropdownMenuTrigger asChild>
                <ToolbarButton tooltip="Cell borders">
                  <Grid2X2Icon />
                </ToolbarButton>
              </DropdownMenuTrigger>

              <DropdownMenuPortal>
                <TableBordersDropdownMenuContent />
              </DropdownMenuPortal>
            </DropdownMenu>

            {collapsedInside && (
              <ToolbarGroup>
                <ToolbarButton tooltip="Delete table" {...buttonProps}>
                  <Trash2Icon />
                </ToolbarButton>
              </ToolbarGroup>
            )}
          </ToolbarGroup>

          {collapsedInside && (
            <ToolbarGroup>
              <ToolbarButton
                onClick={() => {
                  tf.insert.tableRow({ before: true });
                }}
                onMouseDown={(e) => e.preventDefault()}
                tooltip="Insert row before"
              >
                <ArrowUp />
              </ToolbarButton>
              <ToolbarButton
                onClick={() => {
                  tf.insert.tableRow();
                }}
                onMouseDown={(e) => e.preventDefault()}
                tooltip="Insert row after"
              >
                <ArrowDown />
              </ToolbarButton>
              <ToolbarButton
                onClick={() => {
                  tf.remove.tableRow();
                }}
                onMouseDown={(e) => e.preventDefault()}
                tooltip="Delete row"
              >
                <XIcon />
              </ToolbarButton>
            </ToolbarGroup>
          )}

          {collapsedInside && (
            <ToolbarGroup>
              <ToolbarButton
                onClick={() => {
                  tf.insert.tableColumn({ before: true });
                }}
                onMouseDown={(e) => e.preventDefault()}
                tooltip="Insert column before"
              >
                <ArrowLeft />
              </ToolbarButton>
              <ToolbarButton
                onClick={() => {
                  tf.insert.tableColumn();
                }}
                onMouseDown={(e) => e.preventDefault()}
                tooltip="Insert column after"
              >
                <ArrowRight />
              </ToolbarButton>
              <ToolbarButton
                onClick={() => {
                  tf.remove.tableColumn();
                }}
                onMouseDown={(e) => e.preventDefault()}
                tooltip="Delete column"
              >
                <XIcon />
              </ToolbarButton>
            </ToolbarGroup>
          )}
        </Toolbar>
      </PopoverContent>
    </Popover>
  );
}

function TableBordersDropdownMenuContent(
  props: React.ComponentProps<typeof DropdownMenuPrimitive.Content>
) {
  const editor = useEditorRef();
  const {
    getOnSelectTableBorder,
    hasBottomBorder,
    hasLeftBorder,
    hasNoBorders,
    hasOuterBorders,
    hasRightBorder,
    hasTopBorder,
  } = useTableBordersDropdownMenuContentState();

  return (
    <DropdownMenuContent
      className="min-w-[220px]"
      onCloseAutoFocus={(e) => {
        e.preventDefault();
        editor.tf.focus();
      }}
      align="start"
      side="right"
      sideOffset={0}
      {...props}
    >
      <DropdownMenuGroup>
        <DropdownMenuCheckboxItem
          checked={hasTopBorder}
          onCheckedChange={getOnSelectTableBorder("top")}
        >
          <BorderTopIcon />
          <div>Top Border</div>
        </DropdownMenuCheckboxItem>
        <DropdownMenuCheckboxItem
          checked={hasRightBorder}
          onCheckedChange={getOnSelectTableBorder("right")}
        >
          <BorderRightIcon />
          <div>Right Border</div>
        </DropdownMenuCheckboxItem>
        <DropdownMenuCheckboxItem
          checked={hasBottomBorder}
          onCheckedChange={getOnSelectTableBorder("bottom")}
        >
          <BorderBottomIcon />
          <div>Bottom Border</div>
        </DropdownMenuCheckboxItem>
        <DropdownMenuCheckboxItem
          checked={hasLeftBorder}
          onCheckedChange={getOnSelectTableBorder("left")}
        >
          <BorderLeftIcon />
          <div>Left Border</div>
        </DropdownMenuCheckboxItem>
      </DropdownMenuGroup>

      <DropdownMenuGroup>
        <DropdownMenuCheckboxItem
          checked={hasNoBorders}
          onCheckedChange={getOnSelectTableBorder("none")}
        >
          <BorderNoneIcon />
          <div>No Border</div>
        </DropdownMenuCheckboxItem>
        <DropdownMenuCheckboxItem
          checked={hasOuterBorders}
          onCheckedChange={getOnSelectTableBorder("outer")}
        >
          <BorderAllIcon />
          <div>Outside Borders</div>
        </DropdownMenuCheckboxItem>
      </DropdownMenuGroup>
    </DropdownMenuContent>
  );
}

function ColorDropdownMenu({ children, tooltip }: { children: React.ReactNode; tooltip: string }) {
  const [open, setOpen] = React.useState(false);

  const editor = useEditorRef();
  const selectedCells = usePluginOption(TablePlugin, "selectedCells");

  const onUpdateColor = React.useCallback(
    (color: string) => {
      setOpen(false);
      setCellBackground(editor, { color, selectedCells: selectedCells ?? [] });
    },
    [selectedCells, editor]
  );

  const onClearColor = React.useCallback(() => {
    setOpen(false);
    setCellBackground(editor, {
      color: null,
      selectedCells: selectedCells ?? [],
    });
  }, [selectedCells, editor]);

  return (
    <DropdownMenu open={open} onOpenChange={setOpen} modal={false}>
      <DropdownMenuTrigger asChild>
        <ToolbarButton tooltip={tooltip}>{children}</ToolbarButton>
      </DropdownMenuTrigger>

      <DropdownMenuContent align="start">
        <ToolbarMenuGroup label="Colors">
          <ColorDropdownMenuItems
            className="px-2"
            colors={DEFAULT_COLORS}
            updateColor={onUpdateColor}
          />
        </ToolbarMenuGroup>
        <DropdownMenuGroup>
          <DropdownMenuItem className="p-2" onClick={onClearColor}>
            <EraserIcon />
            <span>Clear</span>
          </DropdownMenuItem>
        </DropdownMenuGroup>
      </DropdownMenuContent>
    </DropdownMenu>
  );
}

export function TableRowElement({ children, ...props }: PlateElementProps<TTableRowElement>) {
  const { element } = props;
  const readOnly = useReadOnly();
  const selected = useSelected();
  const editor = useEditorRef();
  const isSelectionAreaVisible = usePluginOption(BlockSelectionPlugin, "isSelectionAreaVisible");
  const hasControls = !readOnly && !isSelectionAreaVisible;

  const { isDragging, nodeRef, previewRef, handleRef } = useDraggable({
    element,
    type: element.type,
    canDropNode: ({ dragEntry, dropEntry }) =>
      PathApi.equals(PathApi.parent(dragEntry[1]), PathApi.parent(dropEntry[1])),
    onDropHandler: (_, { dragItem }) => {
      const dragElement = (dragItem as { element: TElement }).element;

      if (dragElement) {
        editor.tf.select(dragElement);
      }
    },
  });

  return (
    <PlateElement
      {...props}
      ref={useComposedRef(props.ref, previewRef, nodeRef)}
      as="tr"
      className={cn("group/row", isDragging && "opacity-50")}
      attributes={{
        ...props.attributes,
        "data-selected": selected ? "true" : undefined,
      }}
    >
      {hasControls && (
        <td className="w-2 select-none" contentEditable={false}>
          <RowDragHandle dragRef={handleRef} />
          <RowDropLine />
        </td>
      )}

      {children}
    </PlateElement>
  );
}

function RowDragHandle({ dragRef }: { dragRef: React.Ref<any> }) {
  const editor = useEditorRef();
  const element = useElement();

  return (
    <Button
      ref={dragRef}
      variant="outline"
      className={cn(
        "absolute top-1/2 left-0 z-51 h-6 w-4 -translate-y-1/2 p-0 focus-visible:ring-0 focus-visible:ring-offset-0",
        "cursor-grab active:cursor-grabbing",
        'opacity-0 transition-opacity duration-100 group-hover/row:opacity-100 group-has-data-[resizing="true"]/row:opacity-0'
      )}
      onClick={() => {
        editor.tf.select(element);
      }}
    >
      <GripVertical className="text-muted-foreground" />
    </Button>
  );
}

function RowDropLine() {
  const { dropLine } = useDropLine();

  if (!dropLine) return null;

  return (
    <div
      className={cn(
        "bg-brand/50 absolute inset-x-0 left-2 z-50 h-0.5",
        dropLine === "top" ? "-top-px" : "-bottom-px"
      )}
    />
  );
}

/**
 * Table GRIDLINES: the line drawn on an edge the document leaves blank, so the
 * author can see where the cells are. It belongs to the view, not to the
 * document — nothing prints it, and the saved report never carries it.
 *
 * The reports these tables come from are full of BORDERLESS ones: Word's usual
 * way of lining up a label / colon / value column is a table with every rule
 * turned off, and the editor drew it exactly as the printed report shows it,
 * which is to say invisibly. There was nothing on screen to tell a three-column
 * table from three runs of text, and nothing to aim a column-resize drag at.
 *
 * A real rule, in pencil. Same 1px solid line the table would have if it were
 * bordered, so the grid reads as the table it is, and light enough — about the
 * grey Bootstrap draws its own table borders in — that nobody mistakes it for a
 * rule the document asked for. Stated against `--foreground` so it stays a
 * pencil line whatever the editor's ink is, rather than a hard-coded grey.
 */
const CELL_GUIDE = "1px solid color-mix(in srgb, var(--foreground) 14%, transparent)";

/** No line: the document draws this edge itself, or it is not ours to draw. */
const NO_GUIDE = "0";

/**
 * A guide only where this cell OWNS the edge and the document draws nothing on
 * it.
 *
 * Plate's collapse emulation hands a cell its bottom and its right, plus a top
 * only in the first row and a left only in the first column, so each shared
 * edge belongs to exactly one cell (`undefined` on the other). Guiding an edge
 * that is not ours would double every interior line and hang a pencil shadow
 * under every real rule, half a pixel off it.
 */
function cellGuide(border: Partial<CellBorder> | undefined): string {
  if (!border) return NO_GUIDE;
  return (border.size ?? 0) > 0 ? NO_GUIDE : CELL_GUIDE;
}

export function TableCellElement({
  isHeader,
  ...props
}: PlateElementProps<TTableCellElement> & {
  isHeader?: boolean;
}) {
  const { api } = useEditorPlugin(TablePlugin);
  const readOnly = useReadOnly();
  const element = props.element;

  const tableId = useElementSelector(([node]) => node.id as string, [], {
    key: KEYS.table,
  });
  const rowId = useElementSelector(([node]) => node.id as string, [], {
    key: KEYS.tr,
  });
  // What every column of this table adds up to, so a cell can state its width
  // as a SHARE of the table rather than a fixed number of pixels.
  const tableWidth: number = useElementSelector(
    ([node]) =>
      (((node as { colSizes?: number[] }).colSizes ?? []) as number[]).reduce(
        (total, size) => total + (Number.isFinite(size) ? size : 0),
        0
      ),
    [],
    { key: KEYS.table }
  );
  const isSelectingTable = useBlockSelected(tableId);
  const isSelectingRow = useBlockSelected(rowId) || isSelectingTable;
  const isSelectionAreaVisible = usePluginOption(BlockSelectionPlugin, "isSelectionAreaVisible");

  const { borders, colIndex, colSpan, minHeight, rowIndex, selected, width } =
    useTableCellElement();
  const hasControls = !readOnly && !isSelectionAreaVisible;

  const { bottomProps, hiddenLeft, leftProps, rightProps } = useTableCellElementResizable({
    colIndex,
    colSpan,
    rowIndex,
  });

  /**
   * The column's width as a SHARE of the table, not a pixel count.
   *
   * A pixel width is a floor: it pins the column open, so the table cannot
   * shrink below the sum of them and a narrow window gets a horizontal
   * scrollbar instead of a table that reflows. A percentage of a
   * `table-layout: fixed` table shrinks with it and holds the column's
   * proportion exactly — which is all a pasted width ever meant, since the
   * table is serialized as `width: 100%` and every print stretches it to the
   * page anyway.
   *
   * The pixel min/max below stays as the fallback for a table carrying no
   * colSizes at all: that is Plate's own default and nothing here should
   * change it.
   */
  const cellStyle: React.CSSProperties & Record<string, unknown> = {
    "--cellBackground": element.background,
    // The whole rule the document draws on each edge — width, style and colour
    // — so a 2.25pt double line under a header renders as a 2.25pt double
    // line. See `cellBorderShorthand`.
    "--rl-border-top": cellBorderShorthand(borders.top),
    "--rl-border-right": cellBorderShorthand(borders.right),
    "--rl-border-bottom": cellBorderShorthand(borders.bottom),
    "--rl-border-left": cellBorderShorthand(borders.left),
    // The editor's own dashes, kept in their own four vars — and drawn on their
    // own pseudo-element — so nothing the DOCUMENT says about this cell is
    // mixed with what the editor adds to help you read it. See `cellGuide`.
    "--rl-guide-top": cellGuide(borders.top),
    "--rl-guide-right": cellGuide(borders.right),
    "--rl-guide-bottom": cellGuide(borders.bottom),
    "--rl-guide-left": cellGuide(borders.left),
    // Alignment the source put on the cell rather than on its paragraphs.
    ...(element.align ? { textAlign: element.align as React.CSSProperties["textAlign"] } : null),
    ...(element.verticalAlign
      ? { verticalAlign: element.verticalAlign as React.CSSProperties["verticalAlign"] }
      : null),
  };
  // The row's own height belongs on the CELL, not on the box inside it: a
  // content box that is itself as tall as the row leaves `vertical-align`
  // nothing to move, and the text sits at the top of a row the document had
  // centred. See the `h-full` note below.
  if (minHeight) cellStyle.height = minHeight;

  const columnWidth = Number(width) || 0;
  // Against the whole table, control column included — the same total
  // TableElement gives the table itself, so at full width a column renders at
  // exactly the pixel count the document stated.
  const trackWidth = tableWidth + (hasControls ? CONTROL_COLUMN_PX : 0);
  if (columnWidth > 0 && tableWidth > 0) {
    cellStyle.width = `${(columnWidth / trackWidth) * 100}%`;
  } else {
    cellStyle.maxWidth = width || 240;
    cellStyle.minWidth = width || 120;
  }

  return (
    <PlateElement
      {...props}
      as={isHeader ? "th" : "td"}
      className={cn(
        "bg-background h-full overflow-visible border-none p-0",
        element.background ? "bg-(--cellBackground)" : "bg-background",
        isHeader && "text-left *:m-0",
        "before:size-full",
        selected && "before:bg-brand/5 before:z-10",
        // `inset-0`, not just `size-full`: an absolutely positioned box with no
        // offsets falls back to its STATIC position, and the content of a cell
        // is vertically CENTRED (as it is in a browser — see the note below on
        // `h-full`). So in a row where one cell wraps to two lines, every
        // one-line cell drew its rules 10px down from the top of the cell and
        // 10px past the bottom: the row's grid came apart, one box floating
        // above another. The overlay has to cover the CELL, wherever its text
        // happens to sit.
        "before:absolute before:inset-0 before:box-border before:content-[''] before:select-none",
        // The DOCUMENT's line, not the theme's, and the WHOLE of it. These used
        // to be `before:border-b` and friends — a Tailwind width utility, which
        // is 1px and solid and nothing else — so every table arrived here
        // flattened to a hairline grid however it was drawn in the source: a
        // 3px frame, a double rule under a header row and a dotted separator
        // all rendered identically, and identically wrong, while the saved
        // report had carried the real widths and styles all along. The editor
        // and the print disagreed about the same table.
        //
        // An arbitrary property rather than a utility because there is no
        // Tailwind spelling of "whatever this cell says": the value is a
        // complete `border` shorthand computed per cell in `cellStyle`.
        "before:[border-top:var(--rl-border-top)]",
        "before:[border-right:var(--rl-border-right)]",
        "before:[border-bottom:var(--rl-border-bottom)]",
        "before:[border-left:var(--rl-border-left)]",
        // Gridlines, on ::after. `before` is the document's own rules and must
        // stay that way: tests/browser/table-border-design.mjs measures it
        // against a bare browser's rendering of the same markup, and a line
        // this editor invented has no business in that comparison — nor in the
        // saved report, which never sees either of them.
        //
        // `inset-0` where ::before needs none: an absolutely positioned box with
        // no offsets falls back to its STATIC position, and ::after's is after
        // the cell's content — so the guide drew one cell-height below the cell
        // it belongs to, showing as a stray tick under the table.
        "after:absolute after:inset-0 after:box-border after:size-full after:content-[''] after:select-none",
        "after:[border-top:var(--rl-guide-top)]",
        "after:[border-right:var(--rl-guide-right)]",
        "after:[border-bottom:var(--rl-guide-bottom)]",
        "after:[border-left:var(--rl-guide-left)]"
      )}
      style={cellStyle}
      attributes={{
        ...props.attributes,
        colSpan: api.table.getColSpan(element),
        rowSpan: api.table.getRowSpan(element),
      }}
    >
      {/*
        NO `h-full`, and the row's height is set on the cell rather than here.
        Both for the same reason: a content box that is as tall as the cell
        leaves `vertical-align` with nothing to move, so every cell drew its
        text at the top.

        That is not the browser's behaviour and not Word's. HTML's own default
        for a `<td>` is `vertical-align: middle`, so a cell that states nothing
        — which is most of them — centres its text, and the printed report does
        too. Here the leftover space all fell BELOW the text instead, which on a
        row whose other cell runs to three lines reads as a bottom padding the
        document never had. Measured against the same markup in a bare browser,
        the short cell of a three-line row: 18px above and 18px below there,
        0 and 40 here. Now 20 and 20. A cell that states `valign=top` still sits
        at the top, one that states `middle` still centres, and a row with a
        height of its own centres inside it.

        The gap around the text is the DOCUMENT's, not this editor's. It used to
        be a hard `px-3 py-2` — 12px each side and 8px above AND below — while
        Word writes `padding:0cm 5.4pt` on every report cell: 7.2px each side and
        nothing at all vertically. That is ~16px of height added to every row,
        which on a ten-row table is a table standing 160px deeper than the one it
        was copied from. The old Summernote editor never had the problem because
        it never rebuilt the cell: the clipboard HTML went into a contenteditable
        and the browser drew the padding as written.

        And with nothing stated — a table created here — it falls back to exactly
        what `serializeCell` writes, so the editor and the printed report finally
        agree about the same table.
      */}
      <div
        className={cn(
          "relative z-20 box-border",
          // A block inside a cell sits as tightly as it did in the document.
          //
          // ParagraphElement carries `py-1` and the editor sets `leading-normal`
          // against an 18px base — 8px of padding plus a 27px line box, where a
          // Word cell's paragraph has no padding at all and a line box its own
          // 11pt text's height. That is another ~16px on every row, on top of
          // the cell padding, and between them a pasted table stood half again
          // as deep as the document it came from.
          //
          // Scoped to the inside of a cell on purpose: the editor's own
          // paragraph spacing outside tables is not this fix's business. A
          // block that states its OWN line-height still wins, because Plate
          // injects that as an inline style.
          // `py-0!` because build/scope-editor-css.mjs marks `.py-1` !important
          // wherever Bootstrap uses that class name too — without matching it
          // this rule loses to a paragraph's own padding and the row stays tall.
          "*:data-[slate-node=element]:py-0! *:data-[slate-node=element]:leading-[normal]"
        )}
        style={{
          padding: (element as { padding?: string }).padding || DEFAULT_CELL_PADDING,
          // KEEP THE RULE OFF THE TEXT.
          //
          // The cell's rules are painted by a `::before` laid `inset-0` over
          // the whole cell, so they cover whatever is underneath rather than
          // taking room of their own. With a document that pads its cells
          // `0cm 5.4pt` — Word's own, and no vertical padding at all — the text
          // reaches the cell's edge, and a thick rule was drawn straight
          // through the descenders: measured on 5px, the bottom rule came 4.69px
          // INTO the ink of a `g` or a `y`.
          //
          // A transparent border of the same width reserves exactly the band
          // the overlay paints, so the text starts where the rule ends. It is a
          // border and not extra padding because the document's own padding is
          // a shorthand string this must not have to parse, and because
          // `box-border` already counts a border inside the box.
          //
          // Full width, not half: a browser's collapsed border is shared
          // between neighbouring cells and each reserves half, but this overlay
          // draws the whole rule inside this one cell, so the whole of it has
          // to be kept clear.
          borderStyle: "solid",
          borderColor: "transparent",
          borderTopWidth: borders.top?.size ?? 0,
          borderRightWidth: borders.right?.size ?? 0,
          borderBottomWidth: borders.bottom?.size ?? 0,
          borderLeftWidth: borders.left?.size ?? 0,
        }}
      >
        {props.children}
      </div>

      {!isSelectionAreaVisible && (
        <div
          className="group absolute top-0 size-full select-none"
          contentEditable={false}
          suppressContentEditableWarning={true}
        >
          {!readOnly && (
            <>
              <ResizeHandle
                {...rightProps}
                className="-top-2 -right-1 h-[calc(100%_+_8px)] w-2"
                data-col={colIndex}
              />
              <ResizeHandle {...bottomProps} className="-bottom-1 h-2" />
              {!hiddenLeft && (
                <ResizeHandle
                  {...leftProps}
                  className="top-0 -left-1 w-2"
                  data-resizer-left={colIndex === 0 ? "true" : undefined}
                />
              )}

              <div
                className={cn(
                  "bg-ring absolute top-0 z-30 hidden h-full w-1",
                  "right-[-1.5px]",
                  columnResizeVariants({ colIndex: colIndex as any })
                )}
              />
              {colIndex === 0 && (
                <div
                  className={cn(
                    "bg-ring absolute top-0 z-30 h-full w-1",
                    "left-[-1.5px]",
                    'fade-in animate-in hidden group-has-[[data-resizer-left]:hover]/table:block group-has-[[data-resizer-left][data-resizing="true"]]/table:block'
                  )}
                />
              )}
            </>
          )}
        </div>
      )}

      {isSelectingRow && <div className={blockSelectionVariants()} contentEditable={false} />}
    </PlateElement>
  );
}

export function TableCellHeaderElement(props: React.ComponentProps<typeof TableCellElement>) {
  return <TableCellElement {...props} isHeader />;
}

const columnResizeVariants = cva("fade-in hidden animate-in", {
  variants: {
    colIndex: {
      0: 'group-has-[[data-col="0"]:hover]/table:block group-has-[[data-col="0"][data-resizing="true"]]/table:block',
      1: 'group-has-[[data-col="1"]:hover]/table:block group-has-[[data-col="1"][data-resizing="true"]]/table:block',
      2: 'group-has-[[data-col="2"]:hover]/table:block group-has-[[data-col="2"][data-resizing="true"]]/table:block',
      3: 'group-has-[[data-col="3"]:hover]/table:block group-has-[[data-col="3"][data-resizing="true"]]/table:block',
      4: 'group-has-[[data-col="4"]:hover]/table:block group-has-[[data-col="4"][data-resizing="true"]]/table:block',
      5: 'group-has-[[data-col="5"]:hover]/table:block group-has-[[data-col="5"][data-resizing="true"]]/table:block',
      6: 'group-has-[[data-col="6"]:hover]/table:block group-has-[[data-col="6"][data-resizing="true"]]/table:block',
      7: 'group-has-[[data-col="7"]:hover]/table:block group-has-[[data-col="7"][data-resizing="true"]]/table:block',
      8: 'group-has-[[data-col="8"]:hover]/table:block group-has-[[data-col="8"][data-resizing="true"]]/table:block',
      9: 'group-has-[[data-col="9"]:hover]/table:block group-has-[[data-col="9"][data-resizing="true"]]/table:block',
      10: 'group-has-[[data-col="10"]:hover]/table:block group-has-[[data-col="10"][data-resizing="true"]]/table:block',
    },
  },
});

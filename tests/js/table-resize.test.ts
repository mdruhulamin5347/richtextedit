/**
 * Column resizing.
 *
 * Every resize path in @platejs/table is gated on the table node having
 * `colSizes`: `getTableCellSize` sums it to derive a cell's width, so with none
 * the width is 0, TableCellElement falls back to its constant
 * minWidth/maxWidth, and dragging a column border writes an override nothing
 * reads — the handle moves, the column does not.
 *
 * Plate only fills colSizes for tables created with `initialTableWidth`. A
 * table that arrives any other way — pasted from Word, deserialized from a
 * legacy Summernote report — has none, which is every table here.
 */
import { describe, expect, it } from "vitest";
import { createPlateEditor } from "platejs/react";
import { deserializeHtml, type Value } from "platejs";
import { buildPlugins } from "@/plugins";

function load(html: string) {
  const editor = createPlateEditor({ plugins: buildPlugins("clean") });
  const value = deserializeHtml(editor, { element: html }) as Value;
  editor.tf.setValue(value);
  editor.tf.normalize({ force: true });
  return (editor.children as any[]).find((n) => n.type === "table");
}

describe("colSizes are always present, so columns can be dragged", () => {
  it("fills them in for a table pasted with no widths at all", () => {
    const table = load(`<table><tr><td>a</td><td>b</td><td>c</td></tr></table>`);
    expect(Array.isArray(table.colSizes)).toBe(true);
    expect(table.colSizes).toHaveLength(3);
    for (const size of table.colSizes) expect(size).toBeGreaterThan(0);
  });

  it("fills them in for legacy Summernote HTML", () => {
    const table = load(
      `<table border="1" style="border-collapse: collapse; width: 100%;">` +
        `<tr><td style="border: 1px solid #ddd; padding: 8px;">Test</td>` +
        `<td style="border: 1px solid #ddd; padding: 8px;">Result</td></tr></table>`
    );
    expect(table.colSizes).toHaveLength(2);
    for (const size of table.colSizes) expect(size).toBeGreaterThan(0);
  });

  it("does NOT flatten widths that the source actually stated", () => {
    const table = load(
      `<table><tr><td style="width:85pt">Rate</td><td style="width:21pt">:</td><td style="width:212pt">85 b/min</td></tr></table>`
    );
    expect(table.colSizes).toEqual([113, 28, 283]);
  });

  it("fills only the gaps when widths are partially stated", () => {
    const table = load(
      `<table><colgroup><col style="width:200px"><col></colgroup><tr><td>a</td><td>b</td></tr></table>`
    );
    expect(table.colSizes).toHaveLength(2);
    expect(table.colSizes[0]).toBe(200);
    expect(table.colSizes[1]).toBeGreaterThan(0);
  });

  it("counts columns through a merged cell", () => {
    const table = load(
      `<table><tr><td colspan="3">header</td></tr><tr><td>a</td><td>b</td><td>c</td></tr></table>`
    );
    expect(table.colSizes).toHaveLength(3);
  });

  it("settles — normalization does not keep rewriting colSizes", () => {
    // A non-idempotent normalizer loops until Slate throws.
    const editor = createPlateEditor({ plugins: buildPlugins("clean") });
    const value = deserializeHtml(editor, {
      element: `<table><tr><td>a</td><td>b</td></tr></table>`,
    }) as Value;
    editor.tf.setValue(value);
    expect(() => {
      editor.tf.normalize({ force: true });
      editor.tf.normalize({ force: true });
    }).not.toThrow();
    const table = (editor.children as any[]).find((n) => n.type === "table");
    const first = [...table.colSizes];
    editor.tf.normalize({ force: true });
    expect((editor.children as any[]).find((n) => n.type === "table").colSizes).toEqual(first);
  });
});

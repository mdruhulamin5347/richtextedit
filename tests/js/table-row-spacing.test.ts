/**
 * The space a paragraph keeps above and below itself INSIDE a table cell.
 *
 * Most of the air between one table row's text and the next is not the cell's
 * padding, it is the paragraph's own margins — `margin-top:6.0pt` and its pair
 * are ordinary on a Word report's table. Plate's deserializer drops a block's
 * margins outright, so every row arrived at the same minimum height whatever
 * the document said, and the saved report lost the gap again on the way out.
 *
 * The other half of these is scope: a paragraph OUTSIDE a table must keep
 * exactly the spacing it has always had here.
 */
import { describe, expect, it } from "vitest";
import { createPlateEditor } from "platejs/react";
import { deserializeHtml, type Value } from "platejs";
import { buildPlugins } from "@/plugins";
import { plateValueToHtml } from "@/lib/html-serializer";

function clipboard(html: string) {
  return {
    types: ["text/html", "text/plain"],
    getData: (type: string) => (type === "text/html" ? html : ""),
    files: [],
    items: [],
  } as unknown as DataTransfer;
}

function paste(html: string, pasteMode: "clean" | "faithful" = "clean") {
  const editor = createPlateEditor({ plugins: buildPlugins(pasteMode) });
  editor.tf.setValue([{ type: "p", children: [{ text: "" }] }] as never);
  editor.tf.select({ anchor: { path: [0, 0], offset: 0 }, focus: { path: [0, 0], offset: 0 } });
  editor.tf.insertData(clipboard(html));
  return { value: editor.children as any[], html: plateValueToHtml(editor.children as never) };
}

function roundTrip(html: string) {
  const editor = createPlateEditor({ plugins: buildPlugins("clean") });
  const value = deserializeHtml(editor, { element: html }) as Value;
  editor.tf.setValue(value);
  editor.tf.normalize({ force: true });
  return { value: editor.children as any[], html: plateValueToHtml(editor.children as never) };
}

const cell = (text: string, style: string) =>
  `<td style='border:solid windowtext 1.0pt;padding:0cm 5.4pt'>` +
  `<p class=MsoNormal style='${style}'><span style='font-size:11.0pt'>${text}</span></p></td>`;

const table = (style: string) =>
  `<html xmlns:w="urn:schemas-microsoft-com:office:word"><body>` +
  `<table class=MsoTableGrid border=1 cellspacing=0 style='border-collapse:collapse'>` +
  `<tr>${cell("Rate", style)}${cell("85 b/min", style)}</tr></table></body></html>`;

const paragraphOf = (value: any[]) => value.find((n) => n.type === "table").children[0].children[0].children[0];

describe("a paragraph inside a pasted cell", () => {
  it("keeps the space the document put ABOVE it", () => {
    // 6pt at 96dpi.
    const p = paragraphOf(paste(table("margin-top:6.0pt;margin-bottom:6.0pt")).value);
    expect(p.marginTop).toBe("8px");
  });

  it("carries it into the saved report, so the print matches the screen", () => {
    const { html } = paste(table("margin-top:6.0pt;margin-bottom:6.0pt"));
    // Alongside the block's own font size, which the same style attribute
    // carries — see tests/js/block-font-size.test.ts.
    expect(html).toMatch(/<p style="[^"]*margin-top: 8px/);
  });

  it("states nothing where the document stated nothing", () => {
    const { value, html } = paste(table("margin:0cm"));
    expect(paragraphOf(value).marginTop).toBeUndefined();
    expect(html).not.toContain("margin-top");
  });

  it("survives a save and reopen unchanged", () => {
    const once = paste(table("margin-top:6.0pt;margin-bottom:6.0pt")).html;
    const twice = roundTrip(once);
    expect(paragraphOf(twice.value).marginTop).toBe("8px");

    // Every paragraph comes back written exactly as it was saved. Compared on
    // the paragraphs alone because a border colour converges from `#000` to the
    // CSSOM's `rgb(0, 0, 0)` on the first re-save — settled, and pinned by
    // tests/js/table-widths.test.ts — after which the whole document is stable.
    const paragraphs = (html: string) => html.match(/<p style="[^"]*"/g);
    expect(paragraphs(twice.html)).toEqual(paragraphs(once));
    expect(roundTrip(twice.html).html).toBe(twice.html);
  });

  it("keeps only the vertical sides — margin-left is how indent is stated", () => {
    const p = paragraphOf(paste(table("margin-top:6.0pt;margin-left:36.0pt")).value);
    expect(p.marginTop).toBe("8px");
    expect(p.marginLeft).toBeUndefined();
  });
});

describe("the gap UNDER a row", () => {
  /** A cell holding several paragraphs, as a multi-line finding arrives. */
  const multiCell = (lines: string[], style: string) =>
    `<td style='border:solid windowtext 1.0pt;padding:0cm 5.4pt'>` +
    lines
      .map((t) => `<p class=MsoNormal style='${style}'><span style='font-size:11.0pt'>${t}</span></p>`)
      .join("") +
    `</td>`;

  const multiTable = (lines: string[], style: string) =>
    `<html xmlns:w="urn:schemas-microsoft-com:office:word"><body>` +
    `<table class=MsoTableGrid border=1 cellspacing=0 style='border-collapse:collapse'>` +
    `<tr>${multiCell(lines, style)}</tr></table></body></html>`;

  const blocksOf = (value: any[]) =>
    value.find((n) => n.type === "table").children[0].children[0].children;

  // Word's `Normal` style, which every unmodified document carries on every
  // paragraph: space AFTER and none before. Not something an author sets.
  const WORD_DEFAULT = "margin-top:0cm;margin-bottom:10.0pt";

  it("is dropped — the last block in a cell states no space after itself", () => {
    // 10pt at 96dpi would be 13px, and it used to land on every pasted row:
    // a row of 20px text stood 33px, nothing above the text and 13px below.
    const p = paragraphOf(paste(table(WORD_DEFAULT)).value);
    expect(p.marginBottom).toBeUndefined();
  });

  it("so the saved report has no space under a row either", () => {
    expect(paste(table(WORD_DEFAULT)).html).not.toContain("margin-bottom");
  });

  it("but the gaps BETWEEN paragraphs in one cell are kept", () => {
    // Those are the spacing inside the cell, not the gap under the row.
    // Collapsing them would run a three-line finding together.
    const blocks = blocksOf(
      paste(multiTable(["Anterior", "No retro-placental collection", "Matuity: Grade-0"], WORD_DEFAULT))
        .value
    );
    expect(blocks).toHaveLength(3);
    expect(blocks[0].marginBottom).toBe("13.33px");
    expect(blocks[1].marginBottom).toBe("13.33px");
    expect(blocks[2].marginBottom).toBeUndefined();
  });

  it("and a space ABOVE still reaches even the last block", () => {
    const p = paragraphOf(paste(table("margin-top:6.0pt;margin-bottom:10.0pt")).value);
    expect(p.marginTop).toBe("8px");
    expect(p.marginBottom).toBeUndefined();
  });

  it("survives a save and reopen unchanged", () => {
    const once = paste(multiTable(["Anterior", "Matuity: Grade-0"], WORD_DEFAULT)).html;
    const twice = roundTrip(once);
    const paragraphs = (html: string) => html.match(/<p style="[^"]*"/g);
    expect(paragraphs(twice.html)).toEqual(paragraphs(once));
    expect(roundTrip(twice.html).html).toBe(twice.html);
  });
});

describe("scope", () => {
  it("a paragraph OUTSIDE a table is spaced by the editor, not by the document", () => {
    const { value, html } = paste(
      `<html xmlns:w="urn:schemas-microsoft-com:office:word"><body>` +
        `<p class=MsoNormal style='margin-top:6.0pt;margin-bottom:6.0pt'>Impression</p></body></html>`
    );
    const p = value.find((n: any) => n.type === "p" && n.children[0]?.text === "Impression");
    expect(p.marginTop).toBeUndefined();
    expect(html).not.toContain("margin-top");
  });
});

/**
 * Centred text through a paste.
 *
 * A document can state alignment two ways, and Word uses both:
 *
 *   <p align=center style='text-align:center'>   the modern pair
 *   <p align=center>                             the attribute alone
 *
 * Only the CSS half was ever read — TextAlignPlugin is configured with
 * `styleKey: "textAlign"` and looks nowhere else — so a heading centred the old
 * way pasted flush left. LibreOffice, which converts the legacy `.doc` uploads,
 * writes the attribute form constantly, and old content still carries
 * `<center>`.
 */
import { describe, expect, it } from "vitest";
import { createPlateEditor } from "platejs/react";
import { deserializeHtml, type Value } from "platejs";
import { buildPlugins } from "@/plugins";
import { inlineLegacyAlignment } from "@/lib/legacy-alignment";
import { plateValueToHtml } from "@/lib/html-serializer";

/**
 * Must return "" for anything but text/html — Slate probes
 * application/x-slate-fragment and runs atob() on whatever comes back.
 */
function clipboard(html: string) {
  return {
    types: ["text/html", "text/plain"],
    getData: (type: string) => (type === "text/html" ? html : ""),
    files: [],
    items: [],
  } as unknown as DataTransfer;
}

function paste(html: string) {
  const editor = createPlateEditor({ plugins: buildPlugins("clean") });
  editor.tf.setValue([{ type: "p", children: [{ text: "" }] }] as never);
  editor.tf.select({ anchor: { path: [0, 0], offset: 0 }, focus: { path: [0, 0], offset: 0 } });
  editor.tf.insertData(clipboard(html));
  const value = editor.children as any[];
  return { blocks: value, html: plateValueToHtml(value as Value) };
}

const alignments = (blocks: any[]) => blocks.map((block) => block.align ?? null);

describe("a centred heading survives however the document states it", () => {
  it("as the attribute alone, which is what a report title arrives as", () => {
    expect(alignments(paste(`<p align=center>TEST RESULT</p>`).blocks)).toEqual(["center"]);
  });

  it("as CSS, as it always did", () => {
    expect(alignments(paste(`<p style='text-align:center'>TEST RESULT</p>`).blocks)).toEqual([
      "center",
    ]);
  });

  it("as both at once, without the two disagreeing", () => {
    expect(
      alignments(paste(`<p align=center style='text-align:center'>TEST RESULT</p>`).blocks)
    ).toEqual(["center"]);
  });

  it("and it reaches the saved report, so the print is centred too", () => {
    expect(paste(`<p align=center>TEST RESULT</p>`).html).toContain("text-align: center");
  });

  it("right and justify come through the same way", () => {
    expect(alignments(paste(`<p align=right>Signed</p>`).blocks)).toEqual(["right"]);
    expect(alignments(paste(`<p align=justify>Body</p>`).blocks)).toEqual(["justify"]);
  });

  it("a rule of dashes is centred like any other line", () => {
    const { blocks, html } = paste(`<p align=center>--------------------</p>`);
    expect(alignments(blocks)).toEqual(["center"]);
    expect(html).toContain("--------------------");
  });
});

describe("wrappers that do not survive as nodes", () => {
  it("a <center> block centres what it holds", () => {
    expect(alignments(paste(`<center>TEST RESULT</center>`).blocks)).toEqual(["center"]);
  });

  it("including every paragraph inside it", () => {
    expect(alignments(paste(`<center><p>One</p><p>Two</p></center>`).blocks)).toEqual([
      "center",
      "center",
    ]);
  });

  it("a <div align> does the same", () => {
    expect(alignments(paste(`<div align=center>TEST RESULT</div>`).blocks)).toEqual(["center"]);
  });

  it("but never overrules a block that states its own", () => {
    expect(
      alignments(
        paste(`<div align=center><p style='text-align:left'>Left</p><p>Centre</p></div>`).blocks
      )
    ).toEqual(["left", "center"]);
  });
});

describe("alignment stated on a row rather than a cell", () => {
  const cells = (html: string) => {
    const table = paste(html).blocks.find((block: any) => block.type === "table");
    return (table?.children?.[0]?.children ?? []).map((cell: any) => cell.align ?? null);
  };

  const row = (attrs: string, cellAttrs = "") =>
    `<table border=1><tr ${attrs}>` +
    `<td ${cellAttrs} style='width:100px'><p>RESULT</p></td>` +
    `<td style='width:100px'><p>X</p></td></tr></table>`;

  it("centres every cell in the row", () => {
    // `text-align` inherits, so a centred row is centred in Word and in a
    // browser. A row is not a node that can carry an alignment here, though,
    // so the value has to be moved onto the cells or it is simply lost.
    expect(cells(row("align=center"))).toEqual(["center", "center"]);
  });

  it("in CSS as well as the attribute", () => {
    expect(cells(row("style='text-align:center'"))).toEqual(["center", "center"]);
  });

  it("and from a <tbody> or the table itself", () => {
    expect(
      cells(
        `<table border=1><tbody align=center><tr><td style='width:100px'><p>R</p></td></tr></tbody></table>`
      )
    ).toEqual(["center"]);
    expect(
      cells(
        `<table border=1 style='text-align:center'><tr><td style='width:100px'><p>R</p></td></tr></table>`
      )
    ).toEqual(["center"]);
  });

  it("but a cell that states its own outranks the row", () => {
    expect(cells(row("align=center", "align=left"))).toEqual(["left", "center"]);
  });
});

describe("what `align` must NOT be read as text alignment", () => {
  it("leaves a table's own placement alone", () => {
    // `<table align=center>` puts the TABLE in the middle of the page. Read as
    // text alignment it would centre every cell in it instead.
    const { html } = paste(
      `<table align=center border=1><tr><td style="width:100px">x</td></tr></table>`
    );
    expect(html).not.toContain("text-align: center");
  });

  it("nor when the table is only placed on the page", () => {
    // `<table align=center>` positions the TABLE. Pushed down as text
    // alignment it would centre every cell in it.
    const table = paste(
      `<table align=center border=1><tr><td style="width:100px">x</td></tr></table>`
    ).blocks.find((block: any) => block.type === "table");
    expect(table?.children?.[0]?.children?.[0]?.align).toBeUndefined();
  });

  it("and does not centre the contents of a table it merely wraps", () => {
    const { blocks } = paste(
      `<div align=center><table border=1><tr><td style="width:100px">x</td></tr></table></div>`
    );
    const table = blocks.find((block: any) => block.type === "table");
    expect(table?.children?.[0]?.children?.[0]?.align).toBeUndefined();
  });

  /**
   * The three-signature block at the foot of a report: a three-column table,
   * cells left aligned, the TABLE centred on the page. A browser stops
   * `text-align` at a table — its own stylesheet says `table { text-align:
   * start }` — so the document shows those cells left aligned. Pushing the
   * wrapper's alignment through the table centred every one of them here.
   */
  const signatureBlock = (wrap: (table: string) => string) =>
    wrap(
      `<table border=0><tr>` +
        `<td style="width:300px"><p>Medical Technologist</p><p>Department of Molecular Biology</p></td>` +
        `<td style="width:300px"><p>Dr. Nahid Sultana</p><p>Professor &amp; Head</p></td>` +
        `</tr></table>`
    );

  /** How every paragraph inside the table's cells ends up aligned. */
  const cellParagraphs = (html: string) => {
    const table = paste(html).blocks.find((block: any) => block.type === "table");
    return (table?.children?.[0]?.children ?? []).flatMap((cell: any) =>
      (cell.children ?? []).map((block: any) => block.align ?? null)
    );
  };

  it("leaves a centred signature table's cells reading as they were written", () => {
    expect(cellParagraphs(signatureBlock((table) => `<div align=center>${table}</div>`))).toEqual([
      null,
      null,
      null,
      null,
    ]);
  });

  it("the same when the document centres it with <center>", () => {
    expect(cellParagraphs(signatureBlock((table) => `<center>${table}</center>`))).toEqual([
      null,
      null,
      null,
      null,
    ]);
  });

  it("the same when it centres it in CSS", () => {
    expect(
      cellParagraphs(signatureBlock((table) => `<div style='text-align:center'>${table}</div>`))
    ).toEqual([null, null, null, null]);
  });

  it("but a paragraph beside the table is aligned, because that one is text", () => {
    const { blocks } = paste(
      `<div align=center><p>END OF REPORT</p>` +
        `<table border=0><tr><td style="width:300px"><p>Dr. Nahid Sultana</p></td></tr></table></div>`
    );
    expect(blocks.find((block: any) => block.type === "p")?.align).toBe("center");
    const table = blocks.find((block: any) => block.type === "table");
    expect(table?.children?.[0]?.children?.[0]?.children?.[0]?.align).toBeUndefined();
  });

  it("and a wrapper INSIDE a cell still centres that cell's own paragraphs", () => {
    const table = paste(
      `<table border=0><tr><td style="width:300px"><div align=center><p>Dr. Nahid Sultana</p></div></td></tr></table>`
    ).blocks.find((block: any) => block.type === "table");
    expect(table?.children?.[0]?.children?.[0]?.children?.[0]?.align).toBe("center");
  });
});

describe("the rewrite leaves the rest of the document alone", () => {
  it("keeps the <style> block Word ships its formatting in", () => {
    // Returning only the body would drop the stylesheet JuicePlugin inlines,
    // taking every font, size and column width with it.
    const { html } = paste(
      `<html><head><style><!-- p.MsoNormal {font-size:16.0pt;} --></style></head>` +
        `<body><p class=MsoNormal align=center>TEST RESULT</p></body></html>`
    );
    expect(html).toContain("text-align: center");
    expect(html).toContain("font-size: 16pt");
  });

  it("passes HTML that mentions no alignment at all straight through", () => {
    const html = `<p><b>Haemoglobin</b>: 13.5 gm/dl</p>`;
    expect(inlineLegacyAlignment(html)).toBe(html);
  });

  it("leaves alignment already stated in CSS exactly as it is", () => {
    // Such a document IS re-serialized — a row can state `text-align` in CSS
    // and that still has to be pushed onto its cells — but nothing about the
    // paragraph's own alignment changes.
    const rewritten = inlineLegacyAlignment(`<p style="text-align: center">Already CSS</p>`);
    expect(rewritten).toContain(`style="text-align: center"`);
    expect(rewritten).toContain("Already CSS");
    expect(rewritten).not.toContain("align=");
  });

  it("ignores an align value that is not an alignment", () => {
    expect(inlineLegacyAlignment(`<p align=middle>x</p>`)).not.toContain("text-align");
  });

  it("applies on the Word-upload path too, which never sees the clipboard", () => {
    // WordImportToolbarButton calls deserializeHtml directly, so it runs the
    // same rewrite itself.
    const editor = createPlateEditor({ plugins: buildPlugins("clean") });
    const value = deserializeHtml(editor, {
      element: inlineLegacyAlignment(`<p align=center>TEST RESULT</p>`),
    }) as Value;
    expect((value as any[])[0]?.align).toBe("center");
  });
});

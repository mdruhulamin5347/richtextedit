/**
 * A box drawn with Borders & Shading rather than with a text box.
 *
 * From the reader's side it is the same thing — a heading in a box at the top
 * of a report — but Word writes it as a DIV around the paragraph
 * (`mso-element:para-border-div`), and LibreOffice writes it as a border on the
 * paragraph itself. Plate drops both, so the box vanished; worse, the fill did
 * not go with it, because a background on a div reaches the RUNS inside it and
 * the heading came back as highlighted text where a box used to be.
 */
import { describe, expect, it } from "vitest";
import { createPlateEditor } from "platejs/react";
import { buildPlugins } from "@/plugins";
import { plateValueToHtml } from "@/lib/html-serializer";
import { inlineWordBorderBoxes } from "@/lib/word-border-box";

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
  return plateValueToHtml(editor.children as never);
}

/** Word's own markup for a heading with a border and a fill. */
const WORD_BOX =
  `<div style='border:solid #2E74B5 1.0pt;padding:1.0pt 4.0pt 1.0pt 4.0pt;` +
  `background:#DEEAF6;mso-element:para-border-div'>` +
  `<p class=MsoNormal align=center style='text-align:center;border:none;padding:0cm'>` +
  `<b><span style='font-size:14.0pt'>CYP2C19 GENOTYPING</span></b></p></div>`;

describe("pasting a boxed heading", () => {
  it("keeps the line, the fill and the inset, as the cell that can hold them", () => {
    const html = paste(`<html xmlns:w="urn:schemas-microsoft-com:office:word"><body>${WORD_BOX}</body></html>`);

    expect(html).toContain("<table");
    expect(html).toContain("1px solid rgb(46, 116, 181)");
    expect(html).toContain("background-color: rgb(222, 234, 246)");
    expect(html).toContain("CYP2C19 GENOTYPING");
    expect(html).toContain("text-align: center");
  });

  it("stops the fill landing on the words instead of behind them", () => {
    // A background on a div reaches the runs; on the box it belongs to the box.
    const html = paste(`<html xmlns:w="urn:schemas-microsoft-com:office:word"><body>${WORD_BOX}</body></html>`);

    expect(html).not.toMatch(/<span[^>]*background-color/);
  });

  it("reads the same box stated on the paragraph, which is what LibreOffice writes", () => {
    const html = paste(
      `<p style="border: 1pt solid #2e74b5; padding: 0.02in; background: #deeaf6"><b>CYP2C19</b></p>`
    );

    expect(html).toContain("<table");
    expect(html).toContain("solid rgb(46, 116, 181)");
    expect(html).toContain("CYP2C19");
  });
});

describe("the rewrite itself", () => {
  it("draws no box for a paragraph that states none", () => {
    const plain = `<p>CYP2C19 GENOTYPING</p>`;

    expect(inlineWordBorderBoxes(plain)).toBe(plain);
    expect(inlineWordBorderBoxes(`<p style="border: none">x</p>`)).toBe(`<p style="border: none">x</p>`);
    expect(inlineWordBorderBoxes(`<p style="border: 0px solid #000">x</p>`)).toBe(
      `<p style="border: 0px solid #000">x</p>`
    );
  });

  it("leaves a bordered paragraph inside a table alone", () => {
    // There it is a cell's own border, not a box to rebuild — and a table
    // nested in a cell is not what the document said.
    const inCell = `<table><tbody><tr><td><p style="border: 1pt solid #000">x</p></td></tr></tbody></table>`;

    expect(inlineWordBorderBoxes(inCell)).toBe(inCell);
  });

  it("builds one box, not two, for Word's div and the paragraph inside it", () => {
    const out = inlineWordBorderBoxes(WORD_BOX);

    expect((out.match(/<table/g) ?? []).length).toBe(1);
  });

  it("returns HTML with no border in it untouched, byte for byte", () => {
    const html = `<p>Gene</p>`;

    expect(inlineWordBorderBoxes(html)).toBe(html);
  });
});

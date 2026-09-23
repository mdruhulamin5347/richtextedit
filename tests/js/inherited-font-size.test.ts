/**
 * A font size the document states on a BLOCK.
 *
 * `font-size` inherits, so `<td style="font-size:12px"><p>Detected</p>` renders
 * at 12px in every browser and in the printed report. A Plate node cannot say
 * that — only a text LEAF carries a size — so the run arrived carrying nothing,
 * fell back to the editable's 18px base, and a report's body text came out
 * BIGGER on screen than in the document, while its header, which states its
 * size on the span, came out right.
 *
 * And on the way out it was worse than a display bug: the serializer writes no
 * block font-size either, so opening a stored report that stated one and saving
 * it wrote the size out of the report for good.
 */
import { describe, expect, it, beforeEach } from "vitest";
import { act } from "react";
import { createPlateEditor } from "platejs/react";
import api from "@/entry";
import { buildPlugins } from "@/plugins";
import { plateValueToHtml } from "@/lib/html-serializer";
import { inlineInheritedFontSize } from "@/lib/inherited-font-size";

function mount(initialHtml: string) {
  document.body.innerHTML = `
    <form id="f"><textarea id="report_body" hidden>${initialHtml}</textarea>
    <div class="rl-editor-scope" id="report_body_editor"></div></form>`;
  return act(async () => {
    api.mount("#report_body_editor", { textarea: "#report_body" });
  });
}

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

const cellTable = (tableStyle: string, cellStyle: string, pStyle: string) =>
  `<table border="1" style="border-collapse: collapse${tableStyle}"><tbody><tr>` +
  `<td style="padding: 2px 5px${cellStyle}"><p${pStyle}>Measles morbillivirus</p></td>` +
  `</tr></tbody></table>`;

describe("opening a stored report", () => {
  beforeEach(() => {
    document.body.innerHTML = "";
  });

  it("keeps a size stated on the paragraph", async () => {
    await mount(cellTable("", "", ` style="font-size: 12px"`));
    expect(api.getHtml("#report_body_editor")).toContain(`font-size: 12px`);
  });

  it("keeps a size stated on the cell", async () => {
    await mount(cellTable("", "; font-size: 12px", ""));
    expect(api.getHtml("#report_body_editor")).toContain(`font-size: 12px`);
  });

  it("keeps a size stated on the table", async () => {
    await mount(cellTable("; font-size: 12px", "", ""));
    expect(api.getHtml("#report_body_editor")).toContain(`font-size: 12px`);
  });

  it("keeps a size stated on a paragraph outside any table", async () => {
    await mount(`<p style="font-size: 12px">A body paragraph.</p>`);
    // On the run, which is what carries a size here — and back on the block
    // too, because a line's height is measured against the block's own font.
    // See tests/js/block-font-size.test.ts.
    expect(api.getHtml("#report_body_editor")).toBe(
      `<p style="font-size: 12px"><span style="font-size: 12px">A body paragraph.</span></p>`
    );
  });

  it("leaves a report that states no size exactly as it was", async () => {
    const plain = "<p>Pathogen Name</p><p>Result</p>";
    await mount(plain);
    expect(api.getHtml("#report_body_editor")).toBe(plain);
  });

  it("is stable once saved: the same report reopens and re-saves unchanged", async () => {
    await mount(`<p style="font-size: 12px">A body paragraph.</p>`);
    const once = api.getHtml("#report_body_editor");
    document.body.innerHTML = "";
    await mount(once);
    expect(api.getHtml("#report_body_editor")).toBe(once);
  });
});

describe("pasting", () => {
  /** Word's own markup: the header states its size, the body leaves it to the style block. */
  const WORD_TABLE =
    `<html xmlns:w="urn:schemas-microsoft-com:office:word"><head><style>` +
    `<!-- p.MsoNormal {margin:0cm; font-size:11.0pt;} --></style></head><body>` +
    `<table class=MsoTableGrid border=1 cellspacing=0 style='border-collapse:collapse'>` +
    `<tr><td style='padding:0cm 5.4pt'><p class=MsoNormal><b><span style='font-size:14.0pt'>Pathogen Name</span></b></p></td></tr>` +
    `<tr><td style='padding:0cm 5.4pt'><p class=MsoNormal>Measles morbillivirus</p></td></tr>` +
    `</table></body></html>`;

  it("keeps the header at its size and the body at the document's, not the editor's", () => {
    const html = paste(WORD_TABLE);
    // The points the document states, kept as points (see lib/font-size.ts).
    // The body used to arrive with no size at all and render at the editable's
    // 18px base — bigger than the header's own text.
    expect(html).toContain("font-size: 14pt");
    expect(html).toContain("font-size: 11pt");
  });

  it("covers HTML that is not Word's, which nothing cleaned before", () => {
    const html = paste(`<html><body>${cellTable("", "; font-size: 12px", "")}</body></html>`);
    expect(html).toContain("font-size: 12px");
  });
});

describe("the pre-CSS spelling", () => {
  beforeEach(() => {
    document.body.innerHTML = "";
  });

  it("reads `<font size>`, which Plate reads off the style and never off the attribute", async () => {
    await mount(`<p><font size="5">Big</font></p><p><font size="2">small</font></p>`);
    // HTML's absolute size table on a browser's 16px base.
    expect(api.getHtml("#report_body_editor")).toBe(
      `<p style="font-size: 24px"><span style="font-size: 24px">Big</span></p>` +
        `<p style="font-size: 13px"><span style="font-size: 13px">small</span></p>`
    );
  });

  it("reads the relative forms too", () => {
    // `+1` is one step up from the 3 a browser treats as normal: size 4, 18px.
    expect(inlineInheritedFontSize(`<p><font size="+1">x</font></p>`)).toContain("18px");
    expect(inlineInheritedFontSize(`<p><font size="-1">x</font></p>`)).toContain("13px");
  });

  it("lets a CSS size on the same element win, as a browser does", () => {
    const out = inlineInheritedFontSize(`<p><font size="7" style="font-size: 11px">x</font></p>`);
    expect(out).toContain("11px");
    expect(out).not.toContain("48px");
  });
});

describe("the rewrite itself", () => {
  it("leaves a run that states its own size alone", () => {
    const html = `<p style="font-size: 20px"><span style="font-size: 9px">small</span></p>`;
    expect(inlineInheritedFontSize(html)).toContain(`font-size: 9px`);
    expect(inlineInheritedFontSize(html)).not.toMatch(/font-size:\s*20px[^<]*>small/);
  });

  it("gives the size to an existing inline rather than adding a wrapper", () => {
    const out = inlineInheritedFontSize(`<p style="font-size: 12px"><b>Detected</b></p>`);
    // The CSSOM writes the declaration back with a trailing semicolon.
    expect(out).toMatch(/<b style="font-size: 12px;?">Detected<\/b>/);
    expect(out).not.toContain("<span");
  });

  it("wraps the text inside a link, which is an element and not a mark", () => {
    const out = inlineInheritedFontSize(`<p style="font-size: 12px"><a href="#">ref</a></p>`);
    expect(out).toMatch(/<a href="#"><span style="font-size: 12px;?">ref<\/span><\/a>/);
  });

  it("does not wrap the source file's own indentation", () => {
    // Word HTML is a pretty-printed tree; the newline runs between its tags are
    // not the report's text and must not become runs of their own.
    const out = inlineInheritedFontSize(`<div style="font-size: 12px">\n\t<p>x</p>\n</div>`);
    expect(out.match(/<span/g)?.length).toBe(1);
  });

  it("returns HTML that states no size untouched, byte for byte", () => {
    const html = `<p>Pathogen Name</p>`;
    expect(inlineInheritedFontSize(html)).toBe(html);
  });
});

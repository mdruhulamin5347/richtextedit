/**
 * A text colour the document states on a BLOCK, or in the pre-CSS spelling.
 *
 * `color` inherits, so `<td style="color:#c00000"><p>POSITIVE</p>` is red in
 * every browser and in the printed report. A Plate node cannot say that — only
 * a text LEAF carries a colour — so the run arrived carrying nothing and the
 * line came out black. And as with a size, it was worse than a display bug:
 * the serializer writes no block colour either, so opening a stored report that
 * stated one and saving it wrote the colour out of the report for good.
 */
import { describe, expect, it, beforeEach } from "vitest";
import { act } from "react";
import { createPlateEditor } from "platejs/react";
import api from "@/entry";
import { buildPlugins } from "@/plugins";
import { plateValueToHtml } from "@/lib/html-serializer";
import { inlineInheritedColor } from "@/lib/inherited-color";

function mount(initialHtml: string) {
  document.body.innerHTML = `
    <form id="f"><textarea id="report_body" hidden>${initialHtml}</textarea>
    <div class="rte-scope" id="report_body_editor"></div></form>`;
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

const cell = (cellAttrs: string, inner: string) =>
  `<table border="1" style="border-collapse: collapse"><tbody><tr>` +
  `<td style="padding: 2px 5px"${cellAttrs}>${inner}</td>` +
  `</tr></tbody></table>`;

describe("opening a stored report", () => {
  beforeEach(() => {
    document.body.innerHTML = "";
  });

  it("keeps a colour stated on the paragraph", async () => {
    await mount(`<p style="color: red">POSITIVE</p>`);
    expect(api.getHtml("#report_body_editor")).toBe(
      `<p><span style="color: red">POSITIVE</span></p>`
    );
  });

  it("keeps a colour stated on the cell", async () => {
    await mount(cell(` bgcolor="#ffff00"`, `<p style="color: red">Abnormal</p>`));
    const html = api.getHtml("#report_body_editor");
    expect(html).toContain(`color: red`);
    // …and the pre-CSS spelling of the cell's own shading with it.
    expect(html).toContain(`background-color: #ffff00`);
  });

  it("reads `<font color>`, which Plate reads off the style and never off the attribute", async () => {
    await mount(`<p><font color="#ff0000">Raised</font></p>`);
    expect(api.getHtml("#report_body_editor")).toBe(
      `<p><span style="color: rgb(255, 0, 0)">Raised</span></p>`
    );
  });

  it("leaves a report that states no colour exactly as it was", async () => {
    const plain = "<p>Pathogen Name</p><p>Result</p>";
    await mount(plain);
    expect(api.getHtml("#report_body_editor")).toBe(plain);
  });

  it("does not hang a span on ordinary black text", async () => {
    // LibreOffice writes `<font color="#000000">` on practically every run it
    // converts; black is the colour the report already renders in.
    await mount(`<p><font color="#000000">Within normal limits</font></p>`);
    expect(api.getHtml("#report_body_editor")).toBe(`<p>Within normal limits</p>`);
  });

  it("is stable once saved: the same report reopens and re-saves unchanged", async () => {
    await mount(`<p style="color: red">POSITIVE</p>`);
    const once = api.getHtml("#report_body_editor");
    document.body.innerHTML = "";
    await mount(once);
    expect(api.getHtml("#report_body_editor")).toBe(once);
  });
});

describe("pasting", () => {
  it("keeps a colour stated on the paragraph, which no pass covered before", () => {
    expect(paste(`<html><body><p style="color:#c00000">POSITIVE</p></body></html>`)).toContain(
      "color: rgb(192, 0, 0)"
    );
  });

  it("keeps a colour stated on the cell", () => {
    const html = paste(`<html><body>${cell(` bgcolor="#ffff00"`, `<p style="color:red">Abnormal</p>`)}</body></html>`);
    expect(html).toContain("color: red");
    expect(html).toContain("background-color: #ffff00");
  });

  it("keeps a colour stated on the list", () => {
    expect(paste(`<ul style="color:#008000"><li>Repeat in 3 months</li></ul>`)).toContain(
      "color: rgb(0, 128, 0)"
    );
  });

  it("keeps a colour Word states in its style block", () => {
    const word =
      `<html xmlns:w="urn:schemas-microsoft-com:office:word"><head><style>` +
      `<!-- p.MsoNormal {margin:0cm; color:#C00000; font-size:11.0pt;} --></style></head>` +
      `<body><p class=MsoNormal>Grossly abnormal</p></body></html>`;
    expect(paste(word)).toContain("color: rgb(192, 0, 0)");
  });

  it("keeps a red heading pasted from Word", () => {
    // The shape Word actually puts on the clipboard for a coloured heading.
    const word =
      `<html xmlns:w="urn:schemas-microsoft-com:office:word"><head><style><!--` +
      ` p.MsoNormal {margin:0cm; font-size:11.0pt;} --></style></head><body>` +
      `<p class=MsoNormal align=center style='text-align:center'><b>` +
      `<span style='font-size:16.0pt;color:red'>MOLECULAR BIOLOGY REPORT<o:p></o:p></span>` +
      `</b></p></body></html>`;
    expect(paste(word)).toContain("color: red");
  });

  it("keeps a `<font color>` that Word's own cleaner has renamed to a span", () => {
    // cleanDocx replaces the tag and leaves the attribute behind, where it means
    // nothing to a browser and nothing to Plate — and it runs BEFORE this pass.
    const word =
      `<html xmlns:w="urn:schemas-microsoft-com:office:word"><body>` +
      `<p class=MsoNormal><font color="#ff0000">Raised</font></p></body></html>`;
    expect(paste(word)).toContain("color: rgb(255, 0, 0)");
  });

  it("keeps a highlight, which is a colour on the run itself", () => {
    expect(paste(`<p><span style='background:yellow'>Check</span></p>`)).toContain(
      "background-color: yellow"
    );
  });

  it("keeps the colours LibreOffice writes, nested font tags and all", () => {
    // What the Word-import button gets back for a legacy .doc.
    const html = paste(
      `<p><font color="#ff0000"><font face="Calibri"><font size="3" style="font-size:12pt">` +
        `Raised</font></font></font></p>`
    );
    expect(html).toContain("color: rgb(255, 0, 0)");
    expect(html).toContain("font-size: 12pt");
  });
});

describe("the rewrite itself", () => {
  it("leaves a run that states its own colour alone", () => {
    const out = inlineInheritedColor(`<p style="color: blue">outer <span style="color: red">inner</span></p>`);
    expect(out).toMatch(/<span style="color: red;?">inner<\/span>/);
    expect(out).not.toMatch(/color:\s*blue[^<]*>inner/);
  });

  it("gives the colour to an existing inline rather than adding a wrapper", () => {
    const out = inlineInheritedColor(`<p style="color: red"><b>Detected</b></p>`);
    expect(out).toMatch(/<b style="color: red;?">Detected<\/b>/);
    expect(out).not.toContain("<span");
  });

  it("wraps the text inside a link, which is an element and not a mark", () => {
    const out = inlineInheritedColor(`<p><a href="#" style="color: red">ref</a></p>`);
    expect(out).toMatch(/<a href="#" style="color: red;?"><span style="color: red;?">ref<\/span><\/a>/);
  });

  it("does not wrap the source file's own indentation", () => {
    const out = inlineInheritedColor(`<div style="color: red">\n\t<p>x</p>\n</div>`);
    expect(out.match(/<span/g)?.length).toBe(1);
  });

  it("normalises black to the keyword Plate drops as its own default", () => {
    expect(inlineInheritedColor(`<p><font color="#000000">x</font></p>`)).toContain("color: black");
    expect(inlineInheritedColor(`<p><font color="#000">x</font></p>`)).toContain("color: black");
  });

  it("restates nothing for a colour that states nothing", () => {
    const out = inlineInheritedColor(`<p style="color: inherit">x</p>`);
    expect(out).not.toContain("<span");
  });

  it("lets a CSS colour on the same element win over the attribute, as a browser does", () => {
    const out = inlineInheritedColor(`<p><font color="#ff0000" style="color: green">x</font></p>`);
    expect(out).toContain("green");
    expect(out).not.toContain("255, 0, 0");
  });

  it("returns HTML that states no colour untouched, byte for byte", () => {
    const html = `<p>Pathogen Name</p>`;
    expect(inlineInheritedColor(html)).toBe(html);
  });

  it("leaves a report alone that only says the WORD colour, or shades a cell", () => {
    // Both get past the guard. Neither is a colour this pass can restate, so
    // neither may come back as a re-serialized copy of itself.
    const text = `<p>Stool color: brown</p>`;
    expect(inlineInheritedColor(text)).toBe(text);

    const shaded = `<table><tbody><tr><td style="background-color: yellow"><p>x</p></td></tr></tbody></table>`;
    expect(inlineInheritedColor(shaded)).toBe(shaded);
  });
});

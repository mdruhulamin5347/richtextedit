/**
 * Word paste, through the real clipboard path.
 *
 * Not `deserializeHtml` this time: `editor.tf.insertData` is what a Ctrl+V
 * actually calls, and it is the path DocxPlugin + JuicePlugin hook into. That
 * pair is the whole reason this editor was ported — Word ships its formatting
 * as a wall of `mso-` CSS that Summernote discarded.
 */
import { describe, expect, it } from "vitest";
import { createPlateEditor } from "platejs/react";
import { buildPlugins } from "@/plugins";
import { plateValueToHtml } from "@/lib/html-serializer";

/** What Word actually puts on the clipboard, trimmed but not sanitised. */
const WORD_HTML = `<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:w="urn:schemas-microsoft-com:office:word"><head><style>
<!-- p.MsoNormal {margin:0cm; font-size:11.0pt; font-family:"Calibri",sans-serif;}
.MsoTableGrid {border:solid windowtext 1.0pt;} --></style></head><body lang=EN-GB>
<p class=MsoNormal style='text-align:center'><b><span style='font-size:14.0pt;
font-family:"Times New Roman",serif'>HAEMATOLOGY REPORT</span></b></p>
<p class=MsoNormal><span style='font-size:11.0pt'>Sample: <i>Whole blood</i></span></p>
<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0
 style='border-collapse:collapse;border:solid windowtext 1.0pt'>
 <tr>
  <td style='border:solid windowtext 1.0pt;padding:0cm 5.4pt'><p class=MsoNormal><b>Test</b></p></td>
  <td style='border:solid windowtext 1.0pt;padding:0cm 5.4pt'><p class=MsoNormal><b>Result</b></p></td>
 </tr>
 <tr>
  <td style='border:solid windowtext 1.0pt;padding:0cm 5.4pt'><p class=MsoNormal>Haemoglobin</p></td>
  <td style='border:solid windowtext 1.0pt;padding:0cm 5.4pt'><p class=MsoNormal>13.5 g/dL</p></td>
 </tr>
 <tr>
  <td colspan=2 style='border:solid windowtext 1.0pt;padding:0cm 5.4pt'><p class=MsoNormal>Within normal limits</p></td>
 </tr>
</table>
<ul style='margin-top:0cm' type=disc>
 <li class=MsoNormal>Repeat in 3 months
   <ul type=circle><li class=MsoNormal>If symptoms persist</li></ul>
 </li>
</ul></body></html>`;

/** Minimal DataTransfer — jsdom does not implement a usable one. */
function clipboard(html: string, text = "") {
  return {
    types: ["text/html", "text/plain"],
    getData: (type: string) => (type === "text/html" ? html : text),
    files: [],
    items: [],
  } as unknown as DataTransfer;
}

function paste(html: string, pasteMode: "clean" | "faithful" = "clean") {
  const editor = createPlateEditor({ plugins: buildPlugins(pasteMode) });
  editor.tf.setValue([{ type: "p", children: [{ text: "" }] }] as never);
  editor.tf.select({ anchor: { path: [0, 0], offset: 0 }, focus: { path: [0, 0], offset: 0 } });
  editor.tf.insertData(clipboard(html));
  return plateValueToHtml(editor.children as never);
}

describe("pasting a Word lab report", () => {
  const html = paste(WORD_HTML);

  it("keeps the text", () => {
    expect(html).toContain("HAEMATOLOGY REPORT");
    expect(html).toContain("Haemoglobin");
    expect(html).toContain("13.5 g/dL");
    expect(html).toContain("Within normal limits");
  });

  it("keeps bold and italic", () => {
    expect(html).toContain("<strong>");
    expect(html).toContain("<em>");
  });

  it("keeps centre alignment", () => {
    expect(html).toContain("text-align: center");
  });

  it("keeps the table structure, including the merged cell", () => {
    expect(html).toContain("<table");
    expect(html).toContain('colspan="2"');
    expect((html.match(/<tr>/g) || []).length).toBe(3);
  });

  it("keeps the source table's borders, per side", () => {
    // Word states borders per side (`border:solid windowtext 1.0pt`), and they
    // are carried through rather than replaced with an invented default.
    expect(html).toContain("border-top: 1px solid #000");
    expect(html).toContain("border-bottom: 1px solid #000");
    expect(html).toContain('<table border="1"');
  });

  it("keeps the nested list nested", () => {
    expect(html).toContain("Repeat in 3 months");
    expect(html).toContain("If symptoms persist");
    const first = html.indexOf("<ul");
    expect(html.indexOf("<ul", first + 1)).toBeGreaterThan(-1);
  });

  it("drops Word's mso junk rather than storing it", () => {
    expect(html).not.toContain("mso-");
    expect(html).not.toContain("MsoNormal");
    expect(html).not.toContain("windowtext");
  });
});

describe("a BORDERLESS Word table — the label/value layout tables reports use", () => {
  // Word emits these for alignment only. They must not print as a grid.
  const BORDERLESS = `<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0
   style='border-collapse:collapse;border:none'>
   <tr>
    <td style='border:none;padding:0cm 5.4pt'><p class=MsoNormal>Rate</p></td>
    <td style='border:none;padding:0cm 5.4pt'><p class=MsoNormal>:</p></td>
    <td style='border:none;padding:0cm 5.4pt'><p class=MsoNormal>85 b/min</p></td>
   </tr>
   <tr>
    <td style='border:none;padding:0cm 5.4pt'><p class=MsoNormal>Rhythm</p></td>
    <td style='border:none;padding:0cm 5.4pt'><p class=MsoNormal>:</p></td>
    <td style='border:none;padding:0cm 5.4pt'><p class=MsoNormal>Regular</p></td>
   </tr></table>`;

  const html = paste(BORDERLESS);

  it("keeps the content and the layout", () => {
    expect(html).toContain("Rate");
    expect(html).toContain("85 b/min");
    expect(html).toContain("Rhythm");
    expect(html).toContain("<table");
  });

  it("draws NO cell borders", () => {
    // Stated as `border: 0` rather than left unsaid: the print pages supply a
    // `border-style: solid` for the cells Word writes without one, and an edge
    // this serializer stays silent about picks that up with CSS's `medium`
    // width — so silence prints as a 3px black box. See cellBorderStyles.
    const declarations = html.match(/border(-top|-right|-bottom|-left)?:[^;"]*/g) ?? [];
    expect(declarations.length).toBeGreaterThan(0);
    expect(declarations.every((d) => /^border:\s*0$/.test(d))).toBe(true);
  });

  it("does not mark the table as bordered", () => {
    // The print stylesheets and dompdf key a visible grid off this attribute,
    // and `border="0"` is what keeps a borderless table borderless there.
    expect(html).not.toContain('border="1"');
    expect(html).toContain('border="0"');
  });
});

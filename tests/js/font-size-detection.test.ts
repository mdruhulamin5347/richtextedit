/**
 * Every way a real clipboard states a font size, and whether it reaches a run.
 *
 * Asked for as "ensure copy data perfectly detect front-size and show our
 * editor properly". A size that reaches nothing is not a small loss: the run
 * falls back to the editable's own base and the paragraph comes out at a size
 * nobody chose — bigger than the document, as often as not, which is what made
 * the inherited-font-size bug look arbitrary ("the table header is
 * perfect but the body is large").
 *
 * So this is a MATRIX rather than a set of scenarios, and it is deliberately
 * broader than the shapes anyone has reported. `font-size` inherits, and the
 * chain it inherits down can start almost anywhere — a span, a `<p>`, a cell, a
 * row, the table, a wrapper div, `div.WordSection1`, `<body>`, a class in the
 * `<style>` block Juice has yet to inline, or a spelling that is not CSS at all
 * (`<font size=N>`, `mso-ansi-font-size`, the `font` shorthand, a CSS keyword).
 * Each one that arrives carrying nothing is one report opening at the wrong
 * size, and the pipeline has enough passes in it (see `buildPlugins`, where
 * registration order is REVERSED) that a shape can start working or stop
 * working without anything nearby being touched.
 *
 * What the expected values mean: a size stated in POINTS stays in points and a
 * legacy report's PIXELS stay in pixels — see lib/font-size.ts for why each.
 */
import { describe, expect, it } from "vitest";
import { createPlateEditor } from "platejs/react";

import { buildPlugins } from "@/plugins";
import { plateValueToHtml } from "@/lib/html-serializer";

function paste(html: string): string {
  const editor = createPlateEditor({ plugins: buildPlugins("clean") });
  editor.tf.setValue([{ type: "p", children: [{ text: "" }] }] as never);
  editor.tf.select({ anchor: { path: [0, 0], offset: 0 }, focus: { path: [0, 0], offset: 0 } });
  editor.tf.insertData({
    types: ["text/html", "text/plain"],
    getData: (type: string) => (type === "text/html" ? html : ""),
    files: [],
    items: [],
  } as unknown as DataTransfer);
  return plateValueToHtml(editor.children as never);
}

/** Every size in the saved HTML, in document order. */
function statedSizes(html: string): string[] {
  return [...html.matchAll(/font-size:\s*([^;"]+)/g)].map((m) => m[1].trim());
}

/** The size that reached the TEXT, or "(none)" — which means the editor's base. */
function detectedSize(clipboardHtml: string): string {
  const sizes = statedSizes(paste(clipboardHtml));
  return sizes.length === 0 ? "(none)" : sizes[sizes.length - 1];
}

/** A Word clipboard: the `xmlns:w` is what makes Plate's docx cleaner run. */
const word = (body: string, styleBlock = "") =>
  `<html xmlns:w="urn:schemas-microsoft-com:office:word"><head><style><!--${styleBlock}--></style>` +
  `</head><body>${body}</body></html>`;

const CELL = (cellStyle = "", tableStyle = "", rowStyle = "") =>
  word(
    `<table ${tableStyle}><tr ${rowStyle}><td ${cellStyle}><p class=MsoNormal>Detected</p></td></tr></table>`
  );

describe("where a Word clipboard states a size", () => {
  const cases: [string, string][] = [
    ["on the run itself", word(`<p class=MsoNormal><span style='font-size:10.0pt'>T</span></p>`)],
    ["on a bold run, which is also a mark carrier", word(`<p class=MsoNormal><b style='font-size:10.0pt'>T</b></p>`)],
    ["on the paragraph, inline", word(`<p class=MsoNormal style='font-size:10.0pt'>T</p>`)],
    ["in p.MsoNormal, its default paragraph style", word(`<p class=MsoNormal>T</p>`, `p.MsoNormal {font-size:10.0pt;}`)],
    ["in span.MsoNormal", word(`<p class=MsoNormal><span class=MsoNormal>T</span></p>`, `span.MsoNormal {font-size:10.0pt;}`)],
    ["in a class on the run", word(`<p class=MsoNormal><span class=s1>T</span></p>`, `.s1 {font-size:10.0pt;}`)],
    ["in table.MsoNormalTable, its default table style", word(`<table class=MsoNormalTable><tr><td><p class=MsoNormal>T</p></td></tr></table>`, `table.MsoNormalTable {font-size:10.0pt;}`)],
    ["on the cell", CELL(`style='font-size:10.0pt'`)],
    ["on the row", CELL("", "", `style='font-size:10.0pt'`)],
    ["on the table", CELL("", `style='font-size:10.0pt'`)],
    ["on a cell, inherited by a nested table", word(`<table><tr><td style='font-size:10.0pt'><table><tr><td><p class=MsoNormal>T</p></td></tr></table></td></tr></table>`)],
    ["on a wrapper div", word(`<div style='font-size:10.0pt'><p class=MsoNormal>T</p></div>`)],
    ["in div.WordSection1, the wrapper Word always writes", word(`<div class=WordSection1><p class=MsoNormal>T</p></div>`, `div.WordSection1 {font-size:10.0pt;}`)],
    ["on <body>", word(`<p class=MsoNormal>T</p>`).replace("<body>", `<body style='font-size:10.0pt'>`)],
    ["in a body rule", word(`<p class=MsoNormal>T</p>`, `body {font-size:10.0pt;}`)],
    ["on a heading, via its style", word(`<h1>T</h1>`, `h1 {font-size:10.0pt;}`)],
    ["on a list item", word(`<ul><li class=MsoNormal style='font-size:10.0pt'>T</li></ul>`)],
    ["beside an empty <o:p>", word(`<p class=MsoNormal><span style='font-size:10.0pt'>T<o:p></o:p></span></p>`)],
  ];

  for (const [where, clipboard] of cases) {
    it(`reaches the run when it is ${where}`, () => {
      expect(detectedSize(clipboard)).toBe("10pt");
    });
  }

  it("keeps a half point, which is half of Word's own size list", () => {
    expect(detectedSize(word(`<p class=MsoNormal><span style='font-size:10.5pt'>T</span></p>`))).toBe(
      "10.5pt"
    );
    expect(detectedSize(word(`<p class=MsoNormal><span style='font-size:7.5pt'>T</span></p>`))).toBe(
      "7.5pt"
    );
  });

  it("lets the nearest statement win, as a browser does", () => {
    // The paragraph says 11pt and the run inside it says 14pt.
    const sizes = statedSizes(
      paste(word(`<p class=MsoNormal style='font-size:11.0pt'><span style='font-size:14.0pt'>A</span>B</p>`))
    );
    expect(sizes).toContain("14pt");
    expect(sizes).toContain("11pt");
  });
});

describe("spellings that are not CSS font-size at all", () => {
  it("reads Word's own mso-ansi-font-size where nothing states a CSS size", () => {
    expect(
      detectedSize(word(`<p class=MsoNormal><span style='mso-ansi-font-size:10.0pt'>T</span></p>`))
    ).toBe("10pt");
  });

  it("reads mso-bidi-font-size the same way", () => {
    expect(
      detectedSize(word(`<p class=MsoNormal><span style='mso-bidi-font-size:10.0pt'>T</span></p>`))
    ).toBe("10pt");
  });

  it("but lets a real CSS size beat it, because that is what a browser draws", () => {
    expect(
      detectedSize(
        word(`<p class=MsoNormal><span style='font-size:10.0pt;mso-ansi-font-size:14.0pt'>T</span></p>`)
      )
    ).toBe("10pt");
  });

  it("reads the `font` shorthand, which CSSOM expands for us", () => {
    expect(detectedSize(`<p><span style="font: bold 10pt Arial">T</span></p>`)).toBe("10pt");
  });

  it("reads <font size=N> off HTML's absolute size table", () => {
    // In px, because that table is defined in px. `size=2` is 13px.
    expect(detectedSize(`<html><body><p><font size="2">T</font></p></body></html>`)).toBe("13px");
  });

  it("prefers a <font>'s stated size to its size attribute", () => {
    expect(
      detectedSize(`<html><body><p><font size="2" style="font-size:10pt">T</font></p></body></html>`)
    ).toBe("10pt");
  });

  it("resolves CSS's absolute size keywords instead of dropping them", () => {
    expect(detectedSize(`<p><span style="font-size:medium">T</span></p>`)).toBe("16px");
    expect(detectedSize(`<p><span style="font-size:large">T</span></p>`)).toBe("18px");
    expect(detectedSize(`<p><span style="font-size:x-small">T</span></p>`)).toBe("10px");
  });

  it("resolves the relative keywords against the base, as it does em", () => {
    // One step of 1.2 off 13.5pt.
    expect(detectedSize(`<p><span style="font-size:smaller">T</span></p>`)).toBe("11.25pt");
    expect(detectedSize(`<p><span style="font-size:larger">T</span></p>`)).toBe("16.2pt");
  });

  it("still ignores what names no size", () => {
    expect(detectedSize(`<p><span style="font-size:inherit">T</span></p>`)).toBe("(none)");
  });
});

describe("the other applications a report arrives from", () => {
  it("LibreOffice Writer, which states it on a <font> with a size attribute beside it", () => {
    expect(
      detectedSize(
        `<html><head><meta name="generator" content="LibreOffice 7.6"></head><body>` +
          `<p style="line-height: 100%"><font face="Liberation Serif, serif">` +
          `<font size="2" style="font-size: 10pt">T</font></font></p></body></html>`
      )
    ).toBe("10pt");
  });

  it("a spreadsheet, whose cells hold bare text and no paragraph at all", () => {
    expect(
      detectedSize(
        `<html xmlns:x="urn:schemas-microsoft-com:office:excel"><body>` +
          `<table><tr><td style='font-size:10.0pt'>T</td></tr></table></body></html>`
      )
    ).toBe("10pt");
  });

  it("a spreadsheet stating it in a td rule rather than inline", () => {
    expect(
      detectedSize(
        `<html xmlns:x="urn:schemas-microsoft-com:office:excel"><head><style><!--td {font-size:10.0pt;}--></style>` +
          `</head><body><table><tr><td>T</td></tr></table></body></html>`
      )
    ).toBe("10pt");
  });

  it("Google Docs, which states it plainly on the run", () => {
    expect(
      detectedSize(`<html><body><p><span style="font-size:10pt;font-family:Arial">T</span></p></body></html>`)
    ).toBe("10pt");
  });

  it("this editor's own saved HTML, which must survive the round trip untouched", () => {
    expect(detectedSize(`<p style="font-size: 15px"><span style="font-size: 15px">T</span></p>`)).toBe(
      "15px"
    );
    expect(detectedSize(`<p style="font-size: 10pt"><span style="font-size: 10pt">T</span></p>`)).toBe(
      "10pt"
    );
  });
});

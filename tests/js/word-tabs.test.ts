/**
 * A tab copied out of a Word or LibreOffice document is still a tab here.
 *
 * Every tab the document had reaches the editor. Where nothing is laid out —
 * as here, in jsdom — the count is exactly the document's; in a browser it is
 * re-worked so the text lands on the document's own stops, which is measured in
 * tests/browser/office-paste-layout.mjs. See lib/word-tabs.ts and
 * lib/whitespace.ts; the on-screen and printed half is also measured in
 * tests/browser/word-tab-paste.mjs.
 */
import { describe, expect, it } from "vitest";
import { createPlateEditor } from "platejs/react";
import { buildPlugins } from "@/plugins";
import { plateValueToHtml } from "@/lib/html-serializer";
import { expandWordTabSpans } from "@/lib/word-tabs";

const nbsp = (count: number) => "&nbsp;".repeat(count);

const word = (body: string) => `<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:w="urn:schemas-microsoft-com:office:word"><head><style>
<!-- p.MsoNormal {margin:0cm; font-size:11.0pt; font-family:"Calibri",sans-serif;} --></style></head>
<body lang=EN-US>${body}</body></html>`;

/** What LibreOffice puts on the clipboard, measured: literal tabs. */
const libreOffice = (body: string) => `<html><head><meta name="generator" content="LibreOffice 24.2.7.2 (Linux)"/>
<style type="text/css">p { line-height: 115%; margin-bottom: 0.1in }</style></head><body lang="en-US">${body}</body></html>`;

/** Minimal DataTransfer — jsdom does not implement a usable one. */
function clipboard(html: string) {
  return {
    types: ["text/html", "text/plain"],
    getData: (type: string) => (type === "text/html" ? html : ""),
    files: [],
    items: [],
  } as unknown as DataTransfer;
}

/** Paste through `insertData`, the path a Ctrl+V takes, and read the lines back. */
function paste(html: string) {
  const editor = createPlateEditor({ plugins: buildPlugins("clean") });
  editor.tf.setValue([{ type: "p", children: [{ text: "" }] }] as never);
  editor.tf.select({ anchor: { path: [0, 0], offset: 0 }, focus: { path: [0, 0], offset: 0 } });
  editor.tf.insertData(clipboard(html));
  const texts = (editor.children as any[])
    .map((block) => (block.children ?? []).map((leaf: any) => leaf.text ?? "").join(""))
    .filter((text) => text.trim());
  return { texts, html: plateValueToHtml(editor.children as never) };
}

describe("pasting a document's tabs", () => {
  it("keeps Word's tabs, exactly as many as it typed", () => {
    const { texts } = paste(word(
      `<p class=MsoNormal>Test Name<span style='mso-tab-count:2'>${nbsp(13)} </span>: RT-PCR</p>` +
      `<p class=MsoNormal>Specimen<span style='mso-tab-count:2'>${nbsp(15)} </span>: Swab</p>`
    ));
    expect(texts).toEqual(["Test Name\t\t: RT-PCR", "Specimen\t\t: Swab"]);
  });

  it("keeps a ruler-stop line's tab count where nothing can be measured", () => {
    const { texts } = paste(word(
      `<p class=MsoNormal style='tab-stops:180.0pt'>Name<span style='mso-tab-count:1'>${nbsp(38)} </span>: Rahim</p>`
    ));
    expect(texts).toEqual(["Name\t: Rahim"]);
  });

  it("keeps a tab span Plate's own cleaner misses, instead of a run of &nbsp;", () => {
    const { texts, html } = paste(word(
      `<p class=MsoNormal><span style='font-size:11.0pt;mso-tab-count:1'>${nbsp(6)} </span>Indented: yes</p>`
    ));
    expect(texts).toEqual(["\tIndented: yes"]);
    expect(html).toContain('<span style="white-space: pre">\t</span>');
    expect(html).not.toContain("&nbsp;");
  });

  it("keeps LibreOffice's tabs, which it writes as they are", () => {
    const { texts } = paste(libreOffice(
      `<p class="western" style="margin-bottom: 0in">\nTest Name\t\t: RT-PCR</p>`
    ));
    expect(texts).toEqual(["Test Name\t\t: RT-PCR"]);
  });

  it("saves them as tabs, which is what the print page lines up on", () => {
    const { html } = paste(word(
      `<p class=MsoNormal>Result<span style='mso-tab-count:2'>${nbsp(14)} </span>: Negative</p>`
    ));
    expect(html).toContain('<span style="white-space: pre">\t\t</span>: Negative');
  });
});

/**
 * LibreOffice's clipboard for a USG report, as captured from LibreOffice itself
 * (trimmed): text/html AND text/rtf. The RTF is what makes Plate's docx cleaner
 * run, and that cleaner deleted every span holding only tabs — so the lines
 * pasted as `BPD89 mm`. Its HTML writer also soft-wraps long source lines by
 * turning a SPACE into a newline, which here lands right before a tab run.
 */
const LIBRE_OFFICE_USG = `<html><head><meta name="generator" content="LibreOffice 24.2.7.2 (Linux)"/></head>
<body lang="en-US" dir="ltr"><p class="western"><font face="Calibri, serif"><font size="3" style="font-size: 12pt"><span style="font-style: normal">Pregnancy\t\t\tIntrauterine
</span></font></font>
</p>
<h3 class="western"><font face="Calibri, serif"><font size="3" style="font-size: 12pt"><span style="font-style: normal">BPD\t\t\t\t</span></font></font><font color="#ff0000"><font face="Calibri, serif"><font size="3" style="font-size: 12pt"><span style="font-style: normal"><b>89</b></span></font></font></font><font face="Calibri, serif"><font size="3" style="font-size: 12pt"><span style="font-style: normal">
mm</span></font></font></h3>
<h3 class="western" style="margin-right: -0.5in"><font face="Calibri, serif"><font size="3" style="font-size: 12pt"><span style="font-style: normal">Number
\t\t\tSingle </span></font></font>
</h3>
<h3 class="western" style="margin-right: -0.5in"><font face="Calibri, serif"><font size="3" style="font-size: 12pt"><span style="font-style: normal"><b>Status
\t\t              Alive</b></span></font></font></h3>
<h3 class="western" style="margin-right: -0.5in"><font face="Calibri, serif"><font size="3" style="font-size: 12pt"><span style="font-style: normal">Cardiac
Activity \t              Present which is regular in rhythm</span></font></font></h3>
<p class="western"><font face="Calibri, serif"><font size="3" style="font-size: 12pt"><span style="font-style: normal">\t\t</span></font></font><font face="Calibri, serif"><font size="3" style="font-size: 12pt"><b>Bold after a tab-only run</b></font></font></p>
</body></html>`;

const LIBRE_OFFICE_USG_LINES = [
  "Pregnancy\t\t\tIntrauterine",
  "BPD\t\t\t\t89 mm",
  "Number \t\t\tSingle",
  "Status \t\t              Alive",
  "Cardiac Activity \t              Present which is regular in rhythm",
  "\t\tBold after a tab-only run",
];

/** Paste with an RTF flavour on the clipboard, the way LibreOffice copies. */
function pasteWithRtf(html: string) {
  const editor = createPlateEditor({ plugins: buildPlugins("clean") });
  editor.tf.setValue([{ type: "p", children: [{ text: "" }] }] as never);
  editor.tf.select({ anchor: { path: [0, 0], offset: 0 }, focus: { path: [0, 0], offset: 0 } });
  editor.tf.insertData({
    types: ["text/html", "text/rtf", "text/plain"],
    getData: (type: string) => (type === "text/html" ? html : type === "text/rtf" ? "{\\rtf1\\ansi USG}" : ""),
    files: [],
    items: [],
  } as unknown as DataTransfer);
  return (editor.children as any[])
    .map((block) => (block.children ?? []).map((leaf: any) => leaf.text ?? "").join("").trimEnd())
    .filter((text) => text.trim());
}

describe("pasting out of LibreOffice", () => {
  it("keeps every tab when the clipboard carries RTF, as LibreOffice's does", () => {
    expect(pasteWithRtf(LIBRE_OFFICE_USG)).toEqual(LIBRE_OFFICE_USG_LINES);
  });

  it("keeps them the same way when it does not", () => {
    expect(paste(LIBRE_OFFICE_USG).texts.map((text) => text.trimEnd())).toEqual(LIBRE_OFFICE_USG_LINES);
  });

  it("keeps a Word paste's tabs with RTF too, as Word always sends it", () => {
    expect(pasteWithRtf(word(
      `<p class=MsoNormal>Test Name<span style='mso-tab-count:2'>${nbsp(13)} </span>: RT-PCR</p>`
    ))).toEqual(["Test Name\t\t: RT-PCR"]);
  });

  it("still does not paste the source file's own indentation", () => {
    expect(pasteWithRtf(`<html><body><div>\n\t\t<p>One</p>\n\t\t<p>Two</p>\n</div></body></html>`))
      .toEqual(["One", "Two"]);
  });
});

describe("expandWordTabSpans", () => {
  it("turns the span into that many tabs and keeps the rest of its style", () => {
    const out = expandWordTabSpans(
      "<p><span style='font-size:12.0pt;mso-tab-count:2'>&nbsp;&nbsp; </span>Indented</p>"
    );
    expect(out).toContain('<span style="font-size:12.0pt">\t\t</span>Indented');
    expect(out).not.toContain("mso-tab-count");
  });

  it("reads the count through a leader", () => {
    const out = expandWordTabSpans("<p>A<span style='mso-tab-count:1 dotted'>. . . . </span>B</p>");
    expect(out).toContain("A<span>\t</span>B");
  });

  it("returns the very same string when there is no tab span", () => {
    const html = '<p>Result<span style="white-space: pre">\t\t</span>: Negative</p>';
    expect(expandWordTabSpans(html)).toBe(html);
  });
});

/**
 * A table copied out of LibreOffice — the case the first LibreOffice fix broke.
 *
 * Its HTML writer indents the markup inside a cell with TABS, and carries that
 * indentation onto every wrapped line: `Inv. Date:\n\t\t\t13-10-2025` is the
 * text `Inv. Date: 13-10-2025`. Read as the document's own tabs, every cell
 * gained a column of them. Cut from LibreOffice's real clipboard (UNO).
 */
const LIBRE_OFFICE_TABLE = `<html><head><meta name="generator" content="LibreOffice 24.2.7.2 (Linux)"/></head><body lang="en-US" dir="ltr">
<table width="675" cellpadding="7" cellspacing="0">
\t<col width="156"/>
\t<col width="189"/>
\t<tr valign="top">
\t\t<td width="156" style="border: 1px solid #000000; padding: 0in 0.08in"><p align="left" style="orphans: 2; widows: 2">
\t\t\t<font face="Calibri, serif"><font size="2" style="font-size: 11pt">Inv.&nbsp;ID&nbsp;:&nbsp;3222222</font></font></p>
\t\t</td>
\t\t<td width="189" style="border: 1px solid #000000; padding: 0in 0.08in"><p align="left" style="orphans: 2; widows: 2">
\t\t\t<font face="Calibri, serif"><font size="2" style="font-size: 11pt">Inv.&nbsp;Date:
\t\t\t13-10-2025&nbsp;</font></font></p>
\t\t</td>
\t</tr>
\t<tr valign="top">
\t\t<td width="156" style="border: 1px solid #000000; padding: 0in 0.08in"><p align="left">
\t\t\t<font face="Calibri, serif"><font size="2" style="font-size: 11pt">Specimen
\t\t\t          : Whole Blood</font></font></p>
\t\t</td>
\t\t<td width="189" style="border: 1px solid #000000; padding: 0in 0.08in"><p align="left">
\t\t\t<font face="Calibri, serif"><font size="2" style="font-size: 11pt">Probe\t: TaqMan</font></font></p>
\t\t</td>
\t</tr>
</table>
</body></html>`;

describe("pasting a table out of LibreOffice", () => {
  /** Each cell's text, as the editor holds it. */
  const cellsOf = (html: string) =>
    Array.from(new DOMParser().parseFromString(html, "text/html").querySelectorAll("td")).map((cell) =>
      (cell.textContent ?? "").replace(/ /g, " ").trim()
    );

  function pasteHtml(html: string, withRtf: boolean) {
    const editor = createPlateEditor({ plugins: buildPlugins("clean") });
    editor.tf.setValue([{ type: "p", children: [{ text: "" }] }] as never);
    editor.tf.select({ anchor: { path: [0, 0], offset: 0 }, focus: { path: [0, 0], offset: 0 } });
    editor.tf.insertData({
      types: ["text/html", "text/plain"],
      getData: (type: string) => (type === "text/html" ? html : type === "text/rtf" && withRtf ? "{\\rtf1\\ansi x}" : ""),
      files: [],
      items: [],
    } as unknown as DataTransfer);
    return plateValueToHtml(editor.children as never);
  }

  for (const withRtf of [true, false]) {
    it(`keeps every cell's text as the document has it${withRtf ? ", with RTF" : ""} — no indentation tabs`, () => {
      expect(cellsOf(pasteHtml(LIBRE_OFFICE_TABLE, withRtf))).toEqual([
        "Inv. ID : 3222222",
        "Inv. Date: 13-10-2025",
        "Specimen           : Whole Blood",
        "Probe\t: TaqMan",
      ]);
    });
  }
});

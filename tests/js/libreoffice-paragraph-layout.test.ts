/**
 * Lines pasted out of LibreOffice, outside a table: spaced and set the way the
 * document has them.
 *
 * A report's "Test platform" block pasted 0.73 of a line apart, in the editor's
 * own face, with 4px of padding above and below each line, where LibreOffice
 * draws the lines of the .docx 14.67px apart in Calibri — and "line spacing 1"
 * then visibly opened them up. Three causes, each pinned here:
 *  - the line gap: LibreOffice's RTF makes the docx cleaner remove the tag that
 *    names LibreOffice, so `line-height: 100%` was read as CSS's 1 x the font
 *    size (lib/word-line-gap.ts);
 *  - the spacing: the editor's `py-1` stood in for the document's own space above
 *    and below each line (`extractDocumentSpacing`, lib/table-widths.ts);
 *  - the font: LibreOffice's HTML names no font for a line in the document's
 *    default one, only its RTF does (lib/font-face.ts).
 *
 * jsdom measures no font, so every natural line here is the fallback 1.2; what
 * the browser draws is measured in tests/browser/libreoffice-paragraph-layout.mjs.
 * Cut from the LibreOffice clipboard (UNO) of the report the complaint was about.
 */
import { describe, expect, it } from "vitest";
import { createPlateEditor } from "platejs/react";
import { deserializeHtml, type Value } from "platejs";
import { buildPlugins } from "@/plugins";
import { plateValueToHtml } from "@/lib/html-serializer";
import { FALLBACK_NATURAL_LINE_HEIGHT, lineGapToCssRatio } from "@/lib/line-gap";

const LIBRE_OFFICE = `<!DOCTYPE html><html><head><meta http-equiv="content-type" content="text/html; charset=utf-8"/>
<meta name="generator" content="LibreOffice 24.2.7.2 (Linux)"/>
<style type="text/css">
p { line-height: 115%; text-align: left; orphans: 2; widows: 2; margin-bottom: 0.1in; direction: ltr; background: transparent }
td p { orphans: 0; widows: 0; margin-bottom: 0in; direction: ltr; background: transparent }
</style></head><body lang="en-US" dir="ltr">
<p style="line-height: 100%; margin-bottom: 0in"><font size="2" style="font-size: 9pt"><u>Test
platform:</u></font></p>
<p style="line-height: 100%; margin-bottom: 0in"><font size="2" style="font-size: 9pt">Probe\t\t:
TaqMan probe</font></p>
<p><font size="2" style="font-size: 9pt">Impression: normal study.</font></p>
<table width="400" cellpadding="4" cellspacing="0">
\t<tr valign="top">
\t\t<td width="100%" style="border: 1px solid #000000; padding: 0.04in"><p>
\t\t\t<font face="Tahoma, serif"><font size="2" style="font-size: 10pt">Result</font></font></p>
\t\t\t<p><br/>
\t\t\t</p>
\t\t</td>
\t</tr>
</table>
</body></html>`;

const RTF = String.raw`{\rtf1\ansi\deff4{\fonttbl{\f0\froman\fprq2\fcharset0 Times New Roman;}{\f4\froman\fprq2\fcharset0 Calibri;}{\f5\froman\fprq2\fcharset0 Tahoma;}}` +
  String.raw`{\stylesheet{\s0 Normal;}{\s16 Table Contents;}}` +
  String.raw`\pard\plain \s0\f4\sl276\slmult1\sa200\sl240\slmult1\sa0{\fs18\ul Test platform:}\par` +
  String.raw`\pard\plain \s0\f4\sl240\slmult1\sa0{\fs18 Probe\tab\tab : TaqMan probe}\par` +
  String.raw`\pard\plain \s0\f4\sl276\slmult1\sa200{\fs18 Impression: normal study.}\par` +
  String.raw`\trowd\cellx8000\pard\plain \s16\intbl\f4\sl240\slmult1{\f5\fs20 Result}\par` +
  String.raw`\pard\plain \s16\intbl\f4\sl240\slmult1\cell\row}`;

function paste(html: string, rtf: string) {
  const editor = createPlateEditor({ plugins: buildPlugins("clean") });
  editor.tf.setValue([{ type: "p", children: [{ text: "" }] }] as never);
  editor.tf.select({ anchor: { path: [0, 0], offset: 0 }, focus: { path: [0, 0], offset: 0 } });
  editor.tf.insertData({
    types: ["text/html", "text/rtf", "text/plain"],
    getData: (t: string) => (t === "text/html" ? html : t === "text/rtf" ? rtf : ""),
    files: [],
    items: [],
  } as unknown as DataTransfer);
  const value = editor.children as any[];
  return { value, html: plateValueToHtml(value as never) };
}

/** A saved report opened again: the load path, not the paste path. */
function reopen(html: string) {
  const editor = createPlateEditor({ plugins: buildPlugins("clean") });
  editor.tf.setValue(deserializeHtml(editor, { element: html }) as Value);
  editor.tf.normalize({ force: true });
  const value = editor.children as any[];
  return { value, html: plateValueToHtml(value as never) };
}

const textOf = (node: any): string => (node.children ?? []).map((c: any) => c.text ?? textOf(c)).join("");
const lineStarting = (value: any[], start: string) => value.find((node) => textOf(node).trim().startsWith(start));
const firstLeaf = (node: any): any => (node.children?.[0]?.text !== undefined ? node.children.find((c: any) => c.text?.trim()) ?? node.children[0] : firstLeaf(node.children[0]));
const cellLines = (value: any[]) => value.find((n) => n.type === "table").children[0].children[0].children;
const gap = (lines: number) => lineGapToCssRatio(lines, FALLBACK_NATURAL_LINE_HEIGHT);

describe("a LibreOffice line's gap", () => {
  const { value } = paste(LIBRE_OFFICE, RTF);

  it("is the document's single line, not CSS's 1 x the font size", () => {
    expect(lineStarting(value, "Test platform").lineHeight).toBe(gap(1));
    expect(lineStarting(value, "Probe").lineHeight).toBe(gap(1));
  });

  it("is the document's 115% where its style block says so", () => {
    expect(lineStarting(value, "Impression").lineHeight).toBe(gap(1.15));
  });

  it("gives an empty table line the RTF paragraph in its place", () => {
    // It has no text to be matched by; it sits between two matched lines.
    const [result, blank] = cellLines(value);
    expect(result.lineHeight).toBe(gap(1));
    expect(blank.lineHeight).toBe(gap(1));
  });
});

describe("a LibreOffice line's spacing", () => {
  const { value, html } = paste(LIBRE_OFFICE, RTF);

  it("is the document's own, in place of the editor's padding", () => {
    expect(lineStarting(value, "Probe")).toMatchObject({ documentSpacing: true, marginTop: "0px", marginBottom: "0px" });
    // The style block's space after, 0.1in.
    expect(lineStarting(value, "Impression")).toMatchObject({ documentSpacing: true, marginBottom: "9.6px" });
  });

  it("is saved with its mark, zeros and all", () => {
    expect(html).toMatch(/<p data-spacing="document" style="[^"]*margin-top: 0px; margin-bottom: 0px[^"]*"><span[^>]*>Probe/);
  });

  it("opens again exactly as it was saved", () => {
    const reopened = reopen(html);
    expect(lineStarting(reopened.value, "Probe")).toMatchObject({ documentSpacing: true, marginTop: "0px", marginBottom: "0px" });
    expect(reopen(reopened.html).html).toBe(reopened.html);
  });

  it("leaves a table's lines to the table", () => {
    for (const line of cellLines(value)) expect(line.documentSpacing).toBeUndefined();
  });
});

describe("a LibreOffice line's font", () => {
  const { value } = paste(LIBRE_OFFICE, RTF);

  it("is the one its RTF names when its HTML names none", () => {
    expect(firstLeaf(lineStarting(value, "Probe")).fontFamily).toBe("Calibri");
    expect(firstLeaf(lineStarting(value, "Test platform")).fontFamily).toBe("Calibri");
  });

  it("stays the run's own where the HTML names one", () => {
    expect(firstLeaf(cellLines(value)[0]).fontFamily).toContain("Tahoma");
  });
});

describe("what it leaves alone", () => {
  it("a Word paste: no mark, no font added, its own gap", () => {
    const word =
      `<html xmlns:w="urn:schemas-microsoft-com:office:word"><head><style><!--
p.MsoNormal {margin:0cm; font-size:11.0pt;}
--></style></head><body><p class=MsoNormal style='line-height:150%'><span style='font-size:11.0pt'>Findings</span></p></body></html>`;
    const { value, html } = paste(word, RTF.replace("Probe", "Findings"));
    const line = lineStarting(value, "Findings");
    expect(line.documentSpacing).toBeUndefined();
    expect(line.lineHeight).toBe(gap(1.5));
    expect(firstLeaf(line).fontFamily).toBeUndefined();
    expect(html).not.toContain("data-spacing");
  });

  it("a report saved before: stated margins do not make it the document's spacing", () => {
    // What Summernote saved from Word for years.
    const legacy = `<p class="MsoNormal" style="margin: 0cm 0cm 0.0001pt; line-height: normal">Old report</p>`;
    const { value, html } = reopen(legacy);
    expect(value[0].documentSpacing).toBeUndefined();
    expect(html).not.toContain("data-spacing");
  });

  it("a line typed in the editor", () => {
    const { html } = reopen("<p>Typed here</p>");
    expect(html).toBe("<p>Typed here</p>");
  });

  it("a LibreOffice paste with no RTF: no font from nowhere", () => {
    const { value } = paste(LIBRE_OFFICE, "");
    expect(firstLeaf(lineStarting(value, "Probe")).fontFamily).toBeUndefined();
  });
});

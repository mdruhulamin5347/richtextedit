/**
 * A line gap, the way Word and LibreOffice mean it.
 *
 * Both multiply the font's NATURAL line height; CSS multiplies the font SIZE.
 * On Calibri, whose natural line is 1.2207, a gap of 1.5 is the CSS ratio
 * 1.8310 — so a report typed at "1.5 lines" used to open here 17% tighter than
 * it was written, and the error grew with every line of a table.
 *
 * The arithmetic is tested here; that the browser measures the right FONT is
 * tested in tests/browser/line-gap.mjs, because jsdom does no layout and every
 * rect in it is zero.
 */
import { describe, expect, it } from "vitest";
import { createPlateEditor } from "platejs/react";
import { buildPlugins } from "@/plugins";
import { plateValueToHtml } from "@/lib/html-serializer";
import {
  cssRatioToLineGap,
  DEFAULT_PASTED_LINE_GAP,
  FALLBACK_NATURAL_LINE_HEIGHT,
  lineGapToCssRatio,
  naturalLineHeight,
} from "@/lib/line-gap";
import { inlineWordLineGap, isOfficeClipboard } from "@/lib/word-line-gap";

/** Calibri's true natural line, from its own metrics: 2500/2048 units per em. */
const CALIBRI = 1.2207;

describe("converting a line gap to a CSS ratio", () => {
  it("multiplies the font's natural line, which is what Word does", () => {
    expect(lineGapToCssRatio(1, CALIBRI)).toBeCloseTo(1.2207, 4);
    expect(lineGapToCssRatio(1.2, CALIBRI)).toBeCloseTo(1.4648, 4);
    expect(lineGapToCssRatio(1.5, CALIBRI)).toBeCloseTo(1.8311, 4);
    expect(lineGapToCssRatio(2, CALIBRI)).toBeCloseTo(2.4414, 4);
  });

  it("so a gap of 1 is NOT the CSS ratio 1 — that is the whole bug", () => {
    expect(lineGapToCssRatio(1, CALIBRI)).not.toBe(1);
  });

  it("names a stored ratio back as the gap it came from", () => {
    for (const gap of [1, 1.2, 1.5, 2, 3]) {
      expect(cssRatioToLineGap(lineGapToCssRatio(gap, CALIBRI), CALIBRI)).toBeCloseTo(gap, 3);
    }
  });

  it("falls back rather than dividing by a font that could not be measured", () => {
    expect(lineGapToCssRatio(1.5, 0)).toBeCloseTo(1.5 * FALLBACK_NATURAL_LINE_HEIGHT, 4);
    expect(cssRatioToLineGap(1.8, 0)).toBeCloseTo(1.8 / FALLBACK_NATURAL_LINE_HEIGHT, 4);
  });

  it("answers the fallback where nothing can be laid out", () => {
    // jsdom. A browser measures the real font; see tests/browser/line-gap.mjs.
    expect(naturalLineHeight('"Calibri", sans-serif')).toBe(FALLBACK_NATURAL_LINE_HEIGHT);
  });
});

describe("recognising an Office clipboard", () => {
  it("knows Word by its namespace, its properties and its classes", () => {
    expect(isOfficeClipboard(`<html xmlns:w="urn:schemas-microsoft-com:office:word">`)).toBe(true);
    expect(isOfficeClipboard(`<p style='mso-line-height-rule:exactly'>x</p>`)).toBe(true);
    expect(isOfficeClipboard(`<p class=MsoNormal>x</p>`)).toBe(true);
  });

  it("knows LibreOffice by its generator", () => {
    expect(isOfficeClipboard(`<meta name="generator" content="LibreOffice 7.5">`)).toBe(true);
  });

  it("does not mistake an ordinary web page, or this editor's own HTML", () => {
    expect(isOfficeClipboard(`<p style="line-height: 150%">x</p>`)).toBe(false);
    expect(isOfficeClipboard(`<p style="line-height: 1.8311; font-size: 15px">x</p>`)).toBe(false);
  });
});

describe("rewriting an Office clipboard", () => {
  const word = (style: string) =>
    `<html xmlns:w="urn:schemas-microsoft-com:office:word"><body>` +
    `<p class=MsoNormal style='${style}'><span style='font-size:11.0pt'>Findings</span></p>` +
    `</body></html>`;

  it("turns a proportional spacing into the ratio it draws as", () => {
    const out = inlineWordLineGap(word("line-height:150%"));
    expect(out).toContain(`line-height: ${lineGapToCssRatio(1.5, FALLBACK_NATURAL_LINE_HEIGHT)}`);
    expect(out).not.toContain("150%");
  });

  it("treats `normal` as no gap collected, and gives it the default", () => {
    // `normal` is a real declaration but it names no MULTIPLE — it says "ask
    // the font". A document that only says that is one whose gap was not
    // collected, and it used to open at a gap of 1. Reported 2026-09-16: "doc
    // file any line gap table copy and after paste show our editor still line
    // gap 1".
    const out = inlineWordLineGap(word("line-height:normal"));
    expect(out).toContain(
      `line-height: ${lineGapToCssRatio(DEFAULT_PASTED_LINE_GAP, FALLBACK_NATURAL_LINE_HEIGHT)}`
    );
    expect(out).not.toContain("normal");
  });

  it("gives a block that states nothing at all the same default", () => {
    const bare =
      `<html xmlns:w="urn:schemas-microsoft-com:office:word"><body>` +
      `<p class=MsoNormal><span style='font-size:11.0pt'>Findings</span></p></body></html>`;
    expect(inlineWordLineGap(bare)).toContain(
      `line-height: ${lineGapToCssRatio(DEFAULT_PASTED_LINE_GAP, FALLBACK_NATURAL_LINE_HEIGHT)}`
    );
  });

  it("leaves an absolute leading alone — there is no multiplier in it", () => {
    // What Word writes for "Exactly" spacing, beside mso-line-height-rule.
    expect(inlineWordLineGap(word("line-height:12.0pt"))).toContain("12.0pt");
  });

  it("leaves a BARE ratio alone, so this editor's own HTML is never doubled", () => {
    // A report copied out of this editor and pasted back carries no Office
    // marker, but a Word document that quotes one would — and a bare ratio is
    // this serializer's spelling, never Word's.
    const mine = `<html xmlns:w="urn:schemas-microsoft-com:office:word"><body>` +
      `<p style="line-height: 1.8311">Findings</p></body></html>`;
    expect(inlineWordLineGap(mine)).toContain("1.8311");
  });

  it("returns anything that is not an Office clipboard byte for byte", () => {
    const page = `<p style="line-height: 150%">Findings</p>`;
    expect(inlineWordLineGap(page)).toBe(page);
  });
});

describe("a gap the document states above the paragraph", () => {
  const DEFAULT = lineGapToCssRatio(DEFAULT_PASTED_LINE_GAP, FALLBACK_NATURAL_LINE_HEIGHT);
  const EXPECTED = lineGapToCssRatio(1.5, FALLBACK_NATURAL_LINE_HEIGHT);

  /** The `line-height` the CELL'S PARAGRAPH ends up stating, if any. */
  const gapOnCellBlock = (html: string) =>
    inlineWordLineGap(html)
      .match(/<p[^>]*style="[^"]*line-height:\s*([\d.]+)/)
      ?.[1];

  const table = (on: { table?: string; row?: string; cell?: string; block?: string }) =>
    `<html xmlns:w="urn:schemas-microsoft-com:office:word"><body>` +
    `<table border=1 style='${on.table ?? ""}'><tr style='${on.row ?? ""}'>` +
    `<td style='padding:0cm 5.4pt;${on.cell ?? ""}'>` +
    `<p class=MsoNormal style='${on.block ?? ""}'><span style='font-size:11.0pt'>Findings</span></p>` +
    `</td></tr></table></body></html>`;

  /*
   * `line-height` INHERITS, and unlike `text-align` a table does not break the
   * chain — so a gap stated on the cell, the row or the table IS the gap that
   * paragraph is drawn at. Only a BLOCK reaches this editor carrying one
   * though: a `<td>`'s style is read for its padding and borders and nothing
   * else. So these all arrived with no gap at all and were drawn at the cell's
   * own `normal` — a gap of 1, whatever the document said — and an ancestor
   * stating a gap also stopped the default from filling one in.
   */
  it("comes down onto the paragraph from the cell", () => {
    expect(gapOnCellBlock(table({ cell: "line-height:150%" }))).toBe(String(EXPECTED));
  });

  it("comes down from the row and from the table too", () => {
    expect(gapOnCellBlock(table({ row: "line-height:150%" }))).toBe(String(EXPECTED));
    expect(gapOnCellBlock(table({ table: "line-height:150%" }))).toBe(String(EXPECTED));
  });

  it("comes down through a wrapper the editor keeps no node for", () => {
    const wrapped =
      `<html xmlns:w="urn:schemas-microsoft-com:office:word"><body><table><tr><td>` +
      `<div style='line-height:150%'><p class=MsoNormal>Findings</p></div>` +
      `</td></tr></table></body></html>`;
    expect(gapOnCellBlock(wrapped)).toBe(String(EXPECTED));
  });

  it("but the paragraph's own always wins", () => {
    expect(gapOnCellBlock(table({ cell: "line-height:300%", block: "line-height:150%" }))).toBe(
      String(EXPECTED)
    );
  });

  it("and where nothing above states one either, the default still applies", () => {
    expect(gapOnCellBlock(table({}))).toBe(String(DEFAULT));
    expect(gapOnCellBlock(table({ cell: "line-height:normal" }))).toBe(String(DEFAULT));
  });

  it("gives a spreadsheet's loose cell text a paragraph to carry it", () => {
    // Excel and Calc write no block at all — the text sits in the `<td>` — so
    // a pasted spreadsheet had nowhere to put a gap and every row came out at
    // the cell's `normal` while the paragraphs around it took the document's.
    const excel =
      `<html xmlns:x="urn:schemas-microsoft-com:office:excel"><body>` +
      `<table><tr><td style='line-height:150%'>Findings</td></tr></table></body></html>`;
    expect(gapOnCellBlock(excel)).toBe(String(EXPECTED));

    const noGap =
      `<html xmlns:x="urn:schemas-microsoft-com:office:excel"><body>` +
      `<table><tr><td>Findings</td></tr></table></body></html>`;
    expect(gapOnCellBlock(noGap)).toBe(String(DEFAULT));
  });

  it("wraps nothing that is not a line of text", () => {
    const picture =
      `<html xmlns:w="urn:schemas-microsoft-com:office:word"><body>` +
      `<table><tr><td><img src="x.png"></td><td>   </td></tr></table></body></html>`;
    expect(inlineWordLineGap(picture)).not.toMatch(/<p[^>]*>\s*<img/);
  });
});

describe("through a real paste", () => {
  function paste(html: string) {
    const editor = createPlateEditor({ plugins: buildPlugins("clean") });
    editor.tf.setValue([{ type: "p", children: [{ text: "" }] }] as never);
    editor.tf.select({ anchor: { path: [0, 0], offset: 0 }, focus: { path: [0, 0], offset: 0 } });
    editor.tf.insertData({
      types: ["text/html", "text/plain"],
      getData: (t: string) => (t === "text/html" ? html : ""),
      files: [],
      items: [],
    } as unknown as DataTransfer);
    return { value: editor.children as any[], html: plateValueToHtml(editor.children as never) };
  }

  const cellDoc = (style: string) =>
    `<html xmlns:w="urn:schemas-microsoft-com:office:word"><head><style><!--
p.MsoNormal {margin:0cm; font-size:11.0pt; font-family:"Calibri",sans-serif;}
--></style></head><body>` +
    `<table class=MsoTableGrid border=1 cellspacing=0 style='border-collapse:collapse'>` +
    `<tr><td style='border:solid windowtext 1.0pt;padding:0cm 5.4pt'>` +
    `<p class=MsoNormal style='margin:0cm;${style}'><span style='font-size:11.0pt'>Findings</span></p>` +
    `</td></tr></table></body></html>`;

  const blockOf = (value: any[]) =>
    value.find((n) => n.type === "table").children[0].children[0].children[0];

  it("a 1.5-lines cell arrives at the ratio that draws 1.5 lines", () => {
    const { value, html } = paste(cellDoc("line-height:150%"));
    expect(blockOf(value).lineHeight).toBe(lineGapToCssRatio(1.5, FALLBACK_NATURAL_LINE_HEIGHT));
    expect(html).toContain(`line-height: ${lineGapToCssRatio(1.5, FALLBACK_NATURAL_LINE_HEIGHT)}`);
  });

  it("a gap stated on the CELL reaches the paragraph node inside it", () => {
    const onTheCell =
      `<html xmlns:w="urn:schemas-microsoft-com:office:word"><head><style><!--
p.MsoNormal {margin:0cm; font-size:11.0pt; font-family:"Calibri",sans-serif;}
--></style></head><body>` +
      `<table border=1 cellspacing=0 style='border-collapse:collapse'>` +
      `<tr><td style='border:solid windowtext 1.0pt;padding:0cm 5.4pt;line-height:150%'>` +
      `<p class=MsoNormal style='margin:0cm'><span style='font-size:11.0pt'>Findings</span></p>` +
      `</td></tr></table></body></html>`;
    expect(blockOf(paste(onTheCell).value).lineHeight).toBe(
      lineGapToCssRatio(1.5, FALLBACK_NATURAL_LINE_HEIGHT)
    );
  });

  it("and is stable across a save and reopen — never multiplied twice", () => {
    const once = paste(cellDoc("line-height:150%")).html;
    const twice = paste(once).html;
    const ratios = (h: string) => h.match(/line-height: [\d.]+/g);
    expect(ratios(twice)).toEqual(ratios(once));
  });
});

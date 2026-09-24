/**
 * The paragraphs of a clipboard's RTF — text, tab stops, indents and line spacing.
 *
 * LibreOffice's clipboard HTML leaves a paragraph's ruler stops out, and a table
 * cell's line spacing; its RTF states both. See lib/rtf-paragraphs.ts.
 */
import { describe, expect, it } from "vitest";
import { readRtfParagraphs } from "@/lib/rtf-paragraphs";

const rtf = (body: string) =>
  `{\\rtf1\\ansi\\deff0\\deftab720{\\fonttbl{\\f0\\fswiss Calibri;}}{\\colortbl;\\red255\\green0\\blue0;}` +
  `{\\stylesheet{\\s25\\tqc\\tx4680\\tqr\\tx9360 Footer;}}{\\info{\\author Someone}}${body}}`;

describe("reading RTF paragraphs", () => {
  it("reads each paragraph's text, with its tabs", () => {
    const paragraphs = readRtfParagraphs(rtf("\\pard\\plain BPD\\tab\\tab 89 mm\\par\\pard FL\\tab 66 mm\\par"));
    expect(paragraphs.map((p) => p.text)).toEqual(["BPD\t\t89 mm", "FL\t66 mm"]);
  });

  it("reads a paragraph's ruler stops in pt, with their alignment", () => {
    const [paragraph] = readRtfParagraphs(rtf("\\pard\\tx1548\\tqr\\tx9360\\tqdec\\tx5000 Status\\tab Alive\\par"));
    expect(paragraph.stops).toEqual([
      { positionPt: 77.4, alignment: "left" },
      { positionPt: 250, alignment: "decimal" },
      { positionPt: 468, alignment: "right" },
    ]);
  });

  it("skips a bar stop, which draws a rule rather than taking a tab", () => {
    const [paragraph] = readRtfParagraphs(rtf("\\pard\\tb2000\\tx3000 A\\tab B\\par"));
    expect(paragraph.stops).toEqual([{ positionPt: 150, alignment: "left" }]);
  });

  it("reads a hanging indent", () => {
    const [paragraph] = readRtfParagraphs(rtf("\\pard\\li2880\\fi-2880 Placenta\\tab Fundal\\par"));
    expect(paragraph.leftPt).toBe(144);
    expect(paragraph.firstPt).toBe(-144);
  });

  it("keeps a paragraph's properties until \\pard resets them", () => {
    const paragraphs = readRtfParagraphs(rtf("\\pard\\tx1548 One\\par Two\\par\\pard Three\\par"));
    expect(paragraphs.map((p) => p.stops.length)).toEqual([1, 1, 0]);
  });

  it("ignores the stylesheet, font table and info — their stops are not the text's", () => {
    const paragraphs = readRtfParagraphs(rtf("\\pard Body\\par"));
    expect(paragraphs).toHaveLength(1);
    expect(paragraphs[0].stops).toEqual([]);
    expect(paragraphs[0].text).toBe("Body");
  });

  it("reads hex, unicode and symbol escapes, and skips \\* destinations", () => {
    const [paragraph] = readRtfParagraphs(
      rtf("\\pard caf\\'e9 \\u8211? x{\\*\\bkmkstart one}\\~y\\endash z\\par")
    );
    expect(paragraph.text).toBe("café – x y–z");
  });

  it("keeps a last paragraph that has no \\par", () => {
    expect(readRtfParagraphs(rtf("\\pard One\\par\\pard Two")).map((p) => p.text)).toEqual(["One", "Two"]);
  });

  it("is empty for anything that is not RTF", () => {
    expect(readRtfParagraphs("")).toEqual([]);
    expect(readRtfParagraphs(null)).toEqual([]);
    expect(readRtfParagraphs("<p>html</p>")).toEqual([]);
  });
});

describe("reading a paragraph's line spacing", () => {
  const spacingOf = (body: string) => readRtfParagraphs(rtf(body)).map((p) => p.lineSpacing);

  it("reads a multiple of the natural line, which is what LibreOffice's Proportional writes", () => {
    expect(spacingOf("\\pard\\sl480\\slmult1 Double\\par\\pard\\sl360\\slmult1 Half\\par\\pard\\sl276\\slmult1 Wide\\par")).toEqual([
      { rule: "multiple", lines: 2 },
      { rule: "multiple", lines: 1.5 },
      { rule: "multiple", lines: 1.15 },
    ]);
  });

  it("is single where a paragraph states none, or states 0", () => {
    expect(spacingOf("\\pard None\\par\\pard\\sl0 Zero\\par")).toEqual([
      { rule: "multiple", lines: 1 },
      { rule: "multiple", lines: 1 },
    ]);
  });

  it("reads Exactly (negative) and At least (positive) in pt", () => {
    expect(spacingOf("\\pard\\sl-280\\slmult0 Fixed\\par\\pard\\sl360\\slmult0 Least\\par")).toEqual([
      { rule: "exact", pt: 14 },
      { rule: "atLeast", pt: 18 },
    ]);
  });

  it("keeps it until \\pard resets it", () => {
    expect(spacingOf("\\pard\\sl480\\slmult1 One\\par Two\\par\\pard Three\\par").map((s) => s.rule === "multiple" && s.lines)).toEqual([2, 2, 1]);
  });

  it("takes the LAST one a paragraph states — LibreOffice writes its style's, then the paragraph's own", () => {
    // Cut from a real report copied out of LibreOffice: the style says 1.15, the
    // paragraph itself says single.
    const [cell] = readRtfParagraphs(rtf("\\pard\\plain \\s16\\sl276\\slmult1\\intbl\\sl240\\slmult1 Age :\\cell\\row"));
    expect(cell.text).toBe("Age :");
    expect(cell.lineSpacing).toEqual({ rule: "multiple", lines: 1 });
  });

  it("does not take a style's spacing from the stylesheet", () => {
    const styled = `{\\rtf1\\ansi{\\stylesheet{\\s16\\sl480\\slmult1 Table Contents;}}\\pard\\plain\\s16 Body\\par}`;
    expect(readRtfParagraphs(styled)[0].lineSpacing).toEqual({ rule: "multiple", lines: 1 });
  });
});

describe("reading a paragraph's font", () => {
  const withFonts = (body: string) =>
    String.raw`{\rtf1\ansi\deff4{\fonttbl{\f0\froman\fprq2\fcharset0 Times New Roman;}` +
    String.raw`{\f3\froman\fprq2\fcharset0 Liberation Serif{\*\falt Times New Roman};}{\f4\froman\fprq2\fcharset0 Calibri;}` +
    String.raw`{\f13\fnil\fprq2\fcharset0 ;}}{\stylesheet{\s0\f0 Normal;}}${body}}`;

  it("names the font the paragraph's first character is set in", () => {
    // Cut from a real report copied out of LibreOffice: the paragraph sets \f4,
    // and the run inside it only changes the size.
    const [paragraph] = readRtfParagraphs(
      withFonts(String.raw`\pard\plain \s0\hich\af4\loch\f4\fs22{\rtlch\af15\afs18 \ltrch\loch\fs18 PCR kit\tab : TaqMan}\par`)
    );
    expect(paragraph.font).toBe("Calibri");
  });

  it("falls back to the document's default font after \\plain", () => {
    expect(readRtfParagraphs(withFonts(String.raw`\pard\plain Body\par`))[0].font).toBe("Calibri");
  });

  it("scopes a font to its group, as RTF does", () => {
    const [paragraph] = readRtfParagraphs(withFonts(String.raw`\pard\plain {\f0 }{\f3 Findings}\par`));
    expect(paragraph.font).toBe("Liberation Serif");
    const [after] = readRtfParagraphs(withFonts(String.raw`\pard\plain {\f0 x}\par\pard Next\par`)).slice(1);
    expect(after.font).toBe("Calibri");
  });

  it("names none for a font the table leaves unnamed, or a paragraph with no text", () => {
    const paragraphs = readRtfParagraphs(withFonts(String.raw`\pard\plain\f13 Nameless\par\pard\par`));
    expect(paragraphs[0].font).toBeUndefined();
    expect(paragraphs[1].font).toBeUndefined();
  });

  it("does not take the font table's or the stylesheet's fonts as text", () => {
    const paragraphs = readRtfParagraphs(withFonts(String.raw`\pard\plain Body\par`));
    expect(paragraphs.map((p) => p.text)).toEqual(["Body"]);
  });
});

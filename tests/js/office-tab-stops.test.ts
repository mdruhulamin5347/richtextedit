/**
 * A Word or LibreOffice paste, laid out on the document's own tab stops.
 *
 * Re-counting tabs needs a browser that lays text out, so that half is measured
 * in tests/browser/office-paste-layout.mjs against the document's own layout.
 * What is checked here is everything that does not need layout: reading stops,
 * where Word's next stop is, which clipboards are an office suite's, and the
 * two structural fixes — headings as paragraphs, hanging indents started where
 * the document starts them. See lib/office-tab-stops.ts.
 */
import { describe, expect, it } from "vitest";
import { inlineFontFaces } from "@/lib/font-face";
import {
  isOfficeClipboard,
  layOutOfficePaste,
  nextOfficeTabStop,
  OFFICE_DEFAULT_TAB_STOP_PT,
  parseTabStops,
} from "@/lib/office-tab-stops";

describe("reading a paragraph's tab stops", () => {
  it("reads left stops, alignments and units, and skips leaders", () => {
    expect(parseTabStops("108.0pt center 3.0in right dotted 450.0pt .5in")).toEqual([
      { positionPt: 36, alignment: "left" },
      { positionPt: 108, alignment: "left" },
      { positionPt: 216, alignment: "center" },
      { positionPt: 450, alignment: "right" },
    ]);
  });

  it("drops a bar stop and treats a list stop as a left one", () => {
    expect(parseTabStops("bar 100pt list 36pt")).toEqual([{ positionPt: 36, alignment: "left" }]);
  });
});

describe("where Word's next stop is", () => {
  it("is the next half inch by default, and never the point the text is at", () => {
    expect(OFFICE_DEFAULT_TAB_STOP_PT).toBe(36);
    expect(nextOfficeTabStop(0, []).positionPt).toBe(36);
    expect(nextOfficeTabStop(36, []).positionPt).toBe(72);
  });

  it("is a ruler stop when one lies ahead, clearing the default stops before it", () => {
    const stops = parseTabStops("77.4pt");
    expect(nextOfficeTabStop(30, stops).positionPt).toBe(77.4);
    // Past the ruler stop, default stops again: 77.4 → 108.
    expect(nextOfficeTabStop(77.4, stops).positionPt).toBe(108);
  });

  it("uses the document's own default stop when it states one", () => {
    expect(nextOfficeTabStop(0, [], 35.45).positionPt).toBeCloseTo(35.45);
  });
});

describe("which clipboards are an office suite's", () => {
  it("knows Word and LibreOffice", () => {
    expect(isOfficeClipboard('<html xmlns:w="urn:schemas-microsoft-com:office:word">')).toBe(true);
    expect(isOfficeClipboard("<p class=MsoNormal>x</p>")).toBe(true);
    expect(isOfficeClipboard('<meta name="generator" content="LibreOffice 24.2.7.2 (Linux)"/>')).toBe(true);
  });

  it("leaves a web page alone", () => {
    expect(isOfficeClipboard('<p>Result<span style="white-space:pre">\t</span>: Negative</p>')).toBe(false);
  });
});

describe("laying out an office paste (structure)", () => {
  const body = (html: string) => new DOMParser().parseFromString(html, "text/html").body;

  it("turns a document's heading into a paragraph, keeping its attributes", () => {
    const out = body(layOutOfficePaste('<h3 class="western" style="font-weight: bold">BPD\t89 mm</h3>'));
    expect(out.querySelector("h3")).toBeNull();
    const paragraph = out.querySelector("p")!;
    expect(paragraph.getAttribute("class")).toBe("western");
    expect(paragraph.style.fontWeight).toBe("bold");
    expect(paragraph.textContent).toBe("BPD\t89 mm");
  });

  it("starts a hanging-indent paragraph where the document starts its first line", () => {
    const out = body(layOutOfficePaste('<p style="margin-left: 2in; text-indent: -2in">Placenta\tFundal</p>'));
    const paragraph = out.querySelector("p")!;
    expect(paragraph.style.marginLeft).toBe("");
    expect(paragraph.style.textIndent).toBe("");
  });

  it("keeps a partial hang's first-line offset", () => {
    const out = body(layOutOfficePaste('<p style="margin-left: 72pt; text-indent: -36pt">Line</p>'));
    expect(out.querySelector("p")!.style.marginLeft).toBe("36pt");
  });

  it("leaves an ordinary indent alone", () => {
    const html = '<p style="margin-left: 36pt">Indented</p>';
    expect(layOutOfficePaste(html)).toBe(html);
  });

  it("does not re-count tabs where nothing is laid out", () => {
    const html = '<p style="white-space: pre-wrap">Status \t\t Alive</p>';
    expect(layOutOfficePaste(html, { rtf: "{\\rtf1\\pard\\tx1548 Status \\tab\\tab Alive\\par}" })).toBe(html);
  });
});

describe("a run's font, stated as <font face>", () => {
  it("is restated as CSS, so it survives the docx cleaner", () => {
    const out = inlineFontFaces('<h3><font face="Calibri, serif">BPD</font></h3>');
    const font = new DOMParser().parseFromString(out, "text/html").querySelector("font")!;
    expect(font.style.fontFamily).toBe("Calibri, serif");
  });

  it("keeps a family already stated in CSS", () => {
    const out = inlineFontFaces('<font face="Calibri" style="font-family: Arial">x</font>');
    expect(out).toBe('<font face="Calibri" style="font-family: Arial">x</font>');
  });

  it("returns the very same string when there is no <font face>", () => {
    const html = "<p>No font tags</p>";
    expect(inlineFontFaces(html)).toBe(html);
  });
});

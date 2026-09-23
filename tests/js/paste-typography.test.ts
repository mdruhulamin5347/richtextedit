/**
 * Font size and line spacing through a Word paste.
 *
 * Word states type in POINTS and spacing as a PERCENTAGE — `font-size:11.0pt`,
 * `line-height:115%`. Both used to be thrown away, so a pasted report arrived
 * at one flat 18px with one flat spacing: a 16pt heading, 11pt body and 9pt
 * footnote all came out identical and the document's hierarchy was gone.
 *
 * Points are then KEPT as points, which is the second half of it and was
 * reported separately: converting them to whole px named a 10pt paragraph 13 in
 * the font-size control and drew it a quarter of a pixel small. See
 * lib/font-size.ts.
 */
import { describe, expect, it } from "vitest";
import { createPlateEditor } from "platejs/react";
import { buildPlugins } from "@/plugins";
import { plateValueToHtml } from "@/lib/html-serializer";
import { toLineHeightRatio, toPtFontSize } from "@/components/paste-normalization-plugin";
import {
  cssRatioToLineGap,
  FALLBACK_NATURAL_LINE_HEIGHT as NATURAL,
  lineGapToCssRatio,
} from "@/lib/line-gap";
import { fontSizeLabel } from "@/lib/font-size";
import { toUnitLess } from "@platejs/basic-styles";

/** A molecular-biology report as Word puts it on the clipboard. */
const WORD_REPORT = `<html xmlns:w="urn:schemas-microsoft-com:office:word"><head><style>
<!-- p.MsoNormal {margin:0cm; font-size:11.0pt; font-family:"Calibri",sans-serif;} --></style></head><body lang=EN-GB>
<p class=MsoNormal style='margin-bottom:6.0pt;line-height:115%'><b><span style='font-size:16.0pt;font-family:"Times New Roman",serif'>TEST RESULT</span></b></p>
<p class=MsoNormal><span style='font-size:11.0pt'>Test Name : Multiplex PCR for Detection of STDs Pathogens</span></p>
<p class=MsoNormal style='line-height:150%'><span style='font-size:9.0pt'>Limit of Detection: 200 CFU/mL</span></p>
</body></html>`;

/**
 * Must return "" for anything but text/html — Slate probes
 * application/x-slate-fragment and runs atob() on whatever comes back.
 */
function clipboard(html: string) {
  return {
    types: ["text/html", "text/plain"],
    getData: (type: string) => (type === "text/html" ? html : ""),
    files: [],
    items: [],
  } as unknown as DataTransfer;
}

function paste(html: string, pasteMode: "clean" | "faithful" = "clean") {
  const editor = createPlateEditor({ plugins: buildPlugins(pasteMode) });
  editor.tf.setValue([{ type: "p", children: [{ text: "" }] }] as never);
  editor.tf.select({ anchor: { path: [0, 0], offset: 0 }, focus: { path: [0, 0], offset: 0 } });
  editor.tf.insertData(clipboard(html));
  return { value: editor.children as any[], html: plateValueToHtml(editor.children as never) };
}

describe("pasting a Word report", () => {
  const { value, html } = paste(WORD_REPORT);
  const sizeOf = (i: number) => value[i]?.children?.[0]?.fontSize;

  it("keeps the three sizes the document states, in the points it states them in", () => {
    // Exactly what Word wrote, trimmed of its trailing zero — so the size the
    // author set is the size the control names and the browser resolves it as
    // Word does: 72pt to the inch, 96 CSS px to the same inch.
    expect(sizeOf(0)).toBe("16pt");
    expect(sizeOf(1)).toBe("11pt");
    expect(sizeOf(2)).toBe("9pt");
  });

  it("so the heading is still bigger than the body, and the footnote smaller", () => {
    const pt = (i: number) => Number.parseFloat(sizeOf(i));
    expect(pt(0)).toBeGreaterThan(pt(1));
    expect(pt(2)).toBeLessThan(pt(1));
  });

  it("carries the sizes into the saved HTML", () => {
    expect(html).toContain("font-size: 16pt");
    expect(html).toContain("font-size: 11pt");
    expect(html).toContain("font-size: 9pt");
  });

  it("keeps the line spacing the document states, as WORD draws it", () => {
    // `line-height:115%` out of Word is "Multiple 1.15", and Word multiplies
    // the font's NATURAL line, not its size — so it arrives as the CSS ratio
    // that draws the same line here. jsdom lays nothing out, so the natural
    // line is the documented fallback; a browser measures the real font.
    // See lib/line-gap.ts and tests/browser/line-gap.mjs.
    expect(value[0].lineHeight).toBe(lineGapToCssRatio(1.15, NATURAL));
    expect(value[2].lineHeight).toBe(lineGapToCssRatio(1.5, NATURAL));

    // And the gaps a reader would name are still 1.15 and 1.5.
    expect(cssRatioToLineGap(value[0].lineHeight, NATURAL)).toBeCloseTo(1.15, 4);
    expect(cssRatioToLineGap(value[2].lineHeight, NATURAL)).toBeCloseTo(1.5, 4);
  });

  it("still drops the spacing the saved report cannot carry", () => {
    // The serializer writes no margin, so keeping Word's would style the editor
    // with a gap the print never gets.
    expect(value[0].marginBottom).toBeUndefined();
    expect(value[0].margin).toBeUndefined();
  });

  it("survives a save and reopen unchanged", () => {
    const reopened = paste(html).value;
    expect(reopened[0]?.children?.[0]?.fontSize).toBe("16pt");
    // Reopened from THIS editor's own HTML, which states a bare CSS ratio and
    // carries none of Word's markers — so it is read as CSS and not multiplied
    // a second time. That is what keeps a report stable across saves.
    expect(reopened[0].lineHeight).toBe(lineGapToCssRatio(1.15, NATURAL));
  });
});

/**
 * LibreOffice Writer, which is the other application a report is written in and
 * which states a size differently: on a `<font>` element, with the pre-CSS
 * `size` attribute beside it. The style wins, so the points still reach the run
 * — `lib/inherited-font-size.ts` leaves a carrier that states its own size
 * alone, and `<font>` is a carrier.
 */
const LIBREOFFICE_REPORT =
  `<html><head><meta name="generator" content="LibreOffice 7.6"></head><body>` +
  `<p style="margin-bottom: 0in; line-height: 100%">` +
  `<font face="Liberation Serif, serif"><font size="2" style="font-size: 10pt">` +
  `Impression: no acute abnormality.</font></font></p></body></html>`;

describe("pasting from LibreOffice", () => {
  const { value, html } = paste(LIBREOFFICE_REPORT);

  it("keeps the 10pt it states, so the control names it 10 as Writer does", () => {
    expect(value[0]?.children?.[0]?.fontSize).toBe("10pt");
    expect(fontSizeLabel(value[0]?.children?.[0]?.fontSize)).toBe("10");
  });

  it("carries it into the saved HTML, and onto the block so the line is as tall", () => {
    expect(html).toContain('<p style="line-height: 1.2; font-size: 10pt">');
    expect(html).toContain("font-size: 10pt");
  });

  it("prefers the stated points to the <font size> beside them", () => {
    // HTML's `size="2"` is 13px off the absolute size table — a different size
    // and a different scale. A carrier that states its own size keeps it.
    expect(html).not.toContain("13px");
  });
});

describe("font size conversion", () => {
  it("keeps the points Word writes, so 10 in the document is 10 here", () => {
    // The whole of the complaint: 10pt used to become `13px` — the control then
    // said 13, and 13px is not 10pt either (10pt is 13.333px).
    expect(toPtFontSize("10.0pt")).toBe("10pt");
    expect(toPtFontSize("11pt")).toBe("11pt");
    expect(toPtFontSize("16.0pt")).toBe("16pt");
    expect(toPtFontSize("11.5pt")).toBe("11.5pt");
  });

  it("leaves a legacy report's pixels in pixels", () => {
    // What is already in the database, from the years this editor worked in px.
    // Rewriting these to points would change every stored body the first time
    // it was opened and saved.
    expect(toPtFontSize("15px")).toBe("15px");
    expect(toPtFontSize("15")).toBe("15px");
  });

  it("resolves every other unit onto the points scale", () => {
    expect(toPtFontSize("1in")).toBe("72pt");
    expect(toPtFontSize("1pc")).toBe("12pt");
    expect(toPtFontSize("1cm")).toBe("28.35pt");
  });

  it("resolves a relative size against the editor's own text size", () => {
    // `text-[18px]` on the report variant, which is 13.5pt exactly.
    expect(toPtFontSize("1em")).toBe("13.5pt");
    expect(toPtFontSize("150%")).toBe("20.25pt");
  });

  it("resolves CSS's size keywords rather than dropping them", () => {
    // `medium` and friends are a defined table, kept in the px it is defined
    // in; `smaller` is a step of 1.2 off the base, like `em`. Dropping either
    // left the run at the editor's base, which is a size nobody chose. See
    // tests/js/font-size-detection.test.ts for the whole matrix.
    expect(toPtFontSize("medium")).toBe("16px");
    expect(toPtFontSize("smaller")).toBe("11.25pt");
  });

  it("rejects what names no size at all", () => {
    expect(toPtFontSize("inherit")).toBeNull();
    expect(toPtFontSize("initial")).toBeNull();
    expect(toPtFontSize("")).toBeNull();
  });

  it("rejects a size no report would use", () => {
    expect(toPtFontSize("0.5pt")).toBeNull();
    expect(toPtFontSize("900px")).toBeNull();
  });
});

describe("naming a size on the control's scale", () => {
  it("names a pasted size the number the document set it at", () => {
    expect(fontSizeLabel("10pt")).toBe("10");
    expect(fontSizeLabel("11.5pt")).toBe("11.5");
  });

  it("names a legacy px size in points too, so one report shows one scale", () => {
    expect(fontSizeLabel("15px")).toBe("11.25");
    expect(fontSizeLabel("12px")).toBe("9");
  });

  it("names the editor's own base, which no whole number can", () => {
    // Plate's `toUnitLess` matches the first run of digits, so it reads 13.5pt
    // back as "13" — a control using it cannot show the base size at all.
    expect(fontSizeLabel(undefined)).toBe("13.5");
    expect(toUnitLess("13.5pt")).toBe("13");
  });
});

describe("line height conversion", () => {
  it("reads both notations", () => {
    expect(toLineHeightRatio("115%")).toBeCloseTo(1.15, 5);
    expect(toLineHeightRatio("1.15")).toBeCloseTo(1.15, 5);
  });

  it("rejects a value that is not a ratio", () => {
    // An absolute leading has no font size here to be a ratio of.
    expect(toLineHeightRatio("normal")).toBeNull();
    expect(toLineHeightRatio("18pt")).toBeNull();
  });

  it("is no longer snapped onto the control's five values", () => {
    // A gap of 1.5 on Calibri IS the ratio 1.8306, and the nearest offered
    // value to that is 2 — snapping would have added a third again to every
    // exactly spaced paragraph, on the way in and on every reopen.
    expect(toLineHeightRatio("115%")).toBeCloseTo(1.15, 5);
    expect(toLineHeightRatio("183.06%")).toBeCloseTo(1.8306, 5);
  });
});

/**
 * The whitespace a report lays itself out with.
 *
 * A plain-text lab report is columns of tabs, an indented heading and blank
 * spacer lines. HTML collapses all three, so the report read correctly in the
 * editor — which is `white-space: pre-wrap` — and came out flat everywhere it
 * is actually printed. The way back in was worse: Plate's deserializer collapses
 * whitespace with JavaScript's `\s`, which counts U+00A0, so every legacy
 * Summernote report indented with `&nbsp;` lost its alignment on open.
 *
 * See lib/whitespace.ts; the on-screen half is measured in
 * tests/browser/plain-text-paste.mjs.
 */
import { describe, expect, it } from "vitest";
import { createPlateEditor } from "platejs/react";
import { deserializeHtml, type Value } from "platejs";
import { buildPlugins } from "@/plugins";
import { plateValueToHtml } from "@/lib/html-serializer";
import {
  dropTableMarkupWhitespace,
  NBSP,
  protectWhitespace,
  serializeWhitespace,
} from "@/lib/whitespace";

/** Text as it reaches the editor from `html`, block by block. */
function deserializeText(html: string): string[] {
  const editor = createPlateEditor({ plugins: buildPlugins("clean") });
  const value = deserializeHtml(editor, { element: protectWhitespace(html) }) as Value;
  return value.map((block: any) =>
    (block.children ?? []).map((child: any) => child.text ?? "").join("")
  );
}

function serialize(blocks: any[]): string {
  return plateValueToHtml(blocks as Value);
}

describe("serializeWhitespace", () => {
  it("keeps a space that HTML would drop at the start of a block", () => {
    expect(serializeWhitespace("   RESULT")).toBe("&nbsp; &nbsp;RESULT");
  });

  it("leaves the single spaces between words alone, so a line can still wrap", () => {
    expect(serializeWhitespace("Please correlate clinically.")).toBe(
      "Please correlate clinically."
    );
  });

  it("keeps the second of two spaces", () => {
    expect(serializeWhitespace("Result: Negative", { atBlockEnd: false })).toBe(
      "Result: Negative"
    );
    expect(serializeWhitespace("Result:  Negative")).toBe("Result: &nbsp;Negative");
  });

  it("keeps a trailing space, which HTML drops outright", () => {
    expect(serializeWhitespace("Time : ")).toBe("Time :&nbsp;");
  });

  it("leaves a leaf that starts mid-line alone", () => {
    // `<strong>Result</strong> : Negative` — the space follows text, so it is
    // safe, and making it non-breaking would take away a wrapping point.
    expect(serializeWhitespace(" : Negative", { atBlockStart: false })).toBe(" : Negative");
  });

  it("preserves a tab run in one span, because a tab is a distance to a stop", () => {
    expect(serializeWhitespace("Test Name\t\t: X")).toBe(
      'Test Name<span style="white-space: pre">\t\t</span>: X'
    );
  });

  it("is idempotent — a report does not drift on every save", () => {
    const once = serializeWhitespace(`${NBSP} ${NBSP}RESULT`);
    expect(once).toBe("&nbsp; &nbsp;RESULT");
  });
});

describe("protectWhitespace", () => {
  it("keeps the indentation a legacy Summernote report was written with", () => {
    expect(deserializeText("<p>&nbsp;&nbsp;&nbsp;&nbsp;RESULT</p>")).toEqual([
      `${NBSP}${NBSP}${NBSP}${NBSP}RESULT`,
    ]);
  });

  it("keeps a blank spacer line from becoming an empty block", () => {
    expect(deserializeText("<p>&nbsp;</p>")).toEqual([NBSP]);
  });

  it("keeps the tab run a plain-text paste was saved with", () => {
    expect(deserializeText("<p>Result\t\t: Negative</p>")).toEqual(["Result\t\t: Negative"]);
  });

  it("keeps Word's tabs, which it writes as a span full of &nbsp;", () => {
    const word =
      "<p class=MsoNormal>Test Name<span style='mso-tab-count:1'>&nbsp;&nbsp;&nbsp;&nbsp; </span>: RT-PCR</p>";
    expect(deserializeText(word)[0]).toBe(`Test Name${NBSP}${NBSP}${NBSP}${NBSP} : RT-PCR`);
  });

  it("does not paste the source file's own indentation into the report", () => {
    // Word's HTML is one long pretty-printed tree. Those tabs are how the file
    // is written, not what the report says, and they collapse to nothing.
    expect(deserializeText("<div>\n\t\t<p>One</p>\n\t\t<p>Two</p>\n</div>")).toEqual([
      "One",
      "Two",
    ]);
  });

  it("leaves a run of ordinary spaces to collapse, as a browser showed the author", () => {
    expect(deserializeText("<p>One   Two</p>")).toEqual(["One Two"]);
  });
});

describe("a report that lays itself out with whitespace", () => {
  /** The demo report, as the plain-text paste path builds it. */
  const REPORT: any[] = [
    { type: "p", children: [{ text: "Test Name\t\t:  RT-PCR FOR COVID-19" }] },
    { type: "p", children: [{ text: "Collection Site\t\t: Chattogram Maa Shishu" }] },
    { type: "p", children: [{ text: " " }] },
    { type: "p", children: [{ text: "        RESULT" }] },
    { type: "p", children: [{ text: "Result\t\t: Negative" }] },
  ];

  it("survives being saved", () => {
    const html = serialize(REPORT);
    expect(html).toContain('<span style="white-space: pre">\t\t</span>');
    // The spacer line is a blank line, not an empty block that prints as nothing.
    expect(html).toContain("<p>&nbsp;</p>");
    expect(html).toContain("&nbsp; &nbsp; &nbsp; &nbsp; RESULT");
  });

  it("survives being opened again, and saves back the same bytes", () => {
    const html = serialize(REPORT);
    const editor = createPlateEditor({ plugins: buildPlugins("clean") });
    const reopened = deserializeHtml(editor, { element: protectWhitespace(html) }) as Value;

    // Every gap comes back the width it was written. Half of a run reads back
    // as U+00A0 rather than a space — that is how the run was saved, and how it
    // has to be saved for HTML to keep it — so the comparison is on what the
    // line LOOKS like, which is the thing that has to survive.
    const asWritten = (text: string) => text.replaceAll(NBSP, " ");
    expect(
      reopened.map((block: any) =>
        asWritten(block.children.map((c: any) => c.text ?? "").join(""))
      )
    ).toEqual(REPORT.map((block) => block.children[0].text));

    // And it settles there: the second save is byte for byte the first, so a
    // report reopened a hundred times is still the report that was pasted.
    expect(serialize(reopened as any[])).toBe(html);
  });
});

describe("markup whitespace inside a table", () => {
  /**
   * A table written on ONE line — Word does that to tables it has re-flowed.
   * The single spaces between `</td>` and `<td>` are markup, not the report,
   * but they are text nodes all the same, and Plate turns them into paragraphs
   * INSIDE the table and inside the rows. That shifts every index the table is
   * read by: the first row stops being `children[0]` and the first cell stops
   * being column 0, so the table's top and left edges are never drawn and the
   * column count comes out short. A patient header table pasted in pieces.
   */
  const ONE_LINE =
    `<table border="1"> <tr> <td><p>Name</p></td> <td><p>Mr X</p></td> </tr> </table>`;

  it("is removed, so the rows hold nothing but cells", () => {
    const out = dropTableMarkupWhitespace(ONE_LINE);
    expect(out).toContain("<tr><td>");
    expect(out).toContain("</td><td>");
    expect(out).toContain("</td></tr>");
  });

  it("keeps the text inside the cells exactly as it was", () => {
    expect(dropTableMarkupWhitespace(ONE_LINE)).toContain("<p>Mr X</p>");
  });

  it("does not touch the spacing a CELL lays out with", () => {
    // Nothing stray here, so the input comes back as the very same string —
    // entities and all, un-parsed. That is the point: a document that does not
    // need the repair must not be re-serialized on its way through.
    const spaced = `<table><tr><td><p>Result&nbsp;&nbsp;&nbsp;: Negative</p></td></tr></table>`;
    expect(dropTableMarkupWhitespace(spaced)).toBe(spaced);
  });

  it("keeps a cell's own layout spacing when it DOES repair the table", () => {
    const spaced = `<table> <tr> <td><p>Result&nbsp;&nbsp;&nbsp;: Negative</p></td> </tr> </table>`;
    const out = dropTableMarkupWhitespace(spaced);
    expect(out).toContain("<tr><td>");
    // Written back as the entity, which is what `protectWhitespace` — the very
    // next thing to see this HTML — looks for.
    expect(out).toContain("Result&nbsp;&nbsp;&nbsp;: Negative");
  });

  it("leaves HTML with no table of its own alone, byte for byte", () => {
    const plain = "<p>Pathogen Name</p>";
    expect(dropTableMarkupWhitespace(plain)).toBe(plain);
  });

  it("and a table that has nothing stray in it, byte for byte", () => {
    const clean = `<table border="1"><tr><td><p>Name</p></td></tr></table>`;
    expect(dropTableMarkupWhitespace(clean)).toBe(clean);
  });
});

describe("a run of spaces split across leaves", () => {
  it("keeps every space on the page, not just the first", () => {
    // `Fundal   Ant`: three spaces, the middle one bold. Written as three plain
    // spaces in a row they collapsed to one wherever the report was read.
    const html = serialize([
      { type: "p", children: [{ text: "Fundal " }, { text: " ", bold: true }, { text: " Ant" }] },
    ]);
    expect(html).toBe("<p>Fundal <strong>&nbsp;</strong> Ant</p>");
  });

  it("still leaves a single space between leaves alone", () => {
    const html = serialize([{ type: "p", children: [{ text: "Result", bold: true }, { text: " : Negative" }] }]);
    expect(html).toBe("<p><strong>Result</strong> : Negative</p>");
  });
});

describe("a LibreOffice paste's whitespace", () => {
  /** What LibreOffice's HTML writer produces, soft-wrap newlines and all. */
  const lo = (body: string) =>
    `<html><head><meta name="generator" content="LibreOffice 24.2.7.2 (Linux)"/></head><body>${body}</body></html>`;
  const textsOf = (html: string) =>
    Array.from(new DOMParser().parseFromString(html, "text/html").body.querySelectorAll("p, h3")).map(
      (block) => block.textContent
    );

  it("turns a soft-wrap newline back into the space it replaced", () => {
    const out = protectWhitespace(lo("<p><span>Cardiac\nActivity \tPresent</span></p>"), {
      hasRtf: false,
      libreOffice: true,
    });
    expect(textsOf(out)).toEqual(["Cardiac Activity \tPresent"]);
  });

  it("drops the newline it writes after a block's opening tag and before its closing one", () => {
    const out = protectWhitespace(lo("<p>\n<span>USG Findings :</span>\n</p>"), { hasRtf: false, libreOffice: true });
    expect(textsOf(out)).toEqual(["USG Findings :"]);
  });

  it("keeps a line's leading spaces, inside the run they indent", () => {
    const out = protectWhitespace(lo(`<h3>${" ".repeat(5)}\n${" ".repeat(47)}<font face="Calibri">No retro</font></h3>`), {
      hasRtf: false,
      libreOffice: true,
    });
    const font = new DOMParser().parseFromString(out, "text/html").querySelector("font")!;
    expect(font.textContent).toBe(`${" ".repeat(53)}No retro`);
  });

  it("keeps a run of two or more spaces between words", () => {
    const out = protectWhitespace(lo("<p><span>No  retro-placental    collection</span></p>"), {
      hasRtf: false,
      libreOffice: true,
    });
    expect(textsOf(out)).toEqual(["No  retro-placental    collection"]);
    expect(out).toContain('<span style="white-space: pre">    </span>');
  });

  it("writes a space that is an element's only content so the docx cleaner keeps it", () => {
    const out = protectWhitespace(lo("<p><span>Fundal </span><b>\n</b><b>\nAnt</b></p>"), {
      hasRtf: true,
      libreOffice: true,
    });
    const bold = new DOMParser().parseFromString(out, "text/html").querySelector("b")!;
    expect(bold.textContent).toBe(NBSP);
  });

  it("leaves a non-LibreOffice paste's plain space runs to collapse, as before", () => {
    expect(protectWhitespace("<p>One   Two</p>", { hasRtf: false })).toBe("<p>One   Two</p>");
  });
});

/**
 * HTML -> Plate -> HTML round-trip.
 *
 * This is the test that decides whether the editor is safe to point at existing
 * report_body data: legacy Summernote content is deserialized through Plate's
 * own pipeline (the one that also handles a Word paste) and serialized straight
 * back out. Anything the pair loses would be lost from a real report the first
 * time somebody opened and saved it.
 */
import { describe, expect, it } from "vitest";
import { createPlateEditor } from "platejs/react";
import { deserializeHtml, type Value } from "platejs";
import { buildPlugins } from "@/plugins";
import { plateValueToHtml } from "@/lib/html-serializer";

/**
 * The real path: deserialize, put it INTO an editor (which is what mounting and
 * pasting both do, and what runs normalization), then serialize back out.
 */
function roundTrip(html: string, pasteMode: "clean" | "faithful" = "clean") {
  const editor = createPlateEditor({ plugins: buildPlugins(pasteMode) });
  const value = deserializeHtml(editor, { element: html }) as Value;
  editor.tf.setValue(value);
  editor.tf.normalize({ force: true });
  const out = editor.children as Value;
  return { html: plateValueToHtml(out), value: out };
}

describe("legacy Summernote HTML", () => {
  it("keeps inline marks", () => {
    const { html } = roundTrip("<p><strong>Bold</strong> and <em>italic</em> and <u>under</u></p>");
    expect(html).toContain("<strong>Bold</strong>");
    expect(html).toContain("<em>italic</em>");
    expect(html).toContain("<u>under</u>");
  });

  it("keeps colour and font-family", () => {
    const { html } = roundTrip(
      `<p><span style="color: #ff0000; font-family: 'Times New Roman', serif;">Positive</span></p>`
    );
    // The browser's CSSOM normalises #ff0000 to rgb(255, 0, 0) while parsing the
    // style attribute. Equivalent, so either form is accepted.
    expect(html).toMatch(/color: (#ff0000|rgb\(255, 0, 0\))/);
    expect(html).toMatch(/font-family: '?Times New Roman/);
  });

  it("never emits a double quote inside a style attribute", () => {
    // The fontFamily quoting bug: an inner " terminates style="..." early.
    const { html } = roundTrip(`<p><span style='font-family: "Times New Roman", serif;'>X</span></p>`);
    const styles = [...html.matchAll(/style="([^"]*)"/g)].map((m) => m[1]);
    expect(styles.length).toBeGreaterThan(0);
    for (const s of styles) expect(s).not.toContain('"');
  });

  it("keeps alignment", () => {
    const { html } = roundTrip(`<p style="text-align: center;">Centred</p>`);
    expect(html).toContain("text-align: center");
  });

  it("keeps headings and blockquotes", () => {
    const { html } = roundTrip("<h2>Findings</h2><blockquote>Note</blockquote>");
    expect(html).toContain("<h2");
    expect(html).toContain("<blockquote");
  });
});

describe("tables — the part the old serializer damaged", () => {
  const table = `<table border="1" style="border-collapse: collapse; width: 100%;">
    <tr><th>Test</th><th>Result</th></tr>
    <tr><td colspan="2" style="border: 1px solid #000; padding: 2px 5px;">Haemoglobin</td></tr>
  </table>`;

  it("survives with rows, cells and colspan", () => {
    const { html } = roundTrip(table);
    expect(html).toContain("<table");
    expect(html).toContain("<tr>");
    expect(html).toMatch(/<t[dh][^>]*>/);
    expect(html).toContain('colspan="2"');
  });

  it("emits black borders, not the washed-out default", () => {
    // The upstream serializer hardcoded `1px solid #ddd`, silently restyling
    // every bordered lab table on its first re-save.
    //
    // Written per side rather than as one `border:` shorthand, because each
    // edge carries its own width and style — see cellBorderStyles. The colour
    // is whichever spelling of black the source used; a shorthand comes back
    // out of the CSSOM as `rgb(0, 0, 0)`.
    const { html } = roundTrip(table);
    expect(html).toMatch(/border-(top|right|bottom|left): 1px solid (#000|rgb\(0, 0, 0\))/);
    expect(html).not.toContain("#ddd");
  });

  it("writes column widths into a colgroup when the table carries them", () => {
    const value = [
      {
        type: "table",
        colSizes: [120, 380],
        children: [
          {
            type: "tr",
            children: [
              { type: "td", children: [{ text: "a" }] },
              { type: "td", children: [{ text: "b" }] },
            ],
          },
        ],
      },
    ] as unknown as Value;
    const html = plateValueToHtml(value);
    expect(html).toContain("<colgroup>");
    expect(html).toContain("width: 120px");
    expect(html).toContain("width: 380px");
  });
});

describe("lists", () => {
  it("keeps a flat list", () => {
    const { html } = roundTrip("<ul><li>One</li><li>Two</li></ul>");
    expect(html).toContain("<ul");
    expect((html.match(/<li>/g) || []).length).toBe(2);
  });

  it("keeps a NESTED list inside its parent item", () => {
    // The old serializer flattened these away entirely.
    const { html } = roundTrip("<ul><li>Parent<ul><li>Child</li></ul></li></ul>");
    expect(html).toContain("Child");
    const firstUl = html.indexOf("<ul");
    const secondUl = html.indexOf("<ul", firstUl + 1);
    expect(secondUl).toBeGreaterThan(-1);
  });
});

describe("hard breaks", () => {
  it("preserves a shift+enter as <br/> rather than collapsing it", () => {
    const html = plateValueToHtml([
      { type: "p", children: [{ text: "Line one\nLine two" }] },
    ] as unknown as Value);
    expect(html).toContain("<br/>");
    expect(html).toContain("Line one");
    expect(html).toContain("Line two");
  });
});

describe("media nodes are degraded, never dropped", () => {
  it("turns a legacy file node into a link instead of deleting it", () => {
    const html = plateValueToHtml([
      { type: "file", url: "https://x/y.pdf", name: "y.pdf", children: [{ text: "" }] },
    ] as unknown as Value);
    expect(html).toContain("y.pdf");
    expect(html).toContain("href=");
  });
});

describe("paste normalisation modes", () => {
  const word = `<p style="margin: 0cm; line-height: 1.15;"><span style="font-size: 11pt;">Word text</span></p>`;

  it("clean mode keeps Word's pt sizes and a CSS line-height as stated", () => {
    const { html } = roundTrip(word, "clean");
    // The size the document states, in the unit it states it in — the control
    // then names an 11pt run 11, which is the number its author set. It used to
    // be converted to whole px, so 11pt became 15px and was named 15. See
    // lib/font-size.ts.
    expect(html).toContain("font-size: 11pt");
    // What clean mode is still FOR: Word's spacing, which the saved report
    // cannot carry, is gone.
    expect(html).not.toContain("margin");
    // A BARE ratio is CSS and is meant as CSS — this fixture carries none of
    // Word's markers. It used to be snapped to the nearest of the five gaps the
    // control offered, which an exact line gap cannot survive; and only a
    // PERCENTAGE from a real Word or LibreOffice clipboard is reinterpreted as
    // their multiplier. See lib/word-line-gap.ts.
    expect(html).toContain("line-height: 1.15");
  });

  it("faithful mode keeps them", () => {
    const { html } = roundTrip(word, "faithful");
    expect(html).toContain("11pt");
  });
});

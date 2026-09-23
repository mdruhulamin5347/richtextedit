/**
 * Word text boxes.
 *
 * Insert → Text Box is how a report gets its title bar — a full-width box, a
 * fill behind it, coloured text inside — and pasting one used to put a BROKEN
 * IMAGE in the report: Word ships the box's text and its fill inside a
 * conditional COMMENT no browser parses, and leaves every other renderer a
 * picture of it at a `file:///` path nothing on a web page may load.
 */
import { describe, expect, it, beforeEach } from "vitest";
import {
  DEFAULT_PASTED_LINE_GAP,
  FALLBACK_NATURAL_LINE_HEIGHT,
  lineGapToCssRatio,
} from "@/lib/line-gap";
import { act } from "react";
import { createPlateEditor } from "platejs/react";
import api from "@/entry";
import { buildPlugins } from "@/plugins";
import { plateValueToHtml } from "@/lib/html-serializer";
import { inlineWordTextboxes } from "@/lib/word-textbox";

function clipboard(html: string) {
  return {
    types: ["text/html", "text/plain"],
    getData: (type: string) => (type === "text/html" ? html : ""),
    files: [],
    items: [],
  } as unknown as DataTransfer;
}

function paste(html: string) {
  const editor = createPlateEditor({ plugins: buildPlugins("clean") });
  editor.tf.setValue([{ type: "p", children: [{ text: "" }] }] as never);
  editor.tf.select({ anchor: { path: [0, 0], offset: 0 }, focus: { path: [0, 0], offset: 0 } });
  editor.tf.insertData(clipboard(html));
  return plateValueToHtml(editor.children as never);
}

/** The picture fallback Word leaves for renderers that do not speak VML. */
const FALLBACK =
  `<![if !vml]><span style='mso-ignore:vglayout'>` +
  `<table cellpadding=0 cellspacing=0><tr><td width=602 height=47>` +
  `<img width=602 height=47 src="file:///C:/Users/x/AppData/Local/Temp/msohtmlclip1/01/clip_image001.png">` +
  `</td></tr></table></span><![endif]>`;

/** A text box as Word writes it: VML in a comment, picture fallback beside it. */
function wordTextbox(shapeAttrs: string, inner: string, trailing = "") {
  return (
    `<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="urn:schemas-microsoft-com:office:word">` +
    `<head><style><!-- p.MsoNormal {margin:0cm; font-size:11.0pt;} --></style></head><body>` +
    `<div class=WordSection1><p class=MsoNormal>` +
    `<!--[if gte vml 1]><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202"/>` +
    `<v:shape id="Text_x0020_Box_x0020_2" type="#_x0000_t202" ` +
    `style='position:absolute;width:451.3pt;height:35.3pt' ${shapeAttrs}>` +
    `<v:textbox>${inner}</v:textbox></v:shape><![endif]-->` +
    FALLBACK +
    trailing +
    `</p></div></body></html>`
  );
}

const TITLE_BAR = `<div><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:16.0pt;color:red'>MOLECULAR BIOLOGY REPORT</span></b></p></div>`;

describe("pasting a text box", () => {
  it("becomes the band it looks like: fill, text, colour, size and centring", () => {
    const html = paste(wordTextbox(`fillcolor="#a8cdea" stroked="f"`, TITLE_BAR));

    expect(html).toContain("background-color: rgb(168, 205, 234)");
    expect(html).toContain("MOLECULAR BIOLOGY REPORT");
    expect(html).toContain("color: red");
    expect(html).toContain("<strong>");
    expect(html).toContain("text-align: center");
    expect(html).toContain("font-size: 16pt");
  });

  it("drops the picture of itself, which no web page may load", () => {
    const html = paste(wordTextbox(`fillcolor="#a8cdea"`, TITLE_BAR));

    expect(html).not.toContain("file:///");
    expect(html).not.toContain("<img");
  });

  it("leaves no blank line where Word's anchor paragraph was", () => {
    // The block also picks up the default line gap every pasted block gets when
    // the document states none — see DEFAULT_PASTED_LINE_GAP.
    const gap = lineGapToCssRatio(DEFAULT_PASTED_LINE_GAP, FALLBACK_NATURAL_LINE_HEIGHT);
    expect(paste(wordTextbox(`fillcolor="#a8cdea"`, TITLE_BAR))).toBe(
      `<p style="text-align: center; line-height: ${gap}; font-size: 16pt; background-color: rgb(168, 205, 234)">` +
        `<span style="color: red; font-size: 16pt"><strong>MOLECULAR BIOLOGY REPORT</strong></span></p>`
    );
  });

  it("reads the fill Word states as a child element instead of an attribute", () => {
    const inner = TITLE_BAR;
    const html = paste(
      wordTextbox(`stroked="f"><v:fill color="#a8cdea" on="t"/><v:textbox>${inner}</v:textbox`, inner)
    );
    expect(html).toContain("background-color: rgb(168, 205, 234)");
  });

  it("shades nothing when the box is drawn without a fill", () => {
    const html = paste(wordTextbox(`filled="f" stroked="f"`, TITLE_BAR));
    expect(html).toContain("MOLECULAR BIOLOGY REPORT");
    expect(html).not.toContain("background-color");
  });

  it("unwraps the one-cell table Word wraps the content in for non-VML renderers", () => {
    const wrapped =
      `<![if !mso]><table cellpadding=0 cellspacing=0 width="100%"><tr><td style='padding:0'><![endif]>` +
      TITLE_BAR +
      `<![if !mso]></td></tr></table><![endif]>`;
    const html = paste(wordTextbox(`fillcolor="#a8cdea"`, wrapped));

    expect(html).toContain("background-color: rgb(168, 205, 234)");
    // The wrapper is scaffolding for Word, not a table in the report.
    expect(html).not.toContain("<table");
  });

  it("keeps text the anchor paragraph holds beside the box", () => {
    const html = paste(wordTextbox(`fillcolor="#a8cdea"`, TITLE_BAR, `Reported by the lab.`));
    expect(html).toContain("Reported by the lab.");
    expect(html).toContain("MOLECULAR BIOLOGY REPORT");
  });

  it("recovers every box when a header holds more than one", () => {
    const second = `<div><p class=MsoNormal><span style='color:#00b050'>ACCREDITED</span></p></div>`;
    const two =
      `<html xmlns:v="urn:schemas-microsoft-com:vml"><body><p class=MsoNormal>` +
      `<!--[if gte vml 1]><v:shape type="#_x0000_t202" fillcolor="#a8cdea"><v:textbox>${TITLE_BAR}</v:textbox></v:shape>` +
      `<v:shape type="#_x0000_t202" fillcolor="#ffff00"><v:textbox>${second}</v:textbox></v:shape><![endif]-->` +
      `</p></body></html>`;
    const html = paste(two);

    expect(html).toContain("MOLECULAR BIOLOGY REPORT");
    expect(html).toContain("ACCREDITED");
    expect(html).toContain("background-color: rgb(168, 205, 234)");
    expect(html).toContain("background-color: rgb(255, 255, 0)");
  });

  it("keeps a picture that shares the comment with the box", () => {
    // One comment can hold both. Tearing it out whole to get the box would take
    // the picture's `o:spid` with it — the only thing tying it to the image
    // bytes in the clipboard's RTF flavour — and the image would arrive as the
    // `file:///` placeholder it can never load.
    const both =
      `<html xmlns:v="urn:schemas-microsoft-com:vml"><body><p class=MsoNormal>` +
      `<!--[if gte vml 1]>` +
      `<v:shape id="Text_x0020_Box_x0020_2" type="#_x0000_t202" fillcolor="#a8cdea">` +
      `<v:textbox>${TITLE_BAR}</v:textbox></v:shape>` +
      `<v:shape id="Picture_x0020_1" o:spid="_x0000_s1027" type="#_x0000_t75">` +
      `<v:imagedata src="file:///C:/Temp/clip_image001.png" o:title=""/></v:shape>` +
      `<![endif]-->${FALLBACK}</p></body></html>`;

    const out = inlineWordTextboxes(both);

    expect(out).toContain("MOLECULAR BIOLOGY REPORT");
    expect(out).toContain("background-color: rgb(168, 205, 234)");
    // The picture's shape and the fallback Plate rewrites are both still there.
    expect(out).toContain("v:imagedata");
    expect(out).toContain("_x0000_s1027");
    expect(out).toContain("clip_image001.png");
  });

  it("recovers a filled box with nothing in it, which is a plain colour bar", () => {
    const html = paste(wordTextbox(`fillcolor="#a8cdea" stroked="f"`, `<div><p class=MsoNormal><o:p>&nbsp;</o:p></p></div>`));

    expect(html).toContain("background-color: rgb(168, 205, 234)");
    expect(html).not.toContain("file:///");
  });

  it("keeps the line Word drew around the box, as the one thing that can hold one", () => {
    // A paragraph cannot carry an outline; a table cell carries the line, the
    // fill and the inset together — and it is what an "Interpretation" or
    // "Note" box in a genetics report actually looks like.
    const html = paste(
      wordTextbox(
        `fillcolor="#fff2cc" strokecolor="#bf9000" strokeweight="1pt"`,
        `<div><p class=MsoNormal>Intermediate Metabolizer phenotype.</p>` +
          `<p class=MsoNormal>Interpret with the clinical history.</p></div>`
      )
    );

    expect(html).toContain("<table");
    expect(html).toContain("1px solid rgb(191, 144, 0)");
    expect(html).toContain("background-color: rgb(255, 242, 204)");
    expect(html).toContain("Intermediate Metabolizer phenotype.");
    expect(html).toContain("Interpret with the clinical history.");
    // The fill belongs to the cell now, not to each paragraph inside it.
    expect(html).not.toMatch(/<p[^>]*background-color/);
  });

  it("keeps the gap the box holds between its line and its text", () => {
    const html = paste(
      wordTextbox(
        `strokecolor="#000000" strokeweight="0.5pt"`,
        `<div><p class=MsoNormal>Note</p></div>`
      )
    );
    // Word's own default inset, 0.1in by 0.05in.
    expect(html).toContain("padding: 4.8px 9.6px");
  });

  it("draws no line the document did not ask for", () => {
    // VML's default is a thin black line, but a default is not a decision: a
    // box that says nothing about its stroke stays the band it looks like.
    const silent = paste(wordTextbox(`fillcolor="#deeaf6"`, TITLE_BAR));
    expect(silent).not.toContain("<table");
    expect(silent).toContain("background-color: rgb(222, 234, 246)");

    const off = paste(wordTextbox(`fillcolor="#deeaf6" stroked="f"`, TITLE_BAR));
    expect(off).not.toContain("<table");

    const strokeOff = paste(
      wordTextbox(`fillcolor="#deeaf6"><v:stroke on="f"/><v:textbox>${TITLE_BAR}</v:textbox`, TITLE_BAR)
    );
    expect(strokeOff).not.toContain("<table");
  });

  it("takes back the blank lines a FLOATING box was covering", () => {
    // In the document the box hovers over those paragraphs and nothing of them
    // shows. Pasted inline it takes its own space instead, so left behind they
    // become a second gap under it — the three or four blank lines that appear
    // between a boxed heading and the table beneath it.
    const blank = `<p class=MsoNormal><o:p>&nbsp;</o:p></p>`;
    const table = `<table border=1><tr><td><p class=MsoNormal>Gene</p></td></tr></table>`;
    const floating =
      `<html xmlns:v="urn:schemas-microsoft-com:vml"><body><p class=MsoNormal>` +
      `<!--[if gte vml 1]><v:shape id="TB" type="#_x0000_t202" ` +
      `style='position:absolute;width:451pt;height:36pt' fillcolor="#deeaf6" stroked="f">` +
      `<v:textbox>${TITLE_BAR}</v:textbox></v:shape><![endif]-->` +
      `<o:p>&nbsp;</o:p></p>${blank}${blank}${blank}${table}</body></html>`;

    const html = paste(floating);

    expect(html).toContain("MOLECULAR BIOLOGY REPORT");
    expect(html).toContain("<table");
    expect(html).not.toContain("<br/>");
  });

  it("takes back no more than the box was tall enough to cover", () => {
    // 36pt of box is about three blank lines; a fourth is the document's own.
    const blank = `<p class=MsoNormal><o:p>&nbsp;</o:p></p>`;
    const floating =
      `<html xmlns:v="urn:schemas-microsoft-com:vml"><body><p class=MsoNormal>` +
      `<!--[if gte vml 1]><v:shape type="#_x0000_t202" ` +
      `style='position:absolute;width:451pt;height:36pt' fillcolor="#deeaf6" stroked="f">` +
      `<v:textbox>${TITLE_BAR}</v:textbox></v:shape><![endif]-->` +
      `</p>${blank}${blank}${blank}${blank}<p class=MsoNormal>After</p></body></html>`;

    expect((paste(floating).match(/<br\/>/g) ?? []).length).toBe(1);
  });

  it("leaves the gap after an INLINE box alone", () => {
    // Nothing was floating over it, so those blank lines are the document's.
    const blank = `<p class=MsoNormal><o:p>&nbsp;</o:p></p>`;
    const inline =
      `<html xmlns:v="urn:schemas-microsoft-com:vml"><body><p class=MsoNormal>` +
      `<!--[if gte vml 1]><v:shape type="#_x0000_t202" style='width:451pt;height:36pt' ` +
      `fillcolor="#deeaf6" stroked="f"><v:textbox>${TITLE_BAR}</v:textbox></v:shape><![endif]-->` +
      `</p>${blank}${blank}<p class=MsoNormal>After</p></body></html>`;

    expect((paste(inline).match(/<br\/>/g) ?? []).length).toBe(2);
  });

  it("reads the colours Word actually writes, theme slot and all", () => {
    // Word does not write a plain colour on a shape. It writes the colour AND
    // the theme slot it came from — `fillcolor="#deeaf6 [660]"` — and that
    // bracket invalidates the whole CSS declaration.
    const html = paste(
      wordTextbox(`fillcolor="#deeaf6 [660]" strokecolor="#2e74b5 [2404]" strokeweight="1pt"`, TITLE_BAR)
    );

    expect(html).toContain("1px solid rgb(46, 116, 181)");
    expect(html).toContain("background-color: rgb(222, 234, 246)");
    // Not the default black line it fell back to while the slot was in there.
    expect(html).not.toContain("1px solid #000");
  });

  it("shades a box whose outline is off but whose fill is a theme colour", () => {
    // The paste that showed nothing but text: no outline to draw, and a fill
    // the browser threw away for the bracket after it.
    const html = paste(wordTextbox(`fillcolor="#deeaf6 [660]" stroked="f"`, TITLE_BAR));

    expect(html).toContain("background-color: rgb(222, 234, 246)");
    expect(html).toContain("MOLECULAR BIOLOGY REPORT");
  });

  it("reads a named theme colour too", () => {
    const html = paste(wordTextbox(`fillcolor="white [3201]" strokecolor="black [3213]" strokeweight=".5pt"`, TITLE_BAR));

    expect(html).toContain("background-color: white");
    expect(html).toContain("solid black");
  });

  it("leaves a pasted Word PICTURE alone, which is a shape but not a text box", () => {
    const picture =
      `<html xmlns:v="urn:schemas-microsoft-com:vml"><body><p class=MsoNormal>` +
      `<!--[if gte vml 1]><v:shape id="Picture_x0020_1" type="#_x0000_t75" style='width:100pt;height:50pt'>` +
      `<v:imagedata src="file:///C:/temp/image001.png" o:title=""/></v:shape><![endif]-->` +
      `<![if !vml]><img width=133 height=67 src="data:image/png;base64,iVBORw0KGgo="><![endif]>` +
      `</p></body></html>`;
    // Untouched by this pass: an image is Plate's own path (MediaKit + the RTF
    // flavour), and the fallback IS the picture rather than a stand-in for it.
    expect(inlineWordTextboxes(picture)).toBe(picture);
  });
});

describe("the rewrite itself", () => {
  it("returns HTML with no text box in it untouched, byte for byte", () => {
    const html = `<p>Pathogen Name</p>`;
    expect(inlineWordTextboxes(html)).toBe(html);
  });

  it("wraps a box holding bare runs, so the band has a block to be the width of", () => {
    const bare = `<html xmlns:v="urn:schemas-microsoft-com:vml"><body>` +
      `<!--[if gte vml 1]><v:shape fillcolor="#a8cdea"><v:textbox><div>Plain text</div></v:textbox></v:shape><![endif]-->` +
      `</body></html>`;
    const out = inlineWordTextboxes(bare);
    expect(out).toMatch(/<p style="background-color: [^"]+">Plain text<\/p>/);
  });
});

describe("opening a stored report that kept one", () => {
  beforeEach(() => {
    document.body.innerHTML = "";
  });

  it("recovers a box a legacy body saved whole", async () => {
    // A stored body is the report's markup, not a whole Word document: what
    // Summernote kept is the anchor paragraph with the comment still in it.
    const stored =
      `<p><!--[if gte vml 1]><v:shape type="#_x0000_t202" fillcolor="#a8cdea">` +
      `<v:textbox>${TITLE_BAR}</v:textbox></v:shape><![endif]-->${FALLBACK}</p>`;
    document.body.innerHTML = `
      <form id="f"><textarea id="report_body" hidden></textarea>
      <div class="rl-editor-scope" id="report_body_editor"></div></form>`;
    (document.getElementById("report_body") as HTMLTextAreaElement).value = stored;
    await act(async () => {
      api.mount("#report_body_editor", { textarea: "#report_body" });
    });

    const html = api.getHtml("#report_body_editor");
    expect(html).toContain("MOLECULAR BIOLOGY REPORT");
    expect(html).toContain("background-color: rgb(168, 205, 234)");
    expect(html).not.toContain("file:///");
  });
});

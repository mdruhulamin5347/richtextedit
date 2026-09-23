/**
 * A picture Word names in its clipboard HTML.
 *
 * Word does not put the picture in the HTML. It writes a path on the machine
 * doing the copying — `file:///C:/…/clip_image001.png` — and ships the bytes
 * separately, in the clipboard's RTF flavour. With the RTF, Plate swaps the
 * bytes in and the picture arrives as a data URL. Without it (an image LINKED
 * into the document rather than embedded, or a clipboard that crossed a remote
 * desktop) the path is all there is, and a web page may not load `file:`.
 *
 * The reference is kept either way. A picture that will not render is still the
 * document saying a picture belongs there, and that is the author's to decide
 * about — not something a paste should quietly delete.
 */
import { describe, expect, it, beforeEach } from "vitest";
import { act } from "react";
import { createPlateEditor } from "platejs/react";
import api from "@/entry";
import { buildPlugins } from "@/plugins";
import { plateValueToHtml } from "@/lib/html-serializer";

function clipboard(html: string, rtf = "") {
  return {
    types: rtf ? ["text/html", "text/rtf"] : ["text/html", "text/plain"],
    getData: (type: string) => (type === "text/html" ? html : type === "text/rtf" ? rtf : ""),
    files: [],
    items: [],
  } as unknown as DataTransfer;
}

function paste(html: string, rtf = "") {
  const editor = createPlateEditor({ plugins: buildPlugins("clean") });
  editor.tf.setValue([{ type: "p", children: [{ text: "" }] }] as never);
  editor.tf.select({ anchor: { path: [0, 0], offset: 0 }, focus: { path: [0, 0], offset: 0 } });
  editor.tf.insertData(clipboard(html, rtf));
  return plateValueToHtml(editor.children as never);
}

/** The report header: patient details in one cell, the QR in the next. */
const HEADER = (qr: string) =>
  `<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"` +
  ` xmlns:w="urn:schemas-microsoft-com:office:word"><body><table border=1 style='border-collapse:collapse'><tr>` +
  `<td style='padding:2px'><p class=MsoNormal>MRN: v2609256937</p>` +
  `<p class=MsoNormal>Patient's Name: Hazera</p></td>` +
  `<td style='padding:2px'>${qr}</td></tr></table></body></html>`;

/** Word's markup for an embedded picture: VML in a comment, `<img>` fallback. */
const WORD_PICTURE =
  `<!--[if gte vml 1]><v:shape id="Picture_x0020_1" o:spid="_x0000_s1026" type="#_x0000_t75"` +
  ` style='width:96pt;height:96pt'><v:imagedata src="file:///C:/Temp/clip_image001.png" o:title=""/>` +
  `</v:shape><![endif]--><![if !vml]>` +
  `<img width=96 height=96 src="file:///C:/Temp/clip_image001.png" v:shapes="Picture_x0020_1">` +
  `<![endif]>`;

/** The same picture's bytes, where Word actually puts them. */
const WORD_RTF = String.raw`{\rtf1\ansi{\shp{\*\shpinst shplid1026\pngblip\bliptag-1 4142}}}`;

describe("a Word picture the clipboard carries", () => {
  it("is recovered from the RTF as a data URL", () => {
    const html = paste(HEADER(WORD_PICTURE), WORD_RTF);

    expect(html).toContain("data:image/png;base64,");
    expect(html).not.toContain("file:///");
    expect(html).toContain("MRN: v2609256937");
  });
});

describe("a Word picture the clipboard does not carry", () => {
  it("keeps the reference rather than deleting the picture", () => {
    const html = paste(HEADER(WORD_PICTURE));

    expect(html).toContain("<img");
    expect(html).toContain("file:///C:/Temp/clip_image001.png");
    expect(html).toContain("MRN: v2609256937");
  });

  it("keeps every picture the report can show, unchanged", () => {
    const data = `data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAYAAACp8Z5+AAAAHUlEQVR42mP8z8BQz0AEYBxVSF+F/xkYGP4TowsAeYUH/WrKZ1UAAAAASUVORK5CYII=`;
    expect(paste(HEADER(`<img width=96 src="${data}">`))).toContain("data:image/png;base64,");
    expect(paste(HEADER(`<img width=96 src="https://example.test/qr.png">`))).toContain(
      `src="https://example.test/qr.png"`
    );
  });
});

describe("opening a stored report", () => {
  beforeEach(() => {
    document.body.innerHTML = "";
  });

  it("leaves one that holds such a picture exactly as it was", async () => {
    const stored = `<p><img src="file:///C:/Temp/clip_image001.png" alt="" style="max-width: 100%; height: auto;" /></p>`;
    document.body.innerHTML = `
      <form id="f"><textarea id="report_body" hidden></textarea>
      <div class="rl-editor-scope" id="report_body_editor"></div></form>`;
    (document.getElementById("report_body") as HTMLTextAreaElement).value = stored;
    await act(async () => {
      api.mount("#report_body_editor", { textarea: "#report_body" });
    });

    expect(api.getHtml("#report_body_editor")).toBe(stored);
  });
});

describe("a fill that is a picture rather than a colour", () => {
  it("is not written back out as one", () => {
    // `background: url(qr.png) no-repeat` states no colour at all. Passed
    // through it would be written as `background-color: url(qr.png) no-repeat`,
    // which is not a fill and not even valid CSS.
    expect(paste(`<p style="background: url('qr.png') no-repeat">MRN</p>`)).not.toContain(
      "background-color"
    );
    expect(
      paste(`<table><tr><td style="padding:2px;background:url(qr.png) no-repeat"><p>MRN</p></td></tr></table>`)
    ).not.toContain("background-color");
  });

  it("still reads the colour when the shorthand states one beside the picture", () => {
    expect(
      paste(`<table><tr><td style="padding:2px;background:url(qr.png) no-repeat #d9e2f3"><p>MRN</p></td></tr></table>`)
    ).toContain("background-color: rgb(217, 226, 243)");
  });
});

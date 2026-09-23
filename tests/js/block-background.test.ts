/**
 * Shading — the fill behind the text, in every place a report states one.
 *
 * Three different nodes carry it and they are not interchangeable: a RUN's fill
 * is a highlight, as wide as the text; a PARAGRAPH's is a band, the full width
 * of the block, which is what Word calls paragraph shading and what a section
 * heading on grey is made of; a CELL's is the shading of a banded table. Plate
 * reads a fill onto a text leaf only, so the block and the row lost theirs on
 * the way in — and the serializer wrote neither back, so the first save took
 * them out of the stored report for good.
 */
import { describe, expect, it, beforeEach } from "vitest";
import { act } from "react";
import { createPlateEditor } from "platejs/react";
import api from "@/entry";
import { buildPlugins } from "@/plugins";
import {
  DEFAULT_PASTED_LINE_GAP,
  FALLBACK_NATURAL_LINE_HEIGHT,
  lineGapToCssRatio,
} from "@/lib/line-gap";
import { plateValueToHtml } from "@/lib/html-serializer";

function mount(initialHtml: string) {
  document.body.innerHTML = `
    <form id="f"><textarea id="report_body" hidden>${initialHtml}</textarea>
    <div class="rl-editor-scope" id="report_body_editor"></div></form>`;
  return act(async () => {
    api.mount("#report_body_editor", { textarea: "#report_body" });
  });
}

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

const SHADE = "rgb(217, 226, 243)";

describe("opening a stored report", () => {
  beforeEach(() => {
    document.body.innerHTML = "";
  });

  it("keeps a paragraph's band, and re-saves it unchanged", async () => {
    const stored = `<p style="background-color: ${SHADE}">Whole paragraph shaded</p>`;
    await mount(stored);
    expect(api.getHtml("#report_body_editor")).toBe(stored);
  });

  it("keeps a heading's band", async () => {
    await mount(`<h2 style="background-color: #d9e2f3">Section</h2>`);
    expect(api.getHtml("#report_body_editor")).toBe(
      `<h2 style="background-color: ${SHADE}">Section</h2>`
    );
  });

  it("keeps a cell's shading", async () => {
    await mount(
      `<table border="1" style="border-collapse: collapse"><tbody><tr>` +
        `<td style="padding: 2px 5px; background-color: #d9e2f3"><p>Header cell</p></td>` +
        `</tr></tbody></table>`
    );
    expect(api.getHtml("#report_body_editor")).toContain(`background-color: ${SHADE}`);
  });

  it("keeps a run's highlight, which is not the same thing as a band", async () => {
    const stored = `<p><span style="background-color: yellow">Highlighted</span></p>`;
    await mount(stored);
    expect(api.getHtml("#report_body_editor")).toBe(stored);
  });

  it("leaves a report that shades nothing exactly as it was", async () => {
    const plain = "<p>Pathogen Name</p><p>Result</p>";
    await mount(plain);
    expect(api.getHtml("#report_body_editor")).toBe(plain);
  });

  it("is stable once saved: the same report reopens and re-saves unchanged", async () => {
    await mount(`<h2 style="background: #d9e2f3">Section</h2>`);
    const once = api.getHtml("#report_body_editor");
    document.body.innerHTML = "";
    await mount(once);
    expect(api.getHtml("#report_body_editor")).toBe(once);
  });
});

const DEFAULT_GAP_RATIO = lineGapToCssRatio(DEFAULT_PASTED_LINE_GAP, FALLBACK_NATURAL_LINE_HEIGHT);

describe("pasting", () => {
  it("keeps Word's paragraph shading as a band on the block", () => {
    const word =
      `<html xmlns:w="urn:schemas-microsoft-com:office:word"><body>` +
      `<p class=MsoNormal style='background:#D9E2F3;mso-shading:windowtext'>Section</p></body></html>`;
    // A pasted block that states no gap of its own is given the default one —
    // see DEFAULT_PASTED_LINE_GAP. jsdom lays nothing out, so the natural line
    // it is a multiple of is the documented fallback.
    expect(paste(word)).toBe(
      `<p style="line-height: ${DEFAULT_GAP_RATIO}; background-color: ${SHADE}">Section</p>`
    );
  });

  it("keeps LibreOffice's paragraph shading, which is what a .doc import brings", () => {
    const libre = `<p style="line-height: 100%; margin-bottom: 0in; background: #d9e2f3">Shaded</p>`;
    expect(paste(libre)).toContain(`background-color: ${SHADE}`);
  });

  it("shades the cells of a row the document shaded, in either spelling", () => {
    const attr = `<table><tbody><tr bgcolor="#d9e2f3"><td style="padding:2px"><p>row</p></td></tr></tbody></table>`;
    const style = `<table><tbody><tr style="background-color:#d9e2f3"><td style="padding:2px"><p>row</p></td></tr></tbody></table>`;
    // A cell is the only node in a Plate table that can hold a fill, so a row's
    // has to come to rest there — which is what the reader sees anyway.
    expect(paste(attr)).toContain(`background-color: #d9e2f3`);
    expect(paste(style)).toContain(`background-color: ${SHADE}`);
  });

  it("shades the cells of a table the document shaded", () => {
    const table = `<table style="background-color:#d9e2f3"><tbody><tr><td style="padding:2px"><p>x</p></td></tr></tbody></table>`;
    expect(paste(table)).toContain(`background-color: ${SHADE}`);
  });

  it("lets a cell's own fill win over the row it sits in", () => {
    const banded =
      `<table><tbody><tr bgcolor="#d9e2f3">` +
      `<td style="padding:2px;background:#ffff00"><p>own</p></td>` +
      `<td style="padding:2px"><p>row's</p></td>` +
      `</tr></tbody></table>`;
    const html = paste(banded);
    expect(html).toContain(`background-color: rgb(255, 255, 0)`);
    expect(html).toContain(`background-color: #d9e2f3`);
  });

  it("keeps a highlight on the run, and a `<mark>`", () => {
    expect(paste(`<p><span style='background:yellow;mso-highlight:yellow'>Word</span></p>`)).toContain(
      "background-color: yellow"
    );
    expect(paste(`<p><mark>marked</mark></p>`)).toContain("<mark>");
  });

  it("treats `transparent` as the absence of a fill, not a fill", () => {
    expect(paste(`<p style="background: transparent">nothing</p>`)).toBe(`<p>nothing</p>`);
    expect(
      paste(`<table><tbody><tr><td style="padding:2px;background:transparent"><p>x</p></td></tr></tbody></table>`)
    ).not.toContain("background-color");
  });
});

/**
 * Opening a report that states its alignment the legacy way.
 *
 * `<p align=center>`, `<center>` and `<tr align=center>` are how a report body
 * in this database says "centred" — the old Summernote editor wrote them, Word
 * writes them, and a browser renders every one of them centred. None reaches a
 * Plate node from a plain `text-align` read, so `inlineLegacyAlignment` rewrites
 * them into CSS first.
 *
 * It ran on the clipboard (`LegacyAlignmentPlugin`) and on the Word import, but
 * NOT when a stored report was opened — `transformData` is a clipboard-only
 * seam. So a centred column opened flat left, and the first save then wrote the
 * report back without its alignment, while pasting the identical markup came
 * out centred. That is what made it look intermittent.
 *
 * Driven through the real mount API, so it covers the path a report actually
 * takes: textarea -> editor -> textarea.
 */
import { describe, expect, it, beforeEach } from "vitest";
import { act } from "react";
import api from "@/entry";

function mount(initialHtml: string) {
  document.body.innerHTML = `
    <form id="f"><textarea id="report_body" hidden>${initialHtml}</textarea>
    <div class="rte-scope" id="report_body_editor"></div></form>`;
  return act(async () => {
    api.mount("#report_body_editor", { textarea: "#report_body" });
  });
}

const cell = (inner: string, attrs = "") => `<td ${attrs} style="padding: 2px 5px">${inner}</td>`;
const row = (pathogen: string, result: string, rowAttrs = "", cellAttrs = "", pAttrs = "") =>
  `<tr ${rowAttrs}>${cell(`<p>${pathogen}</p>`)}${cell(`<p ${pAttrs}>${result}</p>`, cellAttrs)}</tr>`;
const table = (rows: string) =>
  `<table border="1" style="border-collapse: collapse"><tbody>${rows}</tbody></table>`;

describe("a centred column in a stored report", () => {
  beforeEach(() => {
    document.body.innerHTML = "";
  });

  it("opens centred when the paragraph says so the legacy way", async () => {
    await mount(table(row("Measles morbillivirus", "Detected", "", "", "align=center")));
    expect(api.getHtml("#report_body_editor")).toContain(
      `<p style="text-align: center">Detected</p>`
    );
  });

  it("opens centred when the CELL says so", async () => {
    await mount(table(row("Mumps virus", "Not Detected", "", "align=center")));
    expect(api.getHtml("#report_body_editor")).toContain("text-align: center");
  });

  it("opens centred when the ROW says so, which is no node at all", async () => {
    // `text-align` inherits, so `<tr align=center>` centres every cell in the
    // row — but a row carries no alignment of its own here, so the value has to
    // be pushed onto the cells or it is simply lost.
    await mount(table(row("Rubella virus", "Not Detected", "align=center")));
    const html = api.getHtml("#report_body_editor");
    expect(html.match(/text-align: center/g)?.length).toBe(2);
  });

  it("and when a <center> block says so, outside any table", async () => {
    await mount("<center><p>END OF REPORT</p></center>");
    expect(api.getHtml("#report_body_editor")).toContain("text-align: center");
  });

  it("still reads the CSS spelling, which always worked", async () => {
    await mount(table(row("Varicella", "Not Detected", "", "", `style="text-align: center"`)));
    expect(api.getHtml("#report_body_editor")).toContain("text-align: center");
  });

  it("replacing the document through setHtml gets the same treatment", async () => {
    await mount("<p>placeholder</p>");
    await act(async () => {
      api.setHtml("#report_body_editor", table(row("Measles", "Detected", "", "", "align=center")));
    });
    expect(api.getHtml("#report_body_editor")).toContain("text-align: center");
  });

  it("leaves a report that states no alignment exactly as it was", async () => {
    const plain = "<p>Pathogen Name</p><p>Result</p>";
    await mount(plain);
    expect(api.getHtml("#report_body_editor")).toBe(plain);
  });

  it("does not centre a table that is merely PLACED in the centre", async () => {
    // `<div align=center>` around a table centres the TABLE on the page — a
    // browser's own `table { text-align: start }` leaves the cells alone.
    await mount(`<div align="center">${table(row("Measles", "Detected"))}</div>`);
    expect(api.getHtml("#report_body_editor")).not.toContain("text-align: center");
  });
});

/**
 * The "Upload Word" button, and which converter it hands the file to.
 *
 * Not a detail: Mammoth reads a .docx for its structure and drops every direct
 * format the document states — colour, size, alignment — so a red heading
 * imported BLACK no matter how carefully the rest of this editor preserves
 * colour, and nothing downstream could put back what never arrived. The server
 * route (LibreOffice) states all three in spellings the pipeline reads.
 */
import { describe, expect, it, beforeEach, vi } from "vitest";
import { act } from "react";
import api from "@/entry";

const { mammothConvert } = vi.hoisted(() => ({ mammothConvert: vi.fn() }));
vi.mock("mammoth", () => ({
  default: { convertToHtml: mammothConvert },
  convertToHtml: mammothConvert,
}));

/** What LibreOffice returns for a .docx whose heading is red. */
const LIBREOFFICE_HTML =
  `<p align="center" style="margin-bottom: 0in">` +
  `<font color="#ff0000"><font size="4" style="font-size: 16pt"><b>MOLECULAR BIOLOGY REPORT</b></font></font></p>`;

/** What Mammoth returns for the very same file. */
const MAMMOTH_HTML = `<p><strong>MOLECULAR BIOLOGY REPORT</strong></p>`;

function mount(options: { docConvertUrl?: string } = {}) {
  document.body.innerHTML = `
    <form id="f"><textarea id="report_body" hidden></textarea>
    <div class="rl-editor-scope" id="report_body_editor"></div></form>`;
  return act(async () => {
    api.mount("#report_body_editor", { textarea: "#report_body", ...options });
  });
}

async function upload(fileName: string) {
  const input = document.querySelector<HTMLInputElement>('input[type="file"]')!;
  const file = new File(["PK"], fileName);
  Object.defineProperty(input, "files", { value: [file], configurable: true });
  await act(async () => {
    input.dispatchEvent(new Event("change", { bubbles: true }));
  });
  // The conversion and the load that follows it are both promises.
  await act(async () => {});
}

function serverReturning(html: string) {
  return vi.fn(async () => ({ ok: true, json: async () => ({ html }) }) as unknown as Response);
}

describe("uploading a Word file", () => {
  beforeEach(() => {
    document.body.innerHTML = "";
    mammothConvert.mockReset();
    mammothConvert.mockResolvedValue({ value: MAMMOTH_HTML, messages: [] });
  });

  it("sends a .docx to LibreOffice, which is the converter that keeps the colour", async () => {
    const fetchMock = serverReturning(LIBREOFFICE_HTML);
    vi.stubGlobal("fetch", fetchMock);

    await mount({ docConvertUrl: "/tools/word-to-html" });
    await upload("report.docx");

    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(mammothConvert).not.toHaveBeenCalled();

    const html = api.getHtml("#report_body_editor");
    expect(html).toContain("MOLECULAR BIOLOGY REPORT");
    // The whole point: red on the way in, red on the way out.
    expect(html).toContain("color: rgb(255, 0, 0)");
    // And the two formats Mammoth drops along with it.
    expect(html).toContain("font-size: 16pt");
    expect(html).toContain("text-align: center");
  });

  it("sends a .doc there too, as it always did", async () => {
    const fetchMock = serverReturning(LIBREOFFICE_HTML);
    vi.stubGlobal("fetch", fetchMock);

    await mount({ docConvertUrl: "/tools/word-to-html" });
    await upload("report.doc");

    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(api.getHtml("#report_body_editor")).toContain("color: rgb(255, 0, 0)");
  });

  it("falls back to Mammoth for a .docx when there is no route to ask", async () => {
    // Formatting is lost that way — but an import still works where the server
    // converter is not configured, which is what it did before.
    await mount();
    await upload("report.docx");

    expect(mammothConvert).toHaveBeenCalledTimes(1);
    expect(api.getHtml("#report_body_editor")).toContain("MOLECULAR BIOLOGY REPORT");
  });

  it("falls back to Mammoth for a .docx when the conversion fails", async () => {
    const fetchMock = vi.fn(
      async () => ({ ok: false, json: async () => ({ message: "no" }) }) as unknown as Response
    );
    vi.stubGlobal("fetch", fetchMock);

    await mount({ docConvertUrl: "/tools/word-to-html" });
    await upload("report.docx");

    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(mammothConvert).toHaveBeenCalledTimes(1);
    expect(api.getHtml("#report_body_editor")).toContain("MOLECULAR BIOLOGY REPORT");
  });

  it("says so for a .doc it cannot convert, which the browser cannot read at all", async () => {
    await mount();
    await upload("report.doc");

    expect(mammothConvert).not.toHaveBeenCalled();
    expect(document.body.textContent).toContain("not configured");
  });
});

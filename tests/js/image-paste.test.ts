/**
 * Pictures, pasted and reopened.
 *
 * Two things were wrong and both showed on the printed page rather than in the
 * editor. A picture states its own WIDTH — `<img width=96>` is what Word and
 * LibreOffice write, and it is how an author shrinks a full-page scan to half a
 * page — and Plate's deserializer read it off neither the attribute nor the
 * style, so every pasted picture came back at its natural size. And a picture
 * that came in as `<p><img></p>`, which is how all of them do, became an image
 * node INSIDE a paragraph node: both wrote a `<p>` of their own and the report
 * was saved as `<p><p><img/></p></p>`, which no browser accepts — the print
 * page reparsed it into three paragraphs and drew a blank line above and below
 * every picture.
 */
import { describe, expect, it, beforeEach } from "vitest";
import { act } from "react";
import { createPlateEditor } from "platejs/react";
import api from "@/entry";
import { buildPlugins } from "@/plugins";
import { plateValueToHtml } from "@/lib/html-serializer";

const PNG =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAYAAACp8Z5+AAAAHUlEQVR42mP8z8BQz0AEYBxVSF+F/xkYGP4TowsAeYUH/WrKZ1UAAAAASUVORK5CYII=";

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

function mount(initialHtml: string) {
  document.body.innerHTML = `
    <form id="f"><textarea id="report_body" hidden></textarea>
    <div class="rte-scope" id="report_body_editor"></div></form>`;
  (document.getElementById("report_body") as HTMLTextAreaElement).value = initialHtml;
  return act(async () => {
    api.mount("#report_body_editor", { textarea: "#report_body" });
  });
}

describe("pasting a picture", () => {
  it("writes one paragraph around it, not a paragraph inside a paragraph", () => {
    const html = paste(`<p><img src="${PNG}"></p>`);

    expect(html).toBe(
      `<p><img src="${PNG}" alt="" style="max-width: 100%; height: auto;" /></p>`
    );
    expect(html).not.toContain("<p><p>");
  });

  it("keeps the width the document showed it at", () => {
    // What LibreOffice writes for an imported .doc, and Word for a paste.
    expect(paste(`<p><img src="${PNG}" width="96" height="96"></p>`)).toContain('width="96"');
    // …and the same size stated in the style instead.
    expect(paste(`<p><img src="${PNG}" style="width:200px;height:auto"></p>`)).toContain(
      'width="200"'
    );
  });

  it("ignores a width that is not a length here", () => {
    expect(paste(`<p><img src="${PNG}" width="100%"></p>`)).not.toContain("width=");
    expect(paste(`<p><img src="${PNG}" width="auto"></p>`)).not.toContain("width=");
  });

  it("keeps the alignment the paragraph centres it with", () => {
    const html = paste(`<p style="text-align:center"><img src="${PNG}" width="96"></p>`);

    expect(html).toContain("text-align: center");
    expect(html).toContain('width="96"');
  });

  it("keeps a picture inside a table cell", () => {
    const html = paste(
      `<table border="1"><tbody><tr><td style="padding:2px"><img src="${PNG}" width="96"></td></tr></tbody></table>`
    );

    expect(html).toContain("<td");
    expect(html).toContain('width="96"');
    expect(html).not.toContain("<p><p>");
  });

  it("leaves a picture served over the network as it is", () => {
    const html = paste(`<p><img src="https://example.test/logo.png" width="120"></p>`);

    expect(html).toContain('src="https://example.test/logo.png"');
    expect(html).toContain('width="120"');
  });
});

describe("opening a stored report with a picture", () => {
  beforeEach(() => {
    document.body.innerHTML = "";
  });

  it("reopens and re-saves it byte for byte", async () => {
    const stored = `<p><img src="${PNG}" width="96" alt="" style="max-width: 100%; height: auto;" /></p>`;
    await mount(stored);

    expect(api.getHtml("#report_body_editor")).toBe(stored);
  });

  it("does not grow a blank line around it on every save", async () => {
    // The nested `<p><p>` did exactly that: reparsed, it became three
    // paragraphs, and the print page drew two of them empty.
    await mount(`<p><img src="${PNG}" width="96"></p>`);
    const once = api.getHtml("#report_body_editor");
    document.body.innerHTML = "";
    await mount(once);

    expect(api.getHtml("#report_body_editor")).toBe(once);
    expect(once.match(/<p>/g)?.length).toBe(1);
  });
});

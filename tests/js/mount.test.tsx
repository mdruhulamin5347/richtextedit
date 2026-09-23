/**
 * Integration: the Blade-facing mount API.
 *
 * Proves the parts that have nothing to do with Plate itself — that React
 * actually attaches to a plain div inside an ordinary form, that the hidden
 * textarea ends up carrying HTML, and that a submit flushes the debounce
 * instead of posting stale content.
 */
import { describe, expect, it, beforeEach } from "vitest";
import { act } from "react";
import api from "@/entry";

function buildPage(initialHtml = "") {
  document.body.innerHTML = `
    <form id="f" action="/store" method="POST">
      <textarea name="report_body" id="report_body" hidden>${initialHtml}</textarea>
      <div class="rte-scope" id="report_body_editor"></div>
      <button type="submit">Create</button>
    </form>`;
  return {
    form: document.getElementById("f") as HTMLFormElement,
    textarea: document.getElementById("report_body") as HTMLTextAreaElement,
  };
}

/**
 * A stored report body with no block wrapper.
 *
 * Plate wraps stray root-level runs in a paragraph only when the fragment mixes
 * blocks and inlines, so a body that is inline all the way down stayed a list of
 * text nodes at the root — not a document. slate-react threw reading children
 * off one while painting, which took the whole editor down and left the page
 * with a dead box where the report should be. See `toRenderableValue`.
 */
describe("opening a report stored without a block wrapper", () => {
  beforeEach(() => {
    document.body.innerHTML = "";
  });

  const BODIES: Record<string, string> = {
    "bare text": "Normal study.",
    "a bare span": "<span>Normal study.</span>",
    "bare bold": "<b>Normal study.</b>",
    "an empty div": "<div></div>",
  };

  for (const [name, body] of Object.entries(BODIES)) {
    it(`renders ${name} instead of taking the editor down`, async () => {
      const { textarea } = buildPage(body);
      await act(async () => {
        api.mount("#report_body_editor", { textarea: "#report_body" });
      });

      expect(document.querySelector("[data-slate-editor]")).not.toBeNull();
      expect(api.getHtml("#report_body_editor")).toContain("<p>");
      expect(textarea).not.toBeNull();
    });
  }

  it("keeps the text, in the paragraph it was missing", async () => {
    buildPage("Normal study.");
    await act(async () => {
      api.mount("#report_body_editor", { textarea: "#report_body" });
    });
    expect(api.getHtml("#report_body_editor")).toBe("<p>Normal study.</p>");
  });
});

describe("RichTextEdit mount API", () => {
  beforeEach(() => {
    document.body.innerHTML = "";
  });

  it("exposes the global the Blade page calls", () => {
    expect(typeof api.mount).toBe("function");
    expect((window as any).RichTextEdit).toBe(api);
  });

  it("mounts into a plain div and renders an editable surface", async () => {
    buildPage("<p>Existing</p>");
    await act(async () => {
      api.mount("#report_body_editor", { textarea: "#report_body" });
    });
    const container = document.getElementById("report_body_editor")!;
    expect(container.querySelector("[data-slate-editor]")).not.toBeNull();
    expect(container.textContent).toContain("Existing");
  });

  it("seeds the editor from the textarea's existing HTML", async () => {
    buildPage("<p><strong>Seeded</strong></p>");
    await act(async () => {
      api.mount("#report_body_editor", { textarea: "#report_body" });
    });
    const html = api.getHtml("#report_body_editor");
    expect(html).toContain("<strong>Seeded</strong>");
  });

  it("writes HTML back into the textarea on submit, bypassing the debounce", async () => {
    const { form, textarea } = buildPage("<p>Report text</p>");
    await act(async () => {
      api.mount("#report_body_editor", { textarea: "#report_body" });
    });
    // Nothing has flushed yet — the debounce has not fired.
    textarea.value = "STALE";
    await act(async () => {
      form.dispatchEvent(new Event("submit", { bubbles: true, cancelable: true }));
    });
    expect(textarea.value).not.toBe("STALE");
    expect(textarea.value).toContain("Report text");
  });

  it("setHtml replaces the document, the summernote('code', …) equivalent", async () => {
    const { textarea } = buildPage("<p>Old</p>");
    await act(async () => {
      api.mount("#report_body_editor", { textarea: "#report_body" });
    });
    await act(async () => {
      api.setHtml("#report_body_editor", "<p>Replaced</p>");
    });
    expect(textarea.value).toContain("Replaced");
    expect(textarea.value).not.toContain("Old");
  });

  it("destroy unmounts and detaches the submit listener", async () => {
    const { form, textarea } = buildPage("<p>X</p>");
    await act(async () => {
      api.mount("#report_body_editor", { textarea: "#report_body" });
    });
    await act(async () => {
      api.destroy("#report_body_editor");
    });
    expect(document.getElementById("report_body_editor")!.innerHTML).toBe("");
    textarea.value = "UNTOUCHED";
    await act(async () => {
      form.dispatchEvent(new Event("submit", { bubbles: true, cancelable: true }));
    });
    expect(textarea.value).toBe("UNTOUCHED");
  });
});

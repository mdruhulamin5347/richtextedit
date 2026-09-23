/**
 * JSON storage: getValue / setValue and the `format: "json"` mount option.
 *
 * The HTML path is what every existing report body uses and is covered by
 * mount.test.tsx. These prove the JSON side is a peer of it — the same
 * document either way — and that the default is still HTML, so a page that
 * does not ask for JSON behaves exactly as it did before JSON existed.
 */
import { describe, expect, it, beforeEach } from "vitest";
import { act } from "react";
import type { Value } from "platejs";
import api from "@/entry";

function buildPage(initial = "") {
  document.body.innerHTML = `
    <form id="f" action="/store" method="POST">
      <textarea name="report_body" id="report_body" hidden></textarea>
      <div class="rl-editor-scope" id="report_body_editor"></div>
      <button type="submit">Create</button>
    </form>`;
  const textarea = document.getElementById("report_body") as HTMLTextAreaElement;
  // Assigned rather than templated, so JSON's quotes need no escaping.
  textarea.value = initial;
  return { form: document.getElementById("f") as HTMLFormElement, textarea };
}

async function mount(options: Partial<Parameters<typeof api.mount>[1]> = {}) {
  await act(async () => {
    api.mount("#report_body_editor", { textarea: "#report_body", ...options });
  });
}

const BOLD_VALUE: Value = [
  { type: "p", children: [{ text: "Seeded", bold: true }] },
] as Value;

const EMPTY = [{ type: "p", children: [{ text: "" }] }];

async function submit(form: HTMLFormElement) {
  await act(async () => {
    form.dispatchEvent(new Event("submit", { bubbles: true, cancelable: true }));
  });
}

describe("getValue / setValue on an HTML-stored body", () => {
  beforeEach(() => {
    document.body.innerHTML = "";
  });

  it("getValue returns the document the HTML was opened into", async () => {
    buildPage("<p><strong>Seeded</strong></p>");
    await mount();

    const value = api.getValue("#report_body_editor") as any[];
    expect(value[0].type).toBe("p");
    expect(value[0].children[0]).toMatchObject({ text: "Seeded", bold: true });
  });

  it("setValue replaces the document, and the textarea still gets HTML", async () => {
    const { textarea } = buildPage("<p>Old</p>");
    await mount();

    await act(async () => {
      api.setValue("#report_body_editor", BOLD_VALUE);
    });

    expect(api.getHtml("#report_body_editor")).toContain("<strong>Seeded</strong>");
    expect(textarea.value).toContain("<strong>Seeded</strong>");
    expect(textarea.value).not.toContain("Old");
  });

  it("setValue accepts the value as a JSON string", async () => {
    buildPage("<p>Old</p>");
    await mount();

    await act(async () => {
      api.setValue("#report_body_editor", JSON.stringify(BOLD_VALUE));
    });

    expect(api.getHtml("#report_body_editor")).toContain("<strong>Seeded</strong>");
  });

  it("HTML -> JSON -> HTML comes back unchanged", async () => {
    const html = '<p style="text-align: center"><strong>Impression</strong></p><p>Normal study.</p>';
    buildPage(html);
    await mount();
    const before = api.getHtml("#report_body_editor");

    const json = JSON.stringify(api.getValue("#report_body_editor"));
    await act(async () => {
      api.setValue("#report_body_editor", json);
    });

    expect(api.getHtml("#report_body_editor")).toBe(before);
  });
});

describe("the value getValue returns is a copy", () => {
  beforeEach(() => {
    document.body.innerHTML = "";
  });

  it("mutating it does not change the editor", async () => {
    buildPage("<p>Original</p>");
    await mount();

    const value = api.getValue("#report_body_editor") as any[];
    value[0].children[0].text = "Mutated";

    expect(api.getHtml("#report_body_editor")).toBe("<p>Original</p>");
  });

  it("mutating the value passed to setValue does not change the editor", async () => {
    buildPage("");
    await mount();
    const value = structuredClone(BOLD_VALUE) as any[];

    await act(async () => {
      api.setValue("#report_body_editor", value);
    });
    value[0].children[0].text = "Mutated";

    expect(api.getHtml("#report_body_editor")).toContain("Seeded");
    expect(api.getHtml("#report_body_editor")).not.toContain("Mutated");
  });
});

describe('format: "json"', () => {
  beforeEach(() => {
    document.body.innerHTML = "";
  });

  it("seeds the editor from JSON in the textarea", async () => {
    buildPage(JSON.stringify(BOLD_VALUE));
    await mount({ format: "json" });

    expect(api.getHtml("#report_body_editor")).toBe("<p><strong>Seeded</strong></p>");
  });

  it("writes JSON back into the textarea on submit", async () => {
    const { form, textarea } = buildPage(JSON.stringify(BOLD_VALUE));
    await mount({ format: "json" });

    textarea.value = "STALE";
    await submit(form);

    expect(JSON.parse(textarea.value)).toEqual(api.getValue("#report_body_editor"));
  });

  it("opens an empty textarea as an empty document", async () => {
    buildPage("");
    await mount({ format: "json" });

    expect(document.querySelector("[data-slate-editor]")).not.toBeNull();
    expect(api.getValue("#report_body_editor")).toEqual(EMPTY);
  });

  it("opens invalid JSON as an empty document instead of taking the editor down", async () => {
    buildPage("{not json");
    await mount({ format: "json" });

    expect(document.querySelector("[data-slate-editor]")).not.toBeNull();
    expect(api.getValue("#report_body_editor")).toEqual(EMPTY);
  });

  it("opens JSON that is not a document array as an empty document", async () => {
    buildPage(JSON.stringify({ type: "p", children: [{ text: "x" }] }));
    await mount({ format: "json" });

    expect(api.getValue("#report_body_editor")).toEqual(EMPTY);
  });

  it("wraps bare root-level text in a paragraph, as the HTML path does", async () => {
    buildPage(JSON.stringify([{ text: "Normal study." }]));
    await mount({ format: "json" });

    expect(document.querySelector("[data-slate-editor]")).not.toBeNull();
    expect(api.getHtml("#report_body_editor")).toBe("<p>Normal study.</p>");
  });

  it("setHtml still works, and the textarea gets JSON", async () => {
    const { textarea } = buildPage("");
    await mount({ format: "json" });

    await act(async () => {
      api.setHtml("#report_body_editor", "<p><strong>From HTML</strong></p>");
    });

    const stored = JSON.parse(textarea.value);
    expect(stored[0].children[0]).toMatchObject({ text: "From HTML", bold: true });
  });
});

describe("the default is still HTML", () => {
  beforeEach(() => {
    document.body.innerHTML = "";
  });

  it("writes HTML into the textarea when no format is given", async () => {
    const { form, textarea } = buildPage("<p>Report text</p>");
    await mount();

    await submit(form);

    expect(textarea.value).toBe("<p>Report text</p>");
  });

  it("reads JSON-looking text in the textarea as HTML, not as a document", async () => {
    buildPage('[{"text":"x"}]');
    await mount();

    expect(api.getHtml("#report_body_editor")).toContain("[{");
  });
});

describe("onChange", () => {
  beforeEach(() => {
    document.body.innerHTML = "";
  });

  it("receives both HTML and JSON", async () => {
    const { form } = buildPage("<p>Report text</p>");
    const changes: { html: string; value: Value }[] = [];
    await mount({ onChange: (change) => changes.push(change) });

    await submit(form);

    const last = changes.at(-1)!;
    expect(last.html).toBe("<p>Report text</p>");
    expect((last.value as any[])[0].children[0].text).toBe("Report text");
  });
});

describe("before anything is mounted", () => {
  it("getValue returns an empty array and setValue does nothing", () => {
    document.body.innerHTML = "";
    expect(api.getValue("#missing")).toEqual([]);
    expect(() => api.setValue("#missing", BOLD_VALUE)).not.toThrow();
  });
});

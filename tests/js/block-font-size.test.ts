/**
 * The size a BLOCK is set in, which is what decides how tall its line is.
 *
 * Every line box carries a STRUT — an invisible inline box in the block's own
 * font — so a paragraph's height comes from the BLOCK, not from the text in it.
 * Here that was always the editable's 18px base, so a row of 10px text and a
 * row of 15px text came out the same height, both taller than the document.
 *
 * These pin the rule the editor applies instead: a paragraph is the size its
 * own runs agree on, a blank line is the size in effect around it, and the
 * saved HTML says so — the print page states `body p {font-size: 15px}` and an
 * inline size outranks it, so the page and the screen keep the same shape.
 */
import { describe, expect, it } from "vitest";
import { act } from "react";
import api from "@/entry";
import { createPlateEditor } from "platejs/react";
import { deserializeHtml, type Value } from "platejs";
import { buildPlugins } from "@/plugins";
import { plateValueToHtml } from "@/lib/html-serializer";

function roundTrip(html: string, pasteMode: "clean" | "faithful" = "clean") {
  const editor = createPlateEditor({ plugins: buildPlugins(pasteMode) });
  const value = deserializeHtml(editor, { element: html }) as Value;
  editor.tf.setValue(value);
  editor.tf.normalize({ force: true });
  const out = editor.children as Value;
  return { html: plateValueToHtml(out), value: out as any[] };
}

const TEN = '<p><span style="font-size: 10px">ten</span></p>';
const THIRTY = '<p><span style="font-size: 30px">thirty</span></p>';
const BLANK = "<p><br/></p>";

describe("a blank line takes the size in effect on it", () => {
  it("from the line above", () => {
    const { value } = roundTrip(TEN + BLANK);
    expect(value[1].fontSize).toBe("10px");
  });

  it("from the line below when there is nothing above", () => {
    const { value } = roundTrip(BLANK + THIRTY);
    expect(value[0].fontSize).toBe("30px");
  });

  it("so two reports set in different sizes get different gaps", () => {
    expect(roundTrip(TEN + BLANK + TEN).value[1].fontSize).toBe("10px");
    expect(roundTrip(THIRTY + BLANK + THIRTY).value[1].fontSize).toBe("30px");
  });

  it("one it states itself beats either neighbour", () => {
    const { value } = roundTrip(
      TEN + '<p><span style="font-size: 24px"><br/></span></p>' + TEN
    );
    expect(value[1].fontSize).toBe("24px");
  });

  it("and a run of blank lines is one height, not alternating", () => {
    const { value } = roundTrip(TEN + BLANK + BLANK + BLANK + THIRTY);
    expect([value[1].fontSize, value[2].fontSize, value[3].fontSize]).toEqual([
      "10px",
      "10px",
      "10px",
    ]);
  });

  it("nothing at all when no line around it states a size", () => {
    const { value } = roundTrip("<p>plain</p>" + BLANK);
    expect(value[1].fontSize).toBeUndefined();
  });

  it("and nothing when the neighbouring line's runs disagree", () => {
    const mixed = '<p><span style="font-size: 10px">a</span><span style="font-size: 30px">b</span></p>';
    expect(roundTrip(mixed + BLANK).value[1].fontSize).toBeUndefined();
  });
});

describe("a line with text is the size of its text", () => {
  it("takes the size its runs agree on", () => {
    const { value, html } = roundTrip(TEN);
    expect(value[0].fontSize).toBe("10px");
    expect(html).toBe(`<p style="font-size: 10px"><span style="font-size: 10px">ten</span></p>`);
  });

  it("so two rows of different text are no longer the same height", () => {
    expect(roundTrip(TEN).value[0].fontSize).toBe("10px");
    expect(roundTrip(THIRTY).value[0].fontSize).toBe("30px");
  });

  it("keeps the editor's base when its runs disagree", () => {
    const mixed = `<p><span style="font-size: 10px">a</span><span style="font-size: 30px">b</span></p>`;
    expect(roundTrip(mixed).value[0].fontSize).toBeUndefined();
  });

  it("is not thrown off by an indentation run, which states nothing and shows nothing", () => {
    // Word writes a line's indentation as a run of its own with no size on it —
    // `<span style='mso-spacerun:yes'>&nbsp;&nbsp; </span>`. Counting that as a
    // run that disagrees sent every indented line back to the editor's base: in
    // a three-signature footer the first column stood taller than the other two
    // and its indentation was drawn in 18px spaces.
    const indented = `<p><span style="white-space: pre">&nbsp;&nbsp;&nbsp;</span><span style="font-size: 15px">Department of Molecular Biology</span></p>`;
    expect(roundTrip(indented).value[0].fontSize).toBe("15px");
  });

  it("but a whitespace run that DOES state a size still counts", () => {
    const disagreeing = `<p><span style="font-size: 30px; white-space: pre">&nbsp;&nbsp;</span><span style="font-size: 15px">x</span></p>`;
    expect(roundTrip(disagreeing).value[0].fontSize).toBeUndefined();
  });

  it("and when any run states no size at all — there is no one size to speak of", () => {
    const partial = `<p><span style="font-size: 10px">Result</span>: Negative</p>`;
    expect(roundTrip(partial).value[0].fontSize).toBeUndefined();
  });

  it("asks its own runs and never its neighbours: one line cannot resize the next", () => {
    const { value } = roundTrip(TEN + "<p>plain</p>");
    expect(value[1].fontSize).toBeUndefined();
  });

  it("stops stating one the moment the size is taken off the text", () => {
    const { value } = roundTrip("<p>plain</p>");
    expect(value[0].fontSize).toBeUndefined();
  });
});

describe("what it does not touch", () => {
  it("a line of whitespace is layout, not a blank line", () => {
    // Tabs and space runs set a report's columns, and a space is as wide as
    // the size it is set in — resizing one would move what it aligns.
    const tabs = '<p><span style="white-space: pre">\t\t</span></p>';
    expect(roundTrip(TEN + tabs).value[1].fontSize).toBeUndefined();
  });

  it("a report that states no size anywhere is left exactly as it was", () => {
    const plain = "<p>Pathogen Name</p><p><br/></p><p>Result</p>";
    expect(roundTrip(plain).html).toBe(plain);
  });
});

describe("inside a table cell, which is where this was asked for", () => {
  it("sizes the paragraph in a cell like any other", () => {
    const { value } = roundTrip(`<table><tr><td>${TEN}</td></tr></table>`);
    const cell = value[0].children[0].children[0];
    expect(cell.children[0].fontSize).toBe("10px");
  });

  it("and a blank line in a cell from the lines around it", () => {
    const { value } = roundTrip(`<table><tr><td>${TEN}${BLANK}</td></tr></table>`);
    const cell = value[0].children[0].children[0];
    expect(cell.children[1].fontSize).toBe("10px");
  });
});

describe("the saved report carries the gap", () => {
  it("states the size on the blank <p>, which is what prints it", () => {
    const { html } = roundTrip(TEN + BLANK);
    expect(html).toContain('<p style="font-size: 10px"><br/></p>');
  });

  it("keeps a deliberately sized blank line on reopening", () => {
    const once = roundTrip(TEN + '<p><span style="font-size: 24px"><br/></span></p>');
    expect(once.html).toContain("font-size: 24px");
    // Plate reads a size off a <span>, never off a <p>: the span is what makes
    // the second open agree with the first.
    const twice = roundTrip(once.html);
    expect(twice.value[1].fontSize).toBe("24px");
    expect(twice.html).toBe(once.html);
  });

  it("and a derived one is byte-stable, so saving twice changes nothing", () => {
    const once = roundTrip(TEN + BLANK + TEN).html;
    expect(roundTrip(once).html).toBe(once);
  });

  it("leaves an unsized report's blank lines exactly as they were", () => {
    const plain = "<p>plain</p><p><br/></p><p>more</p>";
    expect(roundTrip(plain).html).toBe(plain);
  });

  it("states the size on the block, which is what the print page's own rule loses to", () => {
    // `body p { font-size: 15px }` in the print stylesheet is not `!important`,
    // so an inline size on the paragraph wins and the printed line is as tall
    // as the editor's.
    expect(roundTrip(THIRTY).html).toContain(`<p style="font-size: 30px">`);
  });
});

describe("it must not fight the paste normalizer", () => {
  /**
   * A patient header table, sizes in POINTS — which is how Word states every
   * one of them.
   *
   * The block's size is derived from the runs, and the paste normalizer used to
   * convert a block's size from `11pt` to `15px` while the runs it was derived
   * from still said `11pt`. So one normalizer rewrote what the other had just
   * written, forever: Slate gave up after its 42 passes per dirty path and the
   * editor rendered NOTHING — the whole report gone, not merely mis-sized.
   *
   * Driven through the real mount API because that is the path a report takes,
   * and because the failure was a thrown error rather than a wrong number.
   */
  const cell = (t: string) =>
    `<td style='border:solid windowtext 1.0pt;padding:0cm 5.4pt'>` +
    `<p class=MsoNormal><span style='font-size:11.0pt'>${t}</span></p></td>`;
  const HEADER =
    `<table border=1 cellspacing=0 style='border-collapse:collapse;width:100%'>` +
    [
      ["Inv. ID :", "3500000"],
      ["Inv. Date:", "26-11-2025"],
      ["Name :", "MR. ABC"],
      ["Bed/ward:", "ICU-05"],
    ]
      .map(([label, value]) => `<tr>${cell(label)}${cell(value)}</tr>`)
      .join("") +
    `</table>`;

  it("opens a report whose sizes are in points, instead of failing to normalize", async () => {
    document.body.innerHTML = `
      <form id="f"><textarea id="report_body" hidden>${HEADER.replace(/</g, "&lt;")}</textarea>
      <div class="rte-scope" id="report_body_editor"></div></form>`;

    await act(async () => {
      api.mount("#report_body_editor", { textarea: "#report_body" });
    });

    const container = document.getElementById("report_body_editor")!;
    expect(container.querySelector("[data-slate-editor]")).not.toBeNull();
    expect(container.querySelectorAll("tr").length).toBe(4);
    expect(container.textContent).toContain("Bed/ward:");
  });

  it("and the block agrees with its runs rather than being rewritten", () => {
    const { value } = roundTrip(`<p><span style="font-size: 11pt">x</span></p>`);
    expect(value[0].fontSize).toBe(value[0].children[0].fontSize);
  });
});

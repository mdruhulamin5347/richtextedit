/**
 * Pasted table column widths.
 *
 * Plate's table deserializer keeps rows, cells and spans but discards every
 * width — no `<td width>`, no `style="width:85pt"`, no `<colgroup>`. Without
 * `colSizes` on the table node every column renders the same, which is very
 * visible on the label / colon / value tables these reports are built from:
 * the middle column holds one ":" and was coming out a third of the table wide.
 */
import { describe, expect, it } from "vitest";
import { createPlateEditor } from "platejs/react";
import { deserializeHtml, type Value } from "platejs";
import { buildPlugins } from "@/plugins";
import { plateValueToHtml } from "@/lib/html-serializer";
import {
  borderShorthand,
  extractCellAlign,
  extractCellBackground,
  extractCellBorders,
  extractCellPadding,
  extractCellVerticalAlign,
  extractColSizes,
  extractRowSize,
  fitColSizes,
  tableColSizes,
} from "@/lib/table-widths";

/** A real ECG report as Word puts it on the clipboard. */
const ECG_ROWS: [string, string][] = [
  ["Rate", "85 b/min"],
  ["Rhythm", "Regular"],
  ["P-Wave", "Normal"],
  ["P-R Interval", "Normal"],
  ["QRS Complex", "Normal"],
  ["ST. Segment", "Isoelectric"],
  ["T. Wave", "Normal"],
  ["Impression", "Findings are within normal limit."],
];

const ECG_HTML = `<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0
 style='border-collapse:collapse;border:none'>${ECG_ROWS.map(
   ([label, value]) => `<tr>
  <td width=113 valign=top style='width:85.0pt;border:solid windowtext 1.0pt;padding:0cm 5.4pt'><p class=MsoNormal>${label}</p></td>
  <td width=28 valign=top style='width:21.0pt;border:solid windowtext 1.0pt;padding:0cm 5.4pt'><p class=MsoNormal>:</p></td>
  <td width=283 valign=top style='width:212.0pt;border:solid windowtext 1.0pt;padding:0cm 5.4pt'><p class=MsoNormal>${value}</p></td>
 </tr>`
 ).join("")}</table>`;

/**
 * Must return "" for anything but text/html — Slate probes
 * application/x-slate-fragment and runs atob() on whatever comes back.
 */
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
  const value = editor.children as Value;
  const table = (value as any[]).find((n) => n.type === "table");
  return { table, html: plateValueToHtml(value) };
}

describe("pasting the ECG report table", () => {
  const { table, html } = paste(ECG_HTML);

  it("carries the source column widths onto the table node", () => {
    // 85pt / 21pt / 212pt at 96dpi.
    expect(table?.colSizes).toEqual([113, 28, 283]);
  });

  it("keeps the colon column narrow instead of giving it a third of the table", () => {
    const [label, colon, value] = table.colSizes as number[];
    expect(colon).toBeLessThan(label / 2);
    expect(colon).toBeLessThan(value / 5);
  });

  it("keeps every row and the content", () => {
    expect((html.match(/<tr>/g) || []).length).toBe(ECG_ROWS.length);
    expect(html).toContain("85 b/min");
    expect(html).toContain("Isoelectric");
    expect(html).toContain("Findings are within normal limit.");
  });

  it("writes the widths into the saved HTML as a colgroup", () => {
    expect(html).toContain("<colgroup>");
    expect(html).toContain("width: 113px");
    expect(html).toContain("width: 28px");
    expect(html).toContain("width: 283px");
  });

  it("reopens with the same widths — the round trip holds", () => {
    const editor = createPlateEditor({ plugins: buildPlugins("clean") });
    const reopened = deserializeHtml(editor, { element: html }) as Value;
    const table2 = (reopened as any[]).find((n) => n.type === "table");
    expect(table2?.colSizes).toEqual([113, 28, 283]);
  });
});

describe("width extraction", () => {
  const table = (inner: string) => {
    const el = document.createElement("div");
    el.innerHTML = `<table>${inner}</table>`;
    return el.firstElementChild as HTMLElement;
  };

  it("prefers a colgroup when one is present", () => {
    expect(
      extractColSizes(
        table(`<colgroup><col width="100"><col width="40"></colgroup><tr><td>a</td><td>b</td></tr>`)
      )
    ).toEqual([100, 40]);
  });

  it("converts CSS units to px", () => {
    expect(
      extractColSizes(table(`<tr><td style="width:1in">a</td><td style="width:2.54cm">b</td></tr>`))
    ).toEqual([96, 96]);
  });

  it("spreads a merged cell's width across the columns it spans", () => {
    expect(
      extractColSizes(table(`<tr><td colspan="2" style="width:200px">a</td><td style="width:50px">b</td></tr>`))
    ).toEqual([100, 100, 50]);
  });

  it("expands a <col span> into one entry per column", () => {
    expect(
      extractColSizes(table(`<colgroup><col span="2" width="60"><col width="30"></colgroup><tr><td>a</td></tr>`))
    ).toEqual([60, 60, 30]);
  });

  it("scales percentage widths into proportional pixels", () => {
    const sizes = extractColSizes(
      table(`<tr><td style="width:25%">a</td><td style="width:75%">b</td></tr>`)
    )!;
    expect(sizes[1] / sizes[0]).toBeCloseTo(3, 1);
  });

  it("reproduces a stated width EXACTLY, however narrow", () => {
    // No rounding-up to a comfortable minimum: Word uses very narrow columns
    // for the ":" separator in label/value tables, and widening them silently
    // is what made a pasted report stop matching its source.
    expect(
      extractColSizes(
        table(`<tr><td style="width:10px">A</td><td style="width:20px">B</td><td style="width:60px">C</td></tr>`)
      )
    ).toEqual([10, 20, 60]);
  });

  it("returns nothing when the markup states no widths, leaving Plate's defaults", () => {
    expect(extractColSizes(table(`<tr><td>a</td><td>b</td></tr>`))).toBeUndefined();
  });

});

describe("row heights", () => {
  const row = (html: string) => {
    const el = document.createElement("div");
    el.innerHTML = `<table><tbody>${html}</tbody></table>`;
    return el.querySelector("tr") as HTMLElement;
  };

  it("reads a CSS height off the row", () => {
    expect(extractRowSize(row(`<tr style="height:24pt"><td>a</td></tr>`))).toBe(32);
  });

  it("reads a height attribute", () => {
    expect(extractRowSize(row(`<tr height="40"><td>a</td></tr>`))).toBe(40);
  });

  it("falls back to the row's first cell, where Word often puts it", () => {
    expect(extractRowSize(row(`<tr><td style="height:30px">a</td></tr>`))).toBe(30);
  });

  it("returns nothing when no height is stated", () => {
    expect(extractRowSize(row(`<tr><td>a</td></tr>`))).toBeUndefined();
  });

  it("survives paste and a save/reopen round trip", () => {
    const source = `<table><tr style="height:36px"><td style="width:100px">a</td><td style="width:40px">b</td></tr></table>`;
    const { table, html } = paste(source);
    expect(table?.children?.[0]?.size).toBe(36);
    expect(html).toContain("height: 36px");

    const editor = createPlateEditor({ plugins: buildPlugins("clean") });
    const reopened = deserializeHtml(editor, { element: html }) as Value;
    const table2 = (reopened as any[]).find((n: any) => n.type === "table");
    expect(table2?.children?.[0]?.size).toBe(36);
    expect(table2?.colSizes).toEqual([100, 40]);
  });
});

describe("fitting a pasted table to the editor width", () => {
  const table = (inner: string) => {
    const el = document.createElement("div");
    el.innerHTML = `<table>${inner}</table>`;
    return el.firstElementChild as HTMLElement;
  };

  /** The five-column layout of a Word haematology report, stated at page width. */
  const WORD_ROW = `<tr><td style="width:170px">Haemoglobin</td><td style="width:20px">:</td><td style="width:120px">13.5</td><td style="width:120px">gm/dl</td><td style="width:194px">13.0 - 17.0</td></tr>`;

  it("stretches a page-width table across the whole editor", () => {
    const sizes = tableColSizes(table(WORD_ROW), 1276)!;
    expect(sizes.reduce((a, b) => a + b, 0)).toBe(1276);
  });

  it("holds the columns' proportions while doing it", () => {
    const stated = [170, 20, 120, 120, 194];
    const sizes = tableColSizes(table(WORD_ROW), 1276)!;
    const statedTotal = stated.reduce((a, b) => a + b, 0);
    sizes.forEach((size, i) => {
      expect(size / 1276).toBeCloseTo(stated[i] / statedTotal, 3);
    });
  });

  it("fills the width with equal columns when the source states none", () => {
    const sizes = tableColSizes(table(`<tr><td>a</td><td>b</td><td>c</td></tr>`), 900)!;
    expect(sizes).toHaveLength(3);
    expect(sizes.reduce((a, b) => a + b, 0)).toBe(900);
    expect(Math.max(...sizes) - Math.min(...sizes)).toBeLessThanOrEqual(1);
  });

  it("counts a merged cell as the columns it spans", () => {
    const sizes = tableColSizes(table(`<tr><td colspan="2">a</td><td>b</td></tr>`), 900)!;
    expect(sizes).toHaveLength(3);
  });

  it("invents nothing when there is no width to fit", () => {
    // The initial value is deserialized before the editable is mounted, and a
    // test editor has no layout at all. Both must keep the source's widths.
    expect(tableColSizes(table(WORD_ROW), undefined)).toEqual([170, 20, 120, 120, 194]);
    expect(tableColSizes(table(`<tr><td>a</td><td>b</td></tr>`), undefined)).toBeUndefined();
  });

  it("lands on the target exactly, rounding drift included", () => {
    const sizes = fitColSizes([1, 1, 1], 1000)!;
    expect(sizes.reduce((a, b) => a + b, 0)).toBe(1000);
  });

  it("keeps a hair-thin column on the grid rather than at zero", () => {
    const sizes = fitColSizes([1000, 1], 300)!;
    expect(sizes[1]).toBeGreaterThan(0);
  });

  it("leaves a table that already fits alone", () => {
    expect(fitColSizes([300, 400], 700)).toEqual([300, 400]);
  });
});

describe("cell alignment", () => {
  const cell = (attrs: string) => {
    const el = document.createElement("div");
    el.innerHTML = `<table><tr><td ${attrs}>a</td></tr></table>`;
    return el.querySelector("td") as HTMLElement;
  };

  it("reads the align attribute Word puts on the cell", () => {
    // A centred header row often arrives as `<td align=center>` around a plain
    // paragraph. The align plugin only looks at the paragraph, so without this
    // the whole row came back left-aligned.
    expect(extractCellAlign(cell('align="center"'))).toBe("center");
  });

  it("reads a CSS text-align just the same", () => {
    expect(extractCellAlign(cell('style="text-align:right"'))).toBe("right");
  });

  it("prefers the CSS, which is what a browser would apply", () => {
    expect(extractCellAlign(cell('align="left" style="text-align:center"'))).toBe("center");
  });

  it("ignores what is not an alignment", () => {
    expect(extractCellAlign(cell('align="middle"'))).toBeUndefined();
    expect(extractCellAlign(cell('valign="top"'))).toBeUndefined();
    expect(extractCellAlign(cell(""))).toBeUndefined();
  });

  it("survives paste and reaches the saved report", () => {
    const { table, html } = paste(
      `<table><tr><td align="center" style="width:100px">Result</td><td style="width:100px">x</td></tr></table>`
    );
    expect(table?.children?.[0]?.children?.[0]?.align).toBe("center");
    expect(html).toContain("text-align: center");
  });
});

describe("cell shading and vertical alignment", () => {
  const cell = (attrs: string) => {
    const el = document.createElement("div");
    el.innerHTML = `<table><tr><td ${attrs}>a</td></tr></table>`;
    return el.querySelector("td") as HTMLElement;
  };

  it("reads the fill Word shades a header row with", () => {
    expect(extractCellBackground(cell('style="background:#D9E2F3"'))).toBe("rgb(217, 226, 243)");
  });

  it("treats no fill as no fill", () => {
    expect(extractCellBackground(cell('style="background:transparent"'))).toBeUndefined();
    expect(extractCellBackground(cell(""))).toBeUndefined();
  });

  it("reads the valign Word puts on practically every report cell", () => {
    // An HTML cell defaults to `middle`; a Word one sits at the top. Dropping
    // this dropped the text of every wrapped row off its neighbours' line.
    expect(extractCellVerticalAlign(cell('valign="top"'))).toBe("top");
    expect(extractCellVerticalAlign(cell('style="vertical-align:bottom"'))).toBe("bottom");
  });

  it("ignores what is not a vertical alignment", () => {
    expect(extractCellVerticalAlign(cell('valign="centre"'))).toBeUndefined();
    expect(extractCellVerticalAlign(cell(""))).toBeUndefined();
  });

  it("both survive paste and reach the saved report", () => {
    const { table, html } = paste(
      `<table><tr><td valign="top" style="width:100px;background:#D9E2F3">Shaded</td><td style="width:100px">x</td></tr></table>`
    );
    const first = table?.children?.[0]?.children?.[0];
    expect(first?.verticalAlign).toBe("top");
    expect(first?.background).toBe("rgb(217, 226, 243)");
    expect(html).toContain("vertical-align: top");
    expect(html).toContain("background-color: rgb(217, 226, 243)");
  });
});

/**
 * The gap a cell keeps around its text.
 *
 * The one piece of the old Summernote editor's table handling this one had no
 * answer for. That editor never parsed a table: the clipboard HTML went into a
 * contenteditable and the browser obeyed Word's `padding:0cm 5.4pt` as written
 * — nothing above or below the text, 7.2px either side. This editor rebuilt
 * every cell with a fixed `px-3 py-2` instead, so each row gained ~16px of
 * height and a ten-row pasted table stood a long way deeper than the document
 * it came from.
 */
describe("cell padding", () => {
  const cell = (attrs: string, tableAttrs = "") => {
    const el = document.createElement("div");
    el.innerHTML = `<table ${tableAttrs}><tr><td ${attrs}>a</td></tr></table>`;
    return el.querySelector("td") as HTMLElement;
  };

  it("reads the padding Word writes on every report cell", () => {
    // 5.4pt at 96dpi is 7.2px; 0cm is a real zero, not "unstated".
    expect(extractCellPadding(cell("style=\"padding:0cm 5.4pt\""))).toBe("0px 7.2px");
  });

  it("keeps a zero rather than treating it as nothing said", () => {
    // The whole bug: a 0 that reads as absent falls back to the editor's own
    // 8px, which is exactly the height every pasted row was gaining.
    expect(extractCellPadding(cell("style=\"padding:0\""))).toBe("0px");
  });

  it("reads sides stated separately, and leaves the rest at the browser's 0", () => {
    expect(extractCellPadding(cell("style=\"padding-left:10px\""))).toBe("0px 0px 0px 10px");
  });

  it("falls back to the cellpadding attribute a legacy report carries", () => {
    expect(extractCellPadding(cell("", 'cellpadding="6"'))).toBe("6px");
  });

  it("lets the cell's own padding beat the table's attribute", () => {
    // Word writes both: `cellpadding=0` on the table AND the real padding on
    // each cell. The cell is what the browser renders, so it is what wins.
    expect(extractCellPadding(cell("style=\"padding:0cm 5.4pt\"", 'cellpadding="0"'))).toBe(
      "0px 7.2px"
    );
  });

  it("says nothing when the document says nothing", () => {
    expect(extractCellPadding(cell(""))).toBeUndefined();
  });

  it("survives paste and reaches the saved report", () => {
    const { table, html } = paste(
      `<table cellpadding=0 style='border-collapse:collapse'><tr>` +
        `<td style='width:100px;padding:0cm 5.4pt'>Rate</td>` +
        `<td style='width:100px;padding:0cm 5.4pt'>85 b/min</td>` +
        `</tr></table>`
    );
    expect(table?.children?.[0]?.children?.[0]?.padding).toBe("0px 7.2px");
    expect(html).toContain("padding: 0px 7.2px");
    // And nothing left over from the editor's old fixed box.
    expect(html).not.toContain("padding: 2px 5px");
  });

  it("round-trips: a saved report reopens with the padding it was saved with", () => {
    const once = paste(
      `<table><tr><td style='width:100px;padding:0cm 5.4pt'>Rate</td>` +
        `<td style='width:100px;padding:0cm 5.4pt'>85</td></tr></table>`
    );
    const twice = paste(once.html);
    expect(twice.table?.children?.[0]?.children?.[0]?.padding).toBe("0px 7.2px");
    expect(twice.html).toContain("padding: 0px 7.2px");
  });

  it("still writes the default for a table created in the editor", () => {
    const { html } = paste(`<table><tr><td>a</td><td>b</td></tr></table>`);
    expect(html).toContain("padding: 2px 5px");
  });
});

describe("a table's border DESIGN, not just whether it has one", () => {
  /** The cells of a table, as elements, so their borders can be read. */
  function cells(html: string): HTMLElement[] {
    const doc = new DOMParser().parseFromString(html, "text/html");
    return Array.from(doc.querySelectorAll("td, th"));
  }

  const borders = (html: string) => cells(html).map((c) => extractCellBorders(c));

  it("keeps the WIDTH and the STYLE, not only the fact of a line", () => {
    // The whole complaint: a 2.25pt double rule and a 1pt solid one used to
    // become the same 1px solid box, because the size was read as a yes/no and
    // the style was not read at all.
    const [cell] = borders(
      `<table style='border-collapse:collapse'>` +
        `<tr><td style='border:double windowtext 2.25pt'>a</td></tr></table>`
    );
    expect(cell?.top).toEqual({ size: 3, style: "double" });
    expect(borderShorthand(cell?.top)).toBe("3px double #000");
  });

  it("rounds a Word point width to whole pixels", () => {
    // 1.0pt is 1.333px, and a fractional rule prints unevenly.
    const [cell] = borders(
      `<table><tr><td style='border:solid windowtext 1.0pt'>a</td></tr></table>`
    );
    expect(cell?.top?.size).toBe(1);
    // 3.0pt is 4px.
    const [thick] = borders(
      `<table><tr><td style='border:solid windowtext 3.0pt'>a</td></tr></table>`
    );
    expect(thick?.top?.size).toBe(4);
  });

  it("lets the neighbour's declaration draw an edge the cell calls `none`", () => {
    // Exactly how Word writes a grid: the interior verticals live on the LEFT
    // cell's `border-right`, and the right cell says `border-left:none`.
    const [left, right] = borders(
      `<table style='border-collapse:collapse'><tr>` +
        `<td style='border:solid windowtext 1.0pt'>a</td>` +
        `<td style='border:solid windowtext 1.0pt;border-left:none'>b</td>` +
        `</tr></table>`
    );
    expect(left?.right?.size).toBe(1);
    expect(right?.left?.size).toBe(1);
  });

  it("gives a shared edge to the heavier of the two rules that meet there", () => {
    const [head, , body] = borders(
      `<table style='border-collapse:collapse'>` +
        `<tr><td style='border:solid windowtext 1.0pt;border-bottom:double windowtext 2.25pt'>H</td>` +
        `<td style='border:solid windowtext 1.0pt'>H2</td></tr>` +
        `<tr><td style='border:solid windowtext 1.0pt;border-top:none'>B</td>` +
        `<td style='border:solid windowtext 1.0pt'>B2</td></tr></table>`
    );
    expect(head?.bottom).toEqual({ size: 3, style: "double" });
    // The body cell shares that edge, so it carries the same rule.
    expect(body?.top).toEqual({ size: 3, style: "double" });
  });

  it("reads a frame stated on the TABLE and puts it on the outer edges only", () => {
    const [a, b, c, d] = borders(
      `<table style='border:3px double #333;border-collapse:collapse'>` +
        `<tr><td>a</td><td>b</td></tr><tr><td>c</td><td>d</td></tr></table>`
    );
    // The colour comes back out of the CSSOM in its own spelling.
    expect(a?.top).toEqual({ size: 3, style: "double", color: "rgb(51, 51, 51)" });
    expect(a?.left).toEqual({ size: 3, style: "double", color: "rgb(51, 51, 51)" });
    // Nothing inside the frame — the cells themselves declare no rules.
    expect(a?.right?.size).toBe(0);
    expect(a?.bottom?.size).toBe(0);
    expect(b?.right?.size).toBe(3);
    expect(c?.bottom?.size).toBe(3);
    expect(d?.top?.size).toBe(0);
  });

  it("reads a rule stated on the ROW", () => {
    const [a, b] = borders(
      `<table style='border-collapse:collapse'>` +
        `<tr style='border-bottom:2px solid red'><td>a</td><td>b</td></tr>` +
        `<tr><td>c</td><td>d</td></tr></table>`
    );
    expect(a?.bottom).toEqual({ size: 2, color: "red" });
    expect(b?.bottom).toEqual({ size: 2, color: "red" });
    expect(a?.top?.size).toBe(0);
  });

  it("keeps a borderless layout table borderless", () => {
    const [cell] = borders(
      `<table border=0 style='border-collapse:collapse;border:none'>` +
        `<tr><td style='border:none'>Rate</td><td style='border:none'>85</td></tr></table>`
    );
    expect(cell).toEqual({
      top: { size: 0 },
      right: { size: 0 },
      bottom: { size: 0 },
      left: { size: 0 },
    });
  });

  it("keeps a legacy report's grid, which is drawn by the border ATTRIBUTE alone", () => {
    const [cell] = borders(`<table border=1><tr><td>a</td><td>b</td></tr></table>`);
    expect(cell?.top?.size).toBe(1);
    expect(cell?.right?.size).toBe(1);
  });

  it("leaves a table that says nothing at all to Plate's own default", () => {
    // A table created in this editor, and the only case where inventing a 1px
    // line is the right answer rather than a guess.
    expect(borders(`<table><tr><td>a</td><td>b</td></tr></table>`)[0]).toBeUndefined();
  });

  it("treats a WIDTH with no style as a rule, the way the print pages do", () => {
    // A bare browser draws nothing for this, but some of Word's exports write
    // a cell's width and colour and leave the style off, and
    // `_partials/print_richtext_fixes` supplies the `solid` so those rules
    // print. They have to show on screen too.
    const [cell] = borders(
      `<table style='border-collapse:collapse'>` +
        `<tr><td style='border-width:3px'>a</td></tr></table>`
    );
    expect(cell?.top).toEqual({ size: 3 });
  });

  it("round-trips a border design through the saved report", () => {
    const source =
      `<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse'>` +
      `<tr><td style='border:solid windowtext 1.0pt;border-bottom:double windowtext 2.25pt;padding:0cm 5.4pt'>H</td></tr>` +
      `<tr><td style='border:solid windowtext 1.0pt;border-top:none;padding:0cm 5.4pt'>B</td></tr></table>`;
    const once = paste(source);
    expect(once.html).toContain("border-bottom: 3px double #000");

    // Reopened, the rule is unchanged — only the spelling of black settles,
    // from the `#000` this serializer writes to the `rgb(0, 0, 0)` the CSSOM
    // hands back for it. A third pass changes nothing at all, so a report does
    // not drift a little further every time it is saved.
    const twice = paste(once.html);
    expect(twice.table?.children?.[0]?.children?.[0]?.borders?.bottom).toEqual({
      size: 3,
      style: "double",
      color: "rgb(0, 0, 0)",
    });
    expect(twice.html).toContain("border-bottom: 3px double rgb(0, 0, 0)");
    expect(paste(twice.html).html).toBe(twice.html);
  });
});

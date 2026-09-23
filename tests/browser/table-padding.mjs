/**
 * Does a pasted table sit as tightly as it did in the old editor?
 *
 * The reference is the browser's OWN rendering of the same Word markup in a
 * clean iframe — and that is not an approximation of Summernote, it IS
 * Summernote: that editor never parsed a table, it dropped the clipboard HTML
 * into a contenteditable and let the browser draw `padding:0cm 5.4pt` as
 * written. So whatever the iframe measures is exactly what the old editor put
 * on screen.
 *
 * What was wrong: every cell here was rebuilt with a fixed `px-3 py-2` box —
 * 12px each side, 8px above AND below — against Word's 7.2px each side and
 * nothing vertically. 16px per row, and the eight-row table below stood ~128px
 * deeper than the document it was copied from.
 */
import { chromium } from 'playwright';
import { harnessHtml, serve } from './harness.mjs';

/** Kept in step with DEFAULT_PASTED_LINE_GAP in src/lib/line-gap.ts. */
const DEFAULT_PASTED_LINE_GAP = 1.2;

const ROWS = [
  ['Rate', '85 b/min'],
  ['Rhythm', 'Regular'],
  ['P-Wave', 'Normal'],
  ['P-R Interval', 'Normal'],
  ['QRS Complex', 'Normal'],
  ['ST. Segment', 'Isoelectric'],
  ['T. Wave', 'Normal'],
  ['Impression', 'Within normal limits'],
];

const cell = (inner, width) =>
  `<td width=${width} valign=top style='width:${(width * 0.75).toFixed(1)}pt;` +
  `border:solid windowtext 1.0pt;padding:0cm 5.4pt'><p class=MsoNormal>${inner}</p></td>`;

/** A real ECG report, exactly as Word puts it on the clipboard. */
const WORD =
  `<html xmlns:w="urn:schemas-microsoft-com:office:word"><head><style>` +
  `<!-- p.MsoNormal {margin:0cm; font-size:11.0pt; font-family:"Calibri",sans-serif;} --></style></head><body>` +
  `<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 width=624 ` +
  `style='width:468.0pt;border-collapse:collapse'>` +
  ROWS.map(([label, value]) => `<tr>${cell(label, 200)}${cell(value, 424)}</tr>`).join('') +
  `</table></body></html>`;

let fails = 0;
function check(label, ok, detail) {
  if (!ok) fails++;
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${label}${detail ? ` — ${detail}` : ''}`);
}

const { server, port } = await serve(() => harnessHtml(''));
const browser = await chromium.launch({ channel: 'chrome' });
const context = await browser.newContext({
  viewport: { width: 1440, height: 900 },
  permissions: ['clipboard-read', 'clipboard-write'],
});
const page = await context.newPage();
await page.goto(`http://127.0.0.1:${port}/`);
await page.evaluate(() => window.__ready);

/** What the old editor showed: the browser rendering the markup untouched. */
const truth = await page.evaluate(async (html) => {
  const frame = document.createElement('iframe');
  frame.style.cssText = 'position:absolute;left:-9999px;width:1200px;height:1400px';
  document.body.append(frame);
  frame.contentDocument.open();
  frame.contentDocument.write(html);
  frame.contentDocument.close();
  await new Promise((resolve) => setTimeout(resolve, 100));
  const doc = frame.contentDocument;
  const table = doc.querySelector('table');
  const first = table.querySelector('td');
  const style = frame.contentWindow.getComputedStyle(first);
  return {
    padding: [style.paddingTop, style.paddingRight, style.paddingBottom, style.paddingLeft],
    rowHeight: Math.round(table.rows[0].getBoundingClientRect().height),
    // Ratio, not pixels: the editor is wider than a Word page, so the table is
    // scaled to fit and only the proportions can be compared.
    tableHeightPerRow:
      Math.round((table.getBoundingClientRect().height / table.rows.length) * 10) / 10,
  };
}, WORD);

await page.click('[data-slate-editor]');
await page.evaluate(async (html) => {
  await navigator.clipboard.write([
    new ClipboardItem({
      'text/html': new Blob([html], { type: 'text/html' }),
      'text/plain': new Blob([''], { type: 'text/plain' }),
    }),
  ]);
}, WORD);
await page.keyboard.press('Control+V');
await page.waitForTimeout(500);

const got = await page.evaluate(() => {
  const editable = document.querySelector('[data-slate-editor]');
  const table = editable.querySelector('table');
  // The row-control column is chrome, not one of the document's cells.
  const cells = Array.from(table.querySelectorAll('td')).filter(
    (c) => !c.classList.contains('w-2')
  );
  // The padding lives on the cell's content box, which is the element the
  // padding is actually applied to — the <td> itself is `p-0` by design.
  const box = cells[0].querySelector(':scope > div');
  const style = getComputedStyle(box);
  // The paragraph inside the cell: its own padding and leading are the other
  // half of the height a pasted row was gaining.
  const block = getComputedStyle(box.querySelector('[data-slate-node="element"]'));
  const rows = Array.from(table.querySelectorAll('tr'));
  // The font's own line, which is what the block's leading is a multiple OF.
  //
  // Read off the RUN and not the block: Plate drops a `font-family` stated on a
  // `<p>` at parse, so the block computes to the editor's own UI stack (1.362)
  // while the text is the document's Calibri (1.221) — the gap is a multiple of
  // the font the reader sees.
  const run = box.querySelector('[data-slate-string]');
  const probe = document.createElement('div');
  probe.textContent = 'Hg';
  probe.style.cssText =
    `position:absolute;left:-9999px;font-size:1000px;line-height:normal;` +
    `font-family:${getComputedStyle(run ?? box).fontFamily}`;
  document.body.append(probe);
  const naturalLine = probe.getBoundingClientRect().height / 1000;
  probe.remove();

  return {
    padding: [style.paddingTop, style.paddingRight, style.paddingBottom, style.paddingLeft],
    blockPadding: block.paddingTop,
    blockLeading: block.lineHeight,
    blockFontSize: Number.parseFloat(block.fontSize),
    naturalLine,
    rowHeight: Math.round(rows[0].getBoundingClientRect().height),
    tableHeightPerRow: Math.round((table.getBoundingClientRect().height / rows.length) * 10) / 10,
    saved: window.RichTextEdit.getHtml('#report_body_editor'),
  };
});

check(
  'the cell is padded the way the document padded it',
  got.padding.join(' ') === truth.padding.join(' '),
  `word ${truth.padding.join(' ')} | editor ${got.padding.join(' ')}`
);
/*
 * What is left over is NOT the table's.
 *
 * The run inside a cell carries the document's own 11pt (15px), but the
 * paragraph BLOCK around it does not — it sits at the editor's 18px base, and
 * an empty 18px strut is taller than a 15px line. That is the `report` variant
 * in components/ui/editor.tsx, it applies to every paragraph in the document
 * and not just the ones in tables, and it is deliberately not this fix's
 * business. It costs about 5px a row.
 *
 * The bound below is what separates that from a regression: the cell's own
 * contribution — its padding, and the padding and leading of the block inside
 * it — was 16px a row and has to stay gone.
 */
const BASE_FONT_SLACK = 6;

check(
  'a row is no taller than the document, bar the editor base font',
  got.rowHeight - truth.rowHeight <= BASE_FONT_SLACK,
  `word ${truth.rowHeight}px | editor ${got.rowHeight}px`
);
check(
  'and so the whole table is not deeper than the one it was copied from',
  got.tableHeightPerRow - truth.tableHeightPerRow <= BASE_FONT_SLACK,
  `word ${truth.tableHeightPerRow}px/row | editor ${got.tableHeightPerRow}px/row`
);
/*
 * The leading is the DOCUMENT's default, not the editor's own.
 *
 * This used to read `normal`, because a cell that stated no spacing was left
 * stating none. A pasted block is now given `DEFAULT_PASTED_LINE_GAP` (1.2)
 * where the document states nothing — asked for on 2026-09-16, and chosen over
 * reading Word's silence as its "Single", with the ~4px a row that costs put to
 * the user first. See lib/word-line-gap.ts.
 *
 * What must still be gone is the EDITOR's leading: `leading-normal` is
 * Tailwind's 1.5 RATIO against an 18px base, which is what used to make a
 * pasted row taller than the document whatever its text was set in. So the
 * check is that the line is the default GAP of the block's own font — not that
 * it is any particular number of pixels.
 */
const expectedLeading = DEFAULT_PASTED_LINE_GAP * got.naturalLine * got.blockFontSize;
check(
  'the block inside a cell adds no padding, and the default gap rather than the editor’s leading',
  got.blockPadding === '0px' &&
    Math.abs(Number.parseFloat(got.blockLeading) - expectedLeading) <= 0.6,
  `padding ${got.blockPadding}, line-height ${got.blockLeading} ` +
    `(expected ${expectedLeading.toFixed(2)}px = gap ${DEFAULT_PASTED_LINE_GAP} of a ` +
    `${got.naturalLine.toFixed(4)} line at ${got.blockFontSize}px)`
);
check(
  'the saved report prints the same padding the editor drew',
  /<td[^>]*padding: 0px 7\.2px/.test(got.saved),
  (got.saved.match(/padding: [^;"]+/g) ?? []).slice(0, 2).join(' | ') || 'no padding written'
);

await browser.close();
server.close();
console.log(fails ? `\n${fails} check(s) failed` : '\nall checks passed');
process.exit(fails ? 1 : 0);

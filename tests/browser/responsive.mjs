/**
 * Does a pasted table reflow when the window gets small?
 *
 * Column widths are fixed at paste time, in pixels, against the editor as it
 * was then. Rendered as pixels they are a FLOOR: shrink the window and the
 * table keeps its width while the block around it grows a horizontal
 * scrollbar. Rendered as a share of the table they shrink with it.
 *
 * Only a real browser can answer this, and only by resizing it.
 */
import { chromium } from 'playwright';
import { harnessHtml, serve } from './harness.mjs';

const COLUMNS = ['Test', 'Result', 'Unit', 'Biological Reference Interval', 'Method', 'Remarks'];
const WIDTHS = [150, 80, 70, 160, 100, 64];

const cell = (text, width) =>
  `<td width=${width} style='width:${(width * 0.75).toFixed(1)}pt;border:solid windowtext 1.0pt;padding:0cm 5.4pt'>` +
  `<p><span style='font-size:9.0pt'>${text}</span></p></td>`;

const WORD = `<table border=1 cellspacing=0 cellpadding=0 width=624 style='width:468.0pt;border-collapse:collapse'>
 <tr>${COLUMNS.map((c, i) => cell(c, WIDTHS[i])).join('')}</tr>
 <tr>${['Haemoglobin', '13.5', 'gm/dl', '13.0 - 17.0', 'Photometric', 'Normal']
   .map((t, i) => cell(t, WIDTHS[i]))
   .join('')}</tr>
</table>`;

/** The share of the table each column should hold, whatever the window. */
const STATED_TOTAL = WIDTHS.reduce((a, b) => a + b, 0);
const SHARES = WIDTHS.map((w) => w / STATED_TOTAL);

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
await page.waitForTimeout(400);

const measure = () =>
  page.evaluate(() => {
    const editable = document.querySelector('[data-slate-editor]');
    const table = editable.querySelector('table');
    const block = table.closest('[data-slate-node="element"]');
    const cells = Array.from(editable.querySelectorAll('table tr:first-child td:not(.w-2)')).map(
      (c) => Math.round(c.getBoundingClientRect().width)
    );
    return {
      block: block.clientWidth,
      blockScrolls: block.scrollWidth > block.clientWidth,
      pageScrolls: document.documentElement.scrollWidth > document.documentElement.clientWidth,
      table: Math.round(table.getBoundingClientRect().width),
      cells,
    };
  });

// Full width first: the table must render at exactly the widths it was pasted
// with, or "responsive" has been bought by making the normal case wrong.
const wide = await measure();
check(
  'at full width the table fills the editor as before',
  wide.table > wide.block - 24 && !wide.blockScrolls,
  `table ${wide.table}px of block ${wide.block}px`
);

for (const width of [1024, 768, 500, 380]) {
  await page.setViewportSize({ width, height: 900 });
  await page.waitForTimeout(250);
  const got = await measure();
  const total = got.cells.reduce((a, b) => a + b, 0);

  check(
    `@${width}: the table shrinks to the window instead of scrolling`,
    !got.blockScrolls && !got.pageScrolls && got.table <= got.block,
    `table ${got.table}px of block ${got.block}px`
  );
  check(
    `@${width}: and every column keeps its share of it`,
    got.cells.every((cellWidth, i) => Math.abs(cellWidth / total - SHARES[i]) < 0.01),
    got.cells.join(' / ')
  );
}

// A cell that WRAPS keeps its row's borders.
//
// This is the shape a patient header table takes as the window narrows: a field
// that fitted on one line — "Delivery Date: 26-11-2025" — becomes two, the row
// grows, and every border in that row has to grow with it. A cell's rules are
// drawn on a `::before` overlay rather than on the cell itself, so nothing
// guarantees that on its own: an overlay that stayed the height the row used to
// be would draw the row's rules across the middle of it, which reads as a
// broken top border and a row that has lost its box.
{
  const header = (label, value) =>
    `<tr style='height:14.2pt'>` +
    [label, value]
      .map(
        (t) =>
          `<td style='border:solid windowtext 1.0pt;padding:0cm 5.4pt'>` +
          `<p><span style='font-size:11.0pt'>${t}</span></p></td>`
      )
      .join('') +
    `</tr>`;
  const HEADER =
    `<table border=1 cellspacing=0 width=624 style='width:468.0pt;border-collapse:collapse'>` +
    header('Inv. ID :', '3500000') +
    header('Delivery Date:', 'MR ABDUL KARIM CHOWDHURY 26-11-2025 10:45 AM') +
    header('Bed/ward:', 'ICU-05') +
    `</table>`;

  const wrapPage = await context.newPage();
  await wrapPage.goto(`http://127.0.0.1:${port}/`);
  await wrapPage.evaluate(() => window.__ready);
  await wrapPage.click('[data-slate-editor]');
  await wrapPage.evaluate(async (html) => {
    await navigator.clipboard.write([
      new ClipboardItem({
        'text/html': new Blob([html], { type: 'text/html' }),
        'text/plain': new Blob(['x'], { type: 'text/plain' }),
      }),
    ]);
  }, HEADER);
  await wrapPage.keyboard.press('Control+V');
  await wrapPage.waitForTimeout(700);

  /**
   * Every cell: its row's height, and the box its rules are actually drawn on.
   *
   * `offset` is the half that broke the grid. The rules live on an overlay, and
   * an absolutely positioned overlay with no offsets falls back to its STATIC
   * position — which, now that a cell's content is vertically centred, is
   * halfway down a cell whose neighbour has wrapped. So in a row of one-line
   * fields beside a two-line one, every one-line cell drew its box 10px low and
   * 10px past the bottom of the row, and the header table came apart.
   */
  const rules = () =>
    wrapPage.evaluate(() =>
      [...document.querySelectorAll('[data-slate-editor] tr')].flatMap((row) => {
        const height = Math.round(row.getBoundingClientRect().height);
        return [...row.children]
          .filter((c) => /^t[dh]$/i.test(c.tagName) && !c.classList.contains('w-2'))
          .map((cell) => {
            const drawn = getComputedStyle(cell, '::before');
            return {
              row: height,
              drawn: Math.round(parseFloat(drawn.height) || 0),
              offset: Math.round(parseFloat(drawn.top) || 0),
            };
          });
      })
    );

  for (const width of [1280, 560, 380, 320]) {
    // The reported case needs no narrow window: one long field in a row of
    // short ones is enough for the row to be two lines deep while its other
    // cells are one.

    await wrapPage.setViewportSize({ width, height: 900 });
    await wrapPage.waitForTimeout(300);
    const cells = await rules();
    const heights = [...new Set(cells.map((c) => c.row))];
    check(
      `@${width}: every cell's rules are drawn the full height of its row`,
      cells.length > 0 && cells.every((c) => Math.abs(c.drawn - c.row) <= 1),
      cells.map((c) => `${c.row}/${c.drawn}`).join(' ')
    );
    check(
      `@${width}: and from the top of the cell, so the grid still lines up`,
      cells.every((c) => c.offset === 0),
      cells.map((c) => `+${c.offset}`).join(' ')
    );
    if (width <= 380) {
      check(
        `@${width}: and a field that no longer fits really has wrapped`,
        heights.length > 1,
        `row heights ${heights.join(', ')}`
      );
    }
  }

  await wrapPage.close();
}

await browser.close();
server.close();
console.log(fails ? `\n${fails} check(s) failed` : '\nall checks passed');
process.exit(fails ? 1 : 0);

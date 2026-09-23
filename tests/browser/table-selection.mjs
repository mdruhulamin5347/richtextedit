/**
 * Can you select text in a table and see that you have?
 *
 * The selection always WORKED — the range was there, the right text was in it,
 * Ctrl+C copied it. It just never looked selected: the table carried
 * `selection:bg-transparent`, so `::selection` painted nothing and dragging
 * across a row left the text looking untouched, which is indistinguishable from
 * being unselectable. Upstream can afford that class because a cell-selection
 * overlay paints the highlight instead; this editor does not register
 * BlockSelectionPlugin, so nothing did.
 *
 * Needs a real browser twice over: `::selection` is a pseudo-element, and the
 * selection itself only exists once something has been dragged across a layout.
 */
import { chromium } from 'playwright';
import { harnessHtml, serve } from './harness.mjs';

const TABLE = `<table border=1 style='border-collapse:collapse'>
 <tr><td style='width:200px'><p>Neisseria gonorrhoeae</p></td><td style='width:200px'><p>Not Detected</p></td><td style='width:200px'><p>200 CFU</p></td></tr>
 <tr><td style='width:200px'><p>Chlamydia trachomatis</p></td><td style='width:200px'><p>Detected</p></td><td style='width:200px'><p>400 CFU</p></td></tr>
</table><p>After the table</p>`;

/**
 * Well clear of the column-resize handle, which straddles a cell's left edge by
 * 4px each way — a drag started on it resizes the column instead of selecting,
 * and reads as "selection is broken".
 */
const INSET_PX = 20;

let fails = 0;
function check(label, ok, detail) {
  if (!ok) fails++;
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${label}${detail ? ` — ${detail}` : ''}`);
}

const { server, port } = await serve(() => harnessHtml(TABLE));
const browser = await chromium.launch({ channel: 'chrome' });
const context = await browser.newContext({
  viewport: { width: 1280, height: 900 },
  permissions: ['clipboard-read', 'clipboard-write'],
});

/**
 * One drag, on a page of its own.
 *
 * Deliberately not several drags on one page: a previous selection leaves the
 * editor with state that anchors the next one at a block boundary, so the
 * second measurement stops describing what a user dragging once would get.
 */
async function dragBetween(from, to) {
  const page = await context.newPage();
  await page.goto(`http://127.0.0.1:${port}/`);
  await page.evaluate(() => window.__ready);

  const cells = await page.evaluate(() =>
    Array.from(document.querySelectorAll('[data-slate-editor] table td:not(.w-2)')).map((cell) => {
      const box = cell.getBoundingClientRect();
      return {
        left: Math.round(box.left),
        right: Math.round(box.right),
        y: Math.round(box.top + box.height / 2),
      };
    })
  );

  await page.mouse.move(cells[from].left + INSET_PX, cells[from].y);
  await page.mouse.down();
  // To the far side of the target cell, not a fixed 60px in.
  //
  // A cell's text is vertically CENTRED now, as it is in a browser, so a drag
  // taken across the middle of a row lands ON the text rather than below it —
  // and a fixed offset therefore stopped mid-word instead of running past the
  // end of the line.
  await page.mouse.move(cells[to].right - INSET_PX, cells[to].y, { steps: 15 });
  await page.mouse.up();
  await page.waitForTimeout(150);

  const selection = await page.evaluate(() => {
    const selected = window.getSelection();
    return {
      text: selected.toString().replace(/\s+/g, ' ').trim(),
      ranges: selected.rangeCount,
    };
  });

  const highlight = await page.evaluate(() => {
    const cell = document.querySelector('[data-slate-editor] table td:not(.w-2)');
    const paragraph = document.querySelectorAll('[data-slate-editor] > div')[1];
    return {
      cell: getComputedStyle(cell, '::selection').backgroundColor,
      paragraph: getComputedStyle(paragraph, '::selection').backgroundColor,
    };
  });

  await page.keyboard.press('Control+C');
  await page.waitForTimeout(200);
  const clipboard = await page.evaluate(async () =>
    (await navigator.clipboard.readText()).replace(/\s+/g, ' ').trim()
  );

  await page.close();
  return { ...selection, highlight, clipboard };
}

// The highlight, which is the thing that was actually missing.
const withinCell = await dragBetween(0, 0);
check(
  'selected text in a cell is actually tinted',
  !/rgba\(0, 0, 0, 0\)|transparent/.test(withinCell.highlight.cell),
  withinCell.highlight.cell
);
check(
  'and tinted the same as everywhere else in the editor',
  withinCell.highlight.cell === withinCell.highlight.paragraph,
  `cell ${withinCell.highlight.cell} | paragraph ${withinCell.highlight.paragraph}`
);
check('a drag inside one cell selects that cell’s text', withinCell.ranges === 1, withinCell.text);
check(
  'and Ctrl+C copies it',
  withinCell.clipboard.length > 0 && withinCell.text.includes(withinCell.clipboard),
  withinCell.clipboard
);

const row = await dragBetween(0, 2);
check(
  'a drag across a row selects the whole row',
  row.ranges === 1 && row.text.includes('Not Detected') && row.text.includes('200 CFU'),
  row.text
);
check(
  'and copies the whole row',
  row.clipboard.includes('Not Detected') && row.clipboard.includes('200 CFU'),
  row.clipboard
);

const column = await dragBetween(1, 4);
check(
  'a drag down a column selects the whole column',
  column.ranges === 1 && column.text.includes('Detected'),
  column.text
);
check('and copies it', column.clipboard.includes('Detected'), column.clipboard);

await browser.close();
server.close();
console.log(fails ? `\n${fails} check(s) failed` : '\nall checks passed');
process.exit(fails ? 1 : 0);

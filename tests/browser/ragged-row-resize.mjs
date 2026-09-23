/**
 * Can a column that exists in only SOME rows be resized?
 *
 * A report table is often ragged: ten `label | : | value` rows, and two rows in
 * the middle broken into five cells for a measurement pair. Those two rows'
 * extra columns could not be dragged at all, while the three columns every row
 * shares resized normally.
 *
 * Run: node tests/browser/ragged-row-resize.mjs
 */
import { chromium } from 'playwright';
import { harnessHtml, serve } from './harness.mjs';

const border = 'border:1px solid #000';
const wide = (label, value) =>
  `<tr><td style="${border}">${label}</td><td style="${border}">:</td><td style="${border}">${value}</td></tr>`;
/** A row broken into five cells — the two extra columns exist only here. */
const split = (a, b, c, d, e) =>
  `<tr>${[a, b, c, d, e].map((t) => `<td style="${border}">${t}</td>`).join('')}</tr>`;

const TABLE =
  `<table border="1" style="border-collapse: collapse;">` +
  [
    wide('Uterus', 'Gravid'),
    wide('No of fetus', 'Single'),
    split('BPD', ':', '3.61 cm', 'FL', '1.98 cm'),
    wide('Placenta', 'Posterior'),
    split('AC', ':', '12.4 cm', 'HC', '15.1 cm'),
    wide('Liquor', 'Adequate'),
    wide('Presentation', 'Cephalic'),
    wide('Cardiac activity', 'Present'),
    wide('Movement', 'Present'),
    wide('EDD', '15/06/26'),
  ].join('') +
  `</table>`;

const DRAG_PX = 60;
let fails = 0;
const check = (name, pass, detail = '') => {
  if (!pass) fails++;
  console.log(`${pass ? 'PASS' : 'FAIL'}  ${name}${detail ? ' — ' + detail : ''}`);
};

const { server, port } = await serve(() => harnessHtml(TABLE));
const browser = await chromium.launch({ channel: 'chrome' });
const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
await page.goto(`http://127.0.0.1:${port}/`);
await page.evaluate(() => window.__ready);
await page.waitForTimeout(400);

/** colSizes as the serializer writes them back out: the <colgroup>. */
const readColSizes = () => page.evaluate(() => {
  const html = window.RichTextEdit.getHtml('#report_body_editor');
  const m = html.match(/<colgroup>([\s\S]*?)<\/colgroup>/);
  return m ? [...m[1].matchAll(/width:\s*([\d.]+)px/g)].map((x) => Math.round(+x[1])) : null;
});
const colSizes = await readColSizes();
console.log('   colSizes on the table node:', JSON.stringify(colSizes));

// The widest row is the grid: 5 columns, so 5 sizes.
check('colSizes covers every column in the table, not just the first row\'s',
  Array.isArray(colSizes) && colSizes.length === 5,
  `length ${colSizes?.length} (want 5)`);

const splitRow = page.locator('[data-slate-editor] table tr').nth(2);
const splitCells = splitRow.locator('td:not(.w-2)');
check('the split row has five cells', (await splitCells.count()) === 5, `got ${await splitCells.count()}`);

const widths = async () => splitCells.evaluateAll((els) => els.map((e) => Math.round(e.getBoundingClientRect().width)));

/** Drag the right-hand handle of column `col` and report what moved. */
async function drag(col) {
  const before = await widths();
  const handle = splitRow.locator(`[data-col="${col}"]`).first();
  if (!(await handle.count())) return { before, after: before, handle: false };
  const box = await handle.boundingBox();
  if (!box) return { before, after: before, handle: false };
  console.log(`   col ${col} handle box: ${Math.round(box.width)}x${Math.round(box.height)} at ${Math.round(box.x)},${Math.round(box.y)}`);
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
  await page.mouse.down();
  for (let i = 1; i <= 6; i++) {
    await page.mouse.move(box.x + box.width / 2 + (i * DRAG_PX) / 6, box.y + box.height / 2);
    await page.waitForTimeout(20);
  }
  await page.mouse.up();
  await page.waitForTimeout(250);
  return { before, after: await widths(), handle: true };
}

for (const col of [0, 3]) {
  const { before, after, handle } = await drag(col);
  const moved = Math.abs(after[col] - before[col]);
  console.log(`   col ${col}: ${before.join('/')}  ->  ${after.join('/')}   colSizes ${JSON.stringify(await readColSizes())}`);
  check(`column ${col} has a resize handle`, handle);
  check(`dragging column ${col} actually resizes it`, moved >= DRAG_PX - 8,
    `moved ${moved}px of ${DRAG_PX} (${before[col]} -> ${after[col]})`);
}

await browser.close();
server.close();
console.log(fails ? `\n${fails} check(s) failed` : '\nAll checks passed');
process.exit(fails ? 1 : 0);

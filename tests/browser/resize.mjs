/**
 * Does dragging a column border actually resize the column?
 *
 * Run: node tests/browser/resize.mjs
 */
import { chromium } from 'playwright';
import { harnessHtml, serve } from './harness.mjs';

/*
 * Plate resizes a PAIR of columns, holding their combined width constant, and
 * will not take a column below its 48px minimum. So the neighbour needs room to
 * give — dragging against a column that is already at the minimum can only move
 * the boundary the other way, which is correct behaviour and not a bug.
 */
const TABLE = `<table border="1" style="border-collapse: collapse;">
<colgroup><col width="150"><col width="200"><col width="300"></colgroup>
<tr><td style="border:1px solid #000">Rate</td><td style="border:1px solid #000">:</td><td style="border:1px solid #000">85 b/min</td></tr>
<tr><td style="border:1px solid #000">Rhythm</td><td style="border:1px solid #000">:</td><td style="border:1px solid #000">Regular</td></tr>
</table>`;

const DRAG_PX = 60;

const results = [];
const check = (name, pass, detail = '') => {
    results.push({ name, pass, detail });
    console.log(`${pass ? 'PASS' : 'FAIL'}  ${name}${detail ? ' — ' + detail : ''}`);
};

const { server, port } = await serve(() => harnessHtml(TABLE));
const browser = await chromium.launch({ channel: 'chrome' });
const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });

const errors = [];
page.on('console', (m) => m.type() === 'error' && errors.push(m.text()));
page.on('pageerror', (e) => errors.push(String(e)));

await page.goto(`http://127.0.0.1:${port}/`);
await page.evaluate(() => window.__ready);
// A table pasted without widths measures itself and switches to fixed layout on
// the frame after mount; let that settle before measuring anything.
await page.waitForTimeout(300);

check('editor mounted', (await page.locator('[data-slate-editor]').count()) > 0);
check('table rendered', (await page.locator('[data-slate-editor] table').count()) > 0);
// The first <td> is the upstream editor's row drag-handle cell, not a content column.
const cells = page.locator('[data-slate-editor] table tr').first().locator('td:not(.w-2)');
check('three content columns', (await cells.count()) === 3, `got ${await cells.count()}`);

const widths = async () =>
    (await cells.evaluateAll((els) => els.map((e) => Math.round(e.getBoundingClientRect().width))));
const before = await widths();
console.log('   widths before:', before.join(' / '));

// The right-hand handle of column 0.
const handle = page.locator('[data-col="0"]').first();
check('resize handle present in the DOM', (await handle.count()) > 0);

if (await handle.count()) {
    const box = await handle.boundingBox();
    check('resize handle has a hit area', !!box && box.width > 0 && box.height > 0,
        box ? `${Math.round(box.width)}x${Math.round(box.height)}` : 'no box');

    const cursor = await handle.evaluate((el) => getComputedStyle(el).cursor);
    check('handle shows the col-resize cursor', cursor === 'col-resize', `cursor: ${cursor}`);

    if (box) {
        await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
        await page.mouse.down();
        // Several small steps: a single jump can be ignored as noise.
        for (let i = 1; i <= 6; i++) {
            await page.mouse.move(box.x + box.width / 2 + (i * DRAG_PX) / 6, box.y + box.height / 2);
            await page.waitForTimeout(20);
        }
        await page.mouse.up();
        await page.waitForTimeout(200);
    }
}

const after = await widths();
console.log('   widths after: ', after.join(' / '));
check('column 0 grew by the drag distance', Math.abs(after[0] - (before[0] + DRAG_PX)) <= 4,
    `${before[0]}px -> ${after[0]}px (expected ~${before[0] + DRAG_PX})`);
check('its neighbour gave up exactly that much', Math.abs(after[1] - (before[1] - DRAG_PX)) <= 4,
    `${before[1]}px -> ${after[1]}px (expected ~${before[1] - DRAG_PX})`);
check('the other column is untouched', after[2] === before[2], `${before[2]}px -> ${after[2]}px`);

const colSizes = await page.evaluate(() => {
    const ta = document.getElementById('report_body');
    return ta ? ta.value.includes('<colgroup>') : false;
});
check('the new width is persisted to the textarea', colSizes);

if (errors.length) console.log('\nBROWSER ERRORS:\n  ' + errors.join('\n  '));

await browser.close();
server.close();

const failed = results.filter((r) => !r.pass);
console.log(`\n${results.length - failed.length}/${results.length} checks passed`);
process.exit(failed.length ? 1 : 0);

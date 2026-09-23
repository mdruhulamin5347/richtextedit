/**
 * Does a pasted Word table fill the editor?
 *
 * Word states a table at the width of a PAGE — the 6.5in text column of a
 * Letter page is 624px — while this editor is as wide as the admin card.
 * Reproducing those numbers literally left a pasted report sitting in the left
 * half of the screen with the right half blank, even though the serializer
 * emits every table as `width: 100%` and each print stretches it across the
 * page. Only a real browser can show that: jsdom does no layout.
 */
import { chromium } from 'playwright';
import { harnessHtml, serve } from './harness.mjs';

const WORD_TABLE = `<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width=624 style='width:468.0pt;border-collapse:collapse'>
 <tr><td width=170 style='width:127.5pt'><p>Haemoglobin</p></td><td width=20 style='width:15.0pt'><p>:</p></td><td width=120 style='width:90.0pt'><p>13.5</p></td><td width=120 style='width:90.0pt'><p>gm/dl</p></td><td width=194 style='width:145.5pt'><p>13.0 - 17.0</p></td></tr>
 <tr><td width=170 style='width:127.5pt'><p>ESR</p></td><td width=20 style='width:15.0pt'><p>:</p></td><td width=120 style='width:90.0pt'><p>12</p></td><td width=120 style='width:90.0pt'><p>mm/1st hr</p></td><td width=194 style='width:145.5pt'><p>0 - 15</p></td></tr>
</table>`;

/** The same table with no width stated anywhere — Word writes plenty of these. */
const NO_WIDTHS = `<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse'>
 <tr><td><p>Test</p></td><td><p>Result</p></td><td><p>Unit</p></td></tr>
 <tr><td><p>Haemoglobin</p></td><td><p>13.5</p></td><td><p>gm/dl</p></td></tr>
</table>`;

const STATED = [170, 20, 120, 120, 194];

let fails = 0;
function check(label, ok, detail) {
  if (!ok) fails++;
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${label}${detail ? ` — ${detail}` : ''}`);
}

/** Paste `html` at `viewport` width and measure what the editor did with it. */
async function pasteInto(browser, html, viewport) {
  const { server, port } = await serve(() => harnessHtml(''));
  const context = await browser.newContext({
    viewport: { width: viewport, height: 900 },
    permissions: ['clipboard-read', 'clipboard-write'],
  });
  const page = await context.newPage();
  await page.goto(`http://127.0.0.1:${port}/`);
  await page.evaluate(() => window.__ready);
  await page.click('[data-slate-editor]');
  await page.evaluate(async (source) => {
    await navigator.clipboard.write([
      new ClipboardItem({
        'text/html': new Blob([source], { type: 'text/html' }),
        'text/plain': new Blob([''], { type: 'text/plain' }),
      }),
    ]);
  }, html);
  await page.keyboard.press('Control+V');
  await page.waitForTimeout(400);

  const measured = await page.evaluate(() => {
    const editable = document.querySelector('[data-slate-editor]');
    const table = editable.querySelector('table');
    const block = table.closest('[data-slate-node="element"]');
    return {
      block: block.clientWidth,
      scrollWidth: block.scrollWidth,
      cells: Array.from(editable.querySelectorAll('table tr:first-child td:not(.w-2)')).map((c) =>
        Math.round(c.getBoundingClientRect().width)
      ),
    };
  });

  await page.close();
  await context.close();
  server.close();
  return measured;
}

const browser = await chromium.launch({ channel: 'chrome' });

for (const viewport of [1440, 1024]) {
  const { block, scrollWidth, cells } = await pasteInto(browser, WORD_TABLE, viewport);
  const total = cells.reduce((a, b) => a + b, 0);

  check(
    `@${viewport}: the pasted table spans the editor, not half of it`,
    total > block - 24,
    `${total}px of ${block}px`
  );
  check(`@${viewport}: nothing overflows into a scrollbar`, scrollWidth <= block, `scrollWidth ${scrollWidth}`);

  const statedTotal = STATED.reduce((a, b) => a + b, 0);
  const held = cells.every((width, i) => Math.abs(width / total - STATED[i] / statedTotal) < 0.01);
  check(`@${viewport}: the columns keep Word's proportions`, held, cells.join(' / '));
}

const wide = await pasteInto(browser, NO_WIDTHS, 1440);
const wideTotal = wide.cells.reduce((a, b) => a + b, 0);
check(
  'a table stating no widths at all also fills the editor',
  wideTotal > wide.block - 24,
  `${wideTotal}px of ${wide.block}px`
);
check(
  'and splits it evenly',
  Math.max(...wide.cells) - Math.min(...wide.cells) <= 1,
  wide.cells.join(' / ')
);

await browser.close();
console.log(fails ? `\n${fails} check(s) failed` : `\nall checks passed`);
process.exit(fails ? 1 : 0);

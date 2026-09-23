/**
 * Are a pasted table's column widths reproduced exactly?
 */
import { chromium } from 'playwright';
import { harnessHtml, serve } from './harness.mjs';

const cases = [
  { name: 'colgroup px 100/40/300',
    html: `<table><colgroup><col width="100"><col width="40"><col width="300"></colgroup><tr><td>Rate</td><td>:</td><td>85 b/min</td></tr></table>`,
    expect: [100, 40, 300] },
  { name: 'Word pt 85/21/212  (=113/28/283px)',
    html: `<table border=1><tr><td style='width:85.0pt'>Rate</td><td style='width:21.0pt'>:</td><td style='width:212.0pt'>85 b/min</td></tr></table>`,
    expect: [113, 28, 283] },
  { name: 'td px 10/20/60 (extreme)',
    html: `<table><tr><td style="width:10px">A</td><td style="width:20px">B</td><td style="width:60px">C</td></tr></table>`,
    expect: [10, 20, 60] },
  { name: 'EQUAL widths stated 200/200/200',
    html: `<table><colgroup><col width="200"><col width="200"><col width="200"></colgroup><tr><td>Rate</td><td>:</td><td>85 b/min</td></tr></table>`,
    expect: [200, 200, 200] },
  { name: 'percent 20/10/70',
    html: `<table><tr><td style="width:20%">Rate</td><td style="width:10%">:</td><td style="width:70%">85 b/min</td></tr></table>`,
    expect: null },
];

const browser = await chromium.launch({ channel: 'chrome' });
let fails = 0;
for (const c of cases) {
  const { server, port } = await serve(() => harnessHtml(c.html));
  const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
  await page.goto(`http://127.0.0.1:${port}/`);
  await page.evaluate(() => window.__ready);
  const got = await page.evaluate(() => {
    const cells = document.querySelectorAll('[data-slate-editor] table tr:first-child td:not(.w-2)');
    return Array.from(cells).map((c) => Math.round(c.getBoundingClientRect().width));
  });
  const colSizes = await page.evaluate(() => {
    const m = window.RadiolensEditor.getHtml('#report_body_editor').match(/width: (\d+)px/g);
    return m ? m.map((x) => parseInt(x.replace(/\D/g, ''), 10)) : null;
  });
  const ok = c.expect ? JSON.stringify(got) === JSON.stringify(c.expect) : true;
  if (!ok) fails++;
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${c.name}`);
  console.log(`        rendered: ${got.join(' / ')}   expected: ${c.expect ? c.expect.join(' / ') : '(proportional)'}`);
  console.log(`        saved colgroup: ${colSizes ? colSizes.join(' / ') : 'none'}`);
  await page.close(); server.close();
}
await browser.close();
console.log(fails ? `\n${fails} case(s) did not match exactly` : '\nall cases exact');

/**
 * Does a pasted table look like the document it came from?
 *
 * Checked against the browser's OWN rendering of the same Word markup, in a
 * clean iframe — that is the only honest reference, and jsdom cannot give it
 * because none of this is visible without layout.
 *
 * Three things this pins down, each of which was wrong:
 *  - the RULE COLOUR. The cell borders were drawn in the editor's theme grey
 *    while the saved report had them black, so the editor and the print
 *    disagreed about the same table.
 *  - alignment stated on the CELL. Word writes `<td align=center>` around a
 *    plain paragraph as readily as it writes it on the paragraph, and only the
 *    paragraph was ever read — so centred header rows came back left.
 *  - the GAP under the table, which was 20px of editor chrome against Word's 0.
 */
import { chromium } from 'playwright';
import { harnessHtml, serve } from './harness.mjs';

const cell = (inner, width, attrs = '') =>
  `<td width=${width} valign=top ${attrs} style='width:${(width * 0.75).toFixed(1)}pt;border:solid windowtext 1.0pt;padding:0cm 5.4pt'>${inner}</td>`;

/** A centred header row written on the CELL, and a data row centred on the P. */
const WORD = `<html xmlns:w="urn:schemas-microsoft-com:office:word"><head><style>
<!-- p.MsoNormal {margin:0cm; line-height:107%; font-size:11.0pt; font-family:"Calibri",sans-serif;} --></style></head><body>
<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 width=624 style='width:468.0pt;border-collapse:collapse'>
 <tr>
  ${cell('<p class=MsoNormal><b>Pathogen Type</b></p>', 130, 'align=center')}
  ${cell('<p class=MsoNormal><b>Tests</b></p>', 320, 'align=center')}
  ${cell('<p class=MsoNormal><b>Result</b></p>', 174, 'align=center')}
 </tr>
 <tr>
  ${cell('<p class=MsoNormal>Bacteria</p>', 130, 'rowspan=2')}
  ${cell('<p class=MsoNormal>Neisseria gonorrhoeae</p>', 320)}
  ${cell("<p class=MsoNormal align=center style='text-align:center'>Not Detected</p>", 174)}
 </tr>
 <tr>
  ${cell('<p class=MsoNormal>Chlamydia trachomatis</p>', 320)}
  ${cell("<p class=MsoNormal align=center style='text-align:center'>Detected</p>", 174)}
 </tr>
</table>
<p class=MsoNormal>Limit of Detection: 200 CFU/mL</p></body></html>`;

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

const truth = await page.evaluate(async (html) => {
  const frame = document.createElement('iframe');
  frame.style.cssText = 'position:absolute;left:-9999px;width:1200px;height:900px';
  document.body.append(frame);
  frame.contentDocument.open();
  frame.contentDocument.write(html);
  frame.contentDocument.close();
  await new Promise((resolve) => setTimeout(resolve, 100));
  const doc = frame.contentDocument;
  const table = doc.querySelector('table');
  const cells = Array.from(table.querySelectorAll('td'));
  return {
    // -webkit-center is what a td-level `align` computes to; the text is centred.
    centred: cells.map((c) =>
      /center/.test(frame.contentWindow.getComputedStyle(c.querySelector('p')).textAlign)
    ),
    borderColor: frame.contentWindow.getComputedStyle(cells[0]).borderTopColor,
    rowspans: cells.map((c) => c.rowSpan).join(','),
    gapUnder: Math.round(
      doc.querySelector('table + p').getBoundingClientRect().top -
        table.getBoundingClientRect().bottom
    ),
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
  const cells = Array.from(table.querySelectorAll('td')).filter(
    (c) => !c.classList.contains('w-2')
  );
  const block = table.closest('[data-slate-node="element"]');
  const saved = window.RichTextEdit.getHtml('#report_body_editor');
  return {
    centred: cells.map((c) =>
      /center/.test(
        getComputedStyle(c.querySelector('[data-slate-node="element"]') ?? c).textAlign
      )
    ),
    // The rule is painted on the cell's ::before, not on the <td> itself.
    borderColor: getComputedStyle(cells[0], '::before').borderTopColor,
    borderWidth: getComputedStyle(cells[0], '::before').borderTopWidth,
    rowspans: cells.map((c) => c.rowSpan).join(','),
    gapUnder: block.nextElementSibling
      ? Math.round(
          block.nextElementSibling.getBoundingClientRect().top -
            table.getBoundingClientRect().bottom
        )
      : null,
    savedCentredCells: (saved.match(/<td[^>]*text-align: center/g) ?? []).length,
  };
});

check(
  'the rules are the colour the document drew them, not the theme grey',
  got.borderColor === truth.borderColor && got.borderWidth === '1px',
  `word ${truth.borderColor} | editor ${got.borderColor} ${got.borderWidth}`
);
check(
  'every cell the document centred is still centred',
  got.centred.length === truth.centred.length &&
    got.centred.every((centred, i) => centred === truth.centred[i]),
  `word ${truth.centred.map(Number).join('')} | editor ${got.centred.map(Number).join('')}`
);
check(
  'alignment stated on the cell reaches the saved report too',
  got.savedCentredCells === 3,
  `${got.savedCentredCells} of 3 header cells`
);
check('merged cells keep their span', got.rowspans === truth.rowspans, got.rowspans);
check(
  'the gap under the table is the handles’ clearance, not 20px of chrome',
  got.gapUnder !== null && got.gapUnder <= 6,
  `word ${truth.gapUnder}px | editor ${got.gapUnder}px`
);

await browser.close();
server.close();
console.log(fails ? `\n${fails} check(s) failed` : '\nall checks passed');
process.exit(fails ? 1 : 0);

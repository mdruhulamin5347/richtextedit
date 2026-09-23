/**
 * Is the text the size the document set it in?
 *
 * Reported as "the table header is perfect but the table body is large after
 * paste", and that is exactly the shape of it: Word states a header's size on
 * the span (it is bold and bigger, so it has a run of its own) and leaves the
 * body to `p.MsoNormal {font-size:11.0pt}`. A size stated on a BLOCK inherits
 * in CSS but reaches no Plate node — only a text leaf carries one — so the body
 * fell back to the editable's 18px base and came out bigger than the header's
 * own text.
 *
 * The reference is the browser's own rendering of the same markup in a clean
 * iframe, which is what the old editor showed and what the printed report
 * shows. Compared EXACTLY, to a hundredth of a pixel: a stated point size is
 * kept in points (see lib/font-size.ts), so the editor hands the browser the
 * same value the iframe got and gets back the same fraction. The tolerance used
 * to be a whole pixel, because every size was converted to whole px first — an
 * 11pt run was 14.67px in the iframe and 15px here, and a 10pt one was named 13
 * by the font-size control.
 */
import { chromium } from 'playwright';
import { harnessHtml, serve } from './harness.mjs';

/** Word's own markup for a results table, header sized on the span, body not. */
const WORD =
  `<html xmlns:w="urn:schemas-microsoft-com:office:word"><head><style>` +
  `<!-- p.MsoNormal {margin:0cm; font-size:11.0pt; font-family:"Times New Roman",serif;} --></style></head><body>` +
  `<table class=MsoTableGrid border=1 cellspacing=0 style='border-collapse:collapse'>` +
  `<tr><td style='border:solid windowtext 1.0pt;padding:0cm 5.4pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:14.0pt'>Pathogen Name</span></b></p></td>` +
  `<td style='border:solid windowtext 1.0pt;padding:0cm 5.4pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:14.0pt'>Result</span></b></p></td></tr>` +
  `<tr><td style='border:solid windowtext 1.0pt;padding:0cm 5.4pt'><p class=MsoNormal>Measles morbillivirus</p></td>` +
  `<td style='border:solid windowtext 1.0pt;padding:0cm 5.4pt'><p class=MsoNormal align=center style='text-align:center'>Detected</p></td></tr>` +
  `<tr><td style='border:solid windowtext 1.0pt;padding:0cm 5.4pt'><p class=MsoNormal><span style='font-size:9.0pt'>Limit of detection</span></p></td>` +
  `<td style='border:solid windowtext 1.0pt;padding:0cm 5.4pt'><p class=MsoNormal><span style='font-size:9.0pt'>200 CFU/mL</span></p></td></tr>` +
  `</table><p class=MsoNormal>Comment: within normal limits.</p></body></html>`;

/**
 * The same table with the body's size only in the `<style>` block — Word's own
 * default table style, which is where a document that names no size on the
 * paragraph gets one. It is not an inline style until JuicePlugin has run, and
 * `transformData` is piped in REVERSE registration order, so this is the case
 * that catches a pass registered in the wrong place.
 */
const WORD_TABLE_STYLE =
  `<html xmlns:w="urn:schemas-microsoft-com:office:word"><head><style>` +
  `<!-- table.MsoNormalTable {mso-style-name:"Table Normal"; font-size:10.0pt;} p.MsoNormal {margin:0cm;} --></style></head><body>` +
  `<table class=MsoNormalTable border=1 cellspacing=0 style='border-collapse:collapse'>` +
  `<tr><td style='border:solid windowtext 1.0pt;padding:0cm 5.4pt'><p class=MsoNormal><b><span style='font-size:14.0pt'>Pathogen Name</span></b></p></td></tr>` +
  `<tr><td style='border:solid windowtext 1.0pt;padding:0cm 5.4pt'><p class=MsoNormal>Measles morbillivirus</p></td></tr>` +
  `</table></body></html>`;

/** A stored report body, which states its sizes on the blocks a browser inherits from. */
const STORED =
  `<p style="font-size: 20px">LABORATORY REPORT</p>` +
  `<table border="1" style="border-collapse: collapse; font-size: 12px"><tbody>` +
  `<tr><td style="padding: 2px 5px"><p>Measles morbillivirus</p></td>` +
  `<td style="padding: 2px 5px; font-size: 16px"><p>Detected</p></td></tr>` +
  `</tbody></table><p style="font-size: 12px">Comment: within normal limits.</p>`;

let fails = 0;
const check = (label, ok, detail) => {
  if (!ok) fails++;
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${label}${detail ? ` — ${detail}` : ''}`);
};

/** Every run of text in the document, with the size it is actually drawn at. */
const MEASURE = (root, win) => {
  const out = [];
  const walker = (root.ownerDocument ?? root).createTreeWalker(root, NodeFilter.SHOW_TEXT);
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const text = node.data.trim();
    if (!text) continue;
    const el = node.parentElement;
    if (!el || el.closest('[contenteditable=false]')) continue;
    // To a hundredth of a pixel, not to the nearest one: the point of this is
    // that a pasted point size is no longer rounded on the way in.
    out.push([
      text.slice(0, 24),
      Math.round(parseFloat(win.getComputedStyle(el).fontSize) * 100) / 100,
    ]);
  }
  return out;
};

const browser = await chromium.launch({ channel: 'chrome' });

async function open(initialHtml) {
  const { server, port } = await serve(() => harnessHtml(initialHtml));
  const context = await browser.newContext({
    viewport: { width: 1280, height: 900 },
    permissions: ['clipboard-read', 'clipboard-write'],
  });
  const page = await context.newPage();
  page.on('pageerror', (e) => console.log('PAGEERROR', String(e)));
  await page.goto(`http://127.0.0.1:${port}/`);
  await page.evaluate(() => window.__ready);
  return { page, context, server };
}

async function truthOf(page, html) {
  return page.evaluate(
    async ([source, measure]) => {
      const frame = document.createElement('iframe');
      frame.style.cssText = 'position:absolute;left:-9999px;width:1200px;height:900px';
      document.body.append(frame);
      frame.contentDocument.open();
      frame.contentDocument.write(source);
      frame.contentDocument.close();
      await new Promise((resolve) => setTimeout(resolve, 80));
      // eslint-disable-next-line no-eval
      const sizes = eval(`(${measure})`)(frame.contentDocument.body, frame.contentWindow);
      frame.remove();
      return sizes;
    },
    [html, MEASURE.toString()]
  );
}

const compare = (label, truth, mine) => {
  check(
    `${label}: every run is the size the document set it in`,
    truth.length === mine.length &&
      truth.every(([text, size], i) => mine[i][0] === text && Math.abs(mine[i][1] - size) <= 0.01),
    truth.map(([t, s], i) => `${t}=${s}/${mine[i]?.[1] ?? '-'}`).join(' ')
  );
};

// Pasted from Word.
{
  const { page, context, server } = await open('');
  const truth = await truthOf(page, WORD);

  await page.click('[data-slate-editor]');
  await page.evaluate(async (source) => {
    await navigator.clipboard.write([
      new ClipboardItem({
        'text/html': new Blob([source], { type: 'text/html' }),
        'text/plain': new Blob(['x'], { type: 'text/plain' }),
      }),
    ]);
  }, WORD);
  await page.keyboard.press('Control+V');
  await page.waitForTimeout(700);

  const mine = await page.evaluate(
    (measure) =>
      // eslint-disable-next-line no-eval
      eval(`(${measure})`)(document.querySelector('[data-slate-editor]'), window),
    MEASURE.toString()
  );

  compare('word paste', truth, mine);
  check(
    'and the body is not drawn bigger than the header above it',
    mine[2][1] < mine[0][1],
    `header ${mine[0][1]}px | body ${mine[2][1]}px`
  );

  await context.close();
  server.close();
}

// Pasted from Word, with the body's size only in the <style> block.
{
  const { page, context, server } = await open('');
  const truth = await truthOf(page, WORD_TABLE_STYLE);

  await page.click('[data-slate-editor]');
  await page.evaluate(async (source) => {
    await navigator.clipboard.write([
      new ClipboardItem({
        'text/html': new Blob([source], { type: 'text/html' }),
        'text/plain': new Blob(['x'], { type: 'text/plain' }),
      }),
    ]);
  }, WORD_TABLE_STYLE);
  await page.keyboard.press('Control+V');
  await page.waitForTimeout(700);

  const mine = await page.evaluate(
    (measure) =>
      // eslint-disable-next-line no-eval
      eval(`(${measure})`)(document.querySelector('[data-slate-editor]'), window),
    MEASURE.toString()
  );

  compare('word paste, size in the style block', truth, mine);

  await context.close();
  server.close();
}

// Opened from the database.
{
  const { page, context, server } = await open(STORED);
  const truth = await truthOf(page, STORED);
  await page.waitForTimeout(200);
  const mine = await page.evaluate(
    (measure) =>
      // eslint-disable-next-line no-eval
      eval(`(${measure})`)(document.querySelector('[data-slate-editor]'), window),
    MEASURE.toString()
  );
  compare('stored report', truth, mine);

  const saved = await page.evaluate(() =>
    window.RichTextEdit.getHtml('#report_body_editor')
  );
  check(
    'and saving it does not write the sizes out of the report',
    ['20px', '12px', '16px'].every((size) => saved.includes(`font-size: ${size}`)),
    saved.slice(0, 120)
  );

  await context.close();
  server.close();
}

await browser.close();
console.log(fails === 0 ? '\nALL PASS' : `\n${fails} FAILED`);
process.exit(fails === 0 ? 0 : 1);

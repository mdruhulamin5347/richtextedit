/**
 * Lines copied out of LibreOffice, outside a table: as far apart as the document has them?
 *
 * A report's "Test platform" block — five 9pt lines in the document's default
 * Calibri, single-spaced, no space between them — stands 14.67px a line in
 * LibreOffice's own layout of the .docx (`--convert-to pdf`, `pdftotext
 * -bbox-layout`). It pasted 20px a line: 0.73 of a line in the editor's own face
 * plus the editor's 4px of padding above and below. "Line spacing 1" then opened
 * it to 24px. See lib/word-line-gap.ts, `extractDocumentSpacing` in
 * lib/table-widths.ts and lib/font-face.ts; the arithmetic is pinned in
 * tests/js/libreoffice-paragraph-layout.test.ts.
 *
 * The HTML is LibreOffice's own clipboard, trimmed — captured through its
 * transferable (UNO). The RTF is cut down to what is read from it. The async
 * clipboard API cannot carry RTF, so the paste is the `beforeinput` Chrome
 * fires for a real one.
 *
 * RichText has no print page — a host renders the saved HTML — so this port
 * leaves out the upstream test's print-page check.
 *
 * Run: node tests/browser/libreoffice-paragraph-layout.mjs
 */
import { chromium } from 'playwright';
import { harnessHtml, serve } from './harness.mjs';

/** LibreOffice's own line pitch for these lines, in px. */
const LIBRE_OFFICE_PITCH = 14.67;

let fails = 0;
const check = (name, pass, detail = '') => {
  if (!pass) fails++;
  console.log(`${pass ? 'PASS' : 'FAIL'}  ${name}${detail ? ' — ' + detail : ''}`);
};

const LINES = [
  ['<u>Test\nplatform:</u>', '{\\ul Test platform:}'],
  ['PCR\nkit\t\t: CE IVD Marked RevoDx STI Pathogen Detection Kit,Turkey', 'PCR kit\\tab\\tab : CE IVD Marked RevoDx STI Pathogen Detection Kit,Turkey'],
  ['DNA\nExtraction\t: CE IVD marked Gene Proof pathogen free DNA isolation\nkit, Czech Republic', 'DNA Extraction\\tab : CE IVD marked Gene Proof pathogen free DNA isolation kit, Czech Republic'],
  ['Instrument\t:\nCFX Opus 96 Real-time PCR System', 'Instrument\\tab : CFX Opus 96 Real-time PCR System'],
  ['Probe\t\t:\nTaqMan probe', 'Probe\\tab\\tab : TaqMan probe'],
];

const LIBRE_OFFICE = `<!DOCTYPE html><html><head><meta http-equiv="content-type" content="text/html; charset=utf-8"/>
<meta name="generator" content="LibreOffice 24.2.7.2 (Linux)"/>
<style type="text/css">
p { line-height: 115%; text-align: left; orphans: 2; widows: 2; margin-bottom: 0.1in; direction: ltr; background: transparent }
</style></head><body lang="en-US" dir="ltr">
${LINES.map(([html]) => `<p style="line-height: 100%; margin-bottom: 0in"><font size="2" style="font-size: 9pt">${html}</font></p>`).join('\n')}
</body></html>`;

const RTF =
  String.raw`{\rtf1\ansi\deff4{\fonttbl{\f0\froman\fprq2\fcharset0 Times New Roman;}{\f4\froman\fprq2\fcharset0 Calibri;}}{\stylesheet{\s0 Normal;}}` +
  LINES.map(([, rtf]) => String.raw`\pard\plain \s0\f4\sl276\slmult1\sa200\sl240\slmult1\sa0{\fs18 ` + rtf + String.raw`}\par`).join('') +
  '}';

/** The same lines as Word writes them: a Word paste keeps the editor's own spacing. */
const WORD = `<html xmlns:w="urn:schemas-microsoft-com:office:word"><head><style><!--
p.MsoNormal {margin:0cm; font-size:9.0pt; font-family:"Calibri",sans-serif;}
--></style></head><body>${LINES.map(([html]) => `<p class=MsoNormal>${html.replace(/\n/g, ' ').replace(/\t/g, '<span style="mso-tab-count:1">&nbsp;&nbsp;&nbsp;</span>')}</p>`).join('')}</body></html>`;

const { server, port } = await serve(() => harnessHtml(''));
const browser = await chromium.launch({ channel: 'chrome' });
const page = await (await browser.newContext({ viewport: { width: 1280, height: 900 } })).newPage();
await page.goto(`http://127.0.0.1:${port}/`);
await page.evaluate(() => window.__ready);

const natural = await page.evaluate(() => {
  const d = document.createElement('div');
  d.textContent = 'Hg';
  d.style.cssText = 'position:absolute;left:-9999px;font-family:Calibri;font-size:1000px;line-height:normal';
  document.body.append(d);
  const ratio = d.getBoundingClientRect().height / 1000;
  d.remove();
  return ratio;
});
check("the browser can measure Calibri's natural line", natural > 1 && natural < 2, `natural = ${natural}`);

async function paste(html, rtf) {
  await page.evaluate(() => window.RichTextEdit.setHtml('#report_body_editor', ''));
  await page.click('[data-slate-editor]');
  await page.evaluate(({ html, rtf }) => {
    const data = new DataTransfer();
    data.setData('text/html', html);
    if (rtf) data.setData('text/rtf', rtf);
    data.setData('text/plain', 'x');
    document.querySelector('[data-slate-editor]').dispatchEvent(
      new InputEvent('beforeinput', { inputType: 'insertFromPaste', dataTransfer: data, bubbles: true, cancelable: true })
    );
  }, { html, rtf });
  await page.waitForTimeout(800);
}

/** Line to line, px, from each line's first character; and the second line's padding and font. */
function measure(frameId = null) {
  return page.evaluate((frameId) => {
    const doc = frameId ? document.getElementById(frameId).contentDocument : document;
    const root = frameId ? doc.querySelector('.rl-report-body') : doc.querySelector('[data-slate-editor]');
    const blocks = Array.from(root.children).filter((b) => /\S/.test(b.textContent));
    const firstText = (b) => {
      const walker = doc.createTreeWalker(b, NodeFilter.SHOW_TEXT);
      for (let n = walker.nextNode(); n; n = walker.nextNode()) if (/\S/.test(n.data)) return n;
      return null;
    };
    const tops = blocks.map((b) => {
      const t = firstText(b);
      const i = t.data.search(/\S/);
      const r = doc.createRange();
      r.setStart(t, i);
      r.setEnd(t, i + 1);
      return r.getBoundingClientRect().top;
    });
    const cs = getComputedStyle(blocks[1]);
    return {
      pitch: tops.slice(1).map((t, i) => Math.round((t - tops[i]) * 100) / 100),
      padding: `${cs.paddingTop}/${cs.paddingBottom}`,
      font: getComputedStyle(firstText(blocks[1]).parentElement).fontFamily,
    };
  }, frameId);
}

const allNear = (pitch, target, slack) => pitch.length === 4 && pitch.every((p) => Math.abs(p - target) <= slack);
const expected = natural * 12; // 9pt, single, in Calibri

await paste(LIBRE_OFFICE, RTF);
const pasted = await measure();
check(
  'pasted: each line one single line of Calibri apart, as in the document',
  allNear(pasted.pitch, expected, 0.1),
  `expected ${expected.toFixed(2)}px | editor ${pasted.pitch.join('/')}px`
);
check(
  "pasted: within 0.1px of LibreOffice's own layout",
  allNear(pasted.pitch, LIBRE_OFFICE_PITCH, 0.1),
  `LibreOffice ${LIBRE_OFFICE_PITCH}px | editor ${pasted.pitch.join('/')}px`
);
check('pasted: no editor padding on a line the document spaces', pasted.padding === '0px/0px', pasted.padding);
check('pasted: set in the document\'s font', /Calibri/.test(pasted.font), pasted.font);

// Choose "1" on all five lines: nothing may move — it is what they already are.
await page.evaluate(() => {
  const blocks = Array.from(document.querySelectorAll('[data-slate-editor] > [data-slate-node="element"]'));
  const strings = (b) => b.querySelectorAll('[data-slate-string]');
  const first = strings(blocks[0])[0].firstChild;
  const lastStrings = strings(blocks[blocks.length - 1]);
  const last = lastStrings[lastStrings.length - 1].firstChild;
  const range = document.createRange();
  range.setStart(first, 0);
  range.setEnd(last, last.length);
  const selection = window.getSelection();
  selection.removeAllRanges();
  selection.addRange(range);
});
await page.waitForTimeout(300);
await page.locator('button:has(svg.lucide-wrap-text)').first().click();
await page.waitForTimeout(300);
const ticked = await page.evaluate(() =>
  [...document.querySelectorAll('[role="menuitemradio"]')].filter((i) => i.getAttribute('aria-checked') === 'true').map((i) => i.textContent.trim())
);
check('the control names the gap 1', ticked.length === 1 && ticked[0] === '1', `menu shows ${JSON.stringify(ticked)}`);
await page.locator('[role="menuitemradio"]', { hasText: /^1$/ }).first().click();
await page.waitForTimeout(400);
const afterOne = await measure();
check('choosing "1" moves nothing', JSON.stringify(afterOne.pitch) === JSON.stringify(pasted.pitch), `before ${pasted.pitch.join('/')} | after ${afterOne.pitch.join('/')}`);

const saved = await page.evaluate(() => window.RichTextEdit.getHtml('#report_body_editor'));
await page.evaluate((html) => window.RichTextEdit.setHtml('#report_body_editor', html), saved);
await page.waitForTimeout(400);
const reopened = await measure();
check('a save and reopen keeps it', allNear(reopened.pitch, expected, 0.1) && reopened.padding === '0px/0px', `reopened ${reopened.pitch.join('/')}px, padding ${reopened.padding}`);

// A Word paste is not LibreOffice's: it keeps the editor's own paragraph spacing.
await paste(WORD, '{\\rtf1\\ansi x}');
const word = await measure();
check('a Word paste keeps the editor\'s own paragraph padding', word.padding === '4px/4px', word.padding);

await browser.close();
server.close();
console.log(fails ? `\n${fails} FAILED` : '\nall passed');
process.exit(fails ? 1 : 0);

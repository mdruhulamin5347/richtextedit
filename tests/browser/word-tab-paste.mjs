/**
 * A tab pasted out of a Word document: is it drawn at the editor's own gap?
 *
 * The editor's stops are Word's default half inch, so a line that uses only
 * default stops keeps Word's tab count and must sit exactly where the same line
 * opened from a saved report sits — and the print page must draw it there too.
 * Ruler stops and hanging indents are measured in office-paste-layout.mjs; the
 * counts themselves are checked in tests/js/word-tabs.test.ts.
 *
 * RichText has no print page — a host renders the saved HTML — so this port
 * leaves out the upstream test's print-page check.
 *
 * Run: node tests/browser/word-tab-paste.mjs
 */
import { chromium } from 'playwright';
import { harnessHtml, serve } from './harness.mjs';

let fails = 0;
const check = (name, pass, detail = '') => {
  if (!pass) fails++;
  console.log(`${pass ? 'PASS' : 'FAIL'}  ${name}${detail ? ' — ' + detail : ''}`);
};

const nbsp = (count) => '&nbsp;'.repeat(count);

const WORD = `<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:w="urn:schemas-microsoft-com:office:word"><head><style>
<!-- p.MsoNormal {margin:0cm; font-size:11.0pt; font-family:"Calibri",sans-serif;} --></style></head>
<body lang=EN-US>
<p class=MsoNormal>Test Name<span style='mso-tab-count:2'>${nbsp(13)} </span>: RT-PCR</p>
<p class=MsoNormal><span style='font-size:11.0pt;mso-tab-count:1'>${nbsp(6)} </span>Indented: yes</p>
</body></html>`;

const { server, port } = await serve(() => harnessHtml(''));
const browser = await chromium.launch({ channel: 'chrome' });
const context = await browser.newContext({
  viewport: { width: 1280, height: 900 },
  permissions: ['clipboard-read', 'clipboard-write'],
});
const page = await context.newPage();

async function fresh() {
  await page.goto(`http://127.0.0.1:${port}/`);
  await page.evaluate(() => window.__ready);
}

/** Where each line's colon sits, in `root`'s blocks, from the block's left edge. */
const colonsIn = (selector) =>
  page.evaluate((selector) => {
    const colonX = (block) => {
      const walker = block.ownerDocument.createTreeWalker(block, NodeFilter.SHOW_TEXT);
      for (let node = walker.nextNode(); node; node = walker.nextNode()) {
        const index = node.data.indexOf(':');
        if (index < 0) continue;
        const range = block.ownerDocument.createRange();
        range.setStart(node, index);
        range.setEnd(node, index + 1);
        return Math.round(range.getBoundingClientRect().left - block.getBoundingClientRect().left);
      }
      return null;
    };
    return Array.from(document.querySelectorAll(selector))
      .filter((block) => block.textContent.includes(':'))
      .map(colonX);
  }, selector);

const saved = () => page.evaluate(() => window.RichTextEdit.getHtml('#report_body_editor'));

// Pasted from Word.
await fresh();
await page.click('[data-slate-editor]');
await page.evaluate(async (h) => {
  await navigator.clipboard.write([
    new ClipboardItem({
      'text/html': new Blob([h], { type: 'text/html' }),
      'text/plain': new Blob(['Test Name\t\t: RT-PCR\n\tIndented: yes'], { type: 'text/plain' }),
    }),
  ]);
}, WORD);
await page.keyboard.press('Control+V');
await page.waitForTimeout(600);
const pasted = await colonsIn('[data-slate-editor] > *');
const pastedHtml = await saved();

// The same lines, as a report already saved with those tabs opens.
const report =
  '<p style="font-size: 11pt"><span style="font-size: 11pt; font-family: Calibri, sans-serif">Test Name<span style="white-space: pre">\t\t</span>: RT-PCR</span></p>' +
  '<p style="font-size: 11pt"><span style="font-size: 11pt; font-family: Calibri, sans-serif"><span style="white-space: pre">\t</span>Indented: yes</span></p>';
await fresh();
await page.evaluate((h) => window.RichTextEdit.setHtml('#report_body_editor', h), report);
await page.waitForTimeout(250);
const opened = await colonsIn('[data-slate-editor] > *');
const reopened = await saved();

check('a pasted Word tab is drawn at the editor\'s standard gap, like a saved one',
  pasted.length === 2 && pasted.every((x, i) => Math.abs(x - opened[i]) <= 1),
  `pasted ${pasted.join('/')}px, saved ${opened.join('/')}px`);

check('the tab Plate\'s cleaner used to leave as &nbsp; is a tab now',
  pastedHtml.includes('<span style="white-space: pre">\t</span>Indented') && !pastedHtml.includes('&nbsp;'),
  pastedHtml);

check('a report being opened keeps exactly the tabs it was saved with', reopened === report, reopened);

// Out of LibreOffice: its clipboard carries RTF as well as HTML, which is what
// makes Plate's docx cleaner run — and that cleaner used to delete every span
// holding only tabs, so these lines pasted as `BPD89 mm`. The async clipboard
// API cannot write RTF, so the paste is built by hand — as the `beforeinput`
// Chrome fires for a real one, which is the event Slate inserts rich HTML from.
const LIBRE_OFFICE = `<html><head><meta name="generator" content="LibreOffice 24.2.7.2 (Linux)"/></head><body lang="en-US">
<h3 class="western"><font face="Calibri, serif"><font size="3" style="font-size: 12pt"><span style="font-style: normal">BPD\t\t\t\t</span></font></font><font color="#ff0000"><font face="Calibri, serif"><font size="3" style="font-size: 12pt"><span style="font-style: normal"><b>89</b></span></font></font></font><font face="Calibri, serif"><font size="3" style="font-size: 12pt"><span style="font-style: normal">
mm</span></font></font></h3>
<h3 class="western"><font face="Calibri, serif"><font size="3" style="font-size: 12pt"><span style="font-style: normal">FL\t\t\t\t</span></font></font><font color="#ff0000"><font face="Calibri, serif"><font size="3" style="font-size: 12pt"><span style="font-style: normal"><b>66</b></span></font></font></font><font face="Calibri, serif"><font size="3" style="font-size: 12pt"><span style="font-style: normal">
mm</span></font></font></h3>
<h3 class="western"><font face="Calibri, serif"><font size="3" style="font-size: 12pt"><span style="font-style: normal">Number
\t\t\tSingle </span></font></font></h3></body></html>`;
await fresh();
await page.click('[data-slate-editor]');
await page.evaluate((html) => {
  const data = new DataTransfer();
  data.setData('text/html', html);
  data.setData('text/rtf', '{\\rtf1\\ansi USG}');
  data.setData('text/plain', 'BPD\t\t\t\t89 mm\nFL\t\t\t\t66 mm\nNumber \t\t\tSingle');
  document.querySelector('[data-slate-editor]').dispatchEvent(
    new InputEvent('beforeinput', { inputType: 'insertFromPaste', dataTransfer: data, bubbles: true, cancelable: true })
  );
}, LIBRE_OFFICE);
await page.waitForTimeout(600);
const libre = await page.evaluate(() =>
  Array.from(document.querySelectorAll('[data-slate-editor] > *')).map((block) => block.textContent)
);
const libreNumbers = await page.evaluate(() =>
  Array.from(document.querySelectorAll('[data-slate-editor] > *')).filter((b) => /\d/.test(b.textContent)).map((block) => {
    const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      const index = node.data.search(/\d/);
      if (index < 0) continue;
      const range = document.createRange();
      range.setStart(node, index);
      range.setEnd(node, index + 1);
      return Math.round(range.getBoundingClientRect().left - block.getBoundingClientRect().left);
    }
    return null;
  })
);
check('a LibreOffice paste, RTF and all, keeps every tab',
  libre.includes('BPD\t\t\t\t89 mm') && libre.includes('FL\t\t\t\t66 mm') && libre.some((t) => t.startsWith('Number \t\t\tSingle')),
  JSON.stringify(libre));
check('…drawn as a gap, not closed up',
  libreNumbers.length === 2 && libreNumbers[0] > 150 && Math.abs(libreNumbers[0] - libreNumbers[1]) <= 1,
  `numbers at ${libreNumbers.join('/')}px`);

await browser.close();
server.close();
console.log(fails ? `\n${fails} check(s) failed` : '\nAll checks passed');
process.exit(fails ? 1 : 0);

/**
 * A USG report copied out of LibreOffice: does it land where the document has it?
 *
 * Measured against the document itself: LibreOffice's own layout of the .docx
 * this fixture was cut from puts every value column at 2in (192px) — by default
 * half-inch stops (`BPD`), by a ruler stop at 1548 twips (`Number`, `Status`),
 * by a hanging indent's own stop (`Placenta`) and by 53 leading spaces (the
 * `No retro-placental` line). Each of those is something the paste used to lose:
 * the ruler stop is only in the RTF, the hanging indent read as a 2in indent,
 * the spaces sat outside the run's `<font face>`, and the lines are "Heading 3"
 * — which the editor drew with its own tight letter-spacing. See
 * lib/office-tab-stops.ts and lib/whitespace.ts.
 *
 * The HTML is LibreOffice's own clipboard, trimmed — captured through its
 * transferable (UNO), soft-wrap newlines and all. The RTF is cut down to what
 * is read from it: each paragraph's text, `\tx` stops and `\li`/`\fi` indents.
 * The async clipboard API cannot carry RTF, so the paste is the `beforeinput`
 * Chrome fires for a real one.
 *
 * RichText has no print page — a host renders the saved HTML — so this port
 * leaves out the upstream test's print-page check.
 *
 * Run: node tests/browser/office-paste-layout.mjs
 */
import { chromium } from 'playwright';
import { harnessHtml, serve } from './harness.mjs';

let fails = 0;
const check = (name, pass, detail = '') => {
  if (!pass) fails++;
  console.log(`${pass ? 'PASS' : 'FAIL'}  ${name}${detail ? ' — ' + detail : ''}`);
};

const F = '<font face="Calibri, serif"><font size="3" style="font-size: 12pt">';
const HTML = `<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN"><html><head>
<meta http-equiv="content-type" content="text/html; charset=utf-8"/><meta name="generator" content="LibreOffice 24.2.7.2 (Linux)"/>
<style type="text/css">h3.western { font-family: "Times New Roman", serif; font-style: italic }
h2.western { font-family: "Times New Roman", serif; font-size: 12pt; font-style: italic; font-weight: bold }
h3 { line-height: 100%; margin-top: 0in; margin-bottom: 0in } h2 { line-height: 100%; margin-top: 0in; margin-bottom: 0in }</style>
</head><body lang="en-US" dir="ltr">
<h3 class="western">${F}<span style="font-style: normal">BPD\t\t\t\t</span></font></font><font color="#ff0000">${F}<span style="font-style: normal"><b>89</b></span></font></font></font>${F}<span style="font-style: normal">
mm</span></font></font></h3>
<h3 class="western" style="margin-right: -0.5in">${F}<span style="font-style: normal">Number
\t\t\tSingle </span></font></font>
</h3>
<h3 class="western" style="margin-right: -0.5in">${F}<span style="font-style: normal"><b>Status
\t\t              Alive</b></span></font></font></h3>
<h3 class="western" style="text-indent: -2in; margin-left: 2in">${F}<span style="font-style: normal">Placenta
\tFundal </span></font></font></h3>
<h3 class="western" style="text-indent: -2in; margin-left: 2in">${" ".repeat(5)}
${" ".repeat(47)}${F}<span style="font-style: normal">No
 retro-placental    collection is seen.   </span></font></font>
</h3>
<h2 class="western"><font color="#000080"><font face="Calibri, serif"><span style="font-style: normal">Amniotic
Fluid\t\t             Adequate in amount</span></font></font></h2>
</body></html>`;

const RTF = String.raw`{\rtf1\ansi\deff0\deftab720{\fonttbl{\f0\fswiss Calibri;}}{\stylesheet{\s3 Heading 3;}}
\pard\plain\s3 BPD\tab\tab\tab\tab 89 mm\par
\pard\plain\s3\tx1548 Number \tab\tab\tab Single \par
\pard\plain\s3\tx1548{\b Status \tab\tab              Alive}\par
\pard\plain\s3\li2880\fi-2880 Placenta \tab Fundal \par
\pard\plain\s3\li2880\fi-2880                                                      No  retro-placental    collection is seen.\par
\pard\plain\s2 Amniotic Fluid\tab\tab             Adequate in amount\par}`;

const { server, port } = await serve(() => harnessHtml(''));
const browser = await chromium.launch({ channel: 'chrome' });
const page = await (await browser.newContext({ viewport: { width: 1280, height: 900 } })).newPage();
await page.goto(`http://127.0.0.1:${port}/`);
await page.evaluate(() => window.__ready);
await page.click('[data-slate-editor]');
await page.evaluate(({ html, rtf }) => {
  const data = new DataTransfer();
  data.setData('text/html', html);
  data.setData('text/rtf', rtf);
  data.setData('text/plain', 'BPD');
  document.querySelector('[data-slate-editor]').dispatchEvent(
    new InputEvent('beforeinput', { inputType: 'insertFromPaste', dataTransfer: data, bubbles: true, cancelable: true })
  );
}, { html: HTML, rtf: RTF });
await page.waitForTimeout(800);

/**
 * Per line, from the block's left edge: where the text after the last tab run
 * starts (`afterTabs`), and where the first word after the line's leading gap
 * starts (`value`) — the column the document lines up.
 */
function lines(rootSelector, frameId) {
  return page.evaluate(({ rootSelector, frameId }) => {
    const doc = frameId ? document.getElementById(frameId).contentDocument : document;
    const root = doc.querySelector(rootSelector);
    return Array.from(root.children).map((block) => {
      const left = block.getBoundingClientRect().left;
      const xAt = (node, i) => {
        const range = doc.createRange();
        range.setStart(node, i);
        range.setEnd(node, i + 1);
        return Math.round(range.getBoundingClientRect().left - left);
      };
      const text = block.textContent;
      const walker = doc.createTreeWalker(block, NodeFilter.SHOW_TEXT);
      const leadingGap = /^[\s\u00a0]/.test(text);
      let afterTabs = null, value = null, inGap = false, lastWasTab = false;
      for (let node = walker.nextNode(); node; node = walker.nextNode()) {
        for (let i = 0; i < node.data.length; i++) {
          const c = node.data[i];
          if (lastWasTab && c !== '\t') afterTabs = xAt(node, i);
          lastWasTab = c === '\t';
          if (c === '\t' || (leadingGap && value === null && /[\s\u00a0]/.test(c))) { inGap = true; continue; }
          if (inGap && value === null && /\S/.test(c) && c !== '\u00a0') value = xAt(node, i);
        }
      }
      const firstRun = walker.root.querySelector('span') ?? block;
      return {
        label: text.trim().split(/[\s\u00a0]+/)[0],
        tag: block.tagName,
        afterTabs,
        value,
        font: getComputedStyle(firstRun).fontFamily,
      };
    });
  }, { rootSelector, frameId });
}

const editor = await lines('[data-slate-editor]');
const byLabel = (rows, label) => rows.find((row) => row.label === label) ?? {};
const near = (x, target) => typeof x === 'number' && Math.abs(x - target) <= 2;

check('BPD: default half-inch stops put the value at 2in (192px)', near(byLabel(editor, 'BPD').value, 192),
  `${byLabel(editor, 'BPD').value}px`);
check('Number: the ruler stop from the RTF puts the value at 192px', near(byLabel(editor, 'Number').value, 192),
  `${byLabel(editor, 'Number').value}px`);
check('Status: the ruler stop takes its tabs to 144px, where its spaces start', near(byLabel(editor, 'Status').afterTabs, 144),
  `${byLabel(editor, 'Status').afterTabs}px`);
check("Placenta: the hanging indent's own stop puts the value at 192px", near(byLabel(editor, 'Placenta').value, 192),
  `${byLabel(editor, 'Placenta').value}px`);
check('…and the line starts where the document starts it, not 2in in',
  editor.every((row) => row.label !== 'Placenta') || (await page.evaluate(() => {
    const block = Array.from(document.querySelectorAll('[data-slate-editor] > *')).find((b) => b.textContent.includes('Placenta'));
    return parseFloat(getComputedStyle(block).marginLeft) === 0;
  })));
check('No retro-placental: 53 leading spaces in the run\'s own font reach 192px', near(byLabel(editor, 'No').value, 192),
  `${byLabel(editor, 'No').value}px`);
check('Amniotic Fluid: its tabs end at 144px, where its spaces start', near(byLabel(editor, 'Amniotic').afterTabs, 144),
  `${byLabel(editor, 'Amniotic').afterTabs}px`);
check('every line is in the run\'s own Calibri, not the style block\'s Times New Roman',
  editor.every((row) => /^Calibri/.test(row.font)), editor.map((row) => row.font.split(',')[0]).join(' / '));
check('the document\'s "Heading" lines are paragraphs, drawn in its own type',
  editor.every((row) => row.tag !== 'H2' && row.tag !== 'H3'), editor.map((row) => row.tag).join(' '));

await browser.close();
server.close();
console.log(fails ? `\n${fails} check(s) failed` : '\nAll checks passed');
process.exit(fails ? 1 : 0);

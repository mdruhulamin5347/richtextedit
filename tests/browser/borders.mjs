/**
 * Does a borderless Word table stay borderless — on screen and in the saved HTML?
 *
 * Word uses borderless tables purely to align label / value columns (an ECG
 * report is exactly that). Drawing a grid for them is wrong in the editor and
 * wrong in print.
 */
import { chromium } from 'playwright';
import { harnessHtml, serve } from './harness.mjs';

const cell = (text) => `<td style='border:none;padding:0cm 5.4pt'><p class=MsoNormal>${text}</p></td>`;
const ROWS = [['Rate','85 b/min'],['Rhythm','Regular'],['P-Wave','Normal'],['Impression','Findings are within normal limit.']];
const BORDERLESS = `<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 style='border-collapse:collapse;border:none'>` +
  ROWS.map(([k,v]) => `<tr>${cell(k)}${cell(':')}${cell(v)}</tr>`).join('') + `</table>`;
const BORDERED = `<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse'>` +
  `<tr><td style='border:solid windowtext 1.0pt'><p>Test</p></td><td style='border:solid windowtext 1.0pt'><p>Result</p></td></tr></table>`;

const results = [];
const check = (n, pass, d='') => { results.push(pass); console.log(`${pass?'PASS':'FAIL'}  ${n}${d?' — '+d:''}`); };

const browser = await chromium.launch({ channel: 'chrome' });

async function borderWidths(html) {
  const { server, port } = await serve(() => harnessHtml(html));
  const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
  await page.goto(`http://127.0.0.1:${port}/`);
  await page.evaluate(() => window.__ready);
  const data = await page.evaluate(() => {
    const cells = Array.from(document.querySelectorAll('[data-slate-editor] table td:not(.w-2)'));
    const widths = cells.map((c) => {
      const cs = getComputedStyle(c);
      // The visible line is drawn on the ::before overlay, not the td itself.
      const pb = getComputedStyle(c, '::before');
      return [cs.borderTopWidth, cs.borderBottomWidth, pb.borderTopWidth, pb.borderBottomWidth]
        .map((v) => parseFloat(v) || 0);
    });
    // Serialize through the editor rather than reading the textarea, which
    // still holds the seeded HTML until a change flushes it.
    return { widths, saved: window.RichTextEdit.getHtml('#report_body_editor') };
  });
  await page.close(); server.close();
  return data;
}

const off = await borderWidths(BORDERLESS);
const maxOff = Math.max(...off.widths.flat());
check('borderless table renders with NO visible cell borders', maxOff === 0, `max border width ${maxOff}px`);
// Every border declaration in the saved HTML is an explicit `border: 0`.
// Stated rather than omitted because the print pages supply the border-style
// Word leaves off, and an edge left unsaid picks that up at CSS's `medium`
// width — silence prints as a 3px black box.
const offDeclarations = off.saved.match(/border(-top|-right|-bottom|-left)?:[^;"]*/g) ?? [];
check('borderless table saves without borders',
  offDeclarations.length > 0 && offDeclarations.every((d) => /^border:\s*0$/.test(d)),
  offDeclarations.join(' | ') || 'no border declarations at all');
check('borderless table saves without border="1"', !off.saved.includes('border="1"'));
check('borderless content survives', off.saved.includes('85 b/min') && off.saved.includes('Impression'));

const on = await borderWidths(BORDERED);
const maxOn = Math.max(...on.widths.flat());
check('a genuinely bordered table still shows borders', maxOn > 0, `max border width ${maxOn}px`);
check('bordered table saves with borders',
  /border(-top|-right|-bottom|-left)?:\s*1px/.test(on.saved),
  on.saved.match(/border[^;"]*/)?.[0] ?? 'none found');

await browser.close();
const failed = results.filter((r) => !r).length;
console.log(`\n${results.length - failed}/${results.length} checks passed`);
process.exit(failed ? 1 : 0);

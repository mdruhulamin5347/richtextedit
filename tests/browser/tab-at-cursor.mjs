/**
 * Where does Tab put a tab?
 *
 * At the caret. `@platejs/indent` binds Tab to move the whole BLOCK — it sets
 * `indent`, which the serializer writes as `margin-left: indent * 40px` — so a
 * Tab pressed in the middle of a line moved the line 40px right and left the
 * text alone. A report is full of lines that put a value on a tab stop part-way
 * along, and typing one was impossible.
 *
 * The three cases that deliberately still reach the indent/list plugins are
 * checked too, because "fixed the tab key" must not mean "took the indent away".
 *
 * Run: node tests/browser/tab-at-cursor.mjs
 */
import { chromium } from 'playwright';
import { harnessHtml, serve } from './harness.mjs';

/** The editor draws a paragraph as a div, so `p` is not a selector here. */
const PARA = '[data-slate-editor] .slate-p';

let fails = 0;
const check = (name, pass, detail = '') => {
  if (!pass) fails++;
  console.log(`${pass ? 'PASS' : 'FAIL'}  ${name}${detail ? ' — ' + detail : ''}`);
};

const { server, port } = await serve(() => harnessHtml(''));
const browser = await chromium.launch({ channel: 'chrome' });
const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });

/**
 * Load `html`, put the caret `right` characters into the block `selector`
 * matches, press `keys`, and hand back what the editor serializes.
 */
async function press(html, selector, right, keys) {
  await page.goto(`http://127.0.0.1:${port}/`);
  await page.evaluate(() => window.__ready);
  await page.evaluate((h) => window.RadiolensEditor.setHtml('#report_body_editor', h), html);
  await page.waitForSelector(selector, { timeout: 5000 });
  await page.waitForTimeout(250);
  await page.click(selector);
  await page.waitForTimeout(120);
  await page.keyboard.press('Home');
  for (let i = 0; i < right; i++) await page.keyboard.press('ArrowRight');
  for (const key of keys) {
    await page.keyboard.press(key);
    await page.waitForTimeout(180);
  }
  return page.evaluate(() => window.RadiolensEditor.getHtml('#report_body_editor'));
}

/** A tab, as the serializer writes one — see lib/whitespace.ts. */
const TAB = '<span style="white-space: pre">\t</span>';

const mid = await press('<p>Findings of the study</p>', PARA, 8, ['Tab']);
check('a Tab lands at the caret, not at the start of the line',
  mid === `<p>Findings${TAB} of the study</p>`, mid);

const start = await press('<p>Findings of the study</p>', PARA, 0, ['Tab']);
check('a Tab at the start of a line is still a tab, not an indent',
  start === `<p>${TAB}Findings of the study</p>`, start);

const heading = await press('<h2>Technique here</h2>', '[data-slate-editor] h2', 4, ['Tab']);
check('a heading takes one the same way', heading === `<h2>Tech${TAB}nique here</h2>`, heading);

// Matched loosely on purpose. Which cell a click lands in shifts with the
// row-control cell the editable prepends, and Home/ArrowRight inside a cell do
// not land on the same offset they do in a paragraph — neither of which this is
// about. What it is about: the tab went in AFTER some text, not in front of the
// line, which is the whole bug.
const cell = await press(
  '<table><tr><td>Alpha</td><td>Beta</td></tr></table>',
  '[data-slate-editor] td:nth-child(2)',
  2,
  ['Tab']
);
check('so does a table cell',
  /<td[^>]*><p>[^<]+<span style="white-space: pre">\t<\/span>/.test(cell), cell);

// One span, two tabs: the serializer wraps a RUN of whitespace, not each
// character. See lib/whitespace.ts.
const twice = await press('<p>AB</p>', PARA, 1, ['Tab', 'Tab']);
check('two of them make two tabs',
  twice === '<p>A<span style="white-space: pre">\t\t</span>B</p>', twice);

// The three that are deliberately left to the plugins underneath.
const nested = await press('<ul><li>Alpha</li><li>Beta</li></ul>', '[data-slate-editor] li:nth-child(2)', 2, ['Tab']);
check('a list item still nests under the one above it',
  nested === '<ul><li>Alpha<ul><li>Beta</li></ul></li></ul>', nested);

const outdent = await press('<p style="margin-left: 80px">Indented line</p>', PARA, 3, ['Shift+Tab']);
check('Shift+Tab still outdents the block',
  outdent === '<p style="margin-left: 40px">Indented line</p>', outdent);

const overSelection = await press('<p>Findings of the study</p>', PARA, 0, [
  'Shift+ArrowRight',
  'Shift+ArrowRight',
  'Tab',
]);
check('Tab over a selection indents rather than eating the selected text',
  overSelection === '<p style="margin-left: 40px">Findings of the study</p>', overSelection);

await browser.close();
server.close();
console.log(fails ? `\n${fails} check(s) failed` : '\nAll checks passed');
process.exit(fails ? 1 : 0);

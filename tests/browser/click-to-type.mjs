/**
 * Does clicking anywhere in the editor give you a caret to type at?
 *
 * The editable is only as tall as its content, so on a short report it filled
 * about 50px of the 600px frame and every click in the space below it landed on
 * the container instead: nothing focused, and typing went nowhere. Only a real
 * browser can show this — it is a question about where a box ends up on screen.
 */
import { chromium } from 'playwright';
import { harnessHtml, serve } from './harness.mjs';

let fails = 0;
function check(label, ok, detail) {
  if (!ok) fails++;
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${label}${detail ? ` — ${detail}` : ''}`);
}

const browser = await chromium.launch({ channel: 'chrome' });

async function open(initialHtml) {
  const { server, port } = await serve(() => harnessHtml(initialHtml));
  const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
  await page.goto(`http://127.0.0.1:${port}/`);
  await page.evaluate(() => window.__ready);
  return { page, server };
}

// The blank space below a short report.
{
  const { page, server } = await open('<p>First line of the report.</p>');

  const geometry = await page.evaluate(() => {
    const editable = document.querySelector('[data-slate-editor]');
    const frame = editable.parentElement.getBoundingClientRect();
    return {
      editable: Math.round(editable.getBoundingClientRect().height),
      frame: Math.round(frame.height),
      x: Math.round(frame.left + frame.width / 2),
      y: Math.round(frame.bottom - 20),
    };
  });

  check(
    'the editable fills the frame rather than just its own text',
    geometry.editable >= geometry.frame,
    `${geometry.editable}px of ${geometry.frame}px`
  );

  await page.mouse.click(geometry.x, geometry.y);
  const focused = await page.evaluate(
    () => document.activeElement?.hasAttribute('data-slate-editor') ?? false
  );
  check('clicking the empty space below the text focuses the editor', focused);

  await page.keyboard.type('TYPED');
  await page.waitForTimeout(250);
  const text = await page.evaluate(() => document.querySelector('[data-slate-editor]').textContent);
  check('and what is typed goes into the report', text.includes('TYPED'), text);

  await page.close();
  server.close();
}

// Clicking ON text still places the caret there, not at the end.
{
  const { page, server } = await open('<p>AAAAAAAAAA</p><p>BBBBBBBBBB</p><p>CCCCCCCCCC</p>');

  const target = await page.evaluate(() => {
    const line = document.querySelectorAll('[data-slate-editor] > div')[1].getBoundingClientRect();
    return { x: Math.round(line.left + 4), y: Math.round(line.top + line.height / 2) };
  });
  await page.mouse.click(target.x, target.y);
  await page.keyboard.type('X');
  await page.waitForTimeout(250);

  const lines = await page.evaluate(() =>
    Array.from(document.querySelectorAll('[data-slate-editor] > div')).map((d) => d.textContent)
  );
  check(
    'clicking on a line still puts the caret on that line',
    lines[1].startsWith('X') && lines[0] === 'AAAAAAAAAA',
    lines.join(' | ')
  );

  await page.close();
  server.close();
}

// The placeholder must not drift into the middle of the taller box.
{
  const { page, server } = await open('');
  const offset = await page.evaluate(() => {
    const editable = document.querySelector('[data-slate-editor]');
    const placeholder = document.querySelector('[data-slate-placeholder]');
    if (!placeholder) return null;
    return Math.round(
      placeholder.getBoundingClientRect().top - editable.getBoundingClientRect().top
    );
  });
  check('the placeholder stays on the first line', offset !== null && offset < 40, `${offset}px down`);
  await page.close();
  server.close();
}

await browser.close();
console.log(fails ? `\n${fails} check(s) failed` : '\nall checks passed');
process.exit(fails ? 1 : 0);

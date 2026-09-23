/**
 * Does a pasted table keep its BORDER DESIGN — not just "has borders"?
 *
 * A table's rules carry meaning in these reports: a double line under a header
 * row, a heavy frame around a summary block, a hairline grid inside it, a
 * dotted separator. All of it used to arrive here as the same flat 1px solid
 * box on every cell, because the cell component drew Tailwind's `border-b`
 * (which is 1px, and solid, and nothing else) and read the size only as a
 * yes/no. The saved report had carried the real widths and styles all along,
 * so the editor and the print disagreed about the same table.
 *
 * The reference is the browser's OWN rendering of the same markup in a bare
 * iframe. That is not an approximation of the truth, it IS the truth: the
 * editor this one replaces never parsed a table at all, it dropped the
 * clipboard HTML into a contenteditable and let the browser draw it. So "what
 * Summernote showed" and "what an iframe shows" are the same thing.
 */
import { chromium } from 'playwright';
import { harnessHtml, serve } from './harness.mjs';

/** Word's own spelling of a bordered grid cell. */
const wordCell = (border, text) =>
  `<td style='${border};padding:0cm 5.4pt'><p class=MsoNormal>${text}</p></td>`;

const CASES = [
  {
    name: 'a double rule under the header row',
    // Word writes the heavy separator as the HEADER's bottom and gives the body
    // row `border-top:none` — an interior edge belongs to whichever of the two
    // cells meeting there claims it.
    html: `<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse'>
      <tr>${wordCell('border:solid windowtext 1.0pt;border-bottom:double windowtext 2.25pt', 'Test')}${wordCell('border:solid windowtext 1.0pt;border-left:none;border-bottom:double windowtext 2.25pt', 'Result')}</tr>
      <tr>${wordCell('border:solid windowtext 1.0pt;border-top:none', 'Haemoglobin')}${wordCell('border:solid windowtext 1.0pt;border-top:none;border-left:none', '13.5 g/dL')}</tr>
    </table>`,
  },
  {
    name: 'a heavy frame around a hairline grid',
    html: `<table style='border:4px solid #000;border-collapse:collapse'>
      <tr><td style='border:1px solid #000;padding:2px'>a</td><td style='border:1px solid #000;padding:2px'>b</td></tr>
      <tr><td style='border:1px solid #000;padding:2px'>c</td><td style='border:1px solid #000;padding:2px'>d</td></tr>
    </table>`,
  },
  {
    name: 'a frame stated on the table, with the cells silent',
    html: `<table style='border:3px double #333;border-collapse:collapse'>
      <tr><td style='padding:2px'>a</td><td style='padding:2px'>b</td></tr>
      <tr><td style='padding:2px'>c</td><td style='padding:2px'>d</td></tr>
    </table>`,
  },
  {
    name: 'a dotted separator and a per-side mix',
    html: `<table style='border-collapse:collapse'>
      <tr><td style='border:2px dotted #000;padding:2px'>a</td><td style='border-top:1px solid #000;border-bottom:5px solid #000;padding:2px'>b</td></tr>
    </table>`,
  },
  {
    name: 'a rule stated on the row',
    html: `<table style='border-collapse:collapse'>
      <tr style='border-bottom:2px solid #000'><td style='padding:2px'>a</td><td style='padding:2px'>b</td></tr>
      <tr><td style='padding:2px'>c</td><td style='padding:2px'>d</td></tr>
    </table>`,
  },
  {
    name: 'a borderless layout table',
    html: `<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 style='border-collapse:collapse;border:none'>
      <tr>${wordCell('border:none', 'Rate')}${wordCell('border:none', ':')}${wordCell('border:none', '85 b/min')}</tr>
    </table>`,
  },
  {
    name: 'a legacy report grid, drawn by the border attribute alone',
    html: `<table border=1 cellspacing=0 cellpadding=2>
      <tr><td>a</td><td>b</td></tr><tr><td>c</td><td>d</td></tr>
    </table>`,
  },
];

/**
 * Every rule the table shows, as one string per EDGE of the grid.
 *
 * Not per cell side: under `border-collapse: collapse` an edge is shared, and
 * the two renderings divide it up differently. The browser leaves both cells'
 * declarations in place and paints the winner; the editor draws each edge once,
 * because Plate's collapse emulation hands a cell its bottom and right, plus a
 * top on the first row and a left on the first column. Comparing resolved edges
 * is what makes them comparable at all.
 *
 * Read in the page so both sides go through the same code.
 */
const MEASURE = (inEditor) => {
  const host = inEditor
    ? { doc: document, win: window, pseudo: '::before' }
    : (() => {
        const frame = document.querySelector('iframe');
        return { doc: frame.contentDocument, win: frame.contentWindow, pseudo: null };
      })();

  const root = inEditor ? host.doc.querySelector('[data-slate-editor]') : host.doc;
  const table = root.querySelector('table');

  // `inset` / `outset` are the 3D bevels the HTML `border` attribute asks a
  // browser for. Nothing in this application draws them: the editor, the
  // serializer and every print stylesheet render a report's rules as flat black
  // lines, which is also what the Summernote editor's own tables looked like.
  // So they count as `solid` on both sides of this comparison.
  const FLAT = { inset: 'solid', outset: 'solid', ridge: 'solid', groove: 'solid' };

  const rule = (el, side) => {
    const cs = host.win.getComputedStyle(el, el === table ? null : host.pseudo);
    const width = Math.round(parseFloat(cs.getPropertyValue(`border-${side}-width`)) || 0);
    const style = cs.getPropertyValue(`border-${side}-style`);
    return width === 0 || style === 'none' ? '-' : `${width}px ${FLAT[style] ?? style}`;
  };

  const strongest = (...rules) =>
    rules.filter(Boolean).reduce((best, r) => (parseFloat(r) > (parseFloat(best) || 0) ? r : best), '-');

  const grid = Array.from(table.querySelectorAll('tr')).map((r) =>
    Array.from(r.children).filter(
      (c) => /^t[dh]$/i.test(c.tagName) && !c.classList.contains('w-2')
    )
  );

  // A rule can be stated on the <table> or the <tr> rather than on a cell, and
  // the browser collapses it down onto the edges those elements cover — a <tr>
  // border draws right across its row, a table border around the whole grid.
  // In the editor that has already been resolved onto the cells; in a plain
  // browser it stays where it was written, so it has to be folded in here or
  // the reference is short of what the screenshot shows.
  const container = (el, side) => (host.pseudo || !el ? '-' : rule(el, side));
  const rowOf = (r) => grid[r]?.[0]?.parentElement;

  const horizontal = [];
  for (let r = 0; r <= grid.length; r++) {
    horizontal.push(
      grid[0].map((_, c) =>
        strongest(
          grid[r - 1]?.[c] && rule(grid[r - 1][c], 'bottom'),
          grid[r]?.[c] && rule(grid[r][c], 'top'),
          container(rowOf(r - 1), 'bottom'),
          container(rowOf(r), 'top'),
          r === 0 ? container(table, 'top') : r === grid.length ? container(table, 'bottom') : null
        )
      )
    );
  }

  const vertical = grid.map((row, r) => {
    const out = [];
    for (let c = 0; c <= row.length; c++) {
      out.push(
        strongest(
          row[c - 1] && rule(row[c - 1], 'right'),
          row[c] && rule(row[c], 'left'),
          c === 0 ? container(rowOf(r), 'left') : c === row.length ? container(rowOf(r), 'right') : null,
          c === 0 ? container(table, 'left') : c === row.length ? container(table, 'right') : null
        )
      );
    }
    return out;
  });

  return { horizontal, vertical };
};

let fails = 0;
const check = (label, ok, detail) => {
  if (!ok) fails++;
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${label}${!ok && detail ? `\n        ${detail}` : ''}`);
};

const fmt = (m) =>
  `h[${m.horizontal.map((r) => r.join(',')).join(' / ')}] v[${m.vertical.map((r) => r.join(',')).join(' / ')}]`;

const browser = await chromium.launch({ channel: 'chrome' });

for (const { name, html } of CASES) {
  const doc = `<html xmlns:w="urn:schemas-microsoft-com:office:word"><body>${html}<p>after</p></body></html>`;
  const { server, port } = await serve(() => harnessHtml(''));
  const context = await browser.newContext({
    viewport: { width: 1440, height: 900 },
    permissions: ['clipboard-read', 'clipboard-write'],
  });
  const page = await context.newPage();
  await page.goto(`http://127.0.0.1:${port}/`);
  await page.evaluate(() => window.__ready);

  await page.evaluate(async (h) => {
    const frame = document.createElement('iframe');
    frame.style.cssText = 'position:absolute;left:-9999px;width:1200px;height:900px';
    document.body.append(frame);
    frame.contentDocument.open();
    frame.contentDocument.write(h);
    frame.contentDocument.close();
    await new Promise((r) => setTimeout(r, 80));
  }, doc);
  const truth = await page.evaluate(`(${MEASURE})(false)`);

  await page.click('[data-slate-editor]');
  await page.evaluate(async (h) => {
    await navigator.clipboard.write([
      new ClipboardItem({
        'text/html': new Blob([h], { type: 'text/html' }),
        'text/plain': new Blob([''], { type: 'text/plain' }),
      }),
    ]);
  }, doc);
  await page.keyboard.press('Control+V');
  await page.waitForTimeout(500);

  const got = await page.evaluate(`(${MEASURE})(true)`);
  const saved = await page.evaluate(() => window.RadiolensEditor.getHtml('#report_body_editor'));

  check(
    `${name}: the editor draws what the browser draws`,
    JSON.stringify(truth) === JSON.stringify(got),
    `browser ${fmt(truth)}\n        editor  ${fmt(got)}`
  );

  // And the saved report has to carry the same design, or the print disagrees
  // with the screen — the exact split this file exists to close.
  const savedRules = new Set(
    (saved.match(/border-(top|right|bottom|left): [^;"]*/g) ?? []).map((d) =>
      d.replace(/^border-\w+: /, '').replace(/ (#\w+|rgb\([^)]*\))$/, '')
    )
  );
  const drawn = [
    ...new Set(truth.horizontal.flat().concat(truth.vertical.flat()).filter((e) => e !== '-')),
  ];
  check(
    `${name}: every rule the browser draws reaches the saved report`,
    drawn.every((r) => savedRules.has(r)),
    `browser draws ${drawn.join(' | ') || 'nothing'}; saved has ${[...savedRules].join(' | ') || 'nothing'}`
  );

  await page.close();
  await context.close();
  server.close();
}

await browser.close();
console.log(fails ? `\n${fails} check(s) failed` : '\nall checks passed');
process.exit(fails ? 1 : 0);

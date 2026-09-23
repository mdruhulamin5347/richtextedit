/**
 * Is a table row as deep as the document made it?
 *
 * A row's height is not the cell's padding alone. Word sets the paragraph
 * inside the cell with space before and after — `margin-top:6.0pt` and its pair
 * are ordinary on a report's table — and that is most of the air between one
 * row's text and the next. Plate's deserializer drops a block's margins
 * outright, so every row arrived here at the same minimum height whatever the
 * document said, and the saved report lost the gap again on the way out.
 *
 * The reference is the browser's OWN rendering of the same markup in a clean
 * iframe, which is not an approximation of the old Summernote editor but
 * literally it: that editor dropped the clipboard HTML into a contenteditable
 * and let the browser draw it.
 *
 * Measured as the DIFFERENCE the spacing makes, not as an absolute height. The
 * editor's paragraphs sit on an 18px base where the document's text is 11pt, so
 * a row here is a few px deeper than Word's whatever happens — that residual is
 * settled ground (see ai/memory/radiolens-editor-paste-scope.md) and it cancels
 * out of a comparison between two documents that differ only in their spacing.
 */
import { chromium } from 'playwright';
import { harnessHtml, serve } from './harness.mjs';

const ROWS = [
  ['Rate', '85 b/min'],
  ['Rhythm', 'Regular'],
  ['P-Wave', 'Normal'],
  ['QRS Complex', 'Normal'],
];

const doc = (paragraphStyle) =>
  `<html xmlns:w="urn:schemas-microsoft-com:office:word"><head><style>` +
  `<!-- p.MsoNormal {margin:0cm; font-size:11.0pt; font-family:"Calibri",sans-serif;} --></style></head><body>` +
  `<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse'>` +
  ROWS.map(
    ([label, value]) =>
      `<tr>` +
      [label, value]
        .map(
          (text) =>
            `<td style='border:solid windowtext 1.0pt;padding:0cm 5.4pt'>` +
            `<p class=MsoNormal style='${paragraphStyle}'><span style='font-size:11.0pt'>${text}</span></p></td>`
        )
        .join('') +
      `</tr>`
  ).join('') +
  `</table></body></html>`;

/** The same table twice: set tight, and set with 6pt above and below each line. */
const TIGHT = doc('margin:0cm');
const SPACED = doc('margin-top:6.0pt;margin-bottom:6.0pt');

let fails = 0;
const check = (label, ok, detail) => {
  if (!ok) fails++;
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${label}${detail ? ` — ${detail}` : ''}`);
};

const browser = await chromium.launch({ channel: 'chrome' });

/** What the old editor showed: the browser rendering the markup untouched. */
async function inIframe(page, html) {
  return page.evaluate(async (source) => {
    const frame = document.createElement('iframe');
    frame.style.cssText = 'position:absolute;left:-9999px;width:1200px;height:900px';
    document.body.append(frame);
    frame.contentDocument.open();
    frame.contentDocument.write(source);
    frame.contentDocument.close();
    await new Promise((resolve) => setTimeout(resolve, 100));
    const table = frame.contentDocument.querySelector('table');
    const height = Math.round(table.rows[0].getBoundingClientRect().height);
    frame.remove();
    return height;
  }, html);
}

/** The editor with a report already in it — the load path, no clipboard. */
async function open(initialHtml) {
  const { server, port } = await serve(() => harnessHtml(initialHtml));
  const context = await browser.newContext({ viewport: { width: 1440, height: 900 } });
  const page = await context.newPage();
  page.on('pageerror', (e) => console.log('PAGEERROR', String(e)));
  await page.goto(`http://127.0.0.1:${port}/`);
  await page.evaluate(() => window.__ready);
  return { page, context, server };
}

async function inEditor(html) {
  const { server, port } = await serve(() => harnessHtml(''));
  const context = await browser.newContext({
    viewport: { width: 1440, height: 900 },
    permissions: ['clipboard-read', 'clipboard-write'],
  });
  const page = await context.newPage();
  page.on('pageerror', (e) => console.log('PAGEERROR', String(e)));
  await page.goto(`http://127.0.0.1:${port}/`);
  await page.evaluate(() => window.__ready);

  const truth = await inIframe(page, html);

  await page.click('[data-slate-editor]');
  await page.evaluate(async (source) => {
    await navigator.clipboard.write([
      new ClipboardItem({
        'text/html': new Blob([source], { type: 'text/html' }),
        'text/plain': new Blob(['x'], { type: 'text/plain' }),
      }),
    ]);
  }, html);
  await page.keyboard.press('Control+V');
  await page.waitForTimeout(600);

  const result = await page.evaluate(() => {
    const table = document.querySelector('[data-slate-editor] table');
    return {
      rowHeight: Math.round(table.rows[0].getBoundingClientRect().height),
      saved: document.querySelector('#report_body').value,
    };
  });

  await context.close();
  server.close();
  return { truth, ...result };
}

/** The same row four times over, set in four sizes. */
const SIZES = [24, 18, 15, 10];
const SCALED =
  `<html><body><table border=1 cellspacing=0 style='border-collapse:collapse'>` +
  SIZES.map(
    (px) =>
      `<tr><td style='border:solid 1px #000;padding:0 5.4pt'>` +
      `<p style='margin-top:8px;margin-bottom:8px'><span style="font-size:${px}px">Measles ${px}px</span></p></td></tr>`
  ).join('') +
  `</table></body></html>`;

/**
 * The three-signature footer at the foot of a report: a borderless table, one
 * row, and a first column whose lines Word indents with a run of spaces of its
 * own — `<span style='mso-spacerun:yes'>` — carrying no font size.
 */
const NBSP = '&nbsp;';
const sigText = (t) => `<span style='font-size:11.0pt;font-family:"Calibri",sans-serif'>${t}</span>`;
const sigCell = (lines) =>
  `<td style='border:none;padding:0cm 5.4pt'>` +
  lines
    .map(
      ([indent, t]) =>
        `<p class=MsoNormal>` +
        (indent ? `<span style='mso-spacerun:yes'>${NBSP.repeat(indent)} </span>` : '') +
        `${sigText(t)}</p>`
    )
    .join('') +
  `</td>`;
const FOOTER =
  `<html xmlns:w="urn:schemas-microsoft-com:office:word"><head><style>` +
  `<!-- p.MsoNormal {margin:0cm;} --></style></head><body>` +
  `<table class=MsoNormalTable border=0 cellspacing=0 style='border-collapse:collapse;border:none'><tr>` +
  sigCell([
    [0, 'Medical Technologist'],
    [4, 'Department of Molecular Biology'],
    [12, 'Chittogram Maa Shishu-O-General Hospital'],
  ]) +
  sigCell([
    [0, 'Dr. Nahid Sultana'],
    [0, 'MBBS. D. Bact, M.Phil (Microbiology)'],
    [0, 'Professor &amp; Head'],
  ]) +
  `</tr></table></body></html>`;

const tight = await inEditor(TIGHT);
const spaced = await inEditor(SPACED);

const wordGap = spaced.truth - tight.truth;
const editorGap = spaced.rowHeight - tight.rowHeight;

check(
  'the document really does state a gap to reproduce',
  wordGap >= 15,
  `${tight.truth}px -> ${spaced.truth}px`
);
// The document states 6pt above AND 6pt below — 8px each. Only the space
// ABOVE is reproduced: the space a cell's LAST paragraph keeps after itself is
// the gap under the whole row, and this editor drops it. See
// `extractBlockSpacing`.
const SPACE_ABOVE = 8;
check(
  'the editor opens the gap the document put ABOVE the text',
  Math.abs(editorGap - SPACE_ABOVE) <= 1,
  `word +${wordGap}px (above AND below) | editor +${editorGap}px (above only)`
);
check(
  'and does NOT open the one it put under the row',
  editorGap < wordGap,
  `word +${wordGap}px | editor +${editorGap}px`
);
check(
  'a row set tight is still tight — no spacing is invented',
  tight.rowHeight - tight.truth <= 6,
  `word ${tight.truth}px | editor ${tight.rowHeight}px`
);
check(
  'the saved report carries the gap above, so the print matches the screen',
  /<p style="[^"]*margin-top: 8px/.test(spaced.saved),
  spaced.saved.match(/<p style="[^"]*"/)?.[0]
);
check(
  'and carries no space under the row, so the print does not open one either',
  !/margin-bottom/.test(spaced.saved),
  spaced.saved.match(/<p style="[^"]*"/)?.[0]
);
check(
  'and states nothing where the document stated nothing',
  !tight.saved.includes('margin-top'),
  tight.saved.match(/<p[^>]*>/)?.[0]
);

// A row is as deep as the text in it, not as deep as the editor's base.
//
// Every line box carries a STRUT in the BLOCK's own font, and that block used to
// sit on the editable's 18px base whatever its text was — so 15px and 10px rows
// measured identically, both taller than the document. The margins were already
// the document's by then; this is the other half of a row's height.
{
  const { page, context, server } = await open(SCALED);
  const truth = await page.evaluate(
    async (source) => {
      const frame = document.createElement('iframe');
      frame.style.cssText = 'position:absolute;left:-9999px;width:1200px;height:900px';
      document.body.append(frame);
      frame.contentDocument.open();
      frame.contentDocument.write(source);
      frame.contentDocument.close();
      await new Promise((resolve) => setTimeout(resolve, 100));
      const heights = [...frame.contentDocument.querySelectorAll('tr')].map((row) =>
        Math.round(row.getBoundingClientRect().height)
      );
      frame.remove();
      return heights;
    },
    SCALED
  );
  await page.waitForTimeout(200);
  const mine = await page.evaluate(() =>
    [...document.querySelectorAll('[data-slate-editor] tr')].map((row) =>
      Math.round(row.getBoundingClientRect().height)
    )
  );

  check(
    'rows of different text are no longer all the same height',
    new Set(mine).size === mine.length,
    `${mine.join(' / ')}px`
  );
  // Every cell here holds ONE paragraph set `margin-top:8px;margin-bottom:8px`,
  // so its margin-bottom is a last block's and is dropped: a row lands 8px
  // shorter than the document's. The +6 band above that is the old residual —
  // the editor's paragraphs sit on an 18px base where the text is smaller.
  const DROPPED_SPACE_AFTER = 8;
  /*
   * And back the other way: the editor RESERVES the rule the document draws.
   *
   * A cell's rules are painted by an overlay laid over the whole cell, so they
   * used to cover whatever was under them — on a document that pads its cells
   * `0cm 5.4pt`, the text reaches the cell's edge and a thick rule was drawn
   * straight through the descenders of a `g` or a `y`. The content box now
   * keeps a transparent band of the same width, which costs each row its top
   * and bottom rule. This fixture draws 1px on each.
   */
  const RESERVED_RULE = 2;
  check(
    'and each one tracks the height the document gives it, less the space under the row',
    truth.every((height, i) => {
      const expected = height - DROPPED_SPACE_AFTER + RESERVED_RULE;
      return mine[i] >= expected - RESERVED_RULE && mine[i] - expected <= 6;
    }),
    `word ${truth.join('/')} | editor ${mine.join('/')} ` +
      `(${DROPPED_SPACE_AFTER}px shorter for the dropped space-after, ${RESERVED_RULE}px taller for the reserved rule)`
  );
  check(
    'smaller text really does make a shorter row',
    mine[0] > mine[1] && mine[1] > mine[2] && mine[2] > mine[3],
    mine.join(' > ')
  );

  await context.close();
  server.close();
}

// An indented line is the same size as the line beside it.
//
// Word writes a line's indentation as a run of its own and puts no size on it,
// so a paragraph holding one looked like a paragraph whose runs disagreed and
// fell back to the editor's base: in this footer the first column's lines stood
// 24px against the other column's 20px, and their indentation was drawn in 18px
// spaces rather than the document's 11pt ones. A space shows nothing and says
// nothing — it takes the size of the line it sits on, as it does in a browser.
{
  const { page, context, server } = await open('');

  /** Per cell: each line's height, and where its visible text begins. */
  const LINES = (root, win, doc) =>
    [...root.querySelectorAll('td')]
      .filter((td) => !td.classList.contains('w-2'))
      .map((td) => {
        const box = td.getBoundingClientRect();
        const paragraphs = td.querySelectorAll(doc ? 'p' : '.slate-p');
        return [...paragraphs].map((p) => {
          const runs = p.querySelectorAll(doc ? 'span' : '[data-slate-string]');
          const run = runs[runs.length - 1] ?? p;
          return {
            height: Math.round(win.getComputedStyle(p).lineHeight === '' ? 0 : p.getBoundingClientRect().height),
            textAt: Math.round(run.getBoundingClientRect().left - box.left),
          };
        });
      });

  const truth = await page.evaluate(
    async ([source, measure]) => {
      const frame = document.createElement('iframe');
      frame.style.cssText = 'position:absolute;left:-9999px;width:1200px;height:900px';
      document.body.append(frame);
      frame.contentDocument.open();
      frame.contentDocument.write(source);
      frame.contentDocument.close();
      await new Promise((resolve) => setTimeout(resolve, 100));
      // eslint-disable-next-line no-eval
      const cells = eval(`(${measure})`)(frame.contentDocument.body, frame.contentWindow, true);
      frame.remove();
      return cells;
    },
    [FOOTER, LINES.toString()]
  );

  await page.click('[data-slate-editor]');
  await page.evaluate(async (source) => {
    await navigator.clipboard.write([
      new ClipboardItem({
        'text/html': new Blob([source], { type: 'text/html' }),
        'text/plain': new Blob(['x'], { type: 'text/plain' }),
      }),
    ]);
  }, FOOTER);
  await page.keyboard.press('Control+V');
  await page.waitForTimeout(700);

  const mine = await page.evaluate(
    (measure) =>
      // eslint-disable-next-line no-eval
      eval(`(${measure})`)(document.querySelector('[data-slate-editor]'), window, false),
    LINES.toString()
  );

  const heights = mine.flat().map((line) => line.height);
  check(
    'an indented line stands the same height as the plain line beside it',
    new Set(heights).size === 1,
    `col1 ${mine[0].map((l) => l.height).join('/')} | col2 ${mine[1].map((l) => l.height).join('/')}`
  );
  check(
    "and its indentation is drawn in the document's own spaces, not the editor's",
    mine[0].every((line, i) => Math.abs(line.textAt - truth[0][i].textAt) <= 2),
    `word ${truth[0].map((l) => l.textAt).join('/')} | editor ${mine[0].map((l) => l.textAt).join('/')}`
  );

  await context.close();
  server.close();
}

// Where the text sits BETWEEN the top and bottom of its row.
//
// HTML's own default for a `<td>` is `vertical-align: middle`, so a cell that
// states nothing — most of them — centres its text, in Word, in a browser and
// in the printed report. Here the content box filled the cell (`h-full`, and
// the row's height set on the box rather than on the cell), which left
// `vertical-align` nothing to move: every cell drew its text at the top and all
// the leftover space fell BELOW it, reading as a bottom padding the document
// never had.
{
  const { page, context, server } = await open('');

  const text = (t) => `<span style='font-size:11.0pt'>${t}</span>`;
  const vRow = (cellAttr, rowAttr = '') =>
    `<tr ${rowAttr}>` +
    `<td ${cellAttr} style='border:solid windowtext 1.0pt;padding:0cm 5.4pt'>` +
    ['Measles morbillivirus', 'second', 'third']
      .map((t) => `<p class=MsoNormal>${text(t)}</p>`)
      .join('') +
    `</td>` +
    `<td ${cellAttr} style='border:solid windowtext 1.0pt;padding:0cm 5.4pt'>` +
    `<p class=MsoNormal>${text('Detected')}</p></td></tr>`;
  const VERTICAL =
    `<html xmlns:w="urn:schemas-microsoft-com:office:word"><head><style>` +
    `<!-- p.MsoNormal {margin:0cm; font-size:11.0pt;} --></style></head><body>` +
    `<table class=MsoTableGrid border=1 cellspacing=0 style='border-collapse:collapse'>` +
    vRow('') +
    vRow('valign=top') +
    vRow(`style='vertical-align:middle'`) +
    vRow('', `style='height:70px'`) +
    `</table></body></html>`;
  const CASES = ['nothing stated', 'valign=top', 'vertical-align:middle', 'a row 70px tall'];

  /** For the SHORT cell of each row: the gap above the text and below it. */
  const GAPS = (root, selector) =>
    [...root.querySelectorAll('tr')].map((tr) => {
      const cells = [...tr.children].filter(
        (c) => /^t[dh]$/i.test(c.tagName) && !c.classList.contains('w-2')
      );
      const box = cells[1].getBoundingClientRect();
      const line = cells[1].querySelector(selector).getBoundingClientRect();
      return [Math.round(line.top - box.top), Math.round(box.bottom - line.bottom)];
    });

  const truth = await page.evaluate(
    async ([source, measure]) => {
      const frame = document.createElement('iframe');
      frame.style.cssText = 'position:absolute;left:-9999px;width:1100px;height:900px';
      document.body.append(frame);
      frame.contentDocument.open();
      frame.contentDocument.write(source);
      frame.contentDocument.close();
      await new Promise((resolve) => setTimeout(resolve, 100));
      // eslint-disable-next-line no-eval
      const gaps = eval(`(${measure})`)(frame.contentDocument.body, 'p');
      frame.remove();
      return gaps;
    },
    [VERTICAL, GAPS.toString()]
  );

  await page.click('[data-slate-editor]');
  await page.evaluate(async (source) => {
    await navigator.clipboard.write([
      new ClipboardItem({
        'text/html': new Blob([source], { type: 'text/html' }),
        'text/plain': new Blob(['x'], { type: 'text/plain' }),
      }),
    ]);
  }, VERTICAL);
  await page.keyboard.press('Control+V');
  await page.waitForTimeout(700);

  const mine = await page.evaluate(
    (measure) =>
      // eslint-disable-next-line no-eval
      eval(`(${measure})`)(document.querySelector('[data-slate-editor]'), '.slate-p'),
    GAPS.toString()
  );

  CASES.forEach((name, i) => {
    const [wordTop, wordBottom] = truth[i];
    const [top, bottom] = mine[i];
    const centred = Math.abs(wordTop - wordBottom) <= 2;
    check(
      `${name}: the text sits where the document puts it, top to bottom`,
      centred ? Math.abs(top - bottom) <= 2 : top < bottom,
      `word ${wordTop}/${wordBottom} | editor ${top}/${bottom}`
    );
  });

  await context.close();
  server.close();
}

await browser.close();
console.log(fails === 0 ? '\nALL PASS' : `\n${fails} FAILED`);
process.exit(fails === 0 ? 0 : 1);

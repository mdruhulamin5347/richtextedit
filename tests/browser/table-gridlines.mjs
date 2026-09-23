/**
 * Can you SEE a borderless table?
 *
 * A gridline is drawn on every edge a table leaves blank, so the author can
 * tell where the cells are, and nothing prints it. The reports these
 * tables come from are full of borderless ones — Word's usual way of lining up
 * a label / colon / value column is a table with every rule turned off — and
 * the editor drew them exactly as the printed report shows them, which is to
 * say invisibly: nothing on screen told a three-column table from three runs of
 * text, and there was nothing to aim a column-resize drag at.
 *
 * A browser test because this is a question about pseudo-elements and computed
 * styles, and because the other half of it is that the SAVED report must not
 * gain a single one of these lines.
 */
import { chromium } from 'playwright';
import { harnessHtml, serve } from './harness.mjs';

let fails = 0;
const check = (label, ok, detail) => {
    if (!ok) fails++;
    console.log(`${ok ? 'PASS' : 'FAIL'}  ${label}${detail ? ` — ${detail}` : ''}`);
};

/** Word's spelling of a borderless layout table: the label / colon / value grid. */
const BORDERLESS = `<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 style='border-collapse:collapse;border:none'>
  <tr><td style='border:none;padding:0cm 5.4pt'><p>Rate</p></td><td style='border:none;padding:0cm 5.4pt'><p>:</p></td><td style='border:none;padding:0cm 5.4pt'><p>85 b/min</p></td></tr>
  <tr><td style='border:none;padding:0cm 5.4pt'><p>Rhythm</p></td><td style='border:none;padding:0cm 5.4pt'><p>:</p></td><td style='border:none;padding:0cm 5.4pt'><p>Regular</p></td></tr>
</table>`;

const BORDERED = `<table border=1 cellspacing=0 cellpadding=2>
  <tr><td>a</td><td>b</td></tr><tr><td>c</td><td>d</td></tr>
</table>`;

/**
 * Each cell's four sides, twice over: what the DOCUMENT draws (::before) and
 * what the editor adds to help you read it (::after). Kept apart on purpose —
 * tests/browser/table-border-design.mjs measures the first against a bare
 * browser's rendering of the same markup, where a dash this editor invented has
 * no business.
 */
const MEASURE = () => {
    const SIDES = ['top', 'right', 'bottom', 'left'];
    const rule = (el, pseudo, side) => {
        const cs = getComputedStyle(el, pseudo);
        const width = Math.round(parseFloat(cs.getPropertyValue(`border-${side}-width`)) || 0);
        const style = cs.getPropertyValue(`border-${side}-style`);
        return width === 0 || style === 'none' ? '-' : `${width}px ${style}`;
    };

    const rows = Array.from(document.querySelectorAll('[data-slate-editor] tr'));
    return rows.map((row) =>
        Array.from(row.children)
            .filter((c) => /^t[dh]$/i.test(c.tagName) && !c.classList.contains('w-2'))
            .map((cell) => {
                const box = cell.getBoundingClientRect();
                const cs = getComputedStyle(cell, '::after');
                return {
                    document: SIDES.map((s) => rule(cell, '::before', s)),
                    guide: SIDES.map((s) => rule(cell, '::after', s)),
                    guideColor: cs.borderBottomColor,
                    // Where the dashes actually land. An absolutely positioned
                    // box with no offsets falls back to its STATIC position, and
                    // ::after's is AFTER the content — so without `inset-0` the
                    // guide is drawn a whole cell-height below its cell, which
                    // reads as a stray tick under the table and is invisible to
                    // any check that only asks what the border looks like.
                    onTheCell:
                        cs.top === '0px' &&
                        cs.left === '0px' &&
                        Math.round(parseFloat(cs.width)) === Math.round(box.width) &&
                        Math.round(parseFloat(cs.height)) === Math.round(box.height),
                };
            })
    );
};

const browser = await chromium.launch({ channel: 'chrome' });

async function open(initialHtml) {
    const { server, port } = await serve(() => harnessHtml(initialHtml));
    const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
    page.on('pageerror', (e) => console.log('PAGEERROR', String(e)));
    await page.goto(`http://127.0.0.1:${port}/`);
    await page.evaluate(() => window.__ready);
    await page.waitForTimeout(250);
    return { page, server };
}

// A borderless table: the document draws nothing, the editor draws dashes.
{
    const { page, server } = await open(BORDERLESS);
    const grid = await page.evaluate(`(${MEASURE})()`);

    check('the table is there', grid.length === 2 && grid[0].length === 3, JSON.stringify(grid.map((r) => r.length)));

    const every = (fn) => grid.every((row, r) => row.every((cell, c) => fn(cell, r, c)));

    check(
        'the document still draws no rule anywhere — the report is unchanged',
        every((cell) => cell.document.every((side) => side === '-')),
        JSON.stringify(grid[0][0].document)
    );

    // Plate's collapse emulation hands a cell its bottom and right, plus a top
    // in the first row and a left in the first column, so every edge of the
    // grid is guided exactly once and no interior line is doubled.
    const [top, right, bottom, left] = [0, 1, 2, 3];
    // The same 1px solid rule the table would have if it were bordered — a
    // pencil line, not a different kind of line. Only its colour says it is the
    // editor's and not the document's.
    const drawn = (s) => s === '1px solid';
    check(
        'every cell guides the two edges it owns',
        every((cell) => drawn(cell.guide[bottom]) && drawn(cell.guide[right])),
        JSON.stringify(grid[0][0].guide)
    );
    check(
        'the first row guides the top of the table, and no other row does',
        every((cell, r) => drawn(cell.guide[top]) === (r === 0))
    );
    check(
        'the first column guides the left of the table, and no other column does',
        every((cell, r, c) => drawn(cell.guide[left]) === (c === 0))
    );

    check(
        'and each cell\'s guide is drawn on that cell, not below it',
        every((cell) => cell.onTheCell)
    );

    // A pencil line, not the document's ink: it has to read as lighter than a
    // real rule or the borderless table looks like a bordered one.
    const ink = await page.evaluate(
        () => getComputedStyle(document.querySelector('[data-slate-editor]')).color
    );
    //
    // Composited over the page, because the guide states its colour as an alpha
    // over the editor's ink — and Chrome reports that as `color(srgb r g b / a)`
    // with the channels 0-1, not as an `rgb()` with them 0-255.
    const luminance = (c) => {
        const n = c.match(/[\d.]+/g).map(Number);
        const scale = c.startsWith('color(') ? 255 : 1;
        const [r, g, b] = n.slice(0, 3).map((v) => v * scale);
        const a = n[3] ?? 1;
        const over = (v) => 255 * (1 - a) + v * a;
        return 0.2126 * over(r) + 0.7152 * over(g) + 0.0722 * over(b);
    };
    check(
        'and it is drawn in pencil, far lighter than the text it sits around',
        luminance(grid[0][0].guideColor) > luminance(ink) + 150,
        `${grid[0][0].guideColor} against ${ink}`
    );

    const html = await page.evaluate(() => window.RichTextEdit.getHtml('#report_body_editor'));
    check(
        'and none of it reaches the saved report',
        html.includes('border: 0') && !/border-(top|right|bottom|left):/.test(html),
        html.slice(0, 120)
    );
    check('which still says the table is borderless', html.includes('<table border="0"'), html.slice(0, 80));

    await page.close();
    server.close();
}

// A table with real rules needs no help and gets none.
{
    const { page, server } = await open(BORDERED);
    const grid = await page.evaluate(`(${MEASURE})()`);

    check(
        'a bordered table draws its own rules',
        grid.every((row) => row.every((cell) => cell.document.some((side) => side !== '-'))),
        JSON.stringify(grid[0][0].document)
    );
    check(
        'and gets no guides on top of them',
        grid.every((row) => row.every((cell) => cell.guide.every((side) => side === '-'))),
        JSON.stringify(grid[0][0].guide)
    );

    await page.close();
    server.close();
}

await browser.close();
console.log(fails === 0 ? '\nALL PASS' : `\n${fails} FAILED`);
process.exit(fails === 0 ? 0 : 1);

/**
 * A table the document only asked to PLACE in the middle of the page.
 *
 * The three-signature block at the foot of a report is a three-column table,
 * each cell left aligned, and the table itself centred on the page — which
 * Word, LibreOffice and every legacy report write the pre-CSS way, as
 * `<div align=center>`, `<center>`, or a `text-align:center` on the wrapper.
 *
 * `inlineLegacyAlignment` restates those as CSS and pushes them down onto the
 * blocks inside, because a `<div>` is not a node in this editor and would take
 * its alignment with it when unwrapped. The selector reached straight through
 * the table into the paragraphs in its cells, so a signature block that reads
 * left aligned in the document arrived centred here.
 *
 * A browser does not do that: its own stylesheet says `table { text-align:
 * start }`, so the alignment stops at the table. That is what is measured
 * below — the same markup in a bare iframe, which is both what the document
 * showed the author and what the old Summernote editor put on screen.
 */
import { chromium } from 'playwright';
import { harnessHtml, serve } from './harness.mjs';

/** Left-aligned cells, three columns, no borders — a signature row. */
const SIGNATURES = [
    ['Medical  Technologist', 'Department of Molecular Biology', 'Chittogram Maa Shishu-O-General Hospital'],
    ['Dr. Nahid Sultana', 'MBBS. D. Bact, M.Phil (Microbiology)', 'Professor &amp; Head'],
    ['DR. Sanjoy Kanti Biswas', 'MBBS, MCPS(Cl. Path), M.Phil (Microbiology)', 'Professor'],
];

const TABLE =
    `<table border=0 cellspacing=0 cellpadding=0 width=900 style='border-collapse:collapse'><tr>` +
    SIGNATURES.map(
        (lines) =>
            `<td valign=top style='padding:0cm 5.4pt'>` +
            lines.map((line) => `<p class=MsoNormal>${line}</p>`).join('') +
            `</td>`
    ).join('') +
    `</tr></table>`;

const CASES = [
    ['a table centred with div align=center', `<html><body><div align=center>${TABLE}</div></body></html>`],
    ['a table centred with <center>', `<html><body><center>${TABLE}</center></body></html>`],
    [
        'a table centred with CSS on the wrapper',
        `<html><body><div style='text-align:center'>${TABLE}</div></body></html>`,
    ],
    [
        // The other half of the rule: alignment stated INSIDE the table is the
        // document aligning its text, and must still come through.
        'a row that really does centre its cells',
        `<html><body><table border=0><tr align=center><td><p class=MsoNormal>Dr. Nahid Sultana</p></td></tr></table></body></html>`,
    ],
    [
        'a paragraph beside the table, which the wrapper really does align',
        `<html><body><div align=center><p class=MsoNormal>END OF REPORT</p>${TABLE}</div></body></html>`,
    ],
];

let fails = 0;
function check(label, ok, detail) {
    if (!ok) fails++;
    console.log(`${ok ? 'PASS' : 'FAIL'}  ${label}${detail ? ` — ${detail}` : ''}`);
}

/** Chrome spells a centred table cell `-webkit-center`; it means `center`. */
const normalize = (align) => (align === '-webkit-center' ? 'center' : align);

const { server, port } = await serve(() => harnessHtml(''));
const browser = await chromium.launch({ channel: 'chrome' });
const context = await browser.newContext({
    viewport: { width: 1200, height: 900 },
    permissions: ['clipboard-read', 'clipboard-write'],
});
const page = await context.newPage();
await page.goto(`http://127.0.0.1:${port}/`);
await page.evaluate(() => window.__ready);

for (const [name, html] of CASES) {
    /** What the document showed: the browser rendering the markup untouched. */
    const document_ = await page.evaluate(async (source) => {
        const frame = document.createElement('iframe');
        frame.style.cssText = 'position:absolute;left:-9999px;width:1200px;height:800px';
        document.body.append(frame);
        frame.contentDocument.open();
        frame.contentDocument.write(source);
        frame.contentDocument.close();
        await new Promise((resolve) => setTimeout(resolve, 80));
        const cell = frame.contentDocument.querySelector('td p');
        const loose = frame.contentDocument.querySelector('div > p');
        const read = (node) => (node ? frame.contentWindow.getComputedStyle(node).textAlign : null);
        const result = { cell: read(cell), loose: read(loose) };
        frame.remove();
        return result;
    }, html);

    await page.evaluate(() => window.RichTextEdit.setHtml('#report_body_editor', '<p></p>'));
    await page.click('[data-slate-editor]');
    await page.evaluate(async (source) => {
        await navigator.clipboard.write([
            new ClipboardItem({
                'text/html': new Blob([source], { type: 'text/html' }),
                'text/plain': new Blob([''], { type: 'text/plain' }),
            }),
        ]);
    }, html);
    await page.keyboard.press('Control+V');
    await page.waitForTimeout(400);

    const editor = await page.evaluate(() => {
        const editable = document.querySelector('[data-slate-editor]');
        const cell = editable.querySelector('td [data-slate-node="element"]');
        const loose = Array.from(editable.querySelectorAll(':scope > [data-slate-node="element"]')).find(
            (block) => block.textContent.trim()
        );
        const read = (node) => (node ? getComputedStyle(node).textAlign : null);
        return { cell: read(cell), loose: read(loose) };
    });

    check(
        `${name}: the cells read as the document reads them`,
        normalize(editor.cell) === normalize(document_.cell),
        `document ${document_.cell} | editor ${editor.cell}`
    );

    if (document_.loose) {
        check(
            `${name}: and a paragraph outside the table still takes the alignment`,
            normalize(editor.loose) === normalize(document_.loose),
            `document ${document_.loose} | editor ${editor.loose}`
        );
    }
}

await browser.close();
server.close();
console.log(fails === 0 ? '\nAll checks passed.' : `\n${fails} check(s) failed.`);
process.exit(fails === 0 ? 0 : 1);

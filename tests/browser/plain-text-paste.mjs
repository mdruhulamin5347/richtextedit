/**
 * A lab report pasted as PLAIN TEXT — does it still say the same thing once it
 * is printed?
 *
 * The reference is the browser's own contenteditable, and that is not an
 * approximation of the old editor, it IS the old editor: Summernote's paste
 * handler in `summernote_table_resize.blade.php` returns early when the
 * clipboard carries no HTML ("plain-text paste -> let the browser handle it").
 * So whatever Chrome writes into a bare contenteditable is exactly what these
 * reports used to be saved as, and what the print pages have always rendered.
 *
 * What was wrong: this editor kept the tabs and the space runs in its model and
 * showed them correctly — it is `white-space: pre-wrap` — and then serialized
 * them raw. HTML collapses every run to one space and drops it at the edges of
 * a block, so `<p> </p>` was an empty block and `<p>        RESULT</p>` was a
 * flush-left heading. The report read right in the editor and printed flat.
 *
 * Everything below is measured in the PRINT context — the report body dropped
 * into `.report_body.main_data` with the rules `ms_format/print.blade.php`
 * applies — because that is where the difference showed.
 */
import { chromium } from 'playwright';
import { harnessHtml, serve } from './harness.mjs';

/** The demo report: two tab-set columns, an indented heading, a spacer line. */
const TEXT = [
    'Test Name\t\t:  RT-PCR FOR COVID-19',
    'Specimen\t\t:  Nasopharyngeal Swab/Oropharyngeal Swab',
    'Collection Site\t\t: Chattogram Maa Shishu-O-General Hospital',
    '-------------------------------------------------------------------',
    ' ',
    '                                                       RESULT',
    'Test Date\t: 26-01-2024',
    'Test Method\t:  rRT-PCR',
    'Result\t\t: Negative',
    'Comment\t: Please correlate clinically.',
].join('\n');

/** A report the old editor saved: its indentation is a run of &nbsp;. */
const LEGACY = '<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;RESULT</p><p>&nbsp;</p>';

let fails = 0;
function check(label, ok, detail) {
    if (!ok) fails++;
    console.log(`${ok ? 'PASS' : 'FAIL'}  ${label}${detail ? ` — ${detail}` : ''}`);
}

const { server, port } = await serve(() => harnessHtml(''));
const browser = await chromium.launch({ channel: 'chrome' });
const context = await browser.newContext({
    viewport: { width: 1440, height: 900 },
    permissions: ['clipboard-read', 'clipboard-write'],
});
const page = await context.newPage();
await page.goto(`http://127.0.0.1:${port}/`);
await page.evaluate(() => window.__ready);

/** Measuring tools, installed once in the page. */
await page.evaluate(() => {
    /** Where a character sits, relative to the block it is in. */
    window.__xOf = (block, char) => {
        const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
        for (let node = walker.nextNode(); node; node = walker.nextNode()) {
            const index = node.data.indexOf(char);
            if (index < 0) continue;
            const range = block.ownerDocument.createRange();
            range.setStart(node, index);
            range.setEnd(node, index + 1);
            return Math.round(range.getBoundingClientRect().left - block.getBoundingClientRect().left);
        }
        return null;
    };

    /** The report body as a print page renders it. */
    window.__printed = async (html) => {
        const frame = document.createElement('iframe');
        frame.style.cssText = 'position:absolute;left:-9999px;width:1000px;height:1400px';
        document.body.append(frame);
        frame.contentDocument.open();
        frame.contentDocument.write(
            `<!doctype html><html><head><style>
              body { margin:0; font-size:15px; font-family:Calibri,"Arial Narrow",Arial,sans-serif }
              .main_data p { margin:0; padding:0 }
            </style></head><body><div class="report_body main_data" style="width:900px">${html}</div></body></html>`
        );
        frame.contentDocument.close();
        await new Promise((resolve) => setTimeout(resolve, 100));
        const root = frame.contentDocument.querySelector('.main_data');
        const lines = Array.from(root.children).map((block) => ({
            text: block.textContent,
            height: Math.round(block.getBoundingClientRect().height),
            colon: window.__xOf(block, ':'),
            firstLetter: window.__xOf(block, 'R'),
        }));
        frame.remove();
        return lines;
    };
});

/** This editor, pasting the report. The caret goes in first: writing to the
 *  clipboard needs the document focused. */
await page.click('[data-slate-editor]');
await page.evaluate(async (text) => {
    await navigator.clipboard.write([
        new ClipboardItem({ 'text/plain': new Blob([text], { type: 'text/plain' }) }),
    ]);
}, TEXT);
await page.keyboard.press('Control+V');
await page.waitForTimeout(600);

const editor = await page.evaluate(() => ({
    saved: window.RadiolensEditor.getHtml('#report_body_editor'),
    onScreen: Array.from(document.querySelectorAll('[data-slate-editor] > *')).map((block) => ({
        text: block.textContent,
        colon: window.__xOf(block, ':'),
    })),
}));

/** What the old editor produced: Chrome pasting the same thing, untouched. */
await page.evaluate(() => {
    const box = document.createElement('div');
    box.id = 'native';
    box.contentEditable = 'true';
    box.style.cssText = 'width:900px;font-size:15px;font-family:Calibri,Arial,sans-serif';
    document.body.append(box);
});
await page.click('#native');
await page.keyboard.press('Control+V');
await page.waitForTimeout(300);
const reference = await page.evaluate(() => {
    const box = document.querySelector('#native');
    const html = box.innerHTML;
    box.remove();
    return html;
});

const printed = await page.evaluate((html) => window.__printed(html), editor.saved);
const asBefore = await page.evaluate((html) => window.__printed(html), reference);

const lines = TEXT.split('\n');

/**
 * A run is saved as alternating U+00A0 and space — that is the only spelling
 * HTML keeps — so the text is compared as the line reads, not byte for byte.
 */
const unpadded = (text) => text.replaceAll('\u00A0', ' ');

check(
    'every line of the report is still a line of the printed report',
    printed.length === lines.length,
    `${printed.length} blocks, ${lines.length} lines`
);

check(
    'no line lost a character of its own spacing',
    printed.every((line, i) => unpadded(line.text) === lines[i]),
    printed.find((line, i) => unpadded(line.text) !== lines[i])?.text
);

const spacer = printed[4];
check(
    'the blank spacer line still opens a gap',
    spacer && spacer.height > 10,
    `${spacer?.height}px (a collapsed one is 0)`
);

check(
    'every line stands the same height',
    new Set(printed.map((line) => line.height)).size === 1,
    [...new Set(printed.map((line) => line.height))].join(', ')
);

check(
    'the lines sit exactly as high as the old editor printed them',
    printed.every((line, i) => Math.abs(line.height - asBefore[i].height) <= 1),
    `${printed[0].height}px vs ${asBefore[0].height}px`
);

const columns = [0, 1, 2, 6, 7, 8, 9];
check(
    'the tabs still carry the value column to where the old editor put it',
    columns.every((i) => Math.abs(printed[i].colon - asBefore[i].colon) <= 1),
    columns.map((i) => `${printed[i].colon}/${asBefore[i].colon}`).join(' ')
);

check(
    'the tabbed columns are a real gap, not the single space HTML collapses them to',
    printed[0].colon > 100,
    `the colon sits ${printed[0].colon}px in`
);

const heading = printed[5];
check(
    'the indented heading is still indented, to the same place',
    Math.abs(heading.firstLetter - asBefore[5].firstLetter) <= 2 && heading.firstLetter > 150,
    `${heading.firstLetter}px vs ${asBefore[5].firstLetter}px`
);

check(
    'what the editor shows is what gets printed: same column, same order',
    editor.onScreen.length === lines.length &&
        editor.onScreen[0].colon > 100 &&
        editor.onScreen[0].colon < editor.onScreen[2].colon,
    editor.onScreen
        .slice(0, 3)
        .map((line) => line.colon)
        .join(' ')
);

// The other half: a report the OLD editor saved, opened here. Plate collapses
// whitespace with JavaScript's \s, which counts U+00A0 where CSS does not, so
// every one of these arrived with its indentation stripped.
const legacy = await page.evaluate(async (html) => {
    window.RadiolensEditor.setHtml('#report_body_editor', html);
    await new Promise((resolve) => setTimeout(resolve, 300));
    const blocks = Array.from(document.querySelectorAll('[data-slate-editor] > *'));
    return {
        indent: window.__xOf(blocks[0], 'R'),
        blocks: blocks.length,
        saved: window.RadiolensEditor.getHtml('#report_body_editor'),
    };
}, LEGACY);

check(
    // Collapsed, a leading run is not narrower — it is GONE, and the heading
    // starts at 0. Eight spaces of the editor's 18px text is about 37.
    'a legacy report keeps the indentation it was written with',
    legacy.indent > 25,
    `heading starts ${legacy.indent}px in`
);

check(
    'and is saved back with it, rather than flattened on the first edit',
    legacy.saved.includes('&nbsp;') && legacy.saved.includes('<p>&nbsp;</p>'),
    legacy.saved
);

await browser.close();
server.close();
console.log(fails === 0 ? '\nAll checks passed.' : `\n${fails} check(s) failed.`);
process.exit(fails === 0 ? 0 : 1);

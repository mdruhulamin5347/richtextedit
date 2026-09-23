/**
 * Does the font-size box show the number the document's author set?
 *
 * Reported by name: "currently main doc has font size 10 but after paste our
 * editor show 13". Both numbers described the same text. Word measures type in
 * POINTS and the editor used to convert every pasted size into whole CSS
 * pixels, so a 10pt paragraph arrived as `13px` — and 13px is not even 10pt,
 * which is 13.333px, so the text was drawn 2.5% small as well as misnamed.
 *
 * Only a browser can answer this. The number comes out of the control's own
 * React state, which needs the caret in real text with a real selection, and
 * the size it should agree with is what `getComputedStyle` says the run is
 * actually drawn at. jsdom has neither.
 *
 * What is pinned here, in the order a reader meets it:
 *
 *  - a Word paste is NAMED on Word's scale (10pt reads 10) and DRAWN at Word's
 *    size (13.33px, to a hundredth of a pixel);
 *  - a legacy report stating px is named on that same one scale, so the box
 *    never shows two different scales in one document;
 *  - setting a size writes points, and the ± steps by a whole point.
 *
 * See lib/font-size.ts, which owns the rule, and
 * tests/browser/font-size-fidelity.mjs, which compares every run in a document
 * against a bare iframe.
 */
import { chromium } from 'playwright';
import { harnessHtml, serve } from './harness.mjs';

let fails = 0;
const check = (label, ok, detail) => {
    if (!ok) fails++;
    console.log(`${ok ? 'PASS' : 'FAIL'}  ${label}${detail ? ` — ${detail}` : ''}`);
};

/** A Word clipboard whose body is 10pt, stated the way Word states it. */
const WORD_10PT =
    `<html xmlns:w="urn:schemas-microsoft-com:office:word"><head><style><!--
p.MsoNormal {margin:0cm; font-size:10.0pt; font-family:"Calibri",sans-serif;}
--></style></head><body>` +
    `<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Calibri",sans-serif'>` +
    `Impression: no acute abnormality.</span></p></body></html>`;

const browser = await chromium.launch({ channel: 'chrome' });

async function open(initialHtml) {
    const { server, port } = await serve(() => harnessHtml(initialHtml));
    const context = await browser.newContext({
        viewport: { width: 1280, height: 900 },
        permissions: ['clipboard-read', 'clipboard-write'],
    });
    const page = await context.newPage();
    page.on('pageerror', (e) => console.log('PAGEERROR', String(e)));
    await page.goto(`http://127.0.0.1:${port}/`);
    await page.evaluate(() => window.__ready);
    await page.waitForTimeout(200);
    return { page, context, server };
}

/**
 * The font-size box. It is the only text input in the toolbar — ToolbarButton
 * puts a control's label in a Radix tooltip, which is not an accessible name.
 */
const BOX = '.rte-scope input[type="text"][data-plate-focus="true"]';

/**
 * Select the first line of text and let the control catch up.
 *
 * A click alone lands the caret wherever the pointer did, which on an empty
 * part of the line is not in the run at all. A SELECTION is what makes
 * `editor.api.marks()` report the run's own size rather than the block's, and
 * taking the whole line means a size set here lands on the whole run rather
 * than on one character.
 */
async function selectFirstLine(page) {
    await page.click('[data-slate-editor] [data-slate-string]');
    await page.keyboard.press('Home');
    await page.keyboard.down('Shift');
    await page.keyboard.press('End');
    await page.keyboard.up('Shift');
    await page.waitForTimeout(250);
}

/**
 * The font-size widget's own + / - buttons.
 *
 * Scoped to the div that holds the box: the Insert menu's trigger is a Plus
 * icon too, and it comes first in the toolbar.
 */
const stepper = (page, icon) =>
    page
        .locator(`.rte-scope div:has(> input[data-plate-focus="true"]) > button:has(svg.lucide-${icon})`)
        .first();

/** The size, to a hundredth of a pixel, the first run is actually drawn at. */
const drawnPx = (page) =>
    page.evaluate(() => {
        const el = document.querySelector('[data-slate-editor] [data-slate-string]');
        return Math.round(parseFloat(getComputedStyle(el).fontSize) * 100) / 100;
    });

async function pasteWord(page, source) {
    await page.click('[data-slate-editor]');
    await page.evaluate(async (src) => {
        await navigator.clipboard.write([
            new ClipboardItem({
                'text/html': new Blob([src], { type: 'text/html' }),
                'text/plain': new Blob(['x'], { type: 'text/plain' }),
            }),
        ]);
    }, source);
    await page.keyboard.press('Control+V');
    await page.waitForTimeout(700);
}

// A 10pt Word paragraph is named 10 and drawn at 10pt. The whole complaint.
{
    const { page, context, server } = await open('');
    await pasteWord(page, WORD_10PT);
    await selectFirstLine(page);

    const shown = await page.inputValue(BOX);
    check('a 10pt Word paragraph is named 10, not 13', shown === '10', `box shows ${shown}`);

    const px = await drawnPx(page);
    check(
        'and is drawn at 10pt exactly, which is 13.33px at 96dpi',
        Math.abs(px - (10 * 96) / 72) <= 0.01,
        `${px}px`
    );

    const saved = await page.evaluate(() =>
        window.RichTextEdit.getHtml('#report_body_editor')
    );
    check(
        'and is SAVED in the points the document stated, so the print matches',
        saved.includes('font-size: 10pt') && !/font-size:\s*13px/.test(saved),
        saved.slice(0, 120)
    );

    await context.close();
    server.close();
}

// A legacy report states px. It is left in px and named on the one scale.
{
    const { page, context, server } = await open(
        '<p style="font-size: 15px"><span style="font-size: 15px">Legacy body text</span></p>'
    );
    await selectFirstLine(page);

    const shown = await page.inputValue(BOX);
    check(
        'a stored 15px run is named 11.25 — what 15px is in points',
        shown === '11.25',
        `box shows ${shown}`
    );

    const px = await drawnPx(page);
    check('and is still drawn at the 15px it was saved at', Math.abs(px - 15) <= 0.01, `${px}px`);

    // The size on the node is untouched, so opening a stored report and saving
    // it does not rewrite every size in it. ~70 live installations have these.
    const saved = await page.evaluate(() =>
        window.RichTextEdit.getHtml('#report_body_editor')
    );
    check(
        'and open + save leaves the px in the report, not points',
        saved.includes('font-size: 15px') && !saved.includes('11.25pt'),
        saved.slice(0, 120)
    );

    await context.close();
    server.close();
}

// Setting a size writes points, and ± steps by a whole point.
{
    const { page, context, server } = await open('<p><span>Findings</span></p>');
    await selectFirstLine(page);

    const base = await page.inputValue(BOX);
    check(
        "an unsized run is named 13.5 — the editable's own 18px, in points",
        base === '13.5',
        `box shows ${base}`
    );

    await page.fill(BOX, '12');
    await page.keyboard.press('Enter');
    await page.waitForTimeout(300);
    await selectFirstLine(page);

    const set = await page.inputValue(BOX);
    const setPx = await drawnPx(page);
    check(
        'typing 12 sets 12 POINTS, drawn at 16px',
        set === '12' && Math.abs(setPx - 16) <= 0.01,
        `box ${set}, drawn ${setPx}px`
    );

    // The + button, which steps the number the box shows.
    await stepper(page, 'plus').click();
    await page.waitForTimeout(300);
    await selectFirstLine(page);

    const stepped = await page.inputValue(BOX);
    const steppedPx = await drawnPx(page);
    check(
        'and + steps it by one point, to 13pt',
        stepped === '13' && Math.abs(steppedPx - (13 * 96) / 72) <= 0.01,
        `box ${stepped}, drawn ${steppedPx}px`
    );

    const saved = await page.evaluate(() =>
        window.RichTextEdit.getHtml('#report_body_editor')
    );
    check('and what it saved is points', saved.includes('font-size: 13pt'), saved.slice(0, 120));

    await context.close();
    server.close();
}

await browser.close();
console.log(fails === 0 ? '\nALL PASS' : `\n${fails} FAILED`);
process.exit(fails === 0 ? 0 : 1);

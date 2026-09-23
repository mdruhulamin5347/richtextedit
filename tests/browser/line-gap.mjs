/**
 * Does a line gap draw what Word and LibreOffice draw?
 *
 * Only a real browser can answer. A gap is a multiple of the font's NATURAL
 * line — what `line-height: normal` resolves to — and that is a property of the
 * font FILE, so it needs a browser with the font installed and a layout engine
 * to ask. jsdom has neither: every rect there is zero, which is why
 * tests/js/line-gap.test.ts can only test the arithmetic.
 *
 * CSS `line-height: N` multiplies the font SIZE instead, which on Calibri is
 * 17% tighter at every step. A report typed at "1.5 lines" opened here at what
 * Word calls 1.23, and the error compounded down a table.
 */
import { chromium } from 'playwright';
import { harnessHtml, serve } from './harness.mjs';

let fails = 0;
function check(label, ok, detail) {
    if (!ok) fails++;
    console.log(`${ok ? 'PASS' : 'FAIL'}  ${label}${detail ? ` — ${detail}` : ''}`);
}

const browser = await chromium.launch({ channel: 'chrome' });

/** What the browser says a line of this font is, the way the editor asks. */
async function naturalLineHeight(page, stack) {
    return page.evaluate((family) => {
        const d = document.createElement('div');
        d.textContent = 'Hg';
        d.style.cssText = `position:absolute;left:-9999px;font-family:${family};font-size:1000px;line-height:normal`;
        document.body.append(d);
        const ratio = d.getBoundingClientRect().height / 1000;
        d.remove();
        return ratio;
    }, stack);
}

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

const CAL = `'Calibri', sans-serif`;

/**
 * The gap the control shows as selected, with the caret in the first table
 * cell. The control is identified by its icon: ToolbarButton puts the label in
 * a Radix tooltip, which is not an accessible name on the button.
 */
async function lineGapMenu(page) {
    await page.click('[data-slate-editor] td [data-slate-node="element"]');
    await page.waitForTimeout(150);
    await page.locator('button:has(svg.lucide-wrap-text)').first().click();
    await page.waitForTimeout(300);
    const ticked = await page.evaluate(() =>
        [...document.querySelectorAll('[role="menuitemradio"]')]
            .filter((item) => item.getAttribute('aria-checked') === 'true')
            .map((item) => item.textContent.trim())
    );
    await page.keyboard.press('Escape');
    await page.waitForTimeout(120);
    return ticked;
}

// A document pasted out of Word opens at the spacing Word drew it at.
{
    const { page, context, server } = await open('');
    const natural = await naturalLineHeight(page, CAL);
    check(
        'the browser can measure a natural line at all',
        natural > 1 && natural < 2,
        `Calibri natural = ${natural}`
    );

    const wordDoc = (lh) =>
        `<html xmlns:w="urn:schemas-microsoft-com:office:word"><head><style><!--
p.MsoNormal {margin:0cm; font-size:11.0pt; font-family:"Calibri",sans-serif;}
--></style></head><body>` +
        `<table class=MsoTableGrid border=1 cellspacing=0 style='border-collapse:collapse'>` +
        `<tr><td style='border:solid windowtext 1.0pt;padding:0cm 5.4pt'>` +
        `<p class=MsoNormal style='margin:0cm;line-height:${lh}'>` +
        `<span style='font-size:11.0pt;font-family:"Calibri",sans-serif'>Gestational Age</span></p>` +
        `</td></tr></table></body></html>`;

    /*
     * `normal` is the one Word spelling this editor does NOT reproduce, and
     * that is a decision rather than a gap in the conversion.
     *
     * It names no multiple — it says "ask the font" — so a document that only
     * says that is one whose gap was not collected, and it opened here at a
     * gap of 1. Asked for on 2026-09-16 ("doc file any line gap table copy and
     * after paste show our editor still line gap 1"), and chosen over reading
     * Word's silence as its Single with the ~4px a row that costs put to the
     * user first. It takes DEFAULT_PASTED_LINE_GAP instead.
     */
    const DEFAULT_PASTED_LINE_GAP = 1.2;

    /*
     * The third column is what the CONTROL has to show as ticked once the
     * paste has landed — the gap the document was written at, snapped to the
     * nearest the menu offers (1.15 is Word's own "Multiple 1.15" and is not
     * one of them). A reader opening a pasted report sees the spacing it is
     * drawn at only if the menu agrees with the page.
     */
    for (const [name, stated, gap, shown] of [
        ['Single (normal, so the default gap)', 'normal', DEFAULT_PASTED_LINE_GAP, '1.2'],
        ['Multiple 1.15', '115%', 1.15, '1.2'],
        ['1.5 lines', '150%', 1.5, '1.5'],
        ['Double', '200%', 2, '2'],
    ]) {
        await page.evaluate(() => window.RadiolensEditor.setHtml('#report_body_editor', ''));
        await page.click('[data-slate-editor]');
        await page.evaluate(async (src) => {
            await navigator.clipboard.write([
                new ClipboardItem({
                    'text/html': new Blob([src], { type: 'text/html' }),
                    'text/plain': new Blob(['x'], { type: 'text/plain' }),
                }),
            ]);
        }, wordDoc(stated));
        await page.keyboard.press('Control+V');
        await page.waitForTimeout(700);

        const got = await page.evaluate(() => {
            const p = document.querySelector('[data-slate-editor] td [data-slate-node="element"]');
            const cs = getComputedStyle(p);
            return { line: cs.lineHeight, size: Number.parseFloat(cs.fontSize) };
        });
        // `normal` is left as `normal` on purpose: it already means the font's
        // natural line in CSS, which is exactly what Single means in Word.
        const drawn =
            got.line === 'normal' ? natural * got.size : Number.parseFloat(got.line);
        const word = gap * natural * got.size;
        check(
            `Word "${name}" opens at the line it should`,
            Math.abs(drawn - word) <= 0.6,
            `Word ${word.toFixed(2)}px | editor ${drawn.toFixed(2)}px`
        );

        const ticked = await lineGapMenu(page);
        check(
            `Word "${name}" is what the control shows as selected`,
            ticked.length === 1 && ticked[0] === shown,
            `expected ${shown}, menu shows ${JSON.stringify(ticked)}`
        );
    }

    await context.close();
    server.close();
}

// Picking a gap in the control draws the same line Word's own setting does.
{
    const cell = `<td style="border: 1px solid #000; padding: 0px 7.2px">` +
        `<p style="font-size: 15px; font-family: ${CAL}">` +
        `<span style="font-size: 15px; font-family: ${CAL}">Gestational Age</span></p></td>`;
    const { page, context, server } = await open(
        `<table border="1" style="border-collapse: collapse; width: 100%;"><tr>${cell}</tr></table>`
    );
    const natural = await naturalLineHeight(page, CAL);

    for (const gap of [1, 1.2, 1.5, 2]) {
        await page.click('[data-slate-editor] td [data-slate-node="element"]');
        await page.waitForTimeout(120);
        // The control is identified by its icon: ToolbarButton puts the label
        // in a Radix tooltip, which is not an accessible name on the button.
        await page.locator('button:has(svg.lucide-wrap-text)').first().click();
        await page.waitForTimeout(300);
        await page
            .locator('[role="menuitemradio"]')
            .filter({ hasText: new RegExp(`^${String(gap).replace('.', '\\.')}$`) })
            .first()
            .click();
        await page.waitForTimeout(350);

        const got = await page.evaluate(() => {
            const p = document.querySelector('[data-slate-editor] td [data-slate-node="element"]');
            const cs = getComputedStyle(p);
            return { line: Number.parseFloat(cs.lineHeight), size: Number.parseFloat(cs.fontSize) };
        });
        const word = gap * natural * got.size;
        check(
            `a line gap of ${gap} draws the line Word draws at ${gap}`,
            Math.abs(got.line - word) <= 0.6,
            `Word ${word.toFixed(2)}px | editor ${got.line.toFixed(2)}px`
        );
    }

    // Whatever the control stored has to be a plain ratio: the printed page
    // runs no script, so it can only be right if the number is already right.
    const saved = await page.evaluate(() =>
        window.RadiolensEditor.getHtml('#report_body_editor')
    );
    const stated = saved.match(/line-height:\s*([\d.]+)/);
    check(
        'and saves it as a plain ratio the print page needs no script to draw',
        !!stated && Math.abs(Number(stated[1]) - 2 * natural) <= 0.01,
        stated?.[0]
    );

    await context.close();
    server.close();
}

/*
 * A gap the document states ABOVE the paragraph — on the cell, the row, the
 * table, a wrapper, or a `<style>` rule Juice inlines onto the cell.
 *
 * `line-height` INHERITS, and unlike `text-align` a table does not break the
 * chain, so all of these ARE the gap the document draws that paragraph at. Only
 * a block reaches this editor carrying one though, so they used to arrive with
 * no gap at all: the paragraph fell back to the cell's own `normal` and every
 * one of them opened at a gap of 1, whatever the document said.
 */
{
    const { page, context, server } = await open('');
    const natural = await naturalLineHeight(page, CAL);
    const RUN = `<span style='font-size:11.0pt;font-family:"Calibri",sans-serif'>Gestational Age</span>`;

    /*
     * The third column is what the CONTROL shows, and it is not always the
     * gap: a gap is a multiple of the font's natural line, so the menu can only
     * name it once the text is drawn in the font the document set. Word's own
     * clipboards restate that font on the run (Plate's docx cleaner copies a
     * paragraph's marks down), and a spreadsheet's does not — it states the
     * font on the `<td>`, which reaches no Plate node, so the run falls back to
     * the editor's own face and the menu measures the gap against THAT. The
     * ROW still stands as tall as the document made it, which is what a table
     * is pasted for; naming it in the menu needs a cell's font-family restated
     * on its runs the way `lib/inherited-font-size.ts` restates its size, and
     * that is a change to make on its own.
     */
    for (const [name, src, shown] of [
        [
            'on the cell',
            `<html xmlns:w="urn:schemas-microsoft-com:office:word"><body><table border=1 cellspacing=0 style='border-collapse:collapse'>` +
                `<tr><td style='padding:0cm 5.4pt;line-height:150%'><p class=MsoNormal>${RUN}</p></td></tr></table></body></html>`,
            '1.5',
        ],
        [
            'on the row',
            `<html xmlns:w="urn:schemas-microsoft-com:office:word"><body><table border=1 cellspacing=0 style='border-collapse:collapse'>` +
                `<tr style='line-height:150%'><td style='padding:0cm 5.4pt'><p class=MsoNormal>${RUN}</p></td></tr></table></body></html>`,
            '1.5',
        ],
        [
            'in a <style> rule on the cell',
            `<html xmlns:w="urn:schemas-microsoft-com:office:word"><head><style><!--
p.MsoNormal {margin:0cm; font-size:11.0pt; font-family:"Calibri",sans-serif;}
td.Spaced {line-height:150%;}
--></style></head><body><table border=1 cellspacing=0 style='border-collapse:collapse'>` +
                `<tr><td class=Spaced style='padding:0cm 5.4pt'><p class=MsoNormal>${RUN}</p></td></tr></table></body></html>`,
            '1.5',
        ],
        [
            "on a spreadsheet's cell, which holds no paragraph at all",
            `<html xmlns:x="urn:schemas-microsoft-com:office:excel"><head><style><!--
td {font-size:11.0pt; font-family:Calibri, sans-serif; line-height:150%;}
--></style></head><body><table border=0 cellspacing=0><tr><td>Gestational Age</td></tr></table></body></html>`,
            // The cell's font reaches no run, so the menu measures the gap
            // against the editor's own face and names 1.2 for the line Word
            // draws at 1.5. The line itself is right, which is the row height.
            '1.2',
        ],
    ]) {
        await page.evaluate(() => window.RadiolensEditor.setHtml('#report_body_editor', ''));
        await page.click('[data-slate-editor]');
        await page.evaluate(async (html) => {
            await navigator.clipboard.write([
                new ClipboardItem({
                    'text/html': new Blob([html], { type: 'text/html' }),
                    'text/plain': new Blob(['x'], { type: 'text/plain' }),
                }),
            ]);
        }, src);
        await page.keyboard.press('Control+V');
        await page.waitForTimeout(700);

        const got = await page.evaluate(() => {
            const p = document.querySelector('[data-slate-editor] td [data-slate-node="element"]');
            if (!p) return null;
            const cs = getComputedStyle(p);
            return { line: cs.lineHeight, size: Number.parseFloat(cs.fontSize) };
        });
        const drawn = !got
            ? NaN
            : got.line === 'normal'
              ? natural * got.size
              : Number.parseFloat(got.line);
        const word = 1.5 * natural * (got?.size ?? 0);
        check(
            `a 1.5-lines gap stated ${name} opens at the line it should`,
            Math.abs(drawn - word) <= 0.6,
            `Word ${word.toFixed(2)}px | editor ${drawn.toFixed(2)}px`
        );

        const ticked = await lineGapMenu(page);
        check(
            `...and the control names it ${shown}`,
            ticked.length === 1 && ticked[0] === shown,
            `expected ${shown}, menu shows ${JSON.stringify(ticked)}`
        );
    }

    await context.close();
    server.close();
}

await browser.close();
console.log(fails ? `\n${fails} FAILED` : '\nall passed');
process.exit(fails ? 1 : 0);

/**
 * Is a blank line as tall as the text around it?
 *
 * Only a real browser can answer: the height of an empty line is its STRUT, the
 * invisible line box the browser gives the element, and jsdom does no layout.
 * Every blank line used to measure the same — the editable's 18px base — so a
 * 10px report and a 30px report had the identical gap between paragraphs.
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
    page.on('pageerror', (e) => console.log('PAGEERROR', String(e)));
    await page.goto(`http://127.0.0.1:${port}/`);
    await page.evaluate(() => window.__ready);
    await page.waitForTimeout(200);
    return { page, server };
}

const lines = (page) =>
    page.evaluate(() =>
        Array.from(document.querySelector('[data-slate-editor]').children).map((el) => ({
            text: el.textContent.trim(),
            height: Math.round(el.getBoundingClientRect().height),
            fontSize: getComputedStyle(el).fontSize,
        }))
    );

const sized = (px, text) => `<p><span style="font-size: ${px}px">${text}</span></p>`;
const BLANK = '<p><br/></p>';

// A blank line between sized text, on the way in.
{
    const { page, server } = await open(
        sized(10, 'ten') + BLANK + sized(30, 'thirty') + BLANK + sized(30, 'thirty again')
    );
    const [, small, , large] = await lines(page);

    check('a blank line after 10px text is drawn at 10px', small.fontSize === '10px', small.fontSize);
    check('a blank line after 30px text is drawn at 30px', large.fontSize === '30px', large.fontSize);
    check(
        'so the two gaps are no longer the same height',
        large.height > small.height,
        `${small.height}px vs ${large.height}px`
    );

    await page.close();
    server.close();
}

// A report that states no sizes is untouched: the base, as before.
{
    const { page, server } = await open('<p>plain</p>' + BLANK + '<p>more</p>');
    const [, blank] = await lines(page);
    check('an unsized report keeps the editor base', blank.fontSize === '18px', blank.fontSize);
    await page.close();
    server.close();
}

// Typing: Enter after sized text, and the toolbar aimed at a blank line.
{
    const { page, server } = await open(sized(10, 'ten'));

    await page.click('[data-slate-editor] > div');
    await page.keyboard.press('End');
    await page.keyboard.press('Enter');
    await page.waitForTimeout(200);

    const typed = (await lines(page))[1];
    check(
        'Enter at the end of a 10px line makes a 10px blank line',
        typed.fontSize === '10px',
        typed.fontSize
    );

    // The caret is on that blank line; set 36 from the toolbar.
    //
    // The box reads and writes POINTS, the scale Word states type on (see
    // lib/font-size.ts), so 36 there is 36pt — which `getComputedStyle` reports
    // back as the 48px it resolves to.
    const input = page.locator('input[data-plate-focus="true"]');
    await input.click();
    await input.fill('36');
    await input.press('Enter');
    await page.waitForTimeout(250);

    const resized = (await lines(page))[1];
    check(
        'setting a size with the caret on a blank line resizes that line',
        resized.fontSize === '48px',
        resized.fontSize
    );
    check(
        'and it is taller than the 10px line it followed',
        resized.height > typed.height,
        `${typed.height}px -> ${resized.height}px`
    );

    const html = await page.evaluate(() =>
        window.RichTextEdit.getHtml('#report_body_editor')
    );
    check(
        'the saved report states the gap it prints with',
        html.includes('<p style="font-size: 36pt">'),
        html
    );

    await page.close();
    server.close();
}

await browser.close();
console.log(fails === 0 ? '\nALL PASS' : `\n${fails} FAILED`);
process.exit(fails === 0 ? 0 : 1);

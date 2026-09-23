/**
 * Pasting a picture, in a real browser.
 *
 * jsdom cannot answer this one: the file path runs through a paste HANDLER, a
 * placeholder node and an async read of the File, none of which happen there.
 *
 * What was wrong: Word puts three things on the clipboard at once — HTML that
 * names the picture by a path on the machine doing the copying
 * (`file:///C:/…/clip_image001.png`), the bytes in the RTF flavour, and an
 * image rendition of the selection. Plate inserts a pasted image file only when
 * the clipboard carries NO html:
 *
 *   if (files.length > 0 && !types.includes('text/html')) …
 *
 * Word's clipboard has html, so the files were ignored; the html named a file
 * the browser may not read; and the paste inserted NOTHING AT ALL, with the
 * bytes sitting on the clipboard the whole time.
 *
 * The rule has to stay narrow, and case C is why: what Windows renders for a
 * selection is a picture of the WHOLE selection, so a header table would come
 * back as one flat image of the entire header in place of the report.
 */
import { chromium } from 'playwright';
import { harnessHtml, serve } from './harness.mjs';

/** A 4x4 PNG — the QR code, as far as this test is concerned. */
const PNG_B64 =
    'iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAYAAACp8Z5+AAAAHUlEQVR42mP8z8BQz0AEYBxVSF+F/xkYGP4TowsAeYUH/WrKZ1UAAAAASUVORK5CYII=';

/** Word's clipboard html for a picture copied on its own. */
const PICTURE_ONLY =
    `<html xmlns:v="urn:schemas-microsoft-com:vml"><body><span style='mso-ignore:vglayout'>` +
    `<!--[if gte vml 1]><v:shape id="Picture_x0020_1" o:spid="_x0000_s1026" type="#_x0000_t75">` +
    `<v:imagedata src="file:///C:/Temp/clip_image001.png"/></v:shape><![endif]-->` +
    `<img width=96 height=96 src="file:///C:/Temp/clip_image001.png" v:shapes="Picture_x0020_1">` +
    `</span></body></html>`;

/** …and for the report header: patient details in one cell, the QR in the next. */
const HEADER =
    `<html xmlns:v="urn:schemas-microsoft-com:vml"><body><table border=1 style='border-collapse:collapse'><tr>` +
    `<td style='padding:2px'><p class=MsoNormal>MRN: v2609256937</p></td>` +
    `<td style='padding:2px'><img width=96 src="file:///C:/Temp/clip_image001.png" v:shapes="Picture_x0020_1"></td>` +
    `</tr></table></body></html>`;

let fails = 0;
function check(label, ok, detail) {
    if (!ok) fails++;
    console.log(`${ok ? 'PASS' : 'FAIL'}  ${label}${detail ? ` — ${detail}` : ''}`);
}

const { server, port } = await serve(() => harnessHtml(''));
const browser = await chromium.launch({ channel: 'chrome' });
const page = await browser.newPage();
await page.goto(`http://127.0.0.1:${port}/`);
await page.evaluate(() => window.__ready);

/** Paste a clipboard built in the page, the way the browser really delivers one. */
async function paste(html, withFile) {
    await page.evaluate(() => window.RadiolensEditor.setHtml('#report_body_editor', '<p>start</p>'));
    await page.waitForTimeout(80);
    await page.click('[data-slate-editor] [data-slate-node="element"]');
    await page.keyboard.press('End');
    await page.waitForTimeout(60);

    return page.evaluate(
        async ({ html, withFile, PNG_B64 }) => {
            const editable = document.querySelector('[data-slate-editor]');
            const dt = new DataTransfer();
            if (html) dt.setData('text/html', html);
            if (withFile) {
                const bytes = Uint8Array.from(atob(PNG_B64), (c) => c.charCodeAt(0));
                dt.items.add(new File([bytes], 'image.png', { type: 'image/png' }));
            }
            editable.dispatchEvent(
                new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true })
            );
            // slate-react ignores `paste` where beforeinput is supported, which
            // is every browser this runs in; a real paste delivers both.
            editable.dispatchEvent(
                new InputEvent('beforeinput', {
                    inputType: 'insertFromPaste',
                    dataTransfer: dt,
                    bubbles: true,
                    cancelable: true,
                })
            );
            await new Promise((r) => setTimeout(r, 700));
            return window.RadiolensEditor.getHtml('#report_body_editor');
        },
        { html, withFile, PNG_B64 }
    );
}

const fileOnly = await paste('', true);
check('a screenshot pastes as a picture, stored as a data URL', fileOnly.includes('data:image/png;base64,'));

const wordPicture = await paste(PICTURE_ONLY, true);
check(
    'a Word picture whose bytes are on the clipboard is pasted from them',
    wordPicture.includes('data:image/png;base64,'),
    wordPicture.includes('data:image') ? 'recovered' : `got ${wordPicture.slice(0, 90)}`
);
check('and not as the file:// path the browser cannot read', !wordPicture.includes('file:///'));

const wordPictureNoBytes = await paste(PICTURE_ONLY, false);
check(
    'with no bytes anywhere, the reference is kept rather than the picture deleted',
    wordPictureNoBytes.includes('file:///'),
    'it will not render, but the document said a picture belongs there'
);

const header = await paste(HEADER, true);
check('a header table still pastes as the table', header.includes('<table') && header.includes('MRN: v2609256937'));
check(
    'and NOT as one flat picture of the whole selection',
    !header.includes('data:image/png;base64,'),
    'the clipboard image renders the entire selection, text and all'
);
check('and its QR keeps the reference Word gave it', header.includes('file:///'));

await browser.close();
server.close();

console.log(fails === 0 ? '\nAll picture-paste checks passed' : `\n${fails} check(s) failed`);
process.exit(fails === 0 ? 0 : 1);

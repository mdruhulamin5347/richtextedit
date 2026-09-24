/**
 * A table copied out of LibreOffice: is it spaced the way the document is?
 *
 * LibreOffice's HTML leaves the line spacing off every paragraph in a table
 * cell — a single-, 1.5- and double-spaced table copy as byte-identical HTML —
 * and the document's default `p { line-height: 115% }` was drawn there instead,
 * so every pasted table came out at the same gap. Only the RTF it copies
 * alongside states each one (`\intbl\sl480\slmult1`). See lib/word-line-gap.ts.
 *
 * Measured against LibreOffice's own layout of the .docx this was cut from
 * (`--convert-to pdf`, `pdftotext -bbox-layout`): 11pt Calibri at single / 1.5
 * / double is a 17.93 / 26.87 / 35.87px line — gap x Calibri's natural line x
 * size, which is what each cell paragraph's `line-height` is checked against.
 * The drawn line is also checked against a Word paste of the same text at the
 * same gap: the two must be spaced alike.
 *
 * The HTML is LibreOffice's own clipboard, trimmed — captured through its
 * transferable (UNO). The RTF is cut down to what is read from it. The async
 * clipboard API cannot carry RTF, so the paste is the `beforeinput` Chrome
 * fires for a real one.
 *
 * Run: node tests/browser/libreoffice-table-line-gap.mjs
 */
import { chromium } from 'playwright';
import { harnessHtml, serve } from './harness.mjs';

let fails = 0;
const check = (name, pass, detail = '') => {
    if (!pass) fails++;
    console.log(`${pass ? 'PASS' : 'FAIL'}  ${name}${detail ? ' — ' + detail : ''}`);
};

const FINDINGS =
    'Liver is normal in size and echotexture. No focal lesion seen in the parenchyma, and the ' +
    'hepatic veins are normal. Gall bladder is well distended with no calculus seen.';

const F = '<font face="Calibri, serif"><font size="2" style="font-size: 11pt">';
const LIBRE_OFFICE = `<!DOCTYPE html><html><head><meta http-equiv="content-type" content="text/html; charset=utf-8"/>
<meta name="generator" content="LibreOffice 24.2.7.2 (Linux)"/>
<style type="text/css">
td p { color: #000000; text-align: left; orphans: 0; widows: 0; margin-bottom: 0in; direction: ltr; background: transparent }
td p.western { font-family: "Liberation Serif", serif; font-size: 12pt; so-language: en-US }
p { color: #000000; line-height: 115%; text-align: left; orphans: 2; widows: 2; margin-bottom: 0.1in; direction: ltr; background: transparent }
p.western { font-family: "Liberation Serif", serif; font-size: 12pt; so-language: en-US }
</style></head>
<body lang="en-US" text="#000000" dir="ltr"><p class="western" align="left" style="line-height: 100%; margin-bottom: 0in">
Before table</p>
<table width="400" cellpadding="4" cellspacing="0">
\t<col width="128*"/>
\t<col width="128*"/>
\t<tr valign="top">
\t\t<td width="50%" style="border: 1px solid #000000; padding: 0.04in"><p class="western" align="left">
\t\t\t${F}Findings
\t\t\t:</font></font></p>
\t\t</td>
\t\t<td width="50%" style="border: 1px solid #000000; padding: 0.04in"><p class="western" align="left">
\t\t\t${F}${FINDINGS.replace(/ (?=hepatic|with)/g, '\n\t\t\t')}</font></font></p>
\t\t</td>
\t</tr>
</table>
<p class="western" align="left" style="line-height: 100%; margin-bottom: 0in">
After table</p>
</body></html>`;

/** LibreOffice's RTF for the same selection, each cell at `spacing`. */
const libreOfficeRtf = (spacing) =>
    String.raw`{\rtf1\ansi\deff3{\fonttbl{\f3\froman Liberation Serif;}{\f4\froman Calibri;}}` +
    String.raw`{\stylesheet{\s0 Normal;}{\s16\sbasedon0 Table Contents;}}` +
    String.raw`\pard\plain \s0\ql{\loch Before table}\par\trowd\cellx2000\cellx4000` +
    String.raw`\pard\plain \s16\intbl${spacing}{\f4\fs22 Findings :}\cell` +
    String.raw`\pard\plain \s16\intbl${spacing}{\f4\fs22 ${FINDINGS}}\cell\row` +
    String.raw`\pard\plain \s0\ql{\loch After table}\par}`;

/** The same table as Word writes it — its HTML states the gap itself. */
const word = (percent) =>
    `<html xmlns:w="urn:schemas-microsoft-com:office:word"><head><style><!--
p.MsoNormal {margin:0cm; font-size:11.0pt; font-family:"Calibri",sans-serif;}
--></style></head><body><table class=MsoTableGrid border=1 cellspacing=0 style='border-collapse:collapse;width:300pt'><tr>` +
    `<td style='padding:0cm 5.4pt'><p class=MsoNormal style='line-height:${percent}'><span style='font-size:11.0pt'>Findings :</span></p></td>` +
    `<td style='padding:0cm 5.4pt'><p class=MsoNormal style='line-height:${percent}'><span style='font-size:11.0pt'>${FINDINGS}</span></p></td>` +
    `</tr></table></body></html>`;

const { server, port } = await serve(() => harnessHtml(''));
const browser = await chromium.launch({ channel: 'chrome' });
const page = await (await browser.newContext({ viewport: { width: 1280, height: 900 } })).newPage();
await page.goto(`http://127.0.0.1:${port}/`);
await page.evaluate(() => window.__ready);

const natural = await page.evaluate(() => {
    const d = document.createElement('div');
    d.textContent = 'Hg';
    d.style.cssText = 'position:absolute;left:-9999px;font-family:Calibri;font-size:1000px;line-height:normal';
    document.body.append(d);
    const ratio = d.getBoundingClientRect().height / 1000;
    d.remove();
    return ratio;
});
check('the browser can measure Calibri\'s natural line', natural > 1 && natural < 2, `natural = ${natural}`);

async function paste(html, rtf) {
    await page.evaluate(() => window.RichTextEdit.setHtml('#report_body_editor', ''));
    await page.click('[data-slate-editor]');
    await page.evaluate(({ html, rtf }) => {
        const data = new DataTransfer();
        data.setData('text/html', html);
        data.setData('text/rtf', rtf);
        data.setData('text/plain', 'x');
        document.querySelector('[data-slate-editor]').dispatchEvent(
            new InputEvent('beforeinput', { inputType: 'insertFromPaste', dataTransfer: data, bubbles: true, cancelable: true })
        );
    }, { html, rtf });
    await page.waitForTimeout(700);
}

/** Each cell paragraph's line, the wrapped cell's line pitch, and the blocks outside the table. */
function measure() {
    return page.evaluate(() => {
        const root = document.querySelector('[data-slate-editor]');
        const cells = Array.from(root.querySelectorAll('td [data-slate-node="element"]')).map((p) => ({
            line: Number.parseFloat(getComputedStyle(p).lineHeight),
            size: Number.parseFloat(getComputedStyle(p).fontSize),
        }));
        let pitch = null;
        for (const s of root.querySelectorAll('[data-slate-string]')) {
            if (!s.textContent.startsWith('Liver')) continue;
            const range = document.createRange();
            range.selectNodeContents(s);
            const tops = [...new Set(Array.from(range.getClientRects()).map((r) => Math.round(r.top * 100) / 100))].sort((a, b) => a - b);
            if (tops.length > 1) pitch = tops[1] - tops[0];
        }
        const outside = Array.from(root.children)
            .filter((block) => !block.querySelector('table'))
            .map((block) => getComputedStyle(block.querySelector('[data-slate-node="element"]') ?? block).lineHeight);
        return { cells, pitch, outside };
    });
}

async function lineGapMenu() {
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

// The paste before this fix, for what must NOT move: an RTF that matches nothing.
await paste(LIBRE_OFFICE, String.raw`{\rtf1\ansi\pard Something else entirely\par}`);
const unmatched = await measure();

for (const [name, spacing, gap, percent, shown] of [
    ['single', String.raw`\sl240\slmult1`, 1, '100%', '1'],
    ['1.5 lines', String.raw`\sl360\slmult1`, 1.5, '150%', '1.5'],
    ['double', String.raw`\sl480\slmult1`, 2, '200%', '2'],
]) {
    await paste(LIBRE_OFFICE, libreOfficeRtf(spacing));
    const pasted = await measure();
    const expected = pasted.cells.map((cell) => gap * natural * cell.size);
    check(
        `LibreOffice ${name}: every cell paragraph is spaced at ${gap} lines`,
        pasted.cells.length === 2 && pasted.cells.every((cell, i) => Math.abs(cell.line - expected[i]) <= 0.1),
        `LibreOffice ${expected.map((px) => px.toFixed(2)).join('/')}px | editor ${pasted.cells.map((c) => c.line.toFixed(2)).join('/')}px`
    );
    check(
        `LibreOffice ${name}: the paragraphs outside the table are untouched`,
        JSON.stringify(pasted.outside) === JSON.stringify(unmatched.outside),
        `before ${unmatched.outside.join(', ')} | now ${pasted.outside.join(', ')}`
    );

    const ticked = await lineGapMenu();
    check(`LibreOffice ${name}: the control names it ${shown}`, ticked.length === 1 && ticked[0] === shown, `menu shows ${JSON.stringify(ticked)}`);

    const saved = await page.evaluate(() => window.RichTextEdit.getHtml('#report_body_editor'));
    await page.evaluate((html) => window.RichTextEdit.setHtml('#report_body_editor', html), saved);
    await page.waitForTimeout(300);
    const reopened = await measure();
    check(
        `LibreOffice ${name}: a save and reopen keeps it`,
        reopened.cells.every((cell, i) => Math.abs(cell.line - pasted.cells[i].line) <= 0.01),
        `pasted ${pasted.cells.map((c) => c.line.toFixed(2)).join('/')} | reopened ${reopened.cells.map((c) => c.line.toFixed(2)).join('/')}`
    );

    await paste(word(percent), '{\\rtf1\\ansi x}');
    const fromWord = await measure();
    check(
        `LibreOffice ${name}: drawn with the same line pitch as the same table pasted from Word`,
        pasted.pitch !== null && fromWord.pitch !== null && Math.abs(pasted.pitch - fromWord.pitch) <= 0.1,
        `Word ${fromWord.pitch?.toFixed(2)}px | LibreOffice ${pasted.pitch?.toFixed(2)}px`
    );
}

await browser.close();
server.close();
console.log(fails ? `\n${fails} FAILED` : '\nall passed');
process.exit(fails ? 1 : 0);

/**
 * Toolbar behaviour on a Bootstrap page.
 *
 * Two failures this guards, both invisible to jsdom because they need layout:
 *
 *  - Text colour. The editor sets none of its own, so every character inherited
 *    Bootstrap's `body { color: #212529 }` — a grey that reads as washed out
 *    where a report should be near-black.
 *
 *  - Dropdowns. Every trigger passes `pressed`, and that branch of ToolbarButton
 *    used to render a Radix ToggleGroup, which does not forward a ref. Radix's
 *    popper ANCHOR ref was swallowed, Floating UI never computed a position, and
 *    the menu mounted at its unpositioned `translate(0, -200%)` far off-screen —
 *    indistinguishable from a dead button.
 */
import { chromium } from 'playwright';
import { harnessHtml, serve } from './harness.mjs';

const results = [];
const check = (name, pass, detail = '') => {
    results.push(pass);
    console.log(`${pass ? 'PASS' : 'FAIL'}  ${name}${detail ? ' — ' + detail : ''}`);
};

const { server, port } = await serve(() => harnessHtml('<p>Normal report text</p>'));
const browser = await chromium.launch({ channel: 'chrome' });
const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
const errors = [];
page.on('pageerror', (e) => errors.push(String(e)));
page.on('console', (m) => m.type() === 'error' && errors.push(m.text()));

await page.goto(`http://127.0.0.1:${port}/`);
await page.evaluate(() => window.__ready);
await page.waitForTimeout(200);

const text = await page.evaluate(() => {
    const el = document.querySelector('[data-slate-editor]');
    const scope = document.querySelector('.rl-editor-scope');
    return {
        color: getComputedStyle(el).color,
        token: getComputedStyle(scope).getPropertyValue('--foreground').trim(),
    };
});
check('editor text uses the editor foreground, not Bootstrap grey',
    text.color === 'rgb(10, 10, 10)', `${text.color} (token ${text.token})`);

const enabled = await page.evaluate(() =>
    Array.from(document.querySelectorAll('button'))
        .filter((b) => !b.disabled)
        .map((b) => getComputedStyle(b).opacity)
);
check('enabled toolbar buttons are not dimmed',
    enabled.length > 0 && enabled.every((o) => Number(o) === 1),
    `${enabled.length} enabled buttons`);

const trigger = page.locator('button:has-text("Text")').first();
check('a dropdown trigger exists and is enabled', (await trigger.count()) > 0 && !(await trigger.isDisabled()));

await trigger.click();
await page.waitForTimeout(400);

const menu = await page.evaluate(() => {
    const w = document.querySelector('[data-radix-popper-content-wrapper]');
    if (!w) return null;
    const r = w.getBoundingClientRect();
    const style = w.getAttribute('style') ?? '';
    return {
        anchored: style.includes('--radix-popper-anchor-width'),
        unpositioned: style.includes('-200%'),
        x: Math.round(r.x), y: Math.round(r.y),
        w: Math.round(r.width), h: Math.round(r.height),
    };
});

check('the menu renders', !!menu);
if (menu) {
    check('Radix resolved the popper anchor', menu.anchored);
    check('the menu is positioned, not parked off-screen', !menu.unpositioned);
    check('the menu is inside the viewport', menu.y >= 0 && menu.x >= 0 && menu.y < 900,
        `at ${menu.x},${menu.y} (${menu.w}x${menu.h})`);
}

check('no console or page errors', errors.length === 0, errors.slice(0, 2).join(' | '));

await browser.close();
server.close();

const failed = results.filter((r) => !r).length;
console.log(`\n${results.length - failed}/${results.length} checks passed`);
process.exit(failed ? 1 : 0);

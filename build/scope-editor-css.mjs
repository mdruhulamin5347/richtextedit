/**
 * Re-roots the editor's stylesheet under `.rte-scope`.
 *
 * The editor imports Tailwind WHOLE, preflight included, because Plate's node
 * components are written against that baseline. The admin panel it is mounted
 * into is a Bootstrap 5 page with ~200 screens, so an unscoped preflight would
 * flatten every heading, list, table and form control on all of them. These
 * transforms run on the BUILT css (see the `scope-editor-css` plugin in
 * vite.config.js), after Tailwind has generated its utilities, so there is
 * nothing left to generate that could escape afterwards.
 *
 * Three things happen here, and the third is the unobvious one:
 *
 *   1. `scopeEditorCss` prefixes every selector with `.rte-scope`.
 *   2. It marks a rule's declarations `!important` when the rule's own class
 *      name also names a Bootstrap utility that Bootstrap declares
 *      `!important` — `.p-0`, `.mb-1`, `.py-1` and ~54 others. Specificity
 *      cannot beat an `!important` declaration, so matching it is the only way
 *      the editor's version of a colliding name wins. (An important LAYERED
 *      declaration does beat an important unlayered one: layer order is
 *      reversed for important declarations, and Tailwind emits inside
 *      `@layer`.)
 *   3. `bootstrapNeutralizerCss` handles the other half of a name collision:
 *      the properties Bootstrap states for that class and the editor does NOT.
 *      Those are unlayered NORMAL declarations, which beat a layered one at any
 *      specificity, so they reach inside the editor untouched. `.table` is the
 *      case that forced this — Tailwind's `.table` is only `display: table`, so
 *      there was no editor declaration for `width` to override Bootstrap's
 *      `width: 100%` with, and a table could not be resized below full width.
 *
 * Element defaults are deliberately NOT neutralised. For every property
 * Bootstrap's reboot states for an element, Bootstrap is what the editor shows
 * — a heading is reboot's 500-weight one, a list is indented reboot's 2rem.
 * That is the contract the web app's report-body print styles partial
 * mirrors for print, and
 * the web app's `tests/browser/print-parity.mjs` measures both sides against it.
 *
 * Guarded by tests/js/editor-css.test.ts, which reads the built css back.
 */
import fs from 'node:fs';
import postcss from 'postcss';

const SCOPE = '.rte-scope';

/**
 * Selectors naming the document root. Inside this bundle they mean the editor
 * container, so they are REPLACED by the scope rather than prefixed with it —
 * this is what puts preflight's `html` rule and editor.css's `:root` tokens
 * onto `.rte-scope`, where portalled menus pick them up (they are
 * rendered into a host that carries the class, see lib/portal-container.ts).
 *
 * `body` is not in this list on purpose. Preflight's only body rule is
 * `line-height: inherit`, and mapping it here would land it on the same element
 * as the `html` rule and, coming later, replace the 1.5 the editor's text is
 * measured at with whatever the surrounding admin page happens to inherit.
 */
const ROOT_SELECTOR = /^(?::root|:host|html)(?![\w(-])/;

/**
 * Class tokens in a selector, escapes included, so Tailwind's own escaping does
 * not read as a shorter name: `.p-0\.5` is the single class `p-0.5`, not `p-0`,
 * and `.md\:p-0` is `md:p-0`. Getting this wrong marks whole families of
 * utilities `!important` off one Bootstrap collision.
 */
const CLASS_TOKEN = /\.((?:[\w-]|\\.)+)/g;

/**
 * @param {string} selector
 * @return {string[]} unescaped class names used in the selector
 */
function selectorClasses(selector) {
    const names = [];
    for (const match of selector.matchAll(CLASS_TOKEN)) {
        names.push(match[1].replace(/\\(.)/g, '$1'));
    }

    return names;
}

/**
 * Keyframe steps (`from`, `to`, `50%`) parse as rules but are not selectors,
 * and prefixing one silently kills the animation.
 *
 * @param {import('postcss').Rule} rule
 */
function insideKeyframes(rule) {
    for (let node = rule.parent; node; node = node.parent) {
        if (node.type === 'atrule' && /keyframes$/.test(node.name)) {
            return true;
        }
    }

    return false;
}

/**
 * @param {string} selector
 * @return {string} the selector confined to the editor
 */
function scopeSelector(selector) {
    const trimmed = selector.trim();

    if (trimmed === SCOPE || trimmed.startsWith(`${SCOPE} `) || trimmed.startsWith(`${SCOPE}.`)
        || trimmed.startsWith(`${SCOPE}:`) || trimmed.startsWith(`${SCOPE}[`)) {
        return trimmed;
    }

    const root = trimmed.match(ROOT_SELECTOR);
    if (root) {
        return SCOPE + trimmed.slice(root[0].length);
    }

    // The portal host carries `rte-scope` alongside its own class, so it
    // is the scope element rather than something inside it — compounded, not
    // descended, or the rule would never match the host itself.
    if (trimmed.startsWith('.rte-')) {
        return SCOPE + trimmed;
    }

    return `${SCOPE} ${trimmed}`;
}

/**
 * Class names Bootstrap declares `!important` on.
 *
 * @param {string} path path to the Bootstrap stylesheet
 * @return {Set<string>}
 */
export function bootstrapImportantClasses(path) {
    const classes = new Set();

    postcss.parse(fs.readFileSync(path, 'utf8')).walkRules((rule) => {
        if (!rule.nodes.some((node) => node.type === 'decl' && node.important)) {
            return;
        }

        for (const selector of rule.selectors) {
            for (const name of selectorClasses(selector)) {
                classes.add(name);
            }
        }
    });

    return classes;
}

/**
 * Confine a stylesheet to `.rte-scope`.
 *
 * @param {string} css the built editor stylesheet
 * @param {Set<string>} important class names Bootstrap declares `!important` on
 * @return {string}
 */
export function scopeEditorCss(css, important = new Set()) {
    const root = postcss.parse(css);

    root.walkRules((rule) => {
        if (insideKeyframes(rule)) {
            return;
        }

        const collides = rule.selectors.some(
            (selector) => selectorClasses(selector).some((name) => important.has(name))
        );

        rule.selectors = [...new Set(rule.selectors.map(scopeSelector))];

        if (collides) {
            rule.walkDecls((decl) => {
                decl.important = true;
            });
        }
    });

    return root.toString();
}

/**
 * Rules that take back the properties Bootstrap states for a class name the
 * editor also uses, but the editor itself states nothing for.
 *
 * Emitted UNLAYERED and last, and at one class more specific than Bootstrap's
 * own rule, so it wins the tie it is meant to win; `unset` hands the property
 * back to the cascade rather than inventing a value, which is why the editor's
 * scoped preflight is what answers for it afterwards. Properties the editor
 * DOES state are left out — they are layered, and an unlayered `unset` here
 * would beat them.
 *
 * Only Bootstrap's unconditional, single-class, non-important rules are read.
 * A media-query rule would be flattened into an unconditional reset, a compound
 * or descendant selector reaches elements whose styling is the document's (a
 * table cell's padding is written inline by the serializer), and an important
 * declaration is already answered by the `!important` marking in
 * `scopeEditorCss`.
 *
 * @param {string} bootstrapCss
 * @param {string} scopedCss output of {@link scopeEditorCss}
 * @return {string}
 */
export function bootstrapNeutralizerCss(bootstrapCss, scopedCss) {
    /** @type {Map<string, Set<string>>} class name to the properties the editor states for it */
    const stated = new Map();

    postcss.parse(scopedCss).walkRules((rule) => {
        if (insideKeyframes(rule)) {
            return;
        }

        const props = rule.nodes
            .filter((node) => node.type === 'decl')
            .map((node) => node.prop.toLowerCase());

        if (props.length === 0) {
            return;
        }

        for (const selector of rule.selectors) {
            for (const name of selectorClasses(selector)) {
                if (name === 'rte-scope' || name === 'rte-portal') {
                    continue;
                }

                let declared = stated.get(name);
                if (!declared) {
                    stated.set(name, (declared = new Set()));
                }

                for (const prop of props) {
                    declared.add(prop);
                }
            }
        }
    });

    /** @type {Map<string, Set<string>>} class name to the properties to hand back */
    const resets = new Map();

    postcss.parse(bootstrapCss).walkRules((rule) => {
        if (rule.parent?.type !== 'root') {
            return;
        }

        for (const selector of rule.selectors) {
            const name = selector.trim().match(/^\.((?:[\w-]|\\.)+)$/)?.[1]?.replace(/\\(.)/g, '$1');
            const declared = name && stated.get(name);
            if (!declared) {
                continue;
            }

            for (const node of rule.nodes) {
                if (node.type !== 'decl' || node.important || node.prop.startsWith('--')) {
                    continue;
                }

                if (declared.has(node.prop.toLowerCase())) {
                    continue;
                }

                let props = resets.get(name);
                if (!props) {
                    resets.set(name, (props = new Set()));
                }

                props.add(node.prop.toLowerCase());
            }
        }
    });

    return [...resets]
        .map(([name, props]) => `${SCOPE} .${name}{${[...props].map((p) => `${p}:unset`).join(';')}}`)
        .join('\n');
}

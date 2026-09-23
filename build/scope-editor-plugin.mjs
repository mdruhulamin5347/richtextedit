/**
 * Vite plugin: confine the editor's stylesheet to `.rl-editor-scope` and let it
 * win against Bootstrap inside it. Shared by both builds.
 *
 * See scope-editor-css.mjs for why this is needed — short version: 57 Tailwind
 * utility names collide with Bootstrap's, and Bootstrap declares nearly all of
 * its own `!important`.
 *
 * Runs on the built CSS, after Tailwind has generated its utilities.
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import {
    bootstrapImportantClasses,
    bootstrapNeutralizerCss,
    scopeEditorCss,
} from './scope-editor-css.mjs';

/** The Bootstrap build the admin panel ships, vendored so the output matches. */
export const BOOTSTRAP_CSS = path.resolve(
    path.dirname(fileURLToPath(import.meta.url)),
    '../vendor/bootstrap.min.css'
);

export function scopeEditorStyles() {
    return {
        name: 'scope-editor-css',
        enforce: 'post',
        generateBundle(_options, bundle) {
            const important = bootstrapImportantClasses(BOOTSTRAP_CSS);
            for (const file of Object.values(bundle)) {
                if (file.type === 'asset' && file.fileName.endsWith('.css')) {
                    const scoped = scopeEditorCss(String(file.source), important);
                    const bootstrap = fs.readFileSync(BOOTSTRAP_CSS, 'utf8');
                    // Appended last so it wins the ties it is meant to win.
                    file.source = scoped + '\n' + bootstrapNeutralizerCss(bootstrap, scoped);
                }
            }
        },
    };
}

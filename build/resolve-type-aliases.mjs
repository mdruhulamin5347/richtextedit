/**
 * Rewrites `@/…` imports in the emitted .d.ts files to relative paths.
 *
 * tsc leaves path aliases untouched in declarations, and a consumer's compiler
 * has no idea what `@/` means in this package — every type would resolve to
 * `any`. The bundlers resolve the alias for the JS; this does it for the types.
 */
import fs from 'node:fs';
import path from 'node:path';

const TYPES = path.resolve('dist/types');

function walk(dir) {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
        const full = path.join(dir, entry.name);
        if (entry.isDirectory()) {
            walk(full);
            continue;
        }
        if (!entry.name.endsWith('.d.ts')) continue;

        const source = fs.readFileSync(full, 'utf8');
        const rewritten = source
            // A stylesheet import is a bundler concern; in a .d.ts it only
            // points at a file the types folder does not have.
            .replace(/^import\s+["'][^"']+\.css["'];\n/gm, '')
            .replace(/(from\s+|import\()(["'])@\/([^"']+)\2/g, (_m, lead, quote, target) => {
            let relative = path.relative(path.dirname(full), path.join(TYPES, target)).split(path.sep).join('/');
            if (!relative.startsWith('.')) relative = './' + relative;
            return `${lead}${quote}${relative}${quote}`;
        });
        if (rewritten !== source) fs.writeFileSync(full, rewritten);
    }
}

walk(TYPES);

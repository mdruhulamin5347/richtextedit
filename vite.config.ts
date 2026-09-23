import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import tailwindcss from '@tailwindcss/vite';
import path from 'node:path';
import pkg from './package.json' with { type: 'json' };
import { scopeEditorStyles } from './build/scope-editor-plugin.mjs';

/**
 * Every dependency and peer stays external, so a host app's bundler dedupes
 * React, Plate and Radix against its own copies instead of shipping two.
 */
const externals = [...Object.keys(pkg.dependencies), ...Object.keys(pkg.peerDependencies)];
const isExternal = (id: string) => externals.some((dep) => id === dep || id.startsWith(dep + '/'));

/** ESM build for React apps: `import { RichTextEdit } from 'rich-text-edit'`. */
export default defineConfig({
    plugins: [react(), tailwindcss(), scopeEditorStyles()],
    resolve: {
        alias: { '@': path.resolve(__dirname, 'src') },
    },
    build: {
        outDir: 'dist',
        emptyOutDir: false,
        sourcemap: true,
        lib: {
            entry: path.resolve(__dirname, 'src/index.ts'),
            formats: ['es'],
            fileName: () => 'index.js',
            cssFileName: 'style',
        },
        rollupOptions: { external: isExternal },
    },
});

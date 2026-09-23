import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import tailwindcss from '@tailwindcss/vite';
import path from 'node:path';
import { scopeEditorStyles } from './build/scope-editor-plugin.mjs';

/**
 * Self-contained build for pages without a bundler (Blade, jQuery, plain HTML):
 * React and every dependency are inlined, and loading the script sets
 * `window.RadiolensEditor` — the same API the web app's Blade pages use.
 */
export default defineConfig({
    plugins: [react(), tailwindcss(), scopeEditorStyles()],
    resolve: {
        alias: { '@': path.resolve(__dirname, 'src') },
    },
    define: {
        'process.env.NODE_ENV': JSON.stringify('production'),
    },
    build: {
        outDir: 'dist/standalone',
        emptyOutDir: true,
        lib: {
            entry: path.resolve(__dirname, 'src/entry.tsx'),
            formats: ['iife'],
            // Not `RadiolensEditor`: the entry assigns that global itself, and an
            // IIFE name would overwrite it with the module namespace afterwards.
            name: 'RadiolensEditorBundle',
            fileName: () => 'radiolens-editor.iife.js',
            cssFileName: 'radiolens-editor',
        },
    },
});

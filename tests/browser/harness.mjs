/**
 * Serves a minimal page that mounts the built editor bundle, so the editor can
 * be driven in a real browser.
 *
 * jsdom does no layout: getBoundingClientRect is all zeros there, so a
 * simulated drag proves nothing about resizing. This harness loads the actual
 * built assets over http, which is the only way to verify that dragging a
 * column border moves the column.
 *
 * The page deliberately loads Bootstrap the way the admin panel does, so
 * conflicts between it and the editor show up here rather than in production.
 */
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';

const ROOT = process.cwd();

const MIME = {
    '.js': 'text/javascript',
    '.css': 'text/css',
    '.html': 'text/html',
    '.json': 'application/json',
    '.woff2': 'font/woff2',
    '.ttf': 'font/ttf',
};

/** The standalone build — run `pnpm run build:standalone` first. */
function assets() {
    return {
        js: '/standalone/radiolens-editor.iife.js',
        css: ['/standalone/radiolens-editor.css'],
    };
}

/** URL path -> file, for the few files that do not live under dist/. */
const ALIASES = {
    '/our_assets/bootstrap/css/bootstrap.min.css': path.join(ROOT, 'vendor/bootstrap.min.css'),
};

export function harnessHtml(initialHtml = '') {
    const { js, css } = assets();
    return `<!doctype html>
<html><head><meta charset="utf-8">
<meta name="csrf-token" content="test-token">
<link rel="stylesheet" href="/our_assets/bootstrap/css/bootstrap.min.css">
${css.map((c) => `<link rel="stylesheet" href="${c}">`).join('\n')}
</head>
<body>
<div class="container">
  <form id="f" action="/store" method="POST">
    <div class="form-group">
      <label for="report_body">Report Body</label>
      <textarea name="report_body" id="report_body" hidden>${initialHtml.replace(/</g, '&lt;')}</textarea>
      <div class="rl-editor-scope" id="report_body_editor"></div>
    </div>
    <button type="submit" class="btn btn-success btn-sm">Create</button>
  </form>
</div>
<script type="module" src="${js}"></script>
<script>
  window.__ready = new Promise((resolve) => {
    (function wait() {
      if (!window.RadiolensEditor) return void setTimeout(wait, 20);
      window.RadiolensEditor.mount('#report_body_editor', {
        textarea: '#report_body', minHeight: 400,
      });
      requestAnimationFrame(() => requestAnimationFrame(resolve));
    })();
  });
</script>
</body></html>`;
}

export function serve(getHtml) {
    const server = http.createServer((req, res) => {
        const url = decodeURIComponent(req.url.split('?')[0]);
        if (url === '/' || url === '/index.html') {
            res.writeHead(200, { 'Content-Type': 'text/html' });
            return res.end(getHtml());
        }
        // The browser asks for one unprompted; the web app's public/ had it.
        if (url === '/favicon.ico') {
            res.writeHead(204);
            return res.end();
        }
        const file = ALIASES[url] ?? path.join(ROOT, 'dist', url);
        if (!(url in ALIASES) && !file.startsWith(path.join(ROOT, 'dist'))) {
            res.writeHead(404);
            return res.end('not found');
        }
        if (!fs.existsSync(file)) {
            res.writeHead(404);
            return res.end('not found');
        }
        res.writeHead(200, { 'Content-Type': MIME[path.extname(file)] ?? 'application/octet-stream' });
        fs.createReadStream(file).pipe(res);
    });
    return new Promise((resolve) => {
        server.listen(0, '127.0.0.1', () => resolve({ server, port: server.address().port }));
    });
}

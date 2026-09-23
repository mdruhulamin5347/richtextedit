# @radiolens/editor

The RadioLens rich text editor as a standalone package. It is built on
[Plate](https://platejs.org) and React, and made for report writing: Word and
Excel paste that keeps its formatting, resizable tables, point-based font sizes
and Word document import.

Each document can be stored as **HTML** or **JSON**, and you can read or write
either one at any time.

It works in two kinds of project:

- **Any page, no bundler needed**: Blade, jQuery, plain HTML. Load one script
  and one stylesheet, then call `RadiolensEditor.mount(...)`.
- **React apps**: import the `<RadiolensEditor>` component.

---

## Contents

- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Quick start: plain HTML / Blade](#quick-start-plain-html--blade)
- [Quick start: React](#quick-start-react)
- [HTML or JSON storage](#html-or-json-storage)
- [API reference](#api-reference)
- [Word import endpoint](#word-import-endpoint)
- [Laravel example](#laravel-example)
- [Styling and Bootstrap](#styling-and-bootstrap)
- [How saving works](#how-saving-works)
- [Development](#development)
- [Project structure](#project-structure)
- [Troubleshooting](#troubleshooting)
- [Known issues](#known-issues)

---

## Features

- **Opens legacy HTML correctly.** Old Summernote report bodies keep their
  alignment, font sizes, colours and whitespace, including legacy markup like
  `<p align=center>`, `<center>` and `<font color>`.
- **Word and Excel paste.** Tables, borders, column widths, line spacing, text
  boxes and pictures (including RTF-embedded ones) survive the paste.
- **Two paste modes.** `clean` (the default) converts Word's sizes and spacing
  onto the editor's own scales. `faithful` keeps every value exactly as pasted.
- **Tables.** Drag to resize columns (works on ragged tables too), set borders
  and cell padding, and select cells.
- **Typography.** Font family, font size in **points**, text and background
  colour, line height, alignment, indentation and lists.
- **Word import.** `.docx` converts in the browser. `.doc` (binary) converts
  through an optional server endpoint.
- **Images.** Paste or upload; images are stored inline as base64.
- **HTML and JSON.** `getHtml / setHtml` and `getValue / setValue`.
- **Safe on Bootstrap pages.** All styles are scoped under `.rl-editor-scope`,
  so the editor and Bootstrap don't restyle each other.

---

## Requirements

| | Version |
|---|---|
| Node.js (for building) | 20 or newer (tested on 24) |
| pnpm | 10.17.0 |
| React (React apps only) | 18.3 |

The standalone build bundles its own React, so plain HTML pages need nothing
installed.

---

## Installation

The package is marked `"private": true` and is not on npm yet. Pick one of these.

### Option 1: copy the built files (plain HTML / Blade)

```bash
cd RichText
pnpm install
pnpm run build
```

Then copy these two files into the other project's public folder:

```
dist/standalone/radiolens-editor.iife.js
dist/standalone/radiolens-editor.css
```

### Option 2: local path dependency (React or bundled apps)

In the other project's `package.json`:

```json
{
  "dependencies": {
    "@radiolens/editor": "file:../RichText"
  }
}
```

Run `pnpm run build` in `RichText` first: only `dist/` is included in the package.

### Option 3: git dependency

Push `RichText` to its own repository, tag a release, then:

```json
{
  "dependencies": {
    "@radiolens/editor": "github:your-org/radiolens-editor#v0.1.0"
  }
}
```

### Option 4: private registry (GitHub Packages / private npm)

1. Remove `"private": true` from `package.json`.
2. Add a `publishConfig` pointing at your registry.
3. Run `pnpm run build`, then `pnpm publish`.

After that, projects install it with `pnpm add @radiolens/editor`.

---

## Quick start: plain HTML / Blade

```html
<link rel="stylesheet" href="/vendor/radiolens-editor.css">

<form method="POST" action="/reports">
  <!-- Carries the value to the server. Hidden; the editor fills it. -->
  <textarea name="report_body" id="report_body" hidden><p>Existing report</p></textarea>

  <!-- The editor mounts here. The rl-editor-scope class is REQUIRED. -->
  <div class="rl-editor-scope" id="report_body_editor"></div>

  <button type="submit">Save</button>
</form>

<script src="/vendor/radiolens-editor.iife.js"></script>
<script>
  RadiolensEditor.mount('#report_body_editor', {
    textarea: '#report_body',
  });
</script>
```

That's all you need. The editor opens whatever is in the textarea and keeps it
up to date as you type. On submit it writes the latest content first, so the
form never posts stale text.

In a Blade view, escape the stored body into the textarea as usual:

```blade
<textarea name="report_body" id="report_body" hidden>{{ old('report_body', $report->report_body) }}</textarea>
```

---

## Quick start: React

```tsx
import { useRef } from "react";
import {
  RadiolensEditor,
  type RadiolensEditorHandle,
  type Value,
} from "@radiolens/editor";
import "@radiolens/editor/style.css";

export function ReportForm({ report }) {
  const editorRef = useRef<RadiolensEditorHandle>(null);

  function save() {
    const html = editorRef.current!.getHtml();
    const json = editorRef.current!.getValue();
    // send html and/or json to the server
  }

  return (
    // The rl-editor-scope class is REQUIRED on a wrapper.
    <div className="rl-editor-scope">
      <RadiolensEditor
        ref={editorRef}
        initialHtml={report.body_html}
        // or: initialValue={report.body_json}
        onChangeHtml={(html) => console.log("html", html)}
        onChangeValue={(value: Value) => console.log("json", value)}
        minHeight={400}
      />
      <button onClick={save}>Save</button>
    </div>
  );
}
```

`initialHtml` and `initialValue` only set the starting content, like
`defaultValue` on an input. To replace the content later, call
`editorRef.current.setHtml(...)` or `setValue(...)`.

---

## HTML or JSON storage

The editor's real document is a Plate **value**: a JSON array of nodes.

```json
[
  { "type": "p", "children": [{ "text": "Impression: ", "bold": true }, { "text": "Normal study." }] }
]
```

HTML is produced from that value on demand. Both forms describe the same
document and you can switch between them at any time:

```js
const value = RadiolensEditor.getValue('#box');   // JSON
RadiolensEditor.setValue('#box', value);          // same document back
RadiolensEditor.getHtml('#box');                  // same document, as HTML
```

### Which one to store

| Store | Use when |
|---|---|
| **HTML** (default) | You print or display the report as HTML, you have existing HTML data (Summernote), or other systems read the column. |
| **JSON** | You want a lossless copy of the editor's own document, to process it in code (search nodes, extract tables), or to reopen it without re-parsing HTML. |
| **Both** | Store JSON to edit and HTML to display. Use `onChange`, which gets both. |

### Choosing what the textarea holds

```js
// Default: the textarea holds HTML. Existing pages keep working unchanged.
RadiolensEditor.mount('#box', { textarea: '#body' });

// The textarea holds the JSON value as a string.
RadiolensEditor.mount('#box', { textarea: '#body', format: 'json' });
```

With `format: 'json'`:

- The textarea's starting text is parsed as JSON.
- An empty textarea opens an empty document.
- Invalid JSON opens an empty document and logs a console error. Valid JSON
  that isn't an array opens an empty document too. The editor never crashes
  on bad data.
- Bare text nodes at the root are wrapped in a paragraph, as the HTML path does.

### Saving both

```html
<textarea name="body_html" id="body_html" hidden></textarea>
<textarea name="body_json" id="body_json" hidden>{{ $report->body_json }}</textarea>
<div class="rl-editor-scope" id="editor"></div>

<script>
  RadiolensEditor.mount('#editor', {
    textarea: '#body_json',
    format: 'json',
    onChange: ({ html }) => { document.getElementById('body_html').value = html; },
  });
</script>
```

`onChange` fires on the debounced publish (see [How saving works](#how-saving-works)),
including the forced one on submit, so `body_html` is current when the form posts.

### Converting existing HTML data to JSON

Open the HTML, then read the JSON back out:

```js
RadiolensEditor.setHtml('#editor', storedHtml);
const json = JSON.stringify(RadiolensEditor.getValue('#editor'));
```

Round-tripping HTML → JSON → HTML gives back identical HTML; there's a test for
it in `tests/js/json-value.test.tsx`.

---

## API reference

### Global / mount API

Available as `window.RadiolensEditor` (standalone build) or as
`RadiolensEditorApi` (ESM build). Every function takes the mount container,
either as a CSS selector or as an `Element`.

| Function | Returns | Description |
|---|---|---|
| `mount(target, options)` | handle ref, or `null` | Mounts an editor in `target`. Mounting the same target twice returns the existing editor. Returns `null`, with a console error, if the target or textarea is not found. |
| `destroy(target)` | `void` | Unmounts the editor and removes its submit listener. The textarea keeps its last value. |
| `getHtml(target)` | `string` | The document as HTML, up to date right now. `""` if nothing is mounted. |
| `setHtml(target, html)` | `void` | Replaces the document from HTML (the same as `$(el).summernote('code', html)`). |
| `getValue(target)` | `Value` | The document as JSON. Returns a **copy**, so changing it does not change the editor. `[]` if nothing is mounted. |
| `setValue(target, value)` | `void` | Replaces the document from JSON: a `Value` array or a JSON string. The editor stores its own copy. |

`setHtml` and `setValue` update the textarea straight away.

### Mount options

| Option | Type | Default | Description |
|---|---|---|---|
| `textarea` | `string \| HTMLTextAreaElement` | **required** | The hidden textarea that carries the value to the server. Its current text is the starting document. |
| `format` | `"html" \| "json"` | `"html"` | What the textarea is read as and written back as. |
| `onChange` | `({ html, value }) => void` | — | Called on every debounced change with both forms. |
| `docConvertUrl` | `string` | — | Server endpoint for converting `.doc` files. See [Word import endpoint](#word-import-endpoint). |
| `csrfToken` | `string` | from `<meta name="csrf-token">` | Sent as `X-CSRF-TOKEN` to `docConvertUrl`. |
| `placeholder` | `string` | `"Type the report here…"` | Shown while the document is empty. |
| `minHeight` | `number` (px) | `600` | Minimum height of the writing area. A click anywhere in it places the caret. |
| `pasteMode` | `"clean" \| "faithful"` | `"clean"` | `clean` drops Word's margins and letter/word spacing and converts its sizes and line heights onto the editor's scales. `faithful` keeps every value exactly as pasted. |

### React component: `<RadiolensEditor>`

| Prop | Type | Default | Description |
|---|---|---|---|
| `initialHtml` | `string` | `""` | Starting document as HTML. |
| `initialValue` | `Value \| string \| null` | — | Starting document as JSON. If both are given, this one wins. |
| `onChangeHtml` | `(html: string) => void` | — | Called with HTML on every debounced change. |
| `onChangeValue` | `(value: Value) => void` | — | Called with JSON on every debounced change. Only pass it if you need it: each call deep-copies the document. |
| `docConvertUrl` | `string` | — | See [Word import endpoint](#word-import-endpoint). |
| `csrfToken` | `string` | — | Sent as `X-CSRF-TOKEN` to `docConvertUrl`. |
| `placeholder` | `string` | `"Type the report here…"` | |
| `minHeight` | `number` | `600` | |
| `pasteMode` | `"clean" \| "faithful"` | `"clean"` | |

### Ref handle: `RadiolensEditorHandle`

| Method | Description |
|---|---|
| `getHtml()` | The document as HTML. |
| `setHtml(html)` | Replace the document from HTML. |
| `getValue()` | The document as JSON (a copy). |
| `setValue(value)` | Replace the document from JSON (an array or a JSON string). |
| `flush()` | Skip the debounce: publish now to `onChangeHtml` / `onChangeValue`, and return the HTML. |
| `focus()` | Focus the editor. |
| `getEditor()` | The underlying Plate editor, for advanced use. |

### Other exports (ESM build)

```ts
import {
  RadiolensEditor,        // React component
  RadiolensEditorApi,     // { mount, destroy, getHtml, setHtml, getValue, setValue }
  plateValueToHtml,       // (value: Value) => string
  plateValueToJson,       // (value: Value) => string
  isPlateJson,            // (text: string) => boolean
  isPlateValueEmpty,      // (value) => boolean
  EMPTY_VALUE,            // an empty document
  type Value,
  type PasteMode,
  type MountOptions,
  type StorageFormat,     // "html" | "json"
  type RadiolensEditorProps,
  type RadiolensEditorHandle,
} from "@radiolens/editor";
```

`plateValueToHtml` turns a stored JSON value into HTML **without mounting an
editor**. That's handy for print pages or emails:

```ts
const html = plateValueToHtml(JSON.parse(report.body_json));
```

---

## Word import endpoint

The toolbar's **Word import** button accepts `.docx` and `.doc`.

| File | Without `docConvertUrl` | With `docConvertUrl` |
|---|---|---|
| `.docx` | Converted in the browser with mammoth. Keeps structure only: colours, sizes and alignment are lost. | Sent to the server first (keeps formatting). If the server fails, it falls back to the browser and shows "formatting reduced". |
| `.doc` | Not supported. The browser can't read binary `.doc` files. | Sent to the server. |

The server endpoint must:

- **Accept** a `POST` with `multipart/form-data` and one field, `file`. The
  `X-CSRF-TOKEN` header is sent when a token is set.
- **On success**, return `200` with `{ "html": "<body inner HTML>" }`.
  Pictures should be inlined as `data:` URLs.
- **On failure**, return a non-2xx status with `{ "message": "Reason shown to the user" }`.

The RadioLens web app does this with LibreOffice, in
`app/Http/Controllers/Tools/DocToHtmlController.php`:
`soffice --headless --convert-to html:HTML`, after which it keeps only the
`<body>` markup and inlines the images. Any server that follows the contract
above works.

---

## Laravel example

**Migration**: store either format, or both.

```php
$table->longText('report_body')->nullable();       // HTML
$table->json('report_body_json')->nullable();      // JSON (optional)
```

**Blade**

```blade
<form method="POST" action="{{ route('reports.store') }}">
    @csrf
    <textarea name="report_body" id="report_body" hidden>{{ old('report_body', $report->report_body ?? '') }}</textarea>
    <div class="rl-editor-scope" id="report_body_editor"></div>
    <button type="submit" class="btn btn-success">Save</button>
</form>

<link rel="stylesheet" href="{{ asset('vendor/radiolens/radiolens-editor.css') }}">
<script src="{{ asset('vendor/radiolens/radiolens-editor.iife.js') }}"></script>
<script>
    RadiolensEditor.mount('#report_body_editor', {
        textarea: '#report_body',
        docConvertUrl: @json(route('tools.word_to_html')),
        minHeight: 600,
    });
</script>
```

`@csrf` gives the form its token. For `docConvertUrl`, the editor reads
`<meta name="csrf-token" content="{{ csrf_token() }}">` from the page layout, or
you can pass `csrfToken` yourself.

**Controller**

```php
public function store(Request $request)
{
    $validated = $request->validate([
        'report_body' => ['required', 'string'],
    ]);

    Report::create($validated);

    return back();
}
```

The editor's HTML output is safe for the editor itself to reopen. If you ever
display HTML that **users can submit** to other users, sanitize it on the
server (for example with HTMLPurifier) as you would any rich text input.

**Dynamic pages** (modals, rows added with JavaScript): call `mount` after the
markup exists and `destroy` before removing it.

```js
$('#reportModal').on('shown.bs.modal', () => RadiolensEditor.mount('#modal_editor', { textarea: '#modal_body' }));
$('#reportModal').on('hidden.bs.modal', () => RadiolensEditor.destroy('#modal_editor'));
```

---

## Styling and Bootstrap

The editor ships Tailwind CSS **precompiled**. The host project doesn't need
Tailwind.

- **Scoped styles.** Every rule is scoped under `.rl-editor-scope`, which is why
  the mount container needs that class. Without it, the editor renders
  unstyled.
- **Menus stay styled.** Dropdowns, tooltips and dialogs render into a
  `div.rl-editor-scope.rl-editor-portal` that the editor adds to `<body>`
  itself.
- **Bootstrap-safe.** The build is tuned for **Bootstrap 5** pages: 57 Tailwind
  class names clash with Bootstrap's (`.p-0`, `.table`, `.mb-1`…). The build
  (`build/scope-editor-css.mjs`) makes the editor's version win inside the
  scope. It reads `vendor/bootstrap.min.css` to know which names clash, so the
  output is identical to the RadioLens web app's.
- **Without Bootstrap.** The CSS works as-is; the Bootstrap-specific rules just
  have nothing to override.

---

## How saving works

1. **While typing.** The editor publishes on a **250 ms debounce**: 250 ms after
   the last change it writes the textarea and calls `onChange` /
   `onChangeHtml` / `onChangeValue`. Moving the caret isn't a change and
   doesn't start the timer.
2. **On submit.** Anything typed within the last 250 ms isn't in the textarea
   yet, so `mount` adds a **capture-phase `submit` listener** to the
   textarea's form that forces a publish first. This works with a jQuery
   `.submit()` too.
3. **Custom saves (AJAX).** Call `getHtml()` / `getValue()` directly. They
   always return the current document and ignore the debounce.

```js
$('#saveBtn').on('click', () => {
  $.post('/reports/1', { report_body: RadiolensEditor.getHtml('#report_body_editor') });
});
```

---

## Development

```bash
pnpm install              # dependency versions match the RadioLens web app's lockfile
pnpm run build            # all three outputs below
pnpm run typecheck        # tsc --noEmit
pnpm run test:js          # vitest (jsdom): 24 files, 400 tests
pnpm run test:browser     # Playwright in real Chrome: 22 files (needs the standalone build)
```

| Script | Output | Used by |
|---|---|---|
| `build:lib` | `dist/index.js`, `dist/style.css` | React / bundled apps. React, Plate and Radix are left as dependencies, so the host app doesn't ship two copies. |
| `build:standalone` | `dist/standalone/radiolens-editor.iife.js` (≈2.2 MB, ≈680 KB gzip), `radiolens-editor.css` | Plain HTML / Blade. Everything bundled; sets `window.RadiolensEditor`. |
| `build:types` | `dist/types/**/*.d.ts` | TypeScript users. `@/` imports are rewritten to relative paths. |

**Browser tests.** `tests/browser/harness.mjs` serves a page that loads
Bootstrap and the standalone build, just like the admin panel, then drives it
in Chrome. **Run `pnpm run build:standalone` before `test:browser`**, and
before `test:js` too: `tests/js/editor-css.test.ts` checks the built CSS.

**Package exports**

| Import | File |
|---|---|
| `@radiolens/editor` | `dist/index.js` (+ types) |
| `@radiolens/editor/style.css` | `dist/style.css` |
| `@radiolens/editor/standalone` | `dist/standalone/radiolens-editor.iife.js` |
| `@radiolens/editor/standalone.css` | `dist/standalone/radiolens-editor.css` |

---

## Project structure

```
RichText/
├── src/
│   ├── index.ts                  ESM package entry (public exports)
│   ├── entry.tsx                 mount API + window.RadiolensEditor (standalone entry)
│   ├── RadiolensEditor.tsx       the React component, toolbar, HTML/JSON load and save
│   ├── plugins.ts                Plate plugin list (buildPlugins(pasteMode))
│   ├── editor.css                Tailwind entry + design tokens
│   ├── components/               Plate kits and paste/import plugins
│   │   ├── *-kit.tsx             feature kits (tables, lists, fonts, media, links…)
│   │   ├── *-plugin.ts           Word/Excel paste fixes (line gap, borders, text boxes, pictures…)
│   │   ├── word-import-toolbar-button.tsx
│   │   └── ui/                   toolbar buttons, node renderers, Radix wrappers
│   ├── lib/                      pure logic: HTML serializer, font size, table widths, paste transforms
│   ├── hooks/                    use-base64-upload
│   └── utils/                    font helpers
├── tests/
│   ├── js/                       vitest unit/integration tests (json-value.test.tsx covers the JSON API)
│   └── browser/                  Playwright tests + harness
├── build/
│   ├── scope-editor-css.mjs      scopes the CSS under .rl-editor-scope, resolves Bootstrap clashes
│   ├── scope-editor-plugin.mjs   the Vite plugin that runs it (shared by both builds)
│   └── resolve-type-aliases.mjs  rewrites @/ imports in .d.ts files
├── vendor/bootstrap.min.css      the Bootstrap build the CSS is tuned against
├── vite.config.ts                ESM library build
├── vite.standalone.config.ts     IIFE standalone build
├── vitest.config.ts
├── tsconfig.json / tsconfig.build.json
└── package.json
```

Inside `src/`, `@/` points to `src/` (for example `@/lib/html-serializer`).

---

## Troubleshooting

| Problem | Cause / fix |
|---|---|
| Editor renders unstyled, or menus look broken | The mount container is missing `class="rl-editor-scope"`, or the CSS file isn't loaded. |
| `[RadiolensEditor] mount target not found` | `mount` ran before the element existed. Call it after the DOM is ready, or after inserting the markup. |
| `[RadiolensEditor] textarea not found` | Wrong `textarea` selector. |
| Server gets the old text | You're sending the textarea from your own AJAX code within 250 ms of typing. Use `getHtml()` / `getValue()` instead of reading the textarea. |
| `.doc` import says "not configured" | Set `docConvertUrl`. `.doc` needs the server. |
| Imported `.docx` lost colours and sizes | No `docConvertUrl`, so it used the in-browser fallback. Configure the endpoint. |
| `setValue` opened an empty document | The JSON was invalid (a console error says so) or wasn't an array of nodes. |
| `editor-css.test.ts` fails with "no such file" | Run `pnpm run build:standalone` first. |
| Two editors on one page | Give each its own container and textarea IDs. Each `mount` is independent. |

---

## Known issues

- **Flaky browser test.** `tests/browser/font-size-control.mjs`: the two
  "+ steps it by one point" checks sometimes fail when the whole suite runs,
  and pass when the file runs alone. The RadioLens web app's copy shows the
  same flake. It's a test-timing issue, not an editor bug.
- **`print-parity.mjs` is not here.** It checks the web app's Blade print
  views, so it stays in the web app.
- **React 18 only.** The peer range is `^18.3.0`; React 19 isn't tested yet.
- **Two copies of the code.** This package was copied from the web app's
  `resources/js/editor`, and the web app still builds its own copy. Until the
  web app installs this package, a fix made in one place must be copied to the
  other.

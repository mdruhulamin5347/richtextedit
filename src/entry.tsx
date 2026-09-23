/**
 * Blade-facing entry point.
 *
 * React is an ISLAND here: it owns the inside of one <div> and nothing else.
 * jQuery, Bootstrap, select2 and the ordinary form POST are untouched — the
 * editor's only job is to keep a hidden <textarea> filled with HTML, exactly
 * where Summernote used to put it.
 */
import { RadiolensEditor, type PasteMode, type RadiolensEditorHandle } from "@/RadiolensEditor";
import type { Value } from "platejs";
import React, { createRef } from "react";
import { createRoot, type Root } from "react-dom/client";
import "@/editor.css";

/**
 * What the hidden <textarea> carries to the server.
 *
 * "html" — serialized HTML, what Summernote stored and every existing report
 * body is. "json" — the Plate value as a JSON string, for a column that stores
 * the editor's own document instead.
 */
export type StorageFormat = "html" | "json";

export interface MountOptions {
  /** Hidden <textarea> that carries the value to the server. */
  textarea: string | HTMLTextAreaElement;
  /** Format the textarea is read in and written back as. Defaults to "html". */
  format?: StorageFormat;
  /** Endpoint converting a legacy binary .doc (route('tools.word_to_html')). */
  docConvertUrl?: string;
  /** CSRF token for docConvertUrl. Defaults to <meta name="csrf-token">. */
  csrfToken?: string;
  placeholder?: string;
  minHeight?: number;
  pasteMode?: PasteMode;
  /** Called on every (debounced) change, with both representations. */
  onChange?: (change: { html: string; value: Value }) => void;
}

interface Instance {
  root: Root;
  handle: React.RefObject<RadiolensEditorHandle | null>;
  container: Element;
  form: HTMLFormElement | null;
  onSubmit?: (e: Event) => void;
}

const instances = new Map<Element, Instance>();

function resolve<T extends Element>(target: string | T): T | null {
  return typeof target === "string" ? document.querySelector<T>(target) : target;
}

function csrf(): string | undefined {
  return document.querySelector<HTMLMetaElement>('meta[name="csrf-token"]')?.content;
}

function mount(selector: string | Element, options: MountOptions) {
  const container = resolve(selector as string);
  if (!container) {
    console.error("[RadiolensEditor] mount target not found:", selector);
    return null;
  }
  if (instances.has(container)) return instances.get(container)!.handle;

  const textarea = resolve(options.textarea as string) as HTMLTextAreaElement | null;
  if (!textarea) {
    console.error("[RadiolensEditor] textarea not found:", options.textarea);
    return null;
  }

  const format: StorageFormat = options.format ?? "html";
  const stored = textarea.value || "";

  const handle = createRef<RadiolensEditorHandle>();
  const root = createRoot(container);

  // Latest HTML from the debounced publish, so onChange can report both forms
  // without serializing a second time. onChangeHtml always fires first.
  let latestHtml = "";

  root.render(
    <RadiolensEditor
      ref={handle}
      initialHtml={format === "html" ? stored : undefined}
      initialValue={format === "json" ? stored : undefined}
      onChangeHtml={(html) => {
        latestHtml = html;
        if (format === "html") textarea.value = html;
      }}
      // Only when something wants the JSON: each call deep-copies the
      // document, which is not free with base64 images in it.
      onChangeValue={
        format === "json" || options.onChange
          ? (value) => {
              if (format === "json") textarea.value = JSON.stringify(value);
              options.onChange?.({ html: latestHtml, value });
            }
          : undefined
      }
      docConvertUrl={options.docConvertUrl}
      csrfToken={options.csrfToken ?? csrf()}
      placeholder={options.placeholder}
      minHeight={options.minHeight}
      pasteMode={options.pasteMode}
    />
  );

  // Force a serialize before the POST.
  //
  // The editor publishes on a 250ms debounce, so clicking Create within a
  // quarter-second of the last keystroke would otherwise post slightly stale
  // HTML. Capture phase, because a jQuery-triggered submit still dispatches a
  // real DOM event but listeners on the form itself may stopPropagation.
  const form = textarea.closest("form");
  const onSubmit = () => {
    handle.current?.flush();
  };
  form?.addEventListener("submit", onSubmit, { capture: true });

  const instance: Instance = { root, handle, container, form, onSubmit };
  instances.set(container, instance);
  return handle;
}

function destroy(selector: string | Element) {
  const container = resolve(selector as string);
  if (!container) return;
  const instance = instances.get(container);
  if (!instance) return;
  if (instance.form && instance.onSubmit) {
    instance.form.removeEventListener("submit", instance.onSubmit, { capture: true });
  }
  instance.root.unmount();
  instances.delete(container);
}

/** Serialize now and return the HTML — useful for custom submit flows. */
function getHtml(selector: string | Element): string {
  const container = resolve(selector as string);
  if (!container) return "";
  return instances.get(container)?.handle.current?.getHtml() ?? "";
}

/** Replace the document — the equivalent of `$(el).summernote('code', html)`. */
function setHtml(selector: string | Element, html: string) {
  const container = resolve(selector as string);
  if (!container) return;
  instances.get(container)?.handle.current?.setHtml(html);
}

/** Serialize now and return the document as JSON (a Plate `Value`). */
function getValue(selector: string | Element): Value {
  const container = resolve(selector as string);
  if (!container) return [];
  return instances.get(container)?.handle.current?.getValue() ?? [];
}

/** Replace the document from JSON — a Plate `Value` or its JSON string. */
function setValue(selector: string | Element, value: Value | string) {
  const container = resolve(selector as string);
  if (!container) return;
  instances.get(container)?.handle.current?.setValue(value);
}

const api = { mount, destroy, getHtml, setHtml, getValue, setValue };

// Guarded so importing the package during SSR does not throw.
if (typeof window !== "undefined") {
  (window as any).RadiolensEditor = api;
}

export default api;

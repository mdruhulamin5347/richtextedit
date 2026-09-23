import { AlignToolbarButton } from "@/components/ui/align-toolbar-button";
import { Editor, EditorContainer } from "@/components/ui/editor";
import { FixedToolbar } from "@/components/ui/fixed-toolbar";
import { FontColorToolbarButton } from "@/components/ui/font-color-toolbar-button";
import { FontSizeToolbarButton } from "@/components/ui/font-size-toolbar-button";
import { RedoToolbarButton, UndoToolbarButton } from "@/components/ui/history-toolbar-button";
import { InsertToolbarButton } from "@/components/ui/insert-toolbar-button";
import { LineHeightToolbarButton } from "@/components/ui/line-height-toolbar-button";
import { LinkToolbarButton } from "@/components/ui/link-toolbar-button";
import {
  BulletedListToolbarButton,
  NumberedListToolbarButton,
} from "@/components/ui/list-toolbar-button";
import { MarkToolbarButton } from "@/components/ui/mark-toolbar-button";
import { MediaToolbarButton } from "@/components/ui/media-toolbar-button";
import { TableToolbarButton } from "@/components/ui/table-toolbar-button";
import { ToolbarGroup } from "@/components/ui/toolbar";
import { TurnIntoToolbarButton } from "@/components/ui/turn-into-toolbar-button";
import { WordImportToolbarButton } from "@/components/word-import-toolbar-button";
import { EMPTY_VALUE, plateValueToHtml } from "@/lib/html-serializer";
import { inlineInheritedColor } from "@/lib/inherited-color";
import { inlineInheritedFontSize } from "@/lib/inherited-font-size";
import { inlineLegacyAlignment } from "@/lib/legacy-alignment";
import { inlineWordTextboxes } from "@/lib/word-textbox";
import { protectWhitespace } from "@/lib/whitespace";
import { buildPlugins, type PasteMode } from "@/plugins";
import { FontFamilyPlugin } from "@platejs/basic-styles/react";
import {
  BaselineIcon,
  BoldIcon,
  ItalicIcon,
  PaintBucketIcon,
  StrikethroughIcon,
  UnderlineIcon,
} from "lucide-react";
import { deserializeHtml, KEYS, type SlateEditor, type Value } from "platejs";
import { Plate, useEditorRef, usePlateEditor, type PlateEditor } from "platejs/react";
import React, { useCallback, useEffect, useImperativeHandle, useMemo, useRef } from "react";

/**
 * Same list the Summernote build offered, so a report author's font choices
 * carry over unchanged.
 */
const FONT_FAMILIES = [
  { label: "Default", value: "" },
  { label: "Calibri", value: "Calibri, sans-serif" },
  { label: "Arial Narrow", value: "'Arial Narrow', Arial, sans-serif" },
  { label: "Arial", value: "Arial, sans-serif" },
  { label: "Arial Unicode MS", value: "'Arial Unicode MS', Arial, sans-serif" },
  { label: "Times New Roman", value: "'Times New Roman', serif" },
  { label: "Tahoma", value: "Tahoma, sans-serif" },
];

/**
 * A stored report body, ready for Plate's deserializer.
 *
 * The clipboard gets both of these from plugins — `LegacyAlignmentPlugin` and
 * `WhitespacePlugin` hook `transformData` — but `transformData` is a
 * CLIPBOARD-ONLY seam, so the load path has to run them itself, exactly as the
 * Word-import button does.
 *
 * `inlineLegacyAlignment` was the half that was missing, and it is not a
 * paste-only concern: a report body in the database states its alignment the
 * legacy way all the time — `<p align=center>`, `<center>`, and `<tr
 * align=center>`, which centres a whole row and survives as no node at all
 * unless it is pushed onto the cells. None of that reaches a node from a plain
 * `text-align` read, so a centred column opened FLAT LEFT, and the first save
 * then wrote the report back without it. Pasting the very same markup came out
 * centred, which is what made it look intermittent.
 *
 * `inlineInheritedFontSize` is here for exactly the same reason, and it is the
 * same story again: a size the report states on a `<p>`, a `<td>` or the
 * `<table>` reached no node, so body text opened at the editor's 18px base —
 * larger than the document — and the first save wrote the size out of the
 * report altogether.
 *
 * `inlineInheritedColor` is the third of the same kind, and the one a reader
 * notices first: a colour stated on a `<p>`, a `<td>` or a `<font color>` — how
 * a legacy report writes an abnormal result — reached no node either, so the
 * report opened in plain black and saving it made that permanent.
 *
 * Alignment and size first, whitespace last: the same order the Word-import
 * button uses. Both of those re-serialize the document, and `protectWhitespace`
 * has to see the result of that, not the other way round.
 */
function prepareHtml(html: string): string {
  // Text boxes first, on the rawest HTML — the same reason the clipboard runs
  // that pass first of all. A stored body can have kept one whole: Summernote
  // put the clipboard into a contenteditable, comments and all.
  return protectWhitespace(
    inlineInheritedColor(
      inlineInheritedFontSize(inlineLegacyAlignment(inlineWordTextboxes(html)))
    )
  );
}

/**
 * A deserialized report body as a value the editor can actually RENDER.
 *
 * Plate wraps stray root-level runs in a paragraph only when the fragment MIXES
 * blocks and inlines — see `normalizeDifferentNodeTypes` in its core. A body
 * that is inline ALL the way down stays a list of text nodes at the root, which
 * is not a document: slate-react reads `children` off one while painting and
 * throws "undefined is not iterable", which takes the whole editor down with it
 * and leaves the page with a dead box where the report should be.
 *
 * Not a corner case in stored data, which is what makes it worth guarding: a
 * body saved as bare text (`Normal study.`), as `<span>…</span>`, as `<b>…</b>`
 * or with an empty `<div></div>` in it all land here. The clipboard never did —
 * an insert goes INTO a block that already exists — so it took a legacy report
 * with no block wrapper to see it.
 */
function toRenderableValue(editor: SlateEditor, value: Value): Value {
  const out: any[] = [];
  let run: any[] | null = null;

  for (const node of value as any[]) {
    if (Array.isArray(node?.children) && !editor.api.isInline(node)) {
      run = null;
      out.push(node);
      continue;
    }
    // Consecutive runs share the paragraph they were missing, so one line does
    // not come back as one paragraph per span.
    if (run) {
      run.push(node);
      continue;
    }
    run = [node];
    out.push({ type: editor.getType(KEYS.p), children: run });
  }

  return (out.length ? out : EMPTY_VALUE) as Value;
}

/** A stored report body, prepared, deserialized and safe to render. */
function deserializeReport(editor: SlateEditor, html?: string): Value {
  if (!html || !html.trim()) return EMPTY_VALUE;

  return toRenderableValue(editor, deserializeHtml(editor, { element: prepareHtml(html) }) as Value);
}

/**
 * A stored JSON body (a Plate `Value`, or that value as a JSON string), made
 * safe to render.
 *
 * JSON skips the HTML pipeline entirely — it is the editor's own document, so
 * there is nothing legacy to inline — but it still goes through
 * `toRenderableValue`: a value built by hand, or saved by another tool, can
 * carry bare text nodes at the root just as a legacy HTML body can, and that
 * takes the editor down the same way.
 *
 * Cloned on the way in, so the caller's object never becomes the editor's live
 * state: Slate mutates-by-replacement, and a caller that kept a reference and
 * later edited it would otherwise be editing a node the editor holds.
 */
function deserializeJson(editor: SlateEditor, value?: Value | string | null): Value {
  if (value == null) return EMPTY_VALUE;

  let parsed: unknown = value;
  if (typeof value === "string") {
    if (!value.trim()) return EMPTY_VALUE;
    try {
      parsed = JSON.parse(value);
    } catch {
      console.error("[RichTextEdit] setValue: not valid JSON");
      return EMPTY_VALUE;
    }
  }
  if (!Array.isArray(parsed) || parsed.length === 0) return EMPTY_VALUE;

  return toRenderableValue(editor, structuredClone(parsed) as Value);
}

/**
 * How long the editor sits on a change before handing HTML to the <textarea>.
 *
 * This is a LAG, not merely a throttle: the form posts whatever the textarea
 * holds, so anything typed inside this window is not yet in what a submit would
 * send. The submit handler force-flushes for exactly that reason — see
 * `entry.tsx`. 250ms matches the upstream editor.
 */
const PUBLISH_DEBOUNCE_MS = 250;

export interface RichTextEditProps {
  /** Existing report_body HTML (legacy Summernote content included). */
  initialHtml?: string;
  /**
   * Existing body as JSON (a Plate `Value` or its JSON string). Takes
   * precedence over `initialHtml` when both are given.
   */
  initialValue?: Value | string | null;
  /** Receives serialized HTML on every (debounced) change. */
  onChangeHtml?: (html: string) => void;
  /** Receives the JSON value on every (debounced) change. */
  onChangeValue?: (value: Value) => void;
  /** Endpoint that converts a legacy binary .doc server-side. */
  docConvertUrl?: string;
  csrfToken?: string;
  placeholder?: string;
  minHeight?: number;
  /**
   * "clean" applies the PasteNormalizationPlugin: Word's margins and
   * letter/word-spacing are dropped, while its font sizes and line-heights are
   * converted onto the scales the editor's own controls use.
   * "faithful" leaves every value exactly as the source stated it.
   */
  pasteMode?: PasteMode;
}

export interface RichTextEditHandle {
  /** Serialize now, bypassing the debounce. Used by the form submit hook. */
  flush: () => string;
  getHtml: () => string;
  setHtml: (html: string) => void;
  /** The document as JSON — a deep copy, safe to store or mutate. */
  getValue: () => Value;
  /** Replace the document from JSON (a Plate `Value` or its JSON string). */
  setValue: (value: Value | string) => void;
  focus: () => void;
  getEditor: () => PlateEditor | null;
}

/**
 * Memoised and prop-free, so it renders once and stays put: the toolbar's ~20
 * buttons must not re-render on every keystroke. Each button holds its own
 * subscription to the marks it reflects, so they still light up on selection.
 */
const ReportToolbar = React.memo(function ReportToolbar({
  docConvertUrl,
  csrfToken,
}: {
  docConvertUrl?: string;
  csrfToken?: string;
}) {
  const editor = useEditorRef();

  const setFontFamily = useCallback(
    (value: string) => {
      if (value) {
        editor.tf.addMark(FontFamilyPlugin.key, value);
      } else {
        editor.tf.removeMark(FontFamilyPlugin.key);
      }
      editor.tf.focus();
    },
    [editor]
  );

  return (
    <div className="flex flex-wrap gap-0 p-0">
      <ToolbarGroup>
        <UndoToolbarButton />
        <RedoToolbarButton />
      </ToolbarGroup>

      <ToolbarGroup>
        <InsertToolbarButton />
        <TurnIntoToolbarButton />
        <select
          className="border-input bg-background text-foreground focus:ring-ring h-7 rounded-md border px-2 text-xs outline-none focus:ring-1"
          onChange={(e) => setFontFamily(e.target.value)}
          defaultValue=""
        >
          {FONT_FAMILIES.map((font) => (
            <option key={font.value} value={font.value} style={{ fontFamily: font.value || "inherit" }}>
              {font.label}
            </option>
          ))}
        </select>
        <FontSizeToolbarButton />
      </ToolbarGroup>

      <ToolbarGroup>
        <MarkToolbarButton nodeType={KEYS.bold} tooltip="Bold (Ctrl+B)">
          <BoldIcon className="size-3" />
        </MarkToolbarButton>
        <MarkToolbarButton nodeType={KEYS.italic} tooltip="Italic (Ctrl+I)">
          <ItalicIcon className="size-3" />
        </MarkToolbarButton>
        <MarkToolbarButton nodeType={KEYS.underline} tooltip="Underline (Ctrl+U)">
          <UnderlineIcon className="size-3" />
        </MarkToolbarButton>
        <MarkToolbarButton nodeType={KEYS.strikethrough} tooltip="Strikethrough">
          <StrikethroughIcon className="size-3" />
        </MarkToolbarButton>
        <FontColorToolbarButton nodeType={KEYS.color} tooltip="Text color">
          <BaselineIcon className="size-3" />
        </FontColorToolbarButton>
        <FontColorToolbarButton nodeType={KEYS.backgroundColor} tooltip="Background color">
          <PaintBucketIcon className="size-3" />
        </FontColorToolbarButton>
      </ToolbarGroup>

      <ToolbarGroup>
        <AlignToolbarButton />
        <LineHeightToolbarButton />
        <NumberedListToolbarButton />
        <BulletedListToolbarButton />
      </ToolbarGroup>

      <ToolbarGroup>
        <LinkToolbarButton />
        <TableToolbarButton />
        <MediaToolbarButton nodeType={KEYS.img} />
      </ToolbarGroup>

      <ToolbarGroup>
        <WordImportToolbarButton docConvertUrl={docConvertUrl} csrfToken={csrfToken} />
      </ToolbarGroup>
    </div>
  );
});

export type { PasteMode };

export const RichTextEdit = React.forwardRef<RichTextEditHandle, RichTextEditProps>(
  function RichTextEdit(
    {
      initialHtml = "",
      initialValue,
      onChangeHtml,
      onChangeValue,
      docConvertUrl,
      csrfToken,
      placeholder = "Type the report here…",
      minHeight = 600,
      pasteMode = "clean",
    },
    ref
  ) {
    const debounceRef = useRef<ReturnType<typeof setTimeout> | null>(null);
    const lastHtmlRef = useRef<string>(initialHtml);
    const onChangeRef = useRef(onChangeHtml);
    onChangeRef.current = onChangeHtml;
    const onChangeValueRef = useRef(onChangeValue);
    onChangeValueRef.current = onChangeValue;

    const plugins = useMemo(() => buildPlugins(pasteMode), [pasteMode]);

    const editor = usePlateEditor({
      plugins,
      // Legacy Summernote HTML is deserialized through Plate's OWN pipeline —
      // the same one that handles a Word paste — rather than a hand-written
      // parser. So existing report bodies get the benefit of the very engine
      // this editor was brought in for.
      // `prepareHtml` for the same reason the clipboard path gets its two
      // transforms from plugins: a legacy Summernote report indents with
      // `&nbsp;` and states its alignment with `align=`, and neither survives
      // Plate's deserializer untouched. See prepareHtml.
      value: (ed) =>
        initialValue != null ? deserializeJson(ed, initialValue) : deserializeReport(ed, initialHtml),
      // Typed as possibly null only because `strict` is off, which makes the
      // hook's `enabled` conditional type resolve to null. It is null only with
      // `enabled: false`, which is never passed.
    }) as PlateEditor;

    const serialize = useCallback((): string => {
      const html = plateValueToHtml(editor.children as Value);
      lastHtmlRef.current = html;
      return html;
    }, [editor]);

    const snapshot = useCallback(
      (): Value => structuredClone(editor.children) as Value,
      [editor]
    );

    /** Hand the current document to both listeners, HTML and JSON. */
    const publish = useCallback((): string => {
      const html = serialize();
      onChangeRef.current?.(html);
      if (onChangeValueRef.current) onChangeValueRef.current(snapshot());
      return html;
    }, [serialize, snapshot]);

    const handleChange = useCallback(
      ({ editor: changed }: { editor: PlateEditor; value: Value }) => {
        // Moving the caret is not a change to the report and must not start the
        // timer: re-rendering inside slate-react's 100ms throttled selection
        // commit drags the caret back to the line just left.
        const ops = changed?.operations;
        if (ops?.length && ops.every((op) => op.type === "set_selection")) return;

        if (debounceRef.current) clearTimeout(debounceRef.current);
        debounceRef.current = setTimeout(publish, PUBLISH_DEBOUNCE_MS);
      },
      [publish]
    );

    useEffect(() => {
      return () => {
        if (debounceRef.current) clearTimeout(debounceRef.current);
      };
    }, []);

    useImperativeHandle(
      ref,
      () => ({
        flush: () => {
          if (debounceRef.current) {
            clearTimeout(debounceRef.current);
            debounceRef.current = null;
          }
          return publish();
        },
        getHtml: () => serialize(),
        setHtml: (html: string) => {
          editor.tf.setValue(deserializeReport(editor, html));
          publish();
        },
        getValue: () => snapshot(),
        setValue: (value: Value | string) => {
          editor.tf.setValue(deserializeJson(editor, value));
          publish();
        },
        focus: () => editor.tf.focus(),
        getEditor: () => editor,
      }),
      [editor, serialize, snapshot, publish]
    );

    return (
      <div className="bg-background flex w-full flex-col overflow-clip rounded-md border">
        <Plate editor={editor} onChange={handleChange}>
          <FixedToolbar className="border-border bg-muted/30 border-b">
            <ReportToolbar docConvertUrl={docConvertUrl} csrfToken={csrfToken} />
          </FixedToolbar>
          <EditorContainer style={{ minHeight }}>
            {/*
              The min-height belongs on the EDITABLE too, not only on the box
              around it. The editable is as tall as its content, so on a short
              report it covered ~50px of a 600px frame and every click in the
              space below it landed on the container: no focus, no caret, and
              typing went nowhere. Sized like this the editable fills the frame,
              so a click anywhere in it puts the caret on the nearest line.
            */}
            <Editor
              variant="report"
              // A tab is as wide as the distance to the next tab stop, and that
              // distance is `tab-size` — 8 spaces by default, which is what the
              // print pages, the PDF and the old Summernote editor all use.
              // Tailwind's preflight sets 4 here, so a pasted report lined its
              // columns up one way on screen and another way on paper.
              className="[tab-size:8]"
              placeholder={placeholder}
              spellCheck={false}
              style={{ minHeight }}
            />
          </EditorContainer>
        </Plate>
      </div>
    );
  }
);

export default RichTextEdit;

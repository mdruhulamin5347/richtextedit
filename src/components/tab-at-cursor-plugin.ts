import { createSlatePlugin, KEYS } from "platejs";

/**
 * Tab puts a tab in AT THE CARET, rather than indenting the whole line.
 *
 * `@platejs/indent` overrides `editor.tf.tab` to move the BLOCK: it sets
 * `indent` on the paragraph, which the serializer writes as `margin-left:
 * indent * 40px`. So the caret's position made no difference — a Tab pressed in
 * the middle of `Findings of the study` moved the whole line 40px right and left
 * the text untouched, which is not what a tab is for in a report. A report is
 * full of lines that line a value up with a tab stop part-way along, and typing
 * one was impossible.
 *
 * What is DELIBERATELY left alone:
 *  - Shift+Tab (`options.reverse`), so outdenting a block still works and a
 *    report pasted from Word with real indentation can still be flattened.
 *  - A list item, where Tab nests the item under the one above it. That is what
 *    Tab means in every editor and nobody asked for it back.
 *  - An EXPANDED selection, which still indents the block. Inserting there would
 *    do what typing any other character does — replace the selected text — and
 *    Tab is not a character anyone means to type over a selection. It also
 *    leaves a way to indent a block from the keyboard.
 * All three fall through to the override chain — the list plugin's `tab`, then
 * the indent plugin's — so those keep whatever behaviour they shipped with.
 *
 * Registered AFTER IndentKit in `plugins.ts`, which is what puts this override
 * outermost: `overrideEditor` wraps what is already there, so the plugin that
 * registers last is the one that decides first.
 *
 * The inserted character is a real `\t`. It survives the round trip — the
 * serializer wraps it in `white-space: pre` (see lib/whitespace.ts) and the
 * editable carries `[tab-size:0.5in]`, Word's stop and the one the print page
 * states too — so what is typed here is what prints.
 */
export const TabAtCursorPlugin = createSlatePlugin({
  key: "tabAtCursor",
}).overrideEditor(({ editor, tf: { tab } }) => ({
  transforms: {
    // `options` is typed by the transform it overrides — `{ reverse: boolean }`,
    // which the chain below requires in full. Annotating it here as optional is
    // what made it unassignable when passed straight back down.
    tab(options) {
      if (options.reverse) return tab(options);
      if (!editor.selection || editor.api.isExpanded()) return tab(options);
      if (editor.api.above({ match: { type: editor.getType(KEYS.li) } })) {
        return tab(options);
      }

      editor.tf.insertText("\t");

      return true;
    },
  },
}));

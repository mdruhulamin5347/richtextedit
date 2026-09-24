import { createSlatePlugin, KEYS, type SlateEditor } from "platejs";

import { isOfficeClipboard, layOutOfficePaste } from "@/lib/office-tab-stops";
import { expandWordTabSpans } from "@/lib/word-tabs";

/**
 * Keep every tab of a Word or LibreOffice paste, where the document puts it.
 *
 * Hooks the HTML plugin's `transformData`, the same seam DocxPlugin uses, and
 * runs after it: whatever `mso-tab-count` span is still there once Plate's docx
 * cleaner has been through is one it did not recognise (lib/word-tabs.ts), and
 * the stops and indents the cleaner and Juice leave on each block are what the
 * layout is worked out from (lib/office-tab-stops.ts).
 *
 * The office check reads the RAW clipboard: by now the cleaner has removed the
 * markers it would look for.
 */
export const WordTabPlugin = createSlatePlugin({
  key: "wordTabs",
  inject: {
    plugins: {
      [KEYS.html]: {
        parser: {
          transformData: ({
            data,
            dataTransfer,
            editor,
          }: {
            data: string;
            dataTransfer: DataTransfer;
            editor: SlateEditor;
          }) => {
            const html = expandWordTabSpans(data);
            if (!isOfficeClipboard(dataTransfer.getData("text/html"))) return html;
            return layOutOfficePaste(html, {
              editable: editableOf(editor),
              rtf: dataTransfer.getData("text/rtf"),
            });
          },
        },
      },
    },
  },
});

/** The contenteditable, whose font and tab stops are measured in — if mounted. */
function editableOf(editor: SlateEditor): HTMLElement | null {
  try {
    return editor.api.toDOMNode(editor) ?? null;
  } catch {
    return null;
  }
}

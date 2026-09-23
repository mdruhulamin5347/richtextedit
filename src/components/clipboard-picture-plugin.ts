"use client";

import { createPlatePlugin } from "platejs/react";

import { imageFilesOf, isLonePicturePaste } from "@/lib/clipboard-picture";

/**
 * Paste the picture the clipboard is carrying when its HTML only names one.
 *
 * A PASTE HANDLER rather than an `insertData` override, because that is where
 * the decision is actually made. Plate's own placeholder plugin inserts a
 * pasted image file only when the clipboard carries NO html:
 *
 *   if (files.length > 0 && !types.includes('text/html')) …
 *
 * — which is the right rule almost always, and the reason a screenshot pastes
 * as a picture while a copied web page pastes as a page. Word's clipboard trips
 * it: there IS html, so the files are ignored, and the html turns out to name a
 * picture at a `file:///` path the browser may not read. Nothing is inserted at
 * all — the bytes were on the clipboard the whole time.
 *
 * So this handles exactly the paste that would otherwise insert nothing: one
 * unreadable picture, no text, one image file. Everything else falls straight
 * through to the rule above. See lib/clipboard-picture.ts for why the case is
 * drawn that narrowly — a selection with text in it renders as a picture of the
 * WHOLE selection, which is not the report.
 */
export const ClipboardPicturePlugin = createPlatePlugin({
  key: "clipboardPicture",
}).extend(() => ({
  handlers: {
    onPaste: ({ editor, event }: { editor: any; event: React.ClipboardEvent }) => {
      const clipboard = event.clipboardData;
      if (!clipboard || !isLonePicturePaste(clipboard.getData("text/html"))) return false;

      const files = imageFilesOf(clipboard);
      if (files.length !== 1) return false;

      event.preventDefault();
      event.stopPropagation();
      // The same call Plate makes for a screenshot: a placeholder that reads the
      // file and stores it as a data URL, which is how every other picture in a
      // report body is kept.
      editor.tf.insert.media(files, { nextBlock: false });

      return true;
    },
  },
}));

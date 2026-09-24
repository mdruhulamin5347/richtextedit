import { createSlatePlugin, KEYS } from "platejs";

import { inlineFontFaces, inlineRtfParagraphFonts } from "@/lib/font-face";
import { isLibreOfficeClipboard } from "@/lib/word-line-gap";

/**
 * Restate `<font face>` as CSS, on the way in from the clipboard.
 *
 * Has to run BEFORE the docx cleaner, which turns every `<font>` into a `<span>`
 * and keeps `face` only as an attribute nothing reads. See lib/font-face.ts.
 */
export const FontFacePlugin = createSlatePlugin({
  key: "fontFace",
  inject: {
    plugins: {
      [KEYS.html]: {
        parser: {
          transformData: ({ data }: { data: string }) => inlineFontFaces(data),
        },
      },
    },
  },
});

/**
 * Give a LibreOffice line that names no font the one its RTF names.
 *
 * Runs AFTER the docx cleaner and Juice, so a family the style block gives a
 * paragraph is already inline and wins; and BEFORE WordTabPlugin, whose tab
 * layout measures each line in the font it will be drawn in. The check for
 * LibreOffice reads the RAW clipboard: by now the cleaner has removed the
 * generator `<meta>` it names itself in. See lib/font-face.ts.
 */
export const RtfParagraphFontPlugin = createSlatePlugin({
  key: "rtfParagraphFont",
  inject: {
    plugins: {
      [KEYS.html]: {
        parser: {
          transformData: ({ data, dataTransfer }: { data: string; dataTransfer: DataTransfer }) =>
            isLibreOfficeClipboard(dataTransfer.getData("text/html"))
              ? inlineRtfParagraphFonts(data, dataTransfer.getData("text/rtf"))
              : data,
        },
      },
    },
  },
});

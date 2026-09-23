import { createSlatePlugin, KEYS } from "platejs";

import { restoreRtfPictures } from "@/lib/rtf-pictures";

/**
 * Put the pictures the clipboard's RTF carries where its HTML only named them.
 *
 * FIRST in the plugin list, which makes it LAST to run — deliberately after
 * Plate's own docx cleaner. That one resolves a FLOATING picture through
 * `v:shapes` → `o:spid` → `shplid`; whatever is still pointing at a `file:///`
 * path afterwards is an INLINE picture, which has none of those and which the
 * cleaner cannot match. Running in this order also keeps the count honest: the
 * pictures already resolved are no longer waiting for one.
 *
 * See lib/rtf-pictures.ts.
 */
export const RtfPicturePlugin = createSlatePlugin({
  key: "rtfPicture",
  inject: {
    plugins: {
      [KEYS.html]: {
        parser: {
          transformData: ({ data, dataTransfer }: { data: string; dataTransfer: DataTransfer }) =>
            restoreRtfPictures(data, dataTransfer?.getData("text/rtf") ?? ""),
        },
      },
    },
  },
});

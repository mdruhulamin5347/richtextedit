"use client";

import { createSlatePlugin, KEYS } from "platejs";

/**
 * Carry a pasted picture's WIDTH onto the node.
 *
 * The size an image is shown at is a decision the author made in the document —
 * a full-page scan dragged down to half a page, a signature block sized to a
 * line — and Plate's deserializer reads none of it: neither `<img width=96>`,
 * which is what LibreOffice and Word both write, nor a `width` in the style.
 * So every pasted picture arrived with no width and rendered at its NATURAL
 * size, which for a phone photo of a form means the whole page.
 *
 * The serializer could always write it back (`n.width` in lib/html-serializer.ts)
 * and the editor's own resize handles have always stored it there. Nothing was
 * putting it there on the way IN — the same gap as a block's colour and size.
 *
 * Height is deliberately not carried: the serializer writes `height: auto` so a
 * picture keeps its aspect ratio at whatever width it is given, which is what
 * the editor's own resizing does too.
 */

/** A width in whole px, or undefined for one that says nothing here. */
function imageWidth(element: HTMLElement): number | undefined {
  // The style wins over the attribute, as it does in a browser.
  const stated = element.style?.width || element.getAttribute("width") || "";
  const match = String(stated).trim().match(/^([\d.]+)\s*(px)?$/i);
  if (!match) return undefined;

  const width = Math.round(Number.parseFloat(match[1]));

  return Number.isFinite(width) && width > 0 ? width : undefined;
}

export const ImageSizePlugin = createSlatePlugin({
  key: "imageSize",
  inject: {
    plugins: {
      [KEYS.img]: {
        parsers: {
          html: {
            deserializer: {
              parse: ({ element }: { element: HTMLElement }) => {
                const width = imageWidth(element);

                return width ? { width } : {};
              },
            },
          },
        },
      },
    },
  },
});

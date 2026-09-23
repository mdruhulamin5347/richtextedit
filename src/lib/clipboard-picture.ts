/**
 * A picture the clipboard's HTML only NAMES, but whose bytes it also carries.
 *
 * Copying a picture in Word puts several things on the clipboard at once: HTML
 * that points at a file on the machine doing the copying
 * (`file:///C:/…/clip_image001.png`), the bytes in the RTF flavour, and — from
 * Windows itself — a plain image rendition of the selection.
 *
 * Plate reads the first two. It deliberately ignores the third for anything
 * that came from Word, and it is right to: what Windows renders is the WHOLE
 * SELECTION, so a header table copied with its QR code would come back as one
 * flat picture of the entire header, text and all, in place of the report.
 *
 * But when the selection was nothing BUT a picture, that rendition is the
 * picture, and it is the only copy of it left when the RTF is missing — which
 * is what happens with an image linked into the document rather than embedded,
 * or a clipboard that crossed a remote desktop. Today that paste inserts
 * NOTHING at all: the HTML names a file the browser may not read, and the bytes
 * sitting right there go unused.
 *
 * So: exactly one picture, that picture unreadable, no text anywhere, and
 * exactly one image on the clipboard. Anything less certain is left alone.
 */

/** Addresses a page can never fetch, whatever the file behind them. */
const UNREACHABLE = /^\s*(file:|cid:|about:|res:|x-raw-image:)/i;

function isDeadSrc(src: string | null): boolean {
  const value = (src ?? "").trim();

  return value === "" || UNREACHABLE.test(value);
}

/**
 * Is this clipboard HTML a single picture the browser cannot load, and nothing
 * else at all?
 */
export function isLonePicturePaste(html: string): boolean {
  if (!html || !/<img/i.test(html)) return false;

  const { body } = new DOMParser().parseFromString(html, "text/html");

  const images = Array.from(body.querySelectorAll("img"));
  if (images.length !== 1 || !isDeadSrc(images[0].getAttribute("src"))) return false;

  // Any text at all and this is a selection, not a picture: whatever Windows
  // rendered for it is a picture of that whole selection.
  if ((body.textContent || "").replace(/\u00a0/g, " ").trim() !== "") return false;

  // Nor anything else that carries content of its own.
  return body.querySelector("table, ul, ol, hr, canvas, svg, video, iframe") === null;
}

/** The images the clipboard carries as files. */
export function imageFilesOf(dataTransfer: DataTransfer): File[] {
  return Array.from(dataTransfer.files ?? []).filter((file) => file.type.startsWith("image/"));
}

/**
 * NOTE on where this is used: components/clipboard-picture-plugin.ts, from a
 * paste HANDLER. An `insertData` override cannot do it — Plate inserts a pasted
 * image file from `onPaste`, reading the raw event rather than the clipboard
 * that reaches `insertData`, so a swap made there is never seen.
 */

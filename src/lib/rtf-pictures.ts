/**
 * The pictures Word ships in the clipboard's RTF, matched to the HTML by ORDER.
 *
 * Word's clipboard HTML never carries a picture. It names one by a path on the
 * machine doing the copying — `file:///C:/…/clip_image001.png` — and puts the
 * bytes in the RTF flavour beside it. Plate reads that flavour already, but
 * only for a FLOATING picture: it matches an `<img>` to its bytes through
 * `v:shapes` → `o:spid` → `shplid`, and an INLINE picture, which is what a
 * picture dropped into a table cell is, has none of those. Word writes it as a
 * plain `<img id="Picture_x0020_1" src="file:///…">` and its bytes as a bare
 * `{\pict}` group.
 *
 * So the QR code in a report header came through as an empty frame while
 * LibreOffice, reading the same clipboard, showed it — the bytes were there the
 * whole time, unmatched.
 *
 * Matching is by document order, which is what Word guarantees and what every
 * RTF→HTML converter relies on: the Nth picture in the RTF is the Nth picture
 * in the document. To keep that honest this only acts when the two sides agree
 * exactly — every picture still unresolved, against every picture the RTF
 * holds. One more or one fewer and nothing is touched, because a picture in the
 * wrong place is worse than a picture that did not arrive.
 */

/** Addresses a page can never fetch, so a picture still waiting for its bytes. */
const UNREACHABLE = /^\s*(file:|cid:|about:|res:|x-raw-image:)/i;

/** How large a single picture may be, as hex — about 6MB of image. */
const MAX_HEX_LENGTH = 12 * 1024 * 1024;

function isUnresolved(img: HTMLImageElement): boolean {
  const src = (img.getAttribute("src") ?? "").trim();

  return src === "" || UNREACHABLE.test(src);
}

/**
 * The `{\pict …}` groups in an RTF, in the order it states them.
 *
 * Brace-matched rather than split on, because a picture group holds groups of
 * its own — `{\*\blipuid …}` — and Word wraps most of them in another
 * (`{\*\shppict{\pict …}}`).
 */
function pictureGroups(rtf: string): string[] {
  const groups: string[] = [];
  const OPEN = "{\\pict";

  for (let start = rtf.indexOf(OPEN); start !== -1; start = rtf.indexOf(OPEN, start + 1)) {
    let depth = 0;
    let end = start;

    for (; end < rtf.length; end++) {
      const character = rtf[end];
      // An escaped brace is text, and a control word's letters are never braces.
      if (character === "\\") {
        end++;
        continue;
      }
      if (character === "{") depth++;
      else if (character === "}") {
        depth--;
        if (depth === 0) break;
      }
    }

    if (depth === 0) groups.push(rtf.slice(start + OPEN.length, end));
  }

  return groups;
}

/**
 * What a picture group is, or null for the formats a report body cannot hold.
 *
 * Word writes a metafile copy of every picture next to the real one
 * (`{\nonshppict{\pict\wmetafile8 …}}`) for readers that cannot do PNG. Leaving
 * those out is also what keeps the count honest: one picture, one group.
 */
function pictureMimeType(group: string): string | null {
  const [beforeData] = group.split("bliptag");

  if (/\\pngblip/.test(beforeData)) return "image/png";
  if (/\\jpegblip/.test(beforeData)) return "image/jpeg";

  return null;
}

/**
 * The picture's bytes, as the hex Word writes them.
 *
 * Everything before the data is control words and metadata groups; the data
 * itself is the run of hex digits the group ends with.
 */
function pictureHex(group: string): string | null {
  const withoutMetadata = group.replace(/\{[^{}]*\}/g, " ");
  const match = withoutMetadata.match(/(?:^|[\s;])([0-9a-f\s]{32,})$/i);
  if (!match) return null;

  const hex = match[1].replace(/\s+/g, "");

  return hex.length % 2 === 0 && hex.length <= MAX_HEX_LENGTH ? hex : null;
}

/** Hex bytes as a data URL the report body can carry. */
function toDataUri(mimeType: string, hex: string): string | null {
  let binary = "";
  for (let i = 0; i < hex.length; i += 2) {
    binary += String.fromCharCode(Number.parseInt(hex.slice(i, i + 2), 16));
  }

  try {
    return `data:${mimeType};base64,${btoa(binary)}`;
  } catch {
    return null;
  }
}

/** Every picture the RTF carries, in order, as data URLs. */
export function rtfPictures(rtf: string): string[] {
  if (!rtf || !rtf.includes("{\\pict")) return [];

  const pictures: string[] = [];
  for (const group of pictureGroups(rtf)) {
    const mimeType = pictureMimeType(group);
    const hex = mimeType ? pictureHex(group) : null;
    const uri = mimeType && hex ? toDataUri(mimeType, hex) : null;
    if (uri) pictures.push(uri);
  }

  return pictures;
}

/**
 * The same HTML with the pictures the clipboard is carrying put where it only
 * named them — or unchanged, whenever the two sides do not line up exactly.
 */
export function restoreRtfPictures(html: string, rtf: string): string {
  if (!html || !rtf || !/<img/i.test(html)) return html;

  const doc = new DOMParser().parseFromString(html, "text/html");
  const waiting = Array.from(doc.querySelectorAll("img")).filter(isUnresolved);
  if (waiting.length === 0) return html;

  const pictures = rtfPictures(rtf);
  if (pictures.length !== waiting.length) return html;

  waiting.forEach((img, index) => {
    img.setAttribute("src", pictures[index]);
  });

  return doc.documentElement.outerHTML;
}

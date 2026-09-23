/**
 * The pictures Word ships in the clipboard's RTF.
 *
 * A QR code in a report header pasted as an empty frame while LibreOffice,
 * reading the same clipboard, showed it — so the bytes were plainly there. They
 * were: Plate matches an `<img>` to its bytes through `v:shapes` → `o:spid` →
 * `shplid`, which is how Word writes a FLOATING picture, and a picture inside a
 * table cell is an INLINE one. Word gives it none of those — a plain
 * `<img id="Picture_x0020_1" src="file:///…">` and a bare `{\pict}` group — so
 * nothing tied the two together and the frame came through empty.
 */
import { describe, expect, it } from "vitest";
import { createPlateEditor } from "platejs/react";
import { buildPlugins } from "@/plugins";
import { plateValueToHtml } from "@/lib/html-serializer";
import { restoreRtfPictures, rtfPictures } from "@/lib/rtf-pictures";

function clipboard(html: string, rtf = "") {
  return {
    types: rtf ? ["text/html", "text/rtf"] : ["text/html", "text/plain"],
    getData: (type: string) => (type === "text/html" ? html : type === "text/rtf" ? rtf : ""),
    files: [],
    items: [],
  } as unknown as DataTransfer;
}

function paste(html: string, rtf = "") {
  const editor = createPlateEditor({ plugins: buildPlugins("clean") });
  editor.tf.setValue([{ type: "p", children: [{ text: "" }] }] as never);
  editor.tf.select({ anchor: { path: [0, 0], offset: 0 }, focus: { path: [0, 0], offset: 0 } });
  editor.tf.insertData(clipboard(html, rtf));
  return plateValueToHtml(editor.children as never);
}

/** A 4x4 PNG's first bytes, in the hex Word writes. */
const PNG_HEX =
  "89504e470d0a1a0a0000000d494844520000000400000004080600000078d5b26f0000001d494441";

/**
 * What Word puts on the clipboard for an INLINE picture: the real one wrapped
 * in `\*\shppict`, and a metafile copy beside it for readers without PNG.
 */
const inlineRtf = (hex = PNG_HEX) =>
  `{\\rtf1\\ansi\\ansicpg1252\\deff0\\trowd\\cellx4000\\cellx8000 MRN\\cell\n` +
  `{\\*\\shppict{\\pict\\pngblip\\picw96\\pich96\\picwgoal1440\\pichgoal1440\\bliptag-1234{\\*\\blipuid 8e4f}\n` +
  `${hex}\n}}{\\nonshppict{\\pict\\wmetafile8\\picw96\\pich96\n0100090000034e00\n}}\\cell\\row}`;

/** …and the HTML beside it: no `v:shapes`, no VML comment, just a path. */
const HEADER =
  `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word">` +
  `<body><table border=1 style='border-collapse:collapse'><tr>` +
  `<td style='padding:2px'><p class=MsoNormal>MRN: v2609256937</p></td>` +
  `<td style='padding:2px'><p class=MsoNormal>` +
  `<img border=0 width=96 height=96 id="Picture_x0020_1" src="file:///C:/Temp/clip_image001.png">` +
  `</p></td></tr></table></body></html>`;

describe("pasting a header table whose QR is inline", () => {
  it("puts the picture where the HTML only named it", () => {
    const html = paste(HEADER, inlineRtf());

    expect(html).toContain("data:image/png;base64,iVBORw0KGgo");
    expect(html).not.toContain("file:///");
    // …and the header itself is untouched.
    expect(html).toContain("MRN: v2609256937");
    expect(html).toContain("<table");
  });

  it("leaves the picture alone when the clipboard has no RTF at all", () => {
    const html = paste(HEADER);

    expect(html).toContain("file:///C:/Temp/clip_image001.png");
  });
});

describe("reading the RTF", () => {
  it("takes the real picture and not the metafile copy beside it", () => {
    const pictures = rtfPictures(inlineRtf());

    expect(pictures).toHaveLength(1);
    expect(pictures[0].startsWith("data:image/png;base64,")).toBe(true);
  });

  it("keeps two pictures in the order the document states them", () => {
    const second = "89504e470d0a1a0a0000000d4948445200000008000000080806000000c4" + "0f be";
    const rtf =
      `{\\rtf1{\\*\\shppict{\\pict\\pngblip\\bliptag-1{\\*\\blipuid a}\n${PNG_HEX}\n}}` +
      `{\\*\\shppict{\\pict\\jpegblip\\bliptag-2{\\*\\blipuid b}\n${second.replace(/\s/g, "")}\n}}}`;

    const pictures = rtfPictures(rtf);

    expect(pictures).toHaveLength(2);
    expect(pictures[0]).toContain("image/png");
    expect(pictures[1]).toContain("image/jpeg");
  });

  it("does nothing when the two sides do not line up exactly", () => {
    // Two frames waiting, one picture on the clipboard: a picture in the wrong
    // cell is worse than a picture that did not arrive.
    const twoImages =
      `<p><img src="file:///C:/Temp/a.png"></p><p><img src="file:///C:/Temp/b.png"></p>`;

    expect(restoreRtfPictures(twoImages, inlineRtf())).toBe(twoImages);
  });

  it("leaves a picture that already has its bytes alone", () => {
    // Plate resolved the floating ones before this ran; they are not waiting.
    const resolved = `<p><img src="data:image/png;base64,iVBORw0KGgo="></p>`;

    expect(restoreRtfPictures(resolved, inlineRtf())).toBe(resolved);
  });

  it("returns HTML with no picture in it untouched, byte for byte", () => {
    const html = `<p>MRN: v2609256937</p>`;

    expect(restoreRtfPictures(html, inlineRtf())).toBe(html);
  });
});

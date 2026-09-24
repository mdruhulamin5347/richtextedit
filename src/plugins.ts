import { AlignKit } from "@/components/align-kit";
import { BasicBlocksKit } from "@/components/basic-blocks-kit";
import { BasicMarksKit } from "@/components/basic-marks-kit";
import { BlockBackgroundPlugin } from "@/components/block-background-plugin";
import { BlockFontSizePlugin } from "@/components/block-font-size-plugin";
import { BlockSpacingPlugin } from "@/components/block-spacing-plugin";
import { ClipboardPicturePlugin } from "@/components/clipboard-picture-plugin";
import { DocxKit } from "@/components/docx-kit";
import { FontFacePlugin, RtfParagraphFontPlugin } from "@/components/font-face-plugin";
import { FontKit } from "@/components/font-kit";
import { ImageSizePlugin } from "@/components/image-size-plugin";
import { IndentKit } from "@/components/indent-kit";
import { InheritedColorPlugin } from "@/components/inherited-color-plugin";
import { InheritedFontSizePlugin } from "@/components/inherited-font-size-plugin";
import { LegacyAlignmentPlugin } from "@/components/legacy-alignment-plugin";
import { LineHeightKit } from "@/components/line-height-kit";
import { LinkKit } from "@/components/link-kit";
import { ListKit } from "@/components/list-classic-kit";
import { MediaKit } from "@/components/media-kit";
import { PasteNormalizationPlugin } from "@/components/paste-normalization-plugin";
import { RtfPicturePlugin } from "@/components/rtf-picture-plugin";
import { TabAtCursorPlugin } from "@/components/tab-at-cursor-plugin";
import { TableColSizesPlugin } from "@/components/table-colsizes-plugin";
import { TableKit } from "@/components/table-kit";
import { WhitespacePlugin } from "@/components/whitespace-plugin";
import { WordBorderBoxPlugin } from "@/components/word-border-box-plugin";
import { WordLineGapPlugin } from "@/components/word-line-gap-plugin";
import { WordTabPlugin } from "@/components/word-tab-plugin";
import { WordTextboxPlugin } from "@/components/word-textbox-plugin";

/**
 * How pasted (and imported) content is treated.
 *
 * "clean" applies the PasteNormalizationPlugin: Word's margins, padding and
 * letter/word-spacing — none of which the saved report can carry — are dropped,
 * while its font sizes are kept in the POINTS it states them in (see
 * lib/font-size.ts) and its proportional line spacing is converted to the CSS
 * ratio that draws the same line (see lib/word-line-gap.ts). So the document's
 * own sizes and spacing survive, and on the document's own scale.
 *
 * "faithful" leaves every value exactly as the source stated it, units and all.
 */
export type PasteMode = "clean" | "faithful";

/**
 * The editor's plugin set. Mirrors the upstream editor's `EditorPlugins` minus the
 * pieces that only make sense there: Yjs collaboration, dictation, spellcheck,
 * short-form autoformat and the selection-actions toolbar (all of which call
 * upstream-only APIs).
 *
 * Kept in its own module so the round-trip tests can build the exact same set
 * without rendering React.
 */
export function buildPlugins(pasteMode: PasteMode = "clean") {
  return [
    // FIRST in the list, which makes it LAST to run.
    //
    // `transformData` is piped in REVERSE registration order (getInjectedPlugins
    // reverses the plugin list), so a pass registered late sees the RAW
    // clipboard and one registered early sees what everything else has already
    // done. This one has to be last of all: a Word document states its body
    // size in a `<style>` block, and that is not an inline style — the only
    // thing a Plate node can be read from — until JuicePlugin has run.
    // Measured, not assumed: registered next to the other passes it saw the
    // `<style>` block still sitting there, untouched.
    InheritedFontSizePlugin,
    // Beside it, and last for the same reason: Word and LibreOffice state a
    // paragraph's proportional line spacing in that same `<style>` block, and
    // the font it is a multiple OF along with it.
    WordLineGapPlugin,
    // Beside it, and last for the same reason: a Word document states a
    // paragraph's colour in that same `<style>` block.
    InheritedColorPlugin,
    // Also last, and for its own reason: Plate's docx cleaner resolves a
    // FLOATING picture from the clipboard's RTF and cannot resolve an inline
    // one, which is what a picture in a table cell is. Whatever is still
    // pointing at a `file:///` path once it has run is one of those, and its
    // bytes are in the RTF too. See lib/rtf-pictures.ts.
    RtfPicturePlugin,
    // Early in the list too, so it runs AFTER DocxKit: a Word tab span the docx
    // cleaner did not turn into tabs becomes one here, instead of a fixed run of
    // `&nbsp;`. One Word tab, one editor tab. See lib/word-tabs.ts.
    WordTabPlugin,
    // Right after it, so it runs just BEFORE it: a LibreOffice line that names no
    // font takes its RTF's, and the tab layout has to measure the text in the
    // font it is drawn in. After DocxKit, so a style block's family is already
    // inline and wins. See lib/font-face.ts.
    RtfParagraphFontPlugin,
    ...BasicBlocksKit,
    ...BasicMarksKit,
    ...ListKit,
    ...MediaKit,
    // When the clipboard's HTML only NAMES a picture but the clipboard itself
    // carries its bytes, use the bytes. Narrowly: see lib/clipboard-picture.ts.
    ClipboardPicturePlugin,
    // The width a picture was shown at, which Plate's deserializer reads off
    // neither the attribute nor the style. See components/image-size-plugin.ts.
    ImageSizePlugin,
    ...TableKit,
    // An addition beyond the upstream editor, kept for a specific reason: Plate only
    // commits a column resize back to the document when the table already has
    // `colSizes`. Without this, dragging a border resizes the column on screen
    // and the new width is lost on save. Verified both ways in
    // tests/browser/resize.mjs.
    TableColSizesPlugin,
    // And the space the document keeps above and below a paragraph INSIDE a
    // cell, which is most of a table row's height and which Plate's own
    // deserializer drops. See components/block-spacing-plugin.ts.
    BlockSpacingPlugin,
    // …and the shading behind it, which the same deserializer drops for the
    // same reason. See components/block-background-plugin.ts.
    BlockBackgroundPlugin,
    ...LinkKit,
    ...IndentKit,
    // Registered after IndentKit so its `tab` override wraps the indent one and
    // gets to decide first: a Tab goes in at the caret instead of moving the
    // whole block. See components/tab-at-cursor-plugin.ts.
    TabAtCursorPlugin,
    ...FontKit,
    ...AlignKit,
    ...LineHeightKit,
    // DocxKit (DocxPlugin + JuicePlugin) is the reason for this port: it parses
    // Word's HTML properly and inlines its CSS instead of discarding it.
    ...DocxKit,
    // Registered after DocxKit, which by the rule above means it runs BEFORE
    // it — on the raw clipboard, `<style>` block and all. That is fine for
    // alignment, which is stated in attributes, and it is why the rewrite has
    // to hand back the whole document rather than just the body.
    LegacyAlignmentPlugin,
    // Beside it, and before the docx cleaner for the same reason: LibreOffice
    // states a run's font as `<font face>`, which that cleaner turns into a
    // `<span>` whose `face` nothing reads. See lib/font-face.ts.
    FontFacePlugin,
    // Last of the clipboard passes: the alignment rewrite works on elements, so
    // it neither reads nor moves the text nodes this one wraps.
    WhitespacePlugin,
    // LAST in the list and so FIRST to run, on the rawest clipboard there is: a
    // Word text box lives inside a conditional COMMENT, which both the docx
    // cleaner and Juice rewrite the document around. Recovered here, it reaches
    // every pass above as ordinary HTML. See lib/word-textbox.ts.
    WordTextboxPlugin,
    // And the other way a document draws a box round a heading: Borders &
    // Shading on the paragraph itself. See lib/word-border-box.ts.
    WordBorderBoxPlugin,
    ...(pasteMode === "clean" ? [PasteNormalizationPlugin] : []),
    // Sizes every paragraph from the text in it — and a blank one from the
    // lines around it — in both paste modes. How tall a line is is not a value
    // any document states, it is one this editor has to work out.
    // See lib/block-font-size.ts.
    BlockFontSizePlugin,
  ];
}

import { createSlatePlugin, KEYS } from "platejs";

import { inlineLibreOfficeCellLineGap, inlineWordLineGap, isLibreOfficeClipboard } from "@/lib/word-line-gap";

/**
 * Correct Word's and LibreOffice's proportional line spacing on the way in.
 *
 * Registered at the FRONT of the plugin list, which makes it LAST to run:
 * `transformData` is piped in reverse registration order, and this pass has to
 * see the HTML after Juice, because a Word document states both the spacing
 * (`p.MsoNormal {line-height:115%}`) and the font it is measured against in a
 * `<style>` block rather than inline. The same reason InheritedFontSizePlugin
 * sits beside it. See lib/word-line-gap.ts.
 *
 * A LibreOffice table's gap is only in the clipboard's RTF, and the check for
 * LibreOffice reads the RAW clipboard: by now the docx cleaner — which that RTF
 * makes run — has removed the generator `<meta>` it names itself in.
 */
export const WordLineGapPlugin = createSlatePlugin({
  key: "wordLineGap",
  inject: {
    plugins: {
      [KEYS.html]: {
        parser: {
          transformData: ({ data, dataTransfer }: { data: string; dataTransfer: DataTransfer }) => {
            const libreOffice = isLibreOfficeClipboard(dataTransfer.getData("text/html"));
            return inlineWordLineGap(
              libreOffice ? inlineLibreOfficeCellLineGap(data, dataTransfer.getData("text/rtf")) : data,
              { libreOffice }
            );
          },
        },
      },
    },
  },
});

import { ToolbarButton } from "@/components/ui/toolbar";
import { EMPTY_VALUE } from "@/lib/html-serializer";
import { inlineInheritedColor } from "@/lib/inherited-color";
import { inlineInheritedFontSize } from "@/lib/inherited-font-size";
import { inlineLegacyAlignment } from "@/lib/legacy-alignment";
import { inlineWordBorderBoxes } from "@/lib/word-border-box";
import { protectWhitespace } from "@/lib/whitespace";
import { deserializeHtml, type Value } from "platejs";
import { useEditorRef } from "platejs/react";
import { FileUpIcon, Loader2Icon } from "lucide-react";
import React, { useCallback, useRef, useState } from "react";

/**
 * "Upload Word (.doc/.docx)" — the same capability the Summernote build had via
 * `_partials/docx_report_body_tools.blade.php`, reimplemented against Plate.
 *
 * BOTH extensions go to the server route (LibreOffice) first, and that is a
 * deliberate change from the Summernote build, which converted a .docx in the
 * browser with Mammoth. Mammoth reads a document for its STRUCTURE and throws
 * away everything a report is written in: the same .docx that gives
 *
 *   LibreOffice: <p align=center><font color="#ff0000"><font size=4><b>…
 *   Mammoth:     <p><strong>…
 *
 * — no colour, no size, no alignment. Nothing downstream can put back what
 * never arrived, so a red heading imported black however carefully the rest of
 * this editor preserves colour. LibreOffice states all three, in spellings the
 * paste pipeline already reads (`lib/inherited-color.ts`,
 * `lib/inherited-font-size.ts`, `lib/legacy-alignment.ts`).
 *
 * Mammoth stays as the FALLBACK for a .docx: it needs no server, so an import
 * still works where the route is not configured or LibreOffice is not installed
 * — with the document's formatting lost, which is what it was doing all along.
 *
 * Like the Summernote version, a successful import REPLACES the document rather
 * than appending to it.
 */
export function WordImportToolbarButton({
  docConvertUrl,
  csrfToken,
}: {
  docConvertUrl?: string;
  csrfToken?: string;
}) {
  const editor = useEditorRef();
  const inputRef = useRef<HTMLInputElement>(null);
  const [status, setStatus] = useState<{ kind: "idle" | "busy" | "ok" | "error"; text: string }>({
    kind: "idle",
    text: "",
  });

  /**
   * Mammoth emits borderless tables; a lab report is a bordered grid. Same
   * treatment the Summernote partial applied.
   */
  const borderizeTables = useCallback((html: string): string => {
    const wrap = document.createElement("div");
    wrap.innerHTML = html;
    wrap.querySelectorAll("table").forEach((t) => {
      t.setAttribute("border", "1");
      t.style.borderCollapse = "collapse";
      if (!t.style.width) t.style.width = "100%";
    });
    wrap.querySelectorAll("td, th").forEach((c) => {
      const cell = c as HTMLElement;
      cell.style.border = "1px solid #000";
      if (!cell.style.padding) cell.style.padding = "2px 5px";
    });
    return wrap.innerHTML;
  }, []);

  const loadHtml = useCallback(
    (html: string, fileName: string, note?: string) => {
      if (!html || !html.trim()) {
        setStatus({ kind: "error", text: "The file appears to be empty." });
        return;
      }
      // The clipboard path gets this from LegacyAlignmentPlugin; an uploaded
      // file never goes through `insertData`, and LibreOffice — which converts
      // the legacy .doc here — states alignment the pre-CSS way constantly.
      // ...and `protectWhitespace` for the same reason: the converters write a
      // tab as a run of `&nbsp;`, which Plate's deserializer would collapse.
      // `inlineInheritedFontSize` likewise: LibreOffice states a paragraph's
      // size on the paragraph, which is a size no Plate node can carry.
      // `inlineInheritedColor` for the same reason twice over — LibreOffice
      // states colour on the paragraph AND in the pre-CSS `<font color>` that
      // Plate reads off no attribute at all.
      // `inlineWordBorderBoxes` because LibreOffice states a boxed heading as a
      // border on the paragraph, which no Plate node can carry.
      const prepared = protectWhitespace(
        inlineInheritedColor(
          inlineInheritedFontSize(
            borderizeTables(inlineWordBorderBoxes(inlineLegacyAlignment(html)))
          )
        )
      );
      const value = deserializeHtml(editor, { element: prepared }) as Value;
      editor.tf.setValue(value.length ? value : EMPTY_VALUE);
      setStatus({ kind: "ok", text: `Loaded: ${fileName}${note ? ` — ${note}` : ""}` });
    },
    [borderizeTables, editor]
  );

  /**
   * LibreOffice, on the server — the converter that keeps the document's own
   * colour, size and alignment. `null` when there is no route to ask.
   */
  const convertOnServer = useCallback(
    async (file: File): Promise<{ html: string } | { failed: string } | null> => {
      if (!docConvertUrl) return null;

      const fd = new FormData();
      fd.append("file", file);
      const res = await fetch(docConvertUrl, {
        method: "POST",
        headers: csrfToken ? { "X-CSRF-TOKEN": csrfToken } : {},
        body: fd,
      }).catch(() => null);

      const body = res ? await res.json().catch(() => null) : null;
      if (res?.ok && body?.html) return { html: body.html as string };

      return { failed: body?.message || "Could not convert this Word file." };
    },
    [csrfToken, docConvertUrl]
  );

  /**
   * Mammoth, in the browser. Structure only — every colour, size and alignment
   * the document states is dropped in the conversion, so this is the fallback
   * and not the path a report should normally take.
   */
  const convertInBrowser = useCallback(async (file: File): Promise<string> => {
    const mammoth = await import("mammoth");
    const arrayBuffer = await file.arrayBuffer();
    const result = await mammoth.convertToHtml(
      { arrayBuffer },
      { styleMap: ["b => strong", "i => em", "u => u"] }
    );
    return result.value || "";
  }, []);

  const handleFile = useCallback(
    async (file: File) => {
      const isDocx = /\.docx$/i.test(file.name);
      if (!isDocx && !/\.doc$/i.test(file.name)) {
        setStatus({ kind: "error", text: "Please upload a .doc or .docx file." });
        return;
      }

      setStatus({ kind: "busy", text: "Converting…" });
      try {
        const converted = await convertOnServer(file);
        if (converted && "html" in converted) {
          loadHtml(converted.html, file.name);
          return;
        }

        // A .doc is binary: the browser cannot read one at all, so there is
        // nothing to fall back TO.
        if (!isDocx) {
          setStatus({
            kind: "error",
            text:
              converted && "failed" in converted
                ? converted.failed
                : "Legacy .doc conversion is not configured.",
          });
          return;
        }

        // Said out loud rather than silently: an import that lands here is
        // missing the document's colours and sizes, and the reader of the
        // report is the last person who should have to work that out.
        loadHtml(
          await convertInBrowser(file),
          file.name,
          "formatting reduced, the server converter was unavailable"
        );
      } catch (error) {
        console.error(error);
        setStatus({ kind: "error", text: "Could not read this Word file." });
      }
    },
    [convertInBrowser, convertOnServer, loadHtml]
  );

  return (
    <div className="flex items-center gap-2">
      <ToolbarButton
        tooltip="Upload a Word (.doc / .docx) file and load it into the editor"
        onClick={() => inputRef.current?.click()}
      >
        {status.kind === "busy" ? (
          <Loader2Icon className="size-3 animate-spin" />
        ) : (
          <FileUpIcon className="size-3" />
        )}
      </ToolbarButton>
      <input
        ref={inputRef}
        type="file"
        accept=".doc,.docx"
        className="hidden"
        onChange={(e) => {
          const file = e.target.files?.[0];
          e.target.value = "";
          if (file) void handleFile(file);
        }}
      />
      {status.text ? (
        <small
          className={
            status.kind === "error"
              ? "text-destructive text-xs"
              : status.kind === "ok"
                ? "text-xs text-green-600"
                : "text-muted-foreground text-xs"
          }
        >
          {status.text}
        </small>
      ) : null}
    </div>
  );
}

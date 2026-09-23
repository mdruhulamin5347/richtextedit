/**
 * Package entry for bundlers (React apps).
 *
 * Exposes both surfaces: the <RichTextEdit> component for apps that render
 * React themselves, and the mount API for pages that only own a <div> and a
 * <textarea>. Styles ship separately as `rich-text-edit/style.css`.
 */
export {
  RichTextEdit,
  type PasteMode,
  type RichTextEditHandle,
  type RichTextEditProps,
} from "@/RichTextEdit";
export {
  default as RichTextEditApi,
  type MountOptions,
  type StorageFormat,
} from "@/entry";
export {
  EMPTY_VALUE,
  isPlateJson,
  isPlateValueEmpty,
  plateValueToHtml,
  plateValueToJson,
} from "@/lib/html-serializer";
export type { Value } from "platejs";

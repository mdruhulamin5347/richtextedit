/**
 * Package entry for bundlers (React apps).
 *
 * Exposes both surfaces: the <RadiolensEditor> component for apps that render
 * React themselves, and the mount API for pages that only own a <div> and a
 * <textarea>. Styles ship separately as `@radiolens/editor/style.css`.
 */
export {
  RadiolensEditor,
  type PasteMode,
  type RadiolensEditorHandle,
  type RadiolensEditorProps,
} from "@/RadiolensEditor";
export {
  default as RadiolensEditorApi,
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

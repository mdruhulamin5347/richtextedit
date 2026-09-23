import type { PlateEditor } from "platejs/react";

import { triggerFloatingLink } from "@platejs/link/react";
import { insertMedia } from "@platejs/media";
import { TablePlugin } from "@platejs/table/react";
import { type NodeEntry, type Path, type TElement, KEYS, PathApi } from "platejs";

/**
 * Trimmed from the upstream editor's `transforms.ts`.
 *
 * The original mapped every block type Plate ships — callout, code block, date,
 * excalidraw, equations, table of contents, column layouts, video/audio/file
 * placeholders. None of those plugins are in this editor's plugin set, so the
 * menu entries were dead, yet the imports still pulled katex, excalidraw and the
 * media stack into the bundle.
 *
 * Lists are deliberately absent too: the original used Plate's MODERN list model
 * (`indent` + `listStyleType`), while this editor uses list-classic
 * (ul/ol/li/lic) so that lists serialize to real <ul>/<ol> for dompdf. The
 * dedicated list toolbar buttons handle them correctly.
 */
const insertBlockMap: Record<string, (editor: PlateEditor, type: string) => void> = {
  [KEYS.img]: (editor) => insertMedia(editor, { select: true, type: KEYS.img }),
  [KEYS.table]: (editor) => editor.getTransforms(TablePlugin).insert.table({}, { select: true }),
};

const insertInlineMap: Record<string, (editor: PlateEditor, type: string) => void> = {
  [KEYS.link]: (editor) => triggerFloatingLink(editor, { focused: true }),
};

type InsertBlockOptions = { upsert?: boolean };

export const insertBlock = (
  editor: PlateEditor,
  type: string,
  options: InsertBlockOptions = {}
) => {
  const { upsert = false } = options;

  editor.tf.withoutNormalizing(() => {
    const block = editor.api.block();
    if (!block) return;

    const [currentNode, path] = block;
    const isCurrentBlockEmpty = editor.api.isEmpty(currentNode);
    const isSameBlockType = type === getBlockType(currentNode);

    if (upsert && isCurrentBlockEmpty && isSameBlockType) return;

    if (type in insertBlockMap) {
      insertBlockMap[type](editor, type);
    } else {
      editor.tf.insertNodes(editor.api.create.block({ type }), {
        at: PathApi.next(path),
        select: true,
      });
    }

    if (!isSameBlockType) {
      editor.tf.removeNodes({ previousEmptyBlock: true });
    }
  });
};

export const insertInlineElement = (editor: PlateEditor, type: string) => {
  if (insertInlineMap[type]) {
    insertInlineMap[type](editor, type);
  }
};

export const setBlockType = (editor: PlateEditor, type: string, { at }: { at?: Path } = {}) => {
  editor.tf.withoutNormalizing(() => {
    const setEntry = (entry: NodeEntry<TElement>) => {
      const [node, path] = entry;
      if (node.type !== type) {
        editor.tf.setNodes({ type }, { at: path });
      }
    };

    if (at) {
      const entry = editor.api.node<TElement>(at);
      if (entry) {
        setEntry(entry);
        return;
      }
    }

    editor.api.blocks({ mode: "lowest" }).forEach(setEntry);
  });
};

export const getBlockType = (block: TElement) => block.type;

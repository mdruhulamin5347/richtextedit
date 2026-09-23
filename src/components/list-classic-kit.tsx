"use client";

import type { TListElement } from "platejs";

import {
  BulletedListPlugin,
  ListItemContentPlugin,
  ListItemPlugin,
  ListPlugin,
  NumberedListPlugin,
  TaskListPlugin,
} from "@platejs/list-classic/react";

import {
  BulletedListElement,
  ListItemElement,
  NumberedListElement,
  TaskListElement,
} from "@/components/ui/list-classic-node";

/** The CSS list-style-type behind a legacy `<ul type>` / `<ol type>` attribute. */
const HTML_LIST_TYPES: Record<string, string> = {
  "1": "decimal",
  a: "lower-alpha",
  A: "upper-alpha",
  circle: "circle",
  disc: "disc",
  i: "lower-roman",
  I: "upper-roman",
  none: "none",
  square: "square",
};

const NO_LIST_STYLE = new Set(["", "inherit", "initial", "unset", "revert"]);

function inlineListStyleType(element: HTMLElement | null): string | undefined {
  const value = element?.style.listStyleType?.trim().toLowerCase();
  return value && !NO_LIST_STYLE.has(value) ? value : undefined;
}

/**
 * The marker the pasted markup asks for, or undefined to leave the list to the
 * editor's own depth cascade.
 *
 * Where it lives depends on who produced the HTML: an inline `list-style-type`
 * on the list itself is the common case, Google Docs puts one on every `<li>`
 * instead, and older markup uses the `type` attribute.
 */
function authoredListStyleType(element: HTMLElement): string | undefined {
  const inline = inlineListStyleType(element);
  if (inline) return inline;

  // Case matters here -- `type="a"` is lower-alpha, `type="A"` is upper-alpha --
  // so the exact spelling is tried before the lower-cased one.
  const type = element.getAttribute("type")?.trim();
  const mapped = type ? (HTML_LIST_TYPES[type] ?? HTML_LIST_TYPES[type.toLowerCase()]) : undefined;
  if (mapped) return mapped;

  return inlineListStyleType(element.querySelector(":scope > li"));
}

/**
 * Keep the marker and start number a pasted list declared.
 *
 * Without this the markup is flattened to a bare `<ul>`/`<ol>` and the list
 * picks up whatever the editor's default cascade gives it, so nested squares
 * come back as circles. Everything downstream reads these same two props --
 * `ListElement` renders them, and the DOCX and HTML exports pin their markers
 * from them -- so capturing them here is what makes a paste round-trip.
 *
 * Supplying `parse` replaces the deserializer's default one, which is why it
 * has to return `type` itself.
 */
const authoredListParsers = {
  html: {
    deserializer: {
      parse: ({ element, type }: { element: HTMLElement; type: string }) => {
        const listStyleType = authoredListStyleType(element);
        const start = Number.parseInt(element.getAttribute("start") ?? "", 10);

        return {
          type,
          ...(listStyleType ? { listStyleType } : {}),
          ...(Number.isFinite(start) && start > 1 ? { listStart: start } : {}),
        } as Partial<TListElement>;
      },
    },
  },
};

export const ListKit = [
  ListPlugin,
  ListItemPlugin,
  ListItemContentPlugin,
  BulletedListPlugin.configure({
    node: { component: BulletedListElement },
    parsers: authoredListParsers,
    shortcuts: { toggle: { keys: "mod+alt+5" } },
  }),
  NumberedListPlugin.configure({
    node: { component: NumberedListElement },
    parsers: authoredListParsers,
    shortcuts: { toggle: { keys: "mod+alt+6" } },
  }),
  TaskListPlugin.configure({
    node: { component: TaskListElement },
    shortcuts: { toggle: { keys: "mod+alt+7" } },
  }),
  ListItemPlugin.withComponent(ListItemElement),
];

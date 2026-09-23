"use client";

import * as React from "react";

import type { DropdownMenuProps } from "@radix-ui/react-dropdown-menu";

import {
  Heading1Icon,
  Heading2Icon,
  Heading3Icon,
  ImageIcon,
  Link2Icon,
  MinusIcon,
  PilcrowIcon,
  PlusIcon,
  QuoteIcon,
  TableIcon,
} from "lucide-react";
import { KEYS } from "platejs";
import { type PlateEditor, useEditorRef } from "platejs/react";

import { insertBlock, insertInlineElement } from "@/components/transforms";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";

import { ToolbarButton, ToolbarMenuGroup } from "./toolbar";

type Group = {
  group: string;
  items: Item[];
};

type Item = {
  icon: React.ReactNode;
  value: string;
  onSelect: (editor: PlateEditor, value: string) => void;
  focusEditor?: boolean;
  label?: string;
};

// Trimmed to the blocks this editor actually has plugins for. Lists are
// deliberately absent: they live on their own toolbar buttons, which speak
// list-classic (real <ul>/<ol>) rather than Plate's modern indent-based lists.
const groups: Group[] = [
  {
    group: "Basic blocks",
    items: [
      { icon: <PilcrowIcon />, label: "Paragraph", value: KEYS.p },
      { icon: <Heading1Icon />, label: "Heading 1", value: "h1" },
      { icon: <Heading2Icon />, label: "Heading 2", value: "h2" },
      { icon: <Heading3Icon />, label: "Heading 3", value: "h3" },
      { icon: <TableIcon />, label: "Table", value: KEYS.table },
      { icon: <QuoteIcon />, label: "Quote", value: KEYS.blockquote },
      { icon: <MinusIcon />, label: "Divider", value: KEYS.hr },
      { icon: <ImageIcon />, label: "Image", value: KEYS.img },
    ].map((item) => ({
      ...item,
      onSelect: (editor, value) => {
        insertBlock(editor, value);
      },
    })),
  },
  {
    group: "Inline",
    items: [{ icon: <Link2Icon />, label: "Link", value: KEYS.link }].map((item) => ({
      ...item,
      onSelect: (editor, value) => {
        insertInlineElement(editor, value);
      },
    })),
  },
];

export function InsertToolbarButton(props: DropdownMenuProps) {
  const editor = useEditorRef();
  const [open, setOpen] = React.useState(false);

  return (
    <DropdownMenu open={open} onOpenChange={setOpen} modal={false} {...props}>
      <DropdownMenuTrigger asChild>
        <ToolbarButton pressed={open} tooltip="Insert" isDropdown>
          <PlusIcon className="size-3" />
        </ToolbarButton>
      </DropdownMenuTrigger>

      <DropdownMenuContent
        className="flex max-h-[500px] min-w-0 flex-col overflow-y-auto"
        align="start"
      >
        {groups.map(({ group, items: nestedItems }) => (
          <ToolbarMenuGroup key={group} label={group}>
            {nestedItems.map(({ icon, label, value, onSelect }) => (
              <DropdownMenuItem
                key={value}
                className="min-w-[100px]"
                onSelect={() => {
                  onSelect(editor, value);
                  editor.tf.focus();
                }}
              >
                {icon}
                {label}
              </DropdownMenuItem>
            ))}
          </ToolbarMenuGroup>
        ))}
      </DropdownMenuContent>
    </DropdownMenu>
  );
}

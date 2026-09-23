"use client";

import * as React from "react";

import type { TListElement } from "platejs";
import type { PlateElementProps } from "platejs/react";

import { useTodoListElement, useTodoListElementState } from "@platejs/list-classic/react";
import { type VariantProps, cva } from "class-variance-authority";
import { PlateElement } from "platejs/react";

import { Checkbox } from "@/components/ui/checkbox";
import { cn } from "@/lib/utils";

const LIST_LAYOUT = "m-0 py-1 ps-6";

/**
 * The marker cascade for a list that names none of its own -- the same one a
 * browser applies by default, spelled out so the DOCX export can mirror it
 * (see `_BULLET_LEVELS` in the backend's docx_generator).
 *
 * `[&_ul_ul]` is a descendant selector, so bullets stay square from the third
 * level down rather than cycling; `<ol>` has no nested override at all, so
 * ordered lists stay decimal at every level.
 */
const listVariants = cva(LIST_LAYOUT, {
  variants: {
    variant: {
      ol: "list-decimal",
      ul: "list-disc [&_ul]:list-[circle] [&_ul_ul]:list-[square]",
    },
  },
});

export function ListElement({
  variant,
  ...props
}: PlateElementProps & VariantProps<typeof listVariants>) {
  const { listStart, listStyleType } = props.element as TListElement;

  return (
    <PlateElement
      as={variant!}
      // A marker the content names outright wins over the cascade above. The
      // cascade has to be dropped rather than layered under it: list-style-type
      // inherits, so an unstyled nested list must keep this one instead of
      // being pulled back to the default for its depth.
      className={listStyleType ? LIST_LAYOUT : listVariants({ variant })}
      style={listStyleType ? { listStyleType } : undefined}
      {...props}
      attributes={{ ...props.attributes, start: listStart }}
    >
      {props.children}
    </PlateElement>
  );
}

export function BulletedListElement(props: PlateElementProps) {
  return <ListElement variant="ul" {...props} />;
}

export function NumberedListElement(props: PlateElementProps) {
  return <ListElement variant="ol" {...props} />;
}

export function TaskListElement(props: PlateElementProps) {
  return (
    <PlateElement as="ul" className="m-0 list-none! py-1 ps-6" {...props}>
      {props.children}
    </PlateElement>
  );
}

export function ListItemElement(props: PlateElementProps) {
  const isTaskList = "checked" in props.element;

  if (isTaskList) {
    return <TaskListItemElement {...props} />;
  }

  return <BaseListItemElement {...props} />;
}

export function BaseListItemElement(props: PlateElementProps) {
  return (
    <PlateElement as="li" {...props}>
      {props.children}
    </PlateElement>
  );
}

export function TaskListItemElement(props: PlateElementProps) {
  const { element } = props;
  const state = useTodoListElementState({ element });
  const { checkboxProps } = useTodoListElement(state);
  const [firstChild, ...otherChildren] = React.Children.toArray(props.children);

  return (
    <BaseListItemElement {...props}>
      <div
        className={cn("flex items-stretch *:nth-[2]:flex-1 *:nth-[2]:focus:outline-none", {
          "*:nth-[2]:text-muted-foreground *:nth-[2]:line-through": state.checked,
        })}
      >
        <div
          className="-ms-5 me-1.5 flex w-fit items-start justify-center pt-[0.275em] select-none"
          contentEditable={false}
        >
          <Checkbox {...checkboxProps} />
        </div>

        {firstChild}
      </div>

      {otherChildren}
    </BaseListItemElement>
  );
}

"use client";

import * as React from "react";

import * as ToolbarPrimitive from "@radix-ui/react-toolbar";
import * as TooltipPrimitive from "@radix-ui/react-tooltip";
import { type VariantProps, cva } from "class-variance-authority";
import { ChevronDown } from "lucide-react";

import {
  DropdownMenuLabel,
  DropdownMenuRadioGroup,
  DropdownMenuSeparator,
} from "@/components/ui/dropdown-menu";
import { Separator } from "@/components/ui/separator";
import { Tooltip, TooltipTrigger } from "@/components/ui/tooltip";
import { cn } from "@/lib/utils";

export function Toolbar({
  className,
  ...props
}: React.ComponentProps<typeof ToolbarPrimitive.Root>) {
  return (
    <ToolbarPrimitive.Root
      className={cn("relative flex items-center select-none", className)}
      {...props}
    />
  );
}

export function ToolbarToggleGroup({
  className,
  ...props
}: React.ComponentProps<typeof ToolbarPrimitive.ToolbarToggleGroup>) {
  return (
    <ToolbarPrimitive.ToolbarToggleGroup
      className={cn("flex items-center", className)}
      {...props}
    />
  );
}

export function ToolbarLink({
  className,
  ...props
}: React.ComponentProps<typeof ToolbarPrimitive.Link>) {
  return (
    <ToolbarPrimitive.Link
      className={cn("font-medium underline underline-offset-4", className)}
      {...props}
    />
  );
}

export function ToolbarSeparator({
  className,
  ...props
}: React.ComponentProps<typeof ToolbarPrimitive.Separator>) {
  return (
    <ToolbarPrimitive.Separator
      className={cn("bg-border mx-2 my-1 w-px shrink-0", className)}
      {...props}
    />
  );
}

// From toggleVariants
const toolbarButtonVariants = cva(
  "inline-flex cursor-pointer items-center justify-center gap-1 whitespace-nowrap rounded-md font-medium text-sm outline-none transition-[color,box-shadow] hover:bg-muted hover:text-muted-foreground focus-visible:border-ring focus-visible:ring-[3px] focus-visible:ring-ring/50 disabled:pointer-events-none disabled:opacity-50 aria-checked:bg-accent aria-checked:text-accent-foreground aria-invalid:border-destructive aria-invalid:ring-destructive/20 dark:aria-invalid:ring-destructive/40 [&_svg:not([class*='size-'])]:size-4 [&_svg]:pointer-events-none [&_svg]:shrink-0",
  {
    defaultVariants: {
      size: "default",
      variant: "default",
    },
    variants: {
      size: {
        default: "h-9 min-w-9 px-1",
        lg: "h-10 min-w-10 px-1.5",
        sm: "h-5 min-w-5 px-0.5",
      },
      variant: {
        default: "bg-transparent",
        outline:
          "border border-input bg-transparent shadow-xs hover:bg-accent hover:text-accent-foreground",
      },
    },
  }
);

const dropdownArrowVariants = cva(
  cn(
    "inline-flex items-center justify-center rounded-r-md font-medium text-foreground text-sm transition-colors disabled:pointer-events-none disabled:opacity-50"
  ),
  {
    defaultVariants: {
      size: "sm",
      variant: "default",
    },
    variants: {
      size: {
        default: "h-9 w-6",
        lg: "h-10 w-8",
        sm: "h-8 w-4",
      },
      variant: {
        default:
          "bg-transparent hover:bg-muted hover:text-muted-foreground aria-checked:bg-accent aria-checked:text-accent-foreground",
        outline:
          "border border-input border-l-0 bg-transparent hover:bg-accent hover:text-accent-foreground",
      },
    },
  }
);

type ToolbarButtonProps = {
  isDropdown?: boolean;
  pressed?: boolean;
} & Omit<React.ComponentPropsWithoutRef<typeof ToolbarToggleItem>, "asChild" | "value"> &
  VariantProps<typeof toolbarButtonVariants>;

/**
 * Deviation from RadioLens, and a necessary one here.
 *
 * Radix attaches a ref through `asChild`: DropdownMenuTrigger uses it to
 * register the button as the popper's ANCHOR. A plain function component cannot
 * receive a ref under React 18 (that only became an ordinary prop in React 19),
 * so the anchor stayed null, Floating UI never computed a position, and every
 * dropdown mounted stuck at its unpositioned `translate(0, -200%)` — far
 * off-screen, which reads as "the toolbar buttons do nothing".
 */
const ToolbarButtonImpl = React.forwardRef<HTMLButtonElement, ToolbarButtonProps>(
  function ToolbarButton(
    { children, className, isDropdown, pressed, size = "sm", variant, ...props },
    ref
  ) {
  return typeof pressed === "boolean" ? (
    // Deviation from RadioLens, and the reason every toolbar dropdown was dead.
    //
    // This branch used to be a ToolbarToggleGroup wrapping a ToolbarToggleItem.
    // Radix's ToggleGroup does not pass a forwarded ref through to the button,
    // so when DropdownMenuTrigger renders this button via `asChild` its popper
    // ANCHOR ref was swallowed. Floating UI then had no reference element,
    // never computed a position, and the menu mounted at its unpositioned
    // `translate(0, -200%)` — far off-screen, which looked exactly like a dead
    // button. Every dropdown trigger here passes `pressed`, so every one of
    // them was affected.
    //
    // A plain toolbar button carrying the same `data-state` renders and styles
    // identically (the variants key off `data-[state=on]`) and forwards the ref.
    <ToolbarPrimitive.Button
      ref={ref}
      data-state={pressed ? "on" : "off"}
      className={cn(
        toolbarButtonVariants({
          size,
          variant,
        }),
        isDropdown && "justify-between gap-1 pr-1",
        className
      )}
      {...props}
    >
        {isDropdown ? (
          <>
            <div className="flex flex-1 items-center gap-1 whitespace-nowrap">{children}</div>
            <div>
              <ChevronDown className="text-muted-foreground size-3.5" data-icon />
            </div>
          </>
        ) : (
          children
        )}
    </ToolbarPrimitive.Button>
  ) : (
    <ToolbarPrimitive.Button
      ref={ref}
      className={cn(
        toolbarButtonVariants({
          size,
          variant,
        }),
        isDropdown && "pr-1",
        className
      )}
      {...props}
    >
      {children}
    </ToolbarPrimitive.Button>
  );
  }
);

export const ToolbarButton = withTooltip(ToolbarButtonImpl);

export function ToolbarSplitButton({
  className,
  ...props
}: React.ComponentPropsWithoutRef<typeof ToolbarButton>) {
  return (
    <ToolbarButton
      className={cn("group flex gap-0 px-0 hover:bg-transparent", className)}
      {...props}
    />
  );
}

type ToolbarSplitButtonPrimaryProps = Omit<
  React.ComponentPropsWithoutRef<typeof ToolbarToggleItem>,
  "value"
> &
  VariantProps<typeof toolbarButtonVariants>;

export function ToolbarSplitButtonPrimary({
  children,
  className,
  size = "sm",
  variant,
  ...props
}: ToolbarSplitButtonPrimaryProps) {
  return (
    <span
      className={cn(
        toolbarButtonVariants({
          size,
          variant,
        }),
        "rounded-r-none",
        "group-data-[pressed=true]:bg-accent group-data-[pressed=true]:text-accent-foreground",
        className
      )}
      {...props}
    >
      {children}
    </span>
  );
}

export function ToolbarSplitButtonSecondary({
  className,
  size,
  variant,
  ...props
}: React.ComponentPropsWithoutRef<"span"> & VariantProps<typeof dropdownArrowVariants>) {
  return (
    <span
      className={cn(
        dropdownArrowVariants({
          size,
          variant,
        }),
        "group-data-[pressed=true]:bg-accent group-data-[pressed=true]:text-accent-foreground",
        className
      )}
      onClick={(e) => e.stopPropagation()}
      role="button"
      {...props}
    >
      <ChevronDown className="text-muted-foreground size-3.5" data-icon />
    </span>
  );
}

export function ToolbarToggleItem({
  className,
  size = "sm",
  variant,
  ...props
}: React.ComponentProps<typeof ToolbarPrimitive.ToggleItem> &
  VariantProps<typeof toolbarButtonVariants>) {
  return (
    <ToolbarPrimitive.ToggleItem
      className={cn(toolbarButtonVariants({ size, variant }), className)}
      {...props}
    />
  );
}

export function ToolbarGroup({ children, className }: React.ComponentProps<"div">) {
  return (
    <div className={cn("group/toolbar-group", "relative hidden has-[button]:flex", className)}>
      <div className="flex items-center">{children}</div>

      <div className="py-0 group-last/toolbar-group:hidden!">
        <Separator orientation="vertical" />
      </div>
    </div>
  );
}

type TooltipProps<T extends React.ElementType> = {
  tooltip?: React.ReactNode;
  tooltipContentProps?: Omit<React.ComponentPropsWithoutRef<typeof TooltipContent>, "children">;
  tooltipProps?: Omit<React.ComponentPropsWithoutRef<typeof Tooltip>, "children">;
  tooltipTriggerProps?: React.ComponentPropsWithoutRef<typeof TooltipTrigger>;
} & React.ComponentProps<T>;

function withTooltip<T extends React.ElementType>(Component: T) {
  // Forwards the ref through to `Component`; without it the wrapper swallows the
  // ref and Radix's `asChild` anchors resolve to null. See ToolbarButtonImpl.
  return React.forwardRef<HTMLElement, TooltipProps<T>>(function ExtendComponent(
    { tooltip, tooltipContentProps, tooltipProps, tooltipTriggerProps, ...props },
    ref
  ) {
    const [mounted, setMounted] = React.useState(false);

    React.useEffect(() => {
      setMounted(true);
    }, []);

    // `Component` is generic, so TS cannot know it takes a ref; the cast is
    // only to satisfy that, the ref itself is what makes Radix's `asChild`
    // anchors resolve.
    const WithRef = Component as React.ElementType;
    const component = <WithRef ref={ref} {...(props as React.ComponentProps<T>)} />;

    if (tooltip && mounted) {
      return (
        <Tooltip {...tooltipProps}>
          <TooltipTrigger asChild {...tooltipTriggerProps}>
            {component}
          </TooltipTrigger>

          <TooltipContent {...tooltipContentProps}>{tooltip}</TooltipContent>
        </Tooltip>
      );
    }

    return component;
  });
}

function TooltipContent({
  children,
  className,
  // CHANGE
  sideOffset = 4,
  ...props
}: React.ComponentProps<typeof TooltipPrimitive.Content>) {
  return (
    <TooltipPrimitive.Portal>
      <TooltipPrimitive.Content
        className={cn(
          "bg-primary text-primary-foreground z-50 w-fit origin-(--radix-tooltip-content-transform-origin) rounded-md px-1.5 py-1.5 text-xs text-balance",
          className
        )}
        data-slot="tooltip-content"
        sideOffset={sideOffset}
        {...props}
      >
        {children}
        {/* CHANGE */}
        {/* <TooltipPrimitive.Arrow className="z-50 size-2.5 translate-y-[calc(-50%_-_2px)] rotate-45 rounded-[2px] bg-primary fill-primary" /> */}
      </TooltipPrimitive.Content>
    </TooltipPrimitive.Portal>
  );
}

export function ToolbarMenuGroup({
  children,
  className,
  label,
  ...props
}: React.ComponentProps<typeof DropdownMenuRadioGroup> & { label?: string }) {
  return (
    <>
      <DropdownMenuSeparator
        className={cn(
          "hidden",
          "mb-0 shrink-0 peer-has-[[role=menuitem]]/menu-group:block peer-has-[[role=menuitemradio]]/menu-group:block peer-has-[[role=option]]/menu-group:block"
        )}
      />

      <DropdownMenuRadioGroup
        {...props}
        className={cn(
          "hidden",
          "peer/menu-group group/menu-group my-1.5 has-[[role=menuitem]]:block has-[[role=menuitemradio]]:block has-[[role=option]]:block",
          className
        )}
      >
        {label && (
          <DropdownMenuLabel className="text-muted-foreground text-xs font-semibold select-none">
            {label}
          </DropdownMenuLabel>
        )}
        {children}
      </DropdownMenuRadioGroup>
    </>
  );
}

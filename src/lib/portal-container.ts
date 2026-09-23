/**
 * A single host element for everything Radix portals to document.body —
 * dropdowns, popovers, tooltips, dialogs.
 *
 * It carries `rte-scope`, which is what every editor style is scoped to.
 * Without it, portalled menus land outside the scope and lose the whole
 * stylesheet: no tokens, no utilities, and Bootstrap's own `.dropdown-menu`
 * (which is `display: none` until Bootstrap's JS opens it) applying instead.
 */
let container: HTMLElement | null = null;

export function editorPortalContainer(): HTMLElement | undefined {
  if (typeof document === "undefined") return undefined;
  if (container?.isConnected) return container;

  container = document.createElement("div");
  container.className = "rte-scope rte-portal";
  document.body.appendChild(container);
  return container;
}

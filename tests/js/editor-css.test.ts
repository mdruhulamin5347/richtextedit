/**
 * Stylesheet guards.
 *
 * Two failure modes this file exists to catch, both invisible in a component
 * test: styling that does not reach Radix's portalled menus (they mount on
 * document.body, outside the editor container), and preflight escaping the
 * editor to flatten Bootstrap across the admin panel.
 */
import { describe, expect, it } from "vitest";
import { readFileSync, readdirSync } from "node:fs";
import { join } from "node:path";
import postcss from "postcss";

const SCOPES = [".rte-scope", "[data-radix-popper-content-wrapper]"];

// The standalone build's stylesheet — the one a Blade page loads.
const built = readdirSync("dist/standalone")
  .filter((f) => f.endsWith(".css"))
  .map((f) => readFileSync(join("dist/standalone", f), "utf8"))
  .join("\n");



describe("design tokens", () => {
  it("are defined on the editor scope, which portalled menus also carry", () => {
    // Radix mounts menus and dialogs on document.body. They reach these tokens
    // because they are portalled into a container that carries the scope class
    // (lib/portal-container.ts) — not because the tokens are global.
    const scoped = built.match(/\.rte-scope\{[^}]*\}/g)?.join("") ?? "";
    for (const token of ["--popover", "--background", "--border", "--muted", "--primary"]) {
      expect(scoped).toContain(token);
    }
  });

  it("lift portalled menus above the admin panel's z-index stack", () => {
    // The admin chrome reaches z-index 20000; Tailwind's z-50 is 50. Applied to
    // the popper wrapper, which is the positioned element the browser stacks —
    // the portal host itself must stay unpositioned, or it becomes the
    // containing block and every menu is offset by the height of the page.
    expect(built).toMatch(/\[data-radix-popper-content-wrapper\]\{[^}]*z-index:200\d\d/);
    expect(built).toMatch(/\.rte-portal\{[^}]*position:static/);
  });
});

describe("stylesheet isolation", () => {
  it("confines EVERY rule to the editor scope", () => {
    // One unscoped rule would restyle headings, lists, tables and form controls
    // on all ~200 admin pages that load this bundle.
    const escaped: string[] = [];
    postcss.parse(built).walkRules((rule) => {
      if (rule.parent?.type === "atrule" && /keyframes$/.test((rule.parent as any).name)) return;
      for (const selector of rule.selectors) {
        if (!selector.trim().startsWith(".rte-scope")) escaped.push(selector);
      }
    });
    expect(escaped).toEqual([]);
  });

  it("ships preflight, scoped — the baseline Plate's components assume", () => {
    expect(built).toMatch(/\.rte-scope \*[^{]*\{[^}]*box-sizing:border-box/);
  });

  it("marks colliding utilities !important, since Bootstrap declares its own that way", () => {
    // Specificity cannot beat an !important declaration, so these must match it.
    expect(built).toMatch(/\.rte-scope \.p-0\{padding:0 ?!important\}/);
  });

  it("neutralises Bootstrap's .table, which forced width:100% and broke column resizing", () => {
    // Tailwind's .table is only `display: table`, so there was no editor
    // declaration for width to override Bootstrap's with.
    expect(built).toMatch(/\.rte-scope \.table\{[^}]*width:unset/);
  });
});

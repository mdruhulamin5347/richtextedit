/**
 * Font Utility Functions
 *
 * Shared utilities for handling font-family values across
 * PDF, DOCX, and HTML export pipelines.
 *
 * ROOT CAUSE: Plate's serializeHtml produces broken HTML when fontFamily
 * values contain double quotes (e.g., '"Times New Roman", serif').
 * The inner " breaks the HTML attribute boundary:
 *   style="font-family:"Times New Roman", serif"
 *
 * FIX: Replace double quotes with single quotes in fontFamily values
 * BEFORE any serialization (HTML or DOCX via @platejs/docx-io which
 * also uses serializeHtml internally).
 */

/**
 * Recursively sanitize fontFamily values in Plate JSON nodes.
 * Replaces double quotes with single quotes to prevent broken HTML attributes.
 *
 * @example
 * // Input:  fontFamily: '"Times New Roman", serif'
 * // Output: fontFamily: "'Times New Roman', serif"
 */
export function sanitizeFontFamilyQuotes(nodes: any[]): any[] {
  return nodes.map((node: any) => {
    const sanitized = { ...node };
    if (sanitized.fontFamily && typeof sanitized.fontFamily === "string") {
      sanitized.fontFamily = sanitized.fontFamily.replace(/"/g, "'");
    }
    if (sanitized.children && Array.isArray(sanitized.children)) {
      sanitized.children = sanitizeFontFamilyQuotes(sanitized.children);
    }
    return sanitized;
  });
}

/**
 * Post-process HTML string to fix any remaining broken font quotes.
 * Safety net for edge cases not caught by pre-processing.
 */
export function fixBrokenFontQuotesInHtml(html: string): string {
  // Fix data-slate-font-family attributes that got broken by &quot;
  return html.replace(/data-slate-font-family="([^"]*)"/g, (match, value) => {
    const fixed = value.replace(/&quot;/g, "'");
    return `data-slate-font-family="${fixed}"`;
  });
}

/**
 * Extract the primary font name from a CSS font-family stack.
 * Strips quotes and returns only the first font name.
 *
 * @example
 * parseFontFamily("'Times New Roman', serif") → "Times New Roman"
 * parseFontFamily("Arial") → "Arial"
 * parseFontFamily("\"Courier New\", monospace") → "Courier New"
 */
export function parseFontFamily(fontFamily: string | undefined): string | undefined {
  if (!fontFamily) return undefined;
  const primaryFont = fontFamily.split(",")[0].trim();
  return primaryFont.replace(/['"]/g, "").trim() || undefined;
}

/**
 * Convert any CSS color value to a 6-digit hex string (without #).
 * Handles: #hex, #shorthand, rgb(), rgba(), and common named colors.
 * Returns undefined if the value cannot be parsed.
 *
 * The `docx` library requires 6-digit hex values (e.g., "E6E6E6").
 *
 * @example
 * cssColorToHex("#E6E6E6") → "E6E6E6"
 * cssColorToHex("rgb(230, 230, 230)") → "E6E6E6"
 * cssColorToHex("rgba(255, 0, 0, 0.5)") → "FF0000"
 * cssColorToHex("red") → "FF0000"
 */
export function cssColorToHex(color: string | undefined): string | undefined {
  if (!color) return undefined;
  const trimmed = color.trim();
  if (!trimmed) return undefined;

  // Already hex format: #RRGGBB or #RGB
  if (trimmed.startsWith("#")) {
    let hex = trimmed.slice(1);
    // Expand shorthand (#RGB → #RRGGBB)
    if (hex.length === 3) {
      hex = hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2];
    }
    // Validate 6-digit hex
    if (/^[0-9a-fA-F]{6}$/.test(hex)) {
      return hex.toUpperCase();
    }
    return undefined;
  }

  // rgb() or rgba() format
  const rgbMatch = trimmed.match(/^rgba?\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})/);
  if (rgbMatch) {
    const r = Math.min(255, parseInt(rgbMatch[1], 10));
    const g = Math.min(255, parseInt(rgbMatch[2], 10));
    const b = Math.min(255, parseInt(rgbMatch[3], 10));
    return (
      r.toString(16).padStart(2, "0") +
      g.toString(16).padStart(2, "0") +
      b.toString(16).padStart(2, "0")
    ).toUpperCase();
  }

  // Common named colors
  const namedColors: Record<string, string> = {
    black: "000000",
    white: "FFFFFF",
    red: "FF0000",
    green: "008000",
    blue: "0000FF",
    yellow: "FFFF00",
    cyan: "00FFFF",
    magenta: "FF00FF",
    gray: "808080",
    grey: "808080",
    orange: "FFA500",
    purple: "800080",
    pink: "FFC0CB",
    brown: "A52A2A",
    navy: "000080",
    teal: "008080",
    maroon: "800000",
    olive: "808000",
    lime: "00FF00",
    aqua: "00FFFF",
    silver: "C0C0C0",
    fuchsia: "FF00FF",
    transparent: undefined!,
  };
  const named = namedColors[trimmed.toLowerCase()];
  if (named) return named;
  if (trimmed.toLowerCase() === "transparent") return undefined;

  // If it looks like a raw hex string without #
  if (/^[0-9a-fA-F]{6}$/.test(trimmed)) {
    return trimmed.toUpperCase();
  }

  return undefined;
}

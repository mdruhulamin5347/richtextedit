/**
 * The paragraphs of a clipboard's RTF: their text, tab stops, indents, line spacing and font.
 *
 * LibreOffice's clipboard HTML leaves a paragraph's tab stops out altogether —
 * `Status<TAB><TAB>…Alive` arrives with no trace of the ruler stop at 1548 twips
 * that put `Alive` where it is — but the RTF it copies alongside states every
 * one (`\tx1548`), as Word's does. This reads just enough RTF to recover them:
 * paragraph text (to match a paragraph to its HTML block), `\tx` stops with
 * their alignment, the `\li` / `\fi` indents, the `\sl` line spacing and the
 * font the paragraph's text starts in. Everything else is skipped.
 */

export type RtfTabAlignment = "left" | "center" | "right" | "decimal";

export interface RtfTabStop {
  /** From the left edge of the text column, in pt. */
  positionPt: number;
  alignment: RtfTabAlignment;
}

/**
 * How far apart a paragraph's lines are, from `\sl` and `\slmult`.
 *
 * LibreOffice's HTML leaves this out for a paragraph inside a table cell, and
 * only its RTF has it (`\intbl\sl480\slmult1` — double). A paragraph that states
 * no `\sl` is single-spaced.
 */
export type RtfLineSpacing =
  /** Word's "Multiple", LibreOffice's "Proportional": lines of the font's natural height. */
  | { rule: "multiple"; lines: number }
  /** Word's "Exactly", LibreOffice's "Fixed". */
  | { rule: "exact"; pt: number }
  /** "At least": the font's natural line, or this, whichever is taller. */
  | { rule: "atLeast"; pt: number };

export interface RtfParagraph {
  text: string;
  stops: RtfTabStop[];
  /** `\li` — the left indent, in pt. */
  leftPt: number;
  /** `\fi` — the first line's indent relative to it, in pt (negative: hanging). */
  firstPt: number;
  lineSpacing: RtfLineSpacing;
  /**
   * The font its first character is set in, by name — `\fN` looked up in the
   * font table. LibreOffice's HTML names no font for a run in the paragraph
   * style's own font, so this is the only place that says what it is.
   */
  font?: string;
}

/** Groups whose content is not the document's text. */
const SKIPPED_DESTINATIONS: ReadonlySet<string> = new Set([
  "fonttbl",
  "colortbl",
  "stylesheet",
  "info",
  "pict",
  "object",
  "header",
  "headerl",
  "headerr",
  "headerf",
  "footer",
  "footerl",
  "footerr",
  "footerf",
  "footnote",
  "listtable",
  "listoverridetable",
  "rsidtbl",
  "generator",
  "xmlnstbl",
  "themedata",
  "colorschememapping",
  "latentstyles",
  "datastore",
  "pgdsctbl",
  "fldinst",
]);

const SYMBOLS: Readonly<Record<string, string>> = {
  emdash: "—",
  endash: "–",
  bullet: "•",
  lquote: "‘",
  rquote: "’",
  ldblquote: "“",
  rdblquote: "”",
  emspace: " ",
  enspace: " ",
};

const TWIPS_PER_PT = 20;

/** `\sl240\slmult1` is one line. */
const TWIPS_PER_LINE = 240;

/**
 * `\slN` read with its `\slmult`: N/240 lines when it is set, else N twips —
 * negative for "exactly", positive for "at least". None, or 0, is single.
 */
function lineSpacingOf(slTwips: number | null, multiple: boolean): RtfLineSpacing {
  if (!slTwips) return { rule: "multiple", lines: 1 };
  if (slTwips < 0) return { rule: "exact", pt: -slTwips / TWIPS_PER_PT };
  return multiple
    ? { rule: "multiple", lines: slTwips / TWIPS_PER_LINE }
    : { rule: "atLeast", pt: slTwips / TWIPS_PER_PT };
}

/** `{\f4\froman\fprq2\fcharset0 Calibri;}` → 4 → Calibri, for every font in the table. */
function readFontTable(source: string): Map<number, string> {
  const fonts = new Map<number, string>();
  const start = source.indexOf("{\\fonttbl");
  if (start === -1) return fonts;
  for (const match of source.slice(start, start + 20000).matchAll(/\{\\f(\d+)((?:\\[a-z]+-?\d*\s?)*)([^;{}\\]*)/g)) {
    const name = match[3].trim();
    if (name && !fonts.has(Number(match[1]))) fonts.set(Number(match[1]), name);
  }
  return fonts;
}

/**
 * Every body paragraph of `rtf`, in order. An empty array for anything that is
 * not RTF, so a caller can treat "no RTF" and "unreadable RTF" the same way.
 */
export function readRtfParagraphs(rtf: string | null | undefined): RtfParagraph[] {
  const source = String(rtf ?? "");
  if (!source.startsWith("{\\rtf")) return [];

  const paragraphs: RtfParagraph[] = [];
  const skipStack: boolean[] = [];
  const fonts = readFontTable(source);
  const defaultFont = Number(/\\deff(\d+)/.exec(source)?.[1] ?? NaN);
  /** `\fN` is a character property, so it is scoped by groups like the rest. */
  const fontStack: number[] = [];
  let font = defaultFont;
  let paragraphFont: string | undefined;
  let skipping = false;
  let groupJustOpened = false;

  let text = "";
  let stops: RtfTabStop[] = [];
  let leftPt = 0;
  let firstPt = 0;
  let pendingAlignment: RtfTabAlignment = "left";
  let pendingBar = false;
  let slTwips: number | null = null;
  let slMultiple = false;
  let unicodeFallback = 1;
  let skipChars = 0;

  const endParagraph = () => {
    paragraphs.push({
      text,
      stops: [...stops].sort((a, b) => a.positionPt - b.positionPt),
      leftPt,
      firstPt,
      lineSpacing: lineSpacingOf(slTwips, slMultiple),
      ...(paragraphFont ? { font: paragraphFont } : {}),
    });
    text = "";
    paragraphFont = undefined;
  };
  const emit = (chars: string) => {
    if (skipping) return;
    if (skipChars > 0) {
      skipChars--;
      return;
    }
    text += chars;
    if (paragraphFont === undefined && /\S/.test(chars)) paragraphFont = fonts.get(font);
  };

  for (let i = 0; i < source.length; i++) {
    const char = source[i];

    if (char === "{") {
      skipStack.push(skipping);
      fontStack.push(font);
      groupJustOpened = true;
      continue;
    }
    if (char === "}") {
      skipping = skipStack.pop() ?? false;
      font = fontStack.pop() ?? font;
      groupJustOpened = false;
      continue;
    }
    if (char === "\r" || char === "\n") continue;

    if (char !== "\\") {
      groupJustOpened = false;
      emit(char);
      continue;
    }

    const next = source[i + 1] ?? "";
    // A control word: letters, an optional signed number, an optional space.
    if (/[a-z]/i.test(next)) {
      let j = i + 1;
      while (j < source.length && /[a-z]/i.test(source[j])) j++;
      const word = source.slice(i + 1, j);
      let k = j;
      if (source[k] === "-" || /\d/.test(source[k] ?? "")) {
        k++;
        while (k < source.length && /\d/.test(source[k])) k++;
      }
      const param = k > j ? Number(source.slice(j, k)) : null;
      if (source[k] === " ") k++;
      i = k - 1;

      if (groupJustOpened && SKIPPED_DESTINATIONS.has(word)) skipping = true;
      groupJustOpened = false;
      if (skipping) continue;

      switch (word) {
        case "par":
        case "cell":
          endParagraph();
          break;
        case "pard":
          stops = [];
          leftPt = 0;
          firstPt = 0;
          pendingAlignment = "left";
          pendingBar = false;
          slTwips = null;
          slMultiple = false;
          break;
        case "tab":
          emit("\t");
          break;
        case "line":
          emit("\n");
          break;
        case "tqr":
          pendingAlignment = "right";
          break;
        case "tqc":
          pendingAlignment = "center";
          break;
        case "tqdec":
          pendingAlignment = "decimal";
          break;
        case "tb":
          pendingBar = true;
        // falls through: a bar stop carries its position the way \tx does
        case "tx":
          if (param !== null && !pendingBar) {
            stops.push({ positionPt: param / TWIPS_PER_PT, alignment: pendingAlignment });
          }
          pendingAlignment = "left";
          pendingBar = false;
          break;
        case "li":
          leftPt = (param ?? 0) / TWIPS_PER_PT;
          break;
        case "fi":
          firstPt = (param ?? 0) / TWIPS_PER_PT;
          break;
        // A paragraph can state these twice — LibreOffice writes its style's
        // spacing, then the paragraph's own — and the last one is the one in force.
        case "sl":
          slTwips = param ?? 0;
          break;
        case "slmult":
          slMultiple = (param ?? 1) !== 0;
          break;
        case "plain":
          font = defaultFont;
          break;
        case "f":
          if (param !== null) font = param;
          break;
        case "uc":
          unicodeFallback = param ?? 1;
          break;
        case "u":
          if (param !== null) {
            emit(String.fromCharCode(param < 0 ? param + 65536 : param));
            skipChars = unicodeFallback;
          }
          break;
        default:
          if (SYMBOLS[word]) emit(SYMBOLS[word]);
      }
      continue;
    }

    // A control symbol.
    i++;
    groupJustOpened = groupJustOpened && next === "*";
    if (next === "*") {
      skipping = true;
      continue;
    }
    if (skipping) {
      if (next === "'") i += 2;
      continue;
    }
    if (next === "'") {
      const code = Number.parseInt(source.slice(i + 1, i + 3), 16);
      i += 2;
      if (Number.isFinite(code)) emit(code === 0xa0 ? " " : String.fromCharCode(code));
    } else if (next === "~") {
      emit(" ");
    } else if (next === "\\" || next === "{" || next === "}") {
      emit(next);
    } else if (next === "\r" || next === "\n") {
      endParagraph();
    }
  }

  // A selection that ends mid-paragraph ends without a `\par`.
  if (text.trim()) endParagraph();

  return paragraphs;
}

/** Text as a match key: what is left once every kind of space is gone. */
function matchKey(text: string): string {
  return text.replace(/[\s ]+/g, "").toLowerCase();
}

/**
 * Each HTML block's RTF paragraph, matched in order by text.
 *
 * The two lists are not index-aligned — the docx cleaner drops empty
 * paragraphs, a selection can start mid-way through one — so this walks both
 * and pairs a block with the next paragraph whose text is the same (or one
 * starts the other), looking a few paragraphs ahead at most.
 */
export function matchRtfParagraphs<T extends Element>(blocks: T[], paragraphs: RtfParagraph[]): Map<T, RtfParagraph> {
  const matched = new Map<T, RtfParagraph>();
  const keyed = paragraphs.map((paragraph) => ({ paragraph, key: matchKey(paragraph.text) })).filter((p) => p.key);
  let cursor = 0;

  for (const block of blocks) {
    const key = matchKey(block.textContent ?? "");
    if (!key) continue;
    for (let i = cursor; i < Math.min(keyed.length, cursor + 8); i++) {
      const candidate = keyed[i].key;
      const shorter = Math.min(candidate.length, key.length);
      if (candidate === key || (shorter >= 8 && (candidate.startsWith(key) || key.startsWith(candidate)))) {
        matched.set(block, keyed[i].paragraph);
        cursor = i + 1;
        break;
      }
    }
  }

  return matched;
}

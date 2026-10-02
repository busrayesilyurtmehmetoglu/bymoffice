// Small tables the layout engine needs: characters of the symbol fonts as
// Unicode, the highlight colours, and the formats of page numbers.

import { formatCounter } from '../../formats/docx/numbering';

/** Which table a font selects: 0 none, 1 Symbol, 2 Wingdings, 3 Wingdings 2, 4 Wingdings 3, 5 Webdings. */
export type SymbolKind = 0 | 1 | 2 | 3 | 4 | 5;

/** Characters of a symbol font -> Unicode: for a code, the characters of that code and the codes after it. */
type Ranges = [start: number, characters: string][];

/**
 * The Symbol font (Adobe Symbol encoding, 0x21..0xFE), in full but for the
 * two glyphs Unicode has no character for (0x60, the piece that lengthens a
 * radical sign, and 0xF0, a logo). The pieces brackets and arrows are built
 * from (0xBD, 0xBE, 0xE6..0xEF, 0xF3..0xFE) are the characters Unicode has
 * for them; few fonts hold them.
 */
const SYMBOL_RANGES: Ranges = [
  [0x21, '!∀#∃%&∋()∗+,−./0123456789:;<=>?≅ΑΒΧΔΕΦΓΗΙϑΚΛΜΝΟΠΘΡΣΤΥςΩΞΨΖ[∴]⊥_'],
  [0x61, 'αβχδεφγηιϕκλμνοπθρστυϖωξψζ{|}∼'],
  [0xa0, '€ϒ′≤⁄∞ƒ♣♦♥♠↔←↑→↓°±″≥×∝∂•÷≠≡≈…⏐⎯↵ℵℑℜ℘⊗⊕∅∩∪⊃⊇⊄⊂⊆∈∉∠∇®©™∏√⋅¬∧∨⇔⇐⇑⇒⇓◊〈®©™∑⎛⎜⎝⎡⎢⎣⎧⎨⎩⎪'],
  [0xf1, '〉∫⌠⎮⌡⎞⎟⎠⎤⎥⎦⎫⎬⎭'],
];

/**
 * Wingdings: the glyphs that have a character of one UTF-16 unit, so that
 * text keeps its length. Where Unicode put the glyph beyond that (pencil,
 * telephone, envelope, the plain and the heavy arrows 0xDF..0xEE, the ballot
 * marks 0xFB..0xFE), the long-standing character of the same picture is used.
 *
 * UNSURE: written from the published mapping from memory, not checked glyph
 * by glyph against the font. The bullets, boxes, check marks, hands, arrows,
 * circled digits and zodiac signs are the well known ones; less certain are
 * 0x21, 0x2B, 0x3E, 0x49, 0x4F, 0x78..0x7C and 0xB1..0xB6.
 */
const WINGDINGS_RANGES: Ranges = [
  [0x21, '✏✂✁'],
  [0x28, '☎✆✉✉'],
  [0x36, '⌛⌨'],
  [0x3e, '✇✍'],
  [0x41, '✌'],
  [0x45, '☜☞☝☟✋☺'],
  [0x4c, '☹'],
  [0x4e, '☠⚐'],
  [0x51, '✈☼'],
  [0x54, '❄'],
  [0x56, '✞'],
  [0x58, '✠✡☪☯ॐ☸♈♉♊♋♌♍♎♏♐♑♒♓'],
  [0x6c, '●❍■□❐❑❒⬧⧫◆❖⬥⌧⍓⌘❀✿❝❞▯'],
  [0x80, '⓪①②③④⑤⑥⑦⑧⑨⑩⓿❶❷❸❹❺❻❼❽❾❿'],
  [0x9e, '·•▪○'],
  [0xa4, '◉◎'],
  [0xa7, '▪◻'],
  [0xaa, '✦★✶✴✹✵'],
  [0xb1, '⌖⟡⌑'],
  [0xb5, '✪✰'],
  [0xd5, '⌫⌦'],
  [0xd8, '➢'],
  [0xdf, '←→↑↓↖↗↙↘⬅➔⬆⬇⬉⬈⬋⬊⇦⇨⇧⇩⬄⇳⬀⬁⬃⬂▭▫✗✓☒☑'],
];

/**
 * Wingdings 2: the check marks and boxes (0x4F..0x54), the circled digits
 * and the round and square bullets.
 * UNSURE: only 0x50 (check mark), 0x52 (checked box), 0x4F, 0x51, 0x53 and
 * 0x54 (crosses, plain and boxed) and 0xA3 (empty box) are certain.
 */
const WINGDINGS2_RANGES: Ranges = [
  [0x4f, '✗✓☒☑☒☒'],
  [0x6a, '⓪①②③④⑤⑥⑦⑧⑨⑩⓿❶❷❸❹❺❻❼❽❾❿'],
  [0x97, '●●○○'],
  [0xa0, '▪■■□'],
  [0xaa, '▪◆◇'],
];

/**
 * Wingdings 3: arrows and the triangles used as bullets.
 * UNSURE: the triangles 0x70, 0x71, 0x74, 0x75 and 0x7D are certain, the rest is from memory.
 */
const WINGDINGS3_RANGES: Ranges = [
  [0x21, '←→↑↓↖↗↙↘'],
  [0x70, '▲▼△▽◀▶◁▷◣◢◤◥◂▸▴▾'],
];

/**
 * Webdings: the few glyphs documents use as marks and bullets.
 * UNSURE: 0x61 (check mark), 0x72 (cross), 0x6E (dot) and the triangles 0x33..0x36 are certain.
 */
const WEBDINGS_RANGES: Ranges = [
  [0x33, '◀▶▲▼'],
  [0x59, '♥'],
  [0x61, '✔'],
  [0x63, '□'],
  [0x67, '■'],
  [0x69, 'ℹ'],
  [0x6e, '●'],
  [0x72, '✖'],
  [0x73, '❓'],
];

function tableOf(ranges: Ranges): (string | undefined)[] {
  const table: (string | undefined)[] = new Array(256);
  for (const [start, characters] of ranges) {
    for (let i = 0; i < characters.length && start + i < 256; i++) table[start + i] = characters.charAt(i);
  }
  return table;
}

/** By SymbolKind, less one. */
const TABLES: (string | undefined)[][] = [SYMBOL_RANGES, WINGDINGS_RANGES, WINGDINGS2_RANGES, WINGDINGS3_RANGES, WEBDINGS_RANGES].map(tableOf);

/** For a character in the private use area whose font is not one of the above: the bullets Word writes most. */
const PRIVATE_USE: Record<number, string> = { 0xb7: '•', 0xa7: '▪', 0xd8: '➢', 0xfc: '✓', 0x6f: '□', 0x6e: '■', 0x76: '❖' };

const KINDS: Record<string, SymbolKind> = { symbol: 1, wingdings: 2, 'wingdings 2': 3, 'wingdings 3': 4, webdings: 5 };

/** Which table a font name selects; 0 for a font that is not one of the symbol fonts the engine knows. */
export function symbolFont(family: string | undefined): SymbolKind {
  if (family === undefined) return 0;
  return KINDS[family.toLowerCase()] || 0;
}

/**
 * Unicode for one character of a symbol font, undefined when the tables do
 * not know it. Symbol fonts are addressed either directly (0x20..0xFF) or
 * through the private use area (0xF020..0xF0FF); both are accepted.
 */
export function symbolChar(kind: SymbolKind, code: number): string | undefined {
  const privateUse = code >= 0xf020 && code <= 0xf0ff;
  const low = privateUse ? code - 0xf000 : code;
  if (low < 0x20 || low > 0xff) return undefined;
  if (kind !== 0) {
    const table = TABLES[kind - 1];
    return table ? table[low] : undefined;
  }
  return privateUse ? PRIVATE_USE[low] : undefined;
}

export interface MappedSymbols {
  text: string;
  /** A character was replaced, or is the same in Unicode as in the font. */
  changed: boolean;
  /**
   * Set when some characters were mapped and others, glyphs of the font the
   * tables do not know, were not: for each character of `text`, whether it
   * was mapped. Unmapped glyphs have to stay in the font of the document.
   */
  mapped?: boolean[];
}

/** Maps the text of a run in a symbol font; `changed` says whether any character was replaced. */
export function mapSymbols(kind: SymbolKind, text: string): MappedSymbols {
  let out = '';
  let changed = false;
  let kept = false;
  const flags: boolean[] | undefined = kind !== 0 ? [] : undefined;
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    const privateUse = code >= 0xf020 && code <= 0xf0ff;
    // Outside a symbol font only the private use area is looked at.
    const mapped = kind !== 0 || privateUse ? symbolChar(kind, code) : undefined;
    if (mapped !== undefined) {
      out += mapped;
      changed = true;
      if (flags) flags.push(true);
      continue;
    }
    out += text.charAt(i);
    if (!flags) continue;
    // A glyph of the font the tables do not know stays in the font; white space and characters that are not
    // glyphs of it go with what is before them.
    if (code > 0x20 && code !== 0xa0 && (privateUse || code <= 0xff)) {
      kept = true;
      flags.push(false);
    } else flags.push(i > 0 ? flags[i - 1] : true);
  }
  const result: MappedSymbols = { text: out, changed };
  if (changed && kept && flags) result.mapped = flags;
  return result;
}

/** True when the text has a character the tables above might replace. */
export function mayHoldSymbols(kind: SymbolKind, text: string): boolean {
  if (kind !== 0) return true;
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    if (code >= 0xf020 && code <= 0xf0ff) return true;
  }
  return false;
}

/** Highlight colour names (ST_HighlightColor) -> "#RRGGBB". */
export const HIGHLIGHT: Record<string, string> = {
  black: '#000000',
  blue: '#0000FF',
  cyan: '#00FFFF',
  green: '#00FF00',
  magenta: '#FF00FF',
  red: '#FF0000',
  yellow: '#FFFF00',
  white: '#FFFFFF',
  darkBlue: '#000080',
  darkCyan: '#008080',
  darkGreen: '#008000',
  darkMagenta: '#800080',
  darkRed: '#800000',
  darkYellow: '#808000',
  darkGray: '#808080',
  lightGray: '#C0C0C0',
};

/** Characters a tab leader is drawn with, by ST_TabTlc. */
export const LEADER: Record<string, string> = { dot: '.', hyphen: '-', underscore: '_', heavy: '_', middleDot: '·' };

/**
 * A page number (or page count) as text. `format` is a number format of the
 * section (ST_NumberFormat) or the argument of a "\*" switch of the field
 * ("roman", "ROMAN", "alphabetic", "ALPHABETIC", "Arabic").
 */
export function formatPageNumber(value: number, format: string | undefined): string {
  switch (format) {
    case 'roman':
      return formatCounter(value, 'lowerRoman');
    case 'ROMAN':
    case 'Roman':
      return formatCounter(value, 'upperRoman');
    case 'alphabetic':
      return formatCounter(value, 'lowerLetter');
    case 'ALPHABETIC':
    case 'Alphabetic':
      return formatCounter(value, 'upperLetter');
    case 'numberInDash':
      return `- ${value} -`;
    case 'Arabic':
    case 'arabic':
    case undefined:
      return String(value);
  }
  return formatCounter(value, format);
}

/** The fields the engine computes itself. */
export type LiveField = 'PAGE' | 'NUMPAGES' | 'SECTIONPAGES';

export interface FieldCode {
  /** Upper case name of the field: "PAGE", "HYPERLINK", "TOC", ... */
  name: string;
  /** Argument of the first "\*" switch that is a number format. */
  format?: string;
  /** HYPERLINK only: the address, or "#bookmark". */
  target?: string;
}

const NUMBER_SWITCHES: Record<string, true> = { roman: true, ROMAN: true, Roman: true, alphabetic: true, ALPHABETIC: true, Alphabetic: true, Arabic: true, arabic: true };

/** Longest field code that is looked at; a longer one is not a field the engine acts on. */
const MAX_FIELD_CODE = 4096;

/** Reads the name and the few switches the engine acts on from a field code. */
export function parseFieldCode(instruction: string): FieldCode {
  const code = instruction.length > MAX_FIELD_CODE ? instruction.slice(0, MAX_FIELD_CODE) : instruction;
  // Tokens: quoted strings (with \" inside) or runs of non-space characters.
  const tokens: string[] = [];
  let i = 0;
  while (i < code.length) {
    const c = code.charCodeAt(i);
    if (c <= 0x20) {
      i++;
      continue;
    }
    if (c === 0x22) {
      let text = '';
      i++;
      while (i < code.length && code.charCodeAt(i) !== 0x22) {
        if (code.charCodeAt(i) === 0x5c && i + 1 < code.length) i++;
        text += code.charAt(i++);
      }
      i++;
      tokens.push('"' + text);
      continue;
    }
    let end = i;
    while (end < code.length && code.charCodeAt(end) > 0x20) end++;
    tokens.push(code.slice(i, end));
    i = end;
  }
  const unquote = (token: string | undefined) => (token === undefined ? undefined : token.charAt(0) === '"' ? token.slice(1) : token);
  const out: FieldCode = { name: tokens.length ? tokens[0].toUpperCase() : '' };
  for (let k = 1; k < tokens.length; k++) {
    const token = tokens[k];
    if (token === '\\*') {
      const argument = unquote(tokens[k + 1]);
      if (argument !== undefined && NUMBER_SWITCHES[argument] && out.format === undefined) out.format = argument;
      k++;
    } else if (out.name === 'HYPERLINK') {
      if (token === '\\l') {
        const anchor = unquote(tokens[k + 1]);
        if (anchor !== undefined && out.target === undefined) out.target = '#' + anchor;
        else if (anchor !== undefined && out.target!.indexOf('#') < 0) out.target += '#' + anchor;
        k++;
      } else if (token === '\\o' || token === '\\t' || token === '\\m') k++;
      else if (token.charAt(0) !== '\\' && out.target === undefined) out.target = unquote(token);
    }
  }
  return out;
}

export const isLiveField = (name: string): name is LiveField => name === 'PAGE' || name === 'NUMPAGES' || name === 'SECTIONPAGES';

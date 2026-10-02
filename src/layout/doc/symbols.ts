// Small tables the layout engine needs: characters of the symbol fonts as
// Unicode, the highlight colours, and the formats of page numbers.

import { formatCounter } from '../../formats/docx/numbering';

/**
 * Characters of the Symbol font by code (0x20..0xFF) -> Unicode. Only the
 * characters documents use for bullets and simple formulas; a character that
 * is not listed stays as it is, in the font the document names.
 */
const SYMBOL: Record<number, string> = {
  0x2d: '−',
  0xa3: '≤',
  0xa5: '∞',
  0xa7: '♣',
  0xa8: '♦',
  0xa9: '♥',
  0xaa: '♠',
  0xab: '↔',
  0xac: '←',
  0xad: '↑',
  0xae: '→',
  0xaf: '↓',
  0xb0: '°',
  0xb1: '±',
  0xb3: '≥',
  0xb4: '×',
  0xb7: '•',
  0xb8: '÷',
  0xb9: '≠',
  0xba: '≡',
  0xbb: '≈',
  0xbc: '…',
  0xd7: '⋅',
  0xde: '⇒',
  0xe0: '◊',
};
const SYMBOL_LOWER = 'αβχδεφγηιϕκλμνοπθρστυϖωξψζ';
const SYMBOL_UPPER = 'ΑΒΧΔΕΦΓΗΙϑΚΛΜΝΟΠΘΡΣΤΥςΩΞΨΖ';

/** The same for Wingdings. */
const WINGDINGS: Record<number, string> = {
  0x22: '✂',
  0x28: '☎',
  0x2a: '✉',
  0x46: '☞',
  0x4a: '☺',
  0x6c: '●',
  0x6e: '■',
  0x6f: '□',
  0x71: '❑',
  0x72: '❒',
  0x73: '⬧',
  0x74: '⧫',
  0x75: '◆',
  0x76: '❖',
  0x77: '⬥',
  0x9f: '•',
  0xa1: '○',
  0xa7: '▪',
  0xa8: '◻',
  0xd8: '➢',
  0xe0: '→',
  0xe8: '➔',
  0xf0: '⇨',
  0xfb: '✗',
  0xfc: '✓',
  0xfd: '☒',
  0xfe: '☑',
};

/** For a character in the private use area whose font is not one of the above: the bullets Word writes most. */
const PRIVATE_USE: Record<number, string> = { 0xb7: '•', 0xa7: '▪', 0xd8: '➢', 0xfc: '✓', 0x6f: '□', 0x6e: '■', 0x76: '❖' };

/** Which table a font name selects: 1 Symbol, 2 Wingdings, 0 none. */
export function symbolFont(family: string | undefined): 0 | 1 | 2 {
  if (family === undefined) return 0;
  const name = family.toLowerCase();
  if (name === 'symbol') return 1;
  if (name === 'wingdings') return 2;
  return 0;
}

/**
 * Unicode for one character of a symbol font, undefined when the tables do
 * not know it. Symbol fonts are addressed either directly (0x20..0xFF) or
 * through the private use area (0xF020..0xF0FF); both are accepted.
 */
export function symbolChar(kind: 0 | 1 | 2, code: number): string | undefined {
  const privateUse = code >= 0xf020 && code <= 0xf0ff;
  const low = privateUse ? code - 0xf000 : code;
  if (low < 0x20 || low > 0xff) return undefined;
  if (kind === 1) {
    if (low >= 0x61 && low <= 0x7a) return SYMBOL_LOWER.charAt(low - 0x61);
    if (low >= 0x41 && low <= 0x5a) return SYMBOL_UPPER.charAt(low - 0x41);
    return SYMBOL[low];
  }
  if (kind === 2) return WINGDINGS[low];
  return privateUse ? PRIVATE_USE[low] : undefined;
}

/** Maps the text of a run in a symbol font; `changed` says whether any character was replaced. */
export function mapSymbols(kind: 0 | 1 | 2, text: string): { text: string; changed: boolean } {
  let out = '';
  let changed = false;
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    // Outside a symbol font only the private use area is looked at.
    const mapped = kind !== 0 || (code >= 0xf020 && code <= 0xf0ff) ? symbolChar(kind, code) : undefined;
    if (mapped === undefined) out += text.charAt(i);
    else {
      out += mapped;
      changed = true;
    }
  }
  return { text: out, changed };
}

/** True when the text has a character the tables above might replace. */
export function mayHoldSymbols(kind: 0 | 1 | 2, text: string): boolean {
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

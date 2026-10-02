// Sprms ("single property modifiers", MS-DOC 2.6) and the structures their
// operands are made of: colours, borders, shading.
//
// A sprm is a 16-bit code followed by an operand. The code itself says how
// long the operand is (the three "spra" bits), so a list of sprms can be
// walked without knowing what each one means: an unknown sprm is passed over,
// and counted by the caller.

import type { Border, DocColor, Shading } from '../../model/doc/types';
import { u16, u32, u8 } from './bytes';

/** Kind of property a sprm changes (the "sgc" bits): 1 paragraph, 2 character, 3 picture, 4 section, 5 table. */
export const sprmKind = (op: number): number => (op >> 10) & 7;

export const SPRM_T_DEF_TABLE = 0xd608;
const SPRM_T_DEF_TABLE_10 = 0xd606;
const SPRM_P_CHG_TABS = 0xc615;

/**
 * Calls `fn` for every sprm in [start, end) of `b` with the code and the
 * place and length of the operand (for an operand of variable size: without
 * the size it starts with). Returns false when the list ends inside a sprm,
 * which the callers report; what came before was still handed out.
 */
export function eachSprm(b: Uint8Array, start: number, end: number, fn: (op: number, at: number, length: number) => void): boolean {
  if (end > b.length) end = b.length;
  let p = start;
  // Every round takes at least three bytes, so the loop is bounded by the length. A list in a PAPX may end with a
  // byte of padding, so what is too short to be a sprm at the end is not an error.
  while (p + 3 <= end) {
    const op = b[p] | (b[p + 1] << 8);
    p += 2;
    let at = p;
    let length: number;
    switch (op >> 13) {
      case 0:
      case 1:
        length = 1;
        break;
      case 2:
      case 4:
      case 5:
        length = 2;
        break;
      case 3:
        length = 4;
        break;
      case 7:
        length = 3;
        break;
      default:
        if (op === SPRM_T_DEF_TABLE || op === SPRM_T_DEF_TABLE_10) {
          // The size is two bytes and counts one byte more than follows.
          if (p + 2 > end) return false;
          length = u16(b, p) - 1;
          at = p + 2;
          if (length < 0) return false;
        } else if (op === SPRM_P_CHG_TABS && u8(b, p) === 255) {
          // Too long for a size of one byte: the length follows from the two counts inside.
          const deleted = u8(b, p + 1);
          const addedAt = p + 2 + deleted * 4;
          if (addedAt >= end) return false;
          length = 1 + deleted * 4 + 1 + u8(b, addedAt) * 3;
          at = p + 1;
        } else {
          length = u8(b, p);
          at = p + 1;
        }
    }
    if (at + length > end) return false;
    fn(op, at, length);
    p = at + length;
  }
  return true;
}

/** The sixteen colours of an "ico" (MS-DOC 2.9.119), index 1..16; 0 is automatic. */
const ICO_RGB = [
  '',
  '000000',
  '0000FF',
  '00FFFF',
  '00FF00',
  'FF00FF',
  'FF0000',
  'FFFF00',
  'FFFFFF',
  '000080',
  '008080',
  '008000',
  '800080',
  '800000',
  '808000',
  '808080',
  'C0C0C0',
];

/** Highlight colour names of the model for an ico. */
const ICO_NAME = [
  'none',
  'black',
  'blue',
  'cyan',
  'green',
  'magenta',
  'red',
  'yellow',
  'white',
  'darkBlue',
  'darkCyan',
  'darkGreen',
  'darkMagenta',
  'darkRed',
  'darkYellow',
  'darkGray',
  'lightGray',
];

const AUTO: DocColor = Object.freeze({ auto: true });

export function icoColor(ico: number): DocColor | undefined {
  if (ico === 0) return AUTO;
  const rgb = ICO_RGB[ico];
  return rgb ? { rgb } : undefined;
}

export const icoName = (ico: number): string | undefined => ICO_NAME[ico];

const HEX = '0123456789ABCDEF';
const hex2 = (value: number) => HEX.charAt((value >> 4) & 15) + HEX.charAt(value & 15);

/** A COLORREF (MS-DOC 2.9.43): red, green, blue and a byte that is 0xFF for the automatic colour. */
export function colorRef(b: Uint8Array, o: number): DocColor {
  if (u8(b, o + 3) === 0xff) return AUTO;
  return { rgb: hex2(u8(b, o)) + hex2(u8(b, o + 1)) + hex2(u8(b, o + 2)) };
}

/** Border styles by BrcType (MS-DOC 2.9.17), as the names of ST_Border. */
const BORDER_STYLES: Record<number, string> = {
  0: 'none',
  1: 'single',
  2: 'thick',
  3: 'double',
  5: 'single',
  6: 'dotted',
  7: 'dashed',
  8: 'dotDash',
  9: 'dotDotDash',
  10: 'triple',
  11: 'thinThickSmallGap',
  12: 'thickThinSmallGap',
  13: 'thinThickThinSmallGap',
  14: 'thinThickMediumGap',
  15: 'thickThinMediumGap',
  16: 'thinThickThinMediumGap',
  17: 'thinThickLargeGap',
  18: 'thickThinLargeGap',
  19: 'thinThickThinLargeGap',
  20: 'wave',
  21: 'doubleWave',
  22: 'dashSmallGap',
  23: 'dashDotStroked',
  24: 'threeDEmboss',
  25: 'threeDEngrave',
  26: 'outset',
  27: 'inset',
};

const NIL_BORDER: Border = Object.freeze({ style: 'nil' });

function makeBorder(type: number, width: number, space: number, color: DocColor | undefined, shadow: boolean, frame: boolean): Border {
  if (type === 0xff) return NIL_BORDER;
  const border: Border = { style: BORDER_STYLES[type] || 'single' };
  if (type !== 0) {
    border.size = width;
    border.space = space;
    if (color) border.color = color;
    if (shadow) border.shadow = true;
    if (frame) border.frame = true;
  }
  return border;
}

/**
 * A Brc80 (MS-DOC 2.9.18), the border of Word 97: width, type, colour index
 * and distance in four bytes. All ones is "no border"; undefined is returned
 * for all zeros when `zeroIsUnset` (cells: nothing is said, the border of the
 * table shows).
 */
export function brc80(b: Uint8Array, o: number, zeroIsUnset = false): Border | undefined {
  const value = u32(b, o);
  if (value === 0xffffffff) return NIL_BORDER;
  if (value === 0 && zeroIsUnset) return undefined;
  return makeBorder((value >>> 8) & 0xff, value & 0xff, (value >>> 24) & 0x1f, icoColor((value >>> 16) & 0xff), (value & 0x20000000) !== 0, (value & 0x40000000) !== 0);
}

/** A Brc (MS-DOC 2.9.16), the border of Word 2000 and later: a full colour, then width, type and distance. */
export function brc(b: Uint8Array, o: number, zeroIsUnset = false): Border | undefined {
  const low = u32(b, o);
  const high = u32(b, o + 4);
  if (low === 0xffffffff && high === 0xffffffff) return NIL_BORDER;
  const type = u8(b, o + 5);
  if (zeroIsUnset && high === 0 && (low === 0 || low === 0xff000000)) return undefined;
  const flags = u16(b, o + 6);
  return makeBorder(type, u8(b, o + 4), flags & 0x1f, colorRef(b, o), (flags & 0x20) !== 0, (flags & 0x40) !== 0);
}

/** Shading patterns by Ipat (MS-DOC 2.9.121), as the names of ST_Shd. */
const PATTERNS: Record<number, string> = {
  0: 'clear',
  1: 'solid',
  2: 'pct5',
  3: 'pct10',
  4: 'pct20',
  5: 'pct25',
  6: 'pct30',
  7: 'pct40',
  8: 'pct50',
  9: 'pct60',
  10: 'pct70',
  11: 'pct75',
  12: 'pct80',
  13: 'pct90',
  14: 'horzStripe',
  15: 'vertStripe',
  16: 'reverseDiagStripe',
  17: 'diagStripe',
  18: 'horzCross',
  19: 'diagCross',
  20: 'thinHorzStripe',
  21: 'thinVertStripe',
  22: 'thinReverseDiagStripe',
  23: 'thinDiagStripe',
  24: 'thinHorzCross',
  25: 'thinDiagCross',
  35: 'pct5',
  36: 'pct10',
  37: 'pct12',
  38: 'pct15',
  39: 'pct20',
  40: 'pct25',
  41: 'pct30',
  42: 'pct35',
  43: 'pct35',
  44: 'pct37',
  45: 'pct40',
  46: 'pct45',
  47: 'pct50',
  48: 'pct55',
  49: 'pct55',
  50: 'pct60',
  51: 'pct62',
  52: 'pct65',
  53: 'pct70',
  54: 'pct75',
  55: 'pct80',
  56: 'pct85',
  57: 'pct85',
  58: 'pct87',
  59: 'pct90',
  60: 'pct95',
  61: 'pct95',
  62: 'pct95',
};

function makeShading(pattern: number, fore: DocColor | undefined, back: DocColor | undefined): Shading {
  const shading: Shading = { pattern: PATTERNS[pattern] || 'clear' };
  if (fore) shading.color = fore;
  if (back) shading.fill = back;
  return shading;
}

/** A Shd80 (MS-DOC 2.9.249): two colour indexes and a pattern in two bytes; undefined for the "nil" value. */
export function shd80(b: Uint8Array, o: number): Shading | undefined {
  const value = u16(b, o);
  if (value === 0xffff) return undefined;
  return makeShading(value >> 10, icoColor(value & 0x1f), icoColor((value >> 5) & 0x1f));
}

/** A SHD (MS-DOC 2.9.248): pattern colour, background colour, pattern; undefined for the "nil" pattern. */
export function shd(b: Uint8Array, o: number): Shading | undefined {
  const pattern = u16(b, o + 8);
  if (pattern === 0xffff) return undefined;
  return makeShading(pattern, colorRef(b, o), colorRef(b, o + 4));
}

/**
 * Number formats by MSONFC (MS-OSHARED 2.2.1.3), as the names of ST_NumberFormat: 0 to 59 in the
 * order of the enumeration. 40 (a fourth Chinese format) has no name there.
 */
const NUMBER_FORMATS: readonly (string | undefined)[] = [
  'decimal',
  'upperRoman',
  'lowerRoman',
  'upperLetter',
  'lowerLetter',
  'ordinal',
  'cardinalText',
  'ordinalText',
  'hex',
  'chicago',
  'ideographDigital',
  'japaneseCounting',
  'aiueo',
  'iroha',
  'decimalFullWidth',
  'decimalHalfWidth',
  'japaneseLegal',
  'japaneseDigitalTenThousand',
  'decimalEnclosedCircle',
  'decimalFullWidth2',
  'aiueoFullWidth',
  'irohaFullWidth',
  'decimalZero',
  'bullet',
  'ganada',
  'chosung',
  'decimalEnclosedFullstop',
  'decimalEnclosedParen',
  'decimalEnclosedCircleChinese',
  'ideographEnclosedCircle',
  'ideographTraditional',
  'ideographZodiac',
  'ideographZodiacTraditional',
  'taiwaneseCounting',
  'ideographLegalTraditional',
  'taiwaneseCountingThousand',
  'taiwaneseDigital',
  'chineseCounting',
  'chineseLegalSimplified',
  'chineseCountingThousand',
  undefined,
  'koreanDigital',
  'koreanCounting',
  'koreanLegal',
  'koreanDigital2',
  'hebrew1',
  'arabicAlpha',
  'hebrew2',
  'arabicAbjad',
  'hindiVowels',
  'hindiConsonants',
  'hindiNumbers',
  'hindiCounting',
  'thaiLetters',
  'thaiNumbers',
  'thaiCounting',
  'vietnameseCounting',
  'numberInDash',
  'russianLower',
  'russianUpper',
];

/** The number format for an MSONFC value; "none" for 255, "decimal" for a value without a name. */
export const numberFormat = (nfc: number): string => (nfc === 0xff ? 'none' : NUMBER_FORMATS[nfc] || 'decimal');

/** Language tags for the language ids documents of this kind carry most; others are written as "x-lcid-<hex>". */
const LANGUAGES: Record<number, string> = {
  0x0401: 'ar-SA',
  0x0402: 'bg-BG',
  0x0404: 'zh-TW',
  0x0405: 'cs-CZ',
  0x0406: 'da-DK',
  0x0407: 'de-DE',
  0x0408: 'el-GR',
  0x0409: 'en-US',
  0x040a: 'es-ES',
  0x040b: 'fi-FI',
  0x040c: 'fr-FR',
  0x040d: 'he-IL',
  0x040e: 'hu-HU',
  0x0410: 'it-IT',
  0x0411: 'ja-JP',
  0x0412: 'ko-KR',
  0x0413: 'nl-NL',
  0x0414: 'nb-NO',
  0x0415: 'pl-PL',
  0x0416: 'pt-BR',
  0x0418: 'ro-RO',
  0x0419: 'ru-RU',
  0x041a: 'hr-HR',
  0x041b: 'sk-SK',
  0x041d: 'sv-SE',
  0x041f: 'tr-TR',
  0x0422: 'uk-UA',
  0x0429: 'fa-IR',
  0x042c: 'az-Latn-AZ',
  0x0804: 'zh-CN',
  0x0807: 'de-CH',
  0x0809: 'en-GB',
  0x080a: 'es-MX',
  0x080c: 'fr-BE',
  0x0816: 'pt-PT',
  0x0c09: 'en-AU',
  0x0c0a: 'es-ES',
  0x1009: 'en-CA',
};

/** Language tag for a language id; undefined for "no language" (0, 0x0400). */
export function languageTag(lid: number): string | undefined {
  if (lid === 0 || lid === 0x0400) return undefined;
  return LANGUAGES[lid] || 'x-lcid-' + lid.toString(16);
}

/** A DTTM (MS-DOC 2.9.71) as ISO 8601 text; undefined when it holds no date. */
export function dttm(value: number): string | undefined {
  if (value === 0) return undefined;
  const minute = value & 0x3f;
  const hour = (value >>> 6) & 0x1f;
  const day = (value >>> 11) & 0x1f;
  const month = (value >>> 16) & 0x0f;
  const year = ((value >>> 20) & 0x1ff) + 1900;
  if (month < 1 || month > 12 || day < 1 || hour > 23 || minute > 59) return undefined;
  const two = (n: number) => (n < 10 ? '0' : '') + n;
  return `${year}-${two(month)}-${two(day)}T${two(hour)}:${two(minute)}:00Z`;
}

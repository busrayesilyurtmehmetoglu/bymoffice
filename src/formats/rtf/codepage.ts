// Code pages of RTF text.
//
// An RTF file is bytes of 7 bits; a character beyond that is written as
// \'hh, a byte of the code page of the font the text is in (\fcharset of the
// font, else \ansicpg of the document), or as \uN with such bytes after it
// for readers that know no Unicode.
//
// The single-byte tables of the xls reader are used where it has them
// (Windows-1250 to -1254, 437, 850); the ones below add what the fonts of a
// Word document name besides: Hebrew, Arabic, Baltic, Vietnamese, Thai and
// the Macintosh set. Each is the upper half (0x80..0xFF) of its code page.
//
// The double-byte code pages (Japanese, Chinese, Korean) have no table here:
// they would be larger than the rest of the reader. Word writes \uN for every
// such character, which needs no table; for a file that has only the bytes,
// the decoder of the host is asked (TextDecoder, where there is one), and a
// character it cannot give reads as U+FFFD. Two bytes are one character
// either way.

import { decodeBytes } from '../xls/codepage';

const CP874 =
  '\u20AC\u0081\u0082\u0083\u0084\u2026\u0086\u0087\u0088\u0089\u008A\u008B\u008C\u008D\u008E\u008F' +
  '\u0090\u2018\u2019\u201C\u201D\u2022\u2013\u2014\u0098\u0099\u009A\u009B\u009C\u009D\u009E\u009F' +
  '\u00A0\u0E01\u0E02\u0E03\u0E04\u0E05\u0E06\u0E07\u0E08\u0E09\u0E0A\u0E0B\u0E0C\u0E0D\u0E0E\u0E0F' +
  '\u0E10\u0E11\u0E12\u0E13\u0E14\u0E15\u0E16\u0E17\u0E18\u0E19\u0E1A\u0E1B\u0E1C\u0E1D\u0E1E\u0E1F' +
  '\u0E20\u0E21\u0E22\u0E23\u0E24\u0E25\u0E26\u0E27\u0E28\u0E29\u0E2A\u0E2B\u0E2C\u0E2D\u0E2E\u0E2F' +
  '\u0E30\u0E31\u0E32\u0E33\u0E34\u0E35\u0E36\u0E37\u0E38\u0E39\u0E3A\uF8C1\uF8C2\uF8C3\uF8C4\u0E3F' +
  '\u0E40\u0E41\u0E42\u0E43\u0E44\u0E45\u0E46\u0E47\u0E48\u0E49\u0E4A\u0E4B\u0E4C\u0E4D\u0E4E\u0E4F' +
  '\u0E50\u0E51\u0E52\u0E53\u0E54\u0E55\u0E56\u0E57\u0E58\u0E59\u0E5A\u0E5B\uF8C5\uF8C6\uF8C7\uF8C8';

const CP1255 =
  '\u20AC\u0081\u201A\u0192\u201E\u2026\u2020\u2021\u02C6\u2030\u008A\u2039\u008C\u008D\u008E\u008F' +
  '\u0090\u2018\u2019\u201C\u201D\u2022\u2013\u2014\u02DC\u2122\u009A\u203A\u009C\u009D\u009E\u009F' +
  '\u00A0\u00A1\u00A2\u00A3\u20AA\u00A5\u00A6\u00A7\u00A8\u00A9\u00D7\u00AB\u00AC\u00AD\u00AE\u00AF' +
  '\u00B0\u00B1\u00B2\u00B3\u00B4\u00B5\u00B6\u00B7\u00B8\u00B9\u00F7\u00BB\u00BC\u00BD\u00BE\u00BF' +
  '\u05B0\u05B1\u05B2\u05B3\u05B4\u05B5\u05B6\u05B7\u05B8\u05B9\uFFFD\u05BB\u05BC\u05BD\u05BE\u05BF' +
  '\u05C0\u05C1\u05C2\u05C3\u05F0\u05F1\u05F2\u05F3\u05F4\uFFFD\uFFFD\uFFFD\uFFFD\uFFFD\uFFFD\uFFFD' +
  '\u05D0\u05D1\u05D2\u05D3\u05D4\u05D5\u05D6\u05D7\u05D8\u05D9\u05DA\u05DB\u05DC\u05DD\u05DE\u05DF' +
  '\u05E0\u05E1\u05E2\u05E3\u05E4\u05E5\u05E6\u05E7\u05E8\u05E9\u05EA\uFFFD\uFFFD\u200E\u200F\uFFFD';

const CP1256 =
  '\u20AC\u067E\u201A\u0192\u201E\u2026\u2020\u2021\u02C6\u2030\u0679\u2039\u0152\u0686\u0698\u0688' +
  '\u06AF\u2018\u2019\u201C\u201D\u2022\u2013\u2014\u06A9\u2122\u0691\u203A\u0153\u200C\u200D\u06BA' +
  '\u00A0\u060C\u00A2\u00A3\u00A4\u00A5\u00A6\u00A7\u00A8\u00A9\u06BE\u00AB\u00AC\u00AD\u00AE\u00AF' +
  '\u00B0\u00B1\u00B2\u00B3\u00B4\u00B5\u00B6\u00B7\u00B8\u00B9\u061B\u00BB\u00BC\u00BD\u00BE\u061F' +
  '\u06C1\u0621\u0622\u0623\u0624\u0625\u0626\u0627\u0628\u0629\u062A\u062B\u062C\u062D\u062E\u062F' +
  '\u0630\u0631\u0632\u0633\u0634\u0635\u0636\u00D7\u0637\u0638\u0639\u063A\u0640\u0641\u0642\u0643' +
  '\u00E0\u0644\u00E2\u0645\u0646\u0647\u0648\u00E7\u00E8\u00E9\u00EA\u00EB\u0649\u064A\u00EE\u00EF' +
  '\u064B\u064C\u064D\u064E\u00F4\u064F\u0650\u00F7\u0651\u00F9\u0652\u00FB\u00FC\u200E\u200F\u06D2';

const CP1257 =
  '\u20AC\u0081\u201A\u0083\u201E\u2026\u2020\u2021\u0088\u2030\u008A\u2039\u008C\u00A8\u02C7\u00B8' +
  '\u0090\u2018\u2019\u201C\u201D\u2022\u2013\u2014\u0098\u2122\u009A\u203A\u009C\u00AF\u02DB\u009F' +
  '\u00A0\uFFFD\u00A2\u00A3\u00A4\uFFFD\u00A6\u00A7\u00D8\u00A9\u0156\u00AB\u00AC\u00AD\u00AE\u00C6' +
  '\u00B0\u00B1\u00B2\u00B3\u00B4\u00B5\u00B6\u00B7\u00F8\u00B9\u0157\u00BB\u00BC\u00BD\u00BE\u00E6' +
  '\u0104\u012E\u0100\u0106\u00C4\u00C5\u0118\u0112\u010C\u00C9\u0179\u0116\u0122\u0136\u012A\u013B' +
  '\u0160\u0143\u0145\u00D3\u014C\u00D5\u00D6\u00D7\u0172\u0141\u015A\u016A\u00DC\u017B\u017D\u00DF' +
  '\u0105\u012F\u0101\u0107\u00E4\u00E5\u0119\u0113\u010D\u00E9\u017A\u0117\u0123\u0137\u012B\u013C' +
  '\u0161\u0144\u0146\u00F3\u014D\u00F5\u00F6\u00F7\u0173\u0142\u015B\u016B\u00FC\u017C\u017E\u02D9';

const CP1258 =
  '\u20AC\u0081\u201A\u0192\u201E\u2026\u2020\u2021\u02C6\u2030\u008A\u2039\u0152\u008D\u008E\u008F' +
  '\u0090\u2018\u2019\u201C\u201D\u2022\u2013\u2014\u02DC\u2122\u009A\u203A\u0153\u009D\u009E\u0178' +
  '\u00A0\u00A1\u00A2\u00A3\u00A4\u00A5\u00A6\u00A7\u00A8\u00A9\u00AA\u00AB\u00AC\u00AD\u00AE\u00AF' +
  '\u00B0\u00B1\u00B2\u00B3\u00B4\u00B5\u00B6\u00B7\u00B8\u00B9\u00BA\u00BB\u00BC\u00BD\u00BE\u00BF' +
  '\u00C0\u00C1\u00C2\u0102\u00C4\u00C5\u00C6\u00C7\u00C8\u00C9\u00CA\u00CB\u0300\u00CD\u00CE\u00CF' +
  '\u0110\u00D1\u0309\u00D3\u00D4\u01A0\u00D6\u00D7\u00D8\u00D9\u00DA\u00DB\u00DC\u01AF\u0303\u00DF' +
  '\u00E0\u00E1\u00E2\u0103\u00E4\u00E5\u00E6\u00E7\u00E8\u00E9\u00EA\u00EB\u0301\u00ED\u00EE\u00EF' +
  '\u0111\u00F1\u0323\u00F3\u00F4\u01A1\u00F6\u00F7\u00F8\u00F9\u00FA\u00FB\u00FC\u01B0\u20AB\u00FF';

const CP10000 =
  '\u00C4\u00C5\u00C7\u00C9\u00D1\u00D6\u00DC\u00E1\u00E0\u00E2\u00E4\u00E3\u00E5\u00E7\u00E9\u00E8' +
  '\u00EA\u00EB\u00ED\u00EC\u00EE\u00EF\u00F1\u00F3\u00F2\u00F4\u00F6\u00F5\u00FA\u00F9\u00FB\u00FC' +
  '\u2020\u00B0\u00A2\u00A3\u00A7\u2022\u00B6\u00DF\u00AE\u00A9\u2122\u00B4\u00A8\u2260\u00C6\u00D8' +
  '\u221E\u00B1\u2264\u2265\u00A5\u00B5\u2202\u2211\u220F\u03C0\u222B\u00AA\u00BA\u03A9\u00E6\u00F8' +
  '\u00BF\u00A1\u00AC\u221A\u0192\u2248\u2206\u00AB\u00BB\u2026\u00A0\u00C0\u00C3\u00D5\u0152\u0153' +
  '\u2013\u2014\u201C\u201D\u2018\u2019\u00F7\u25CA\u00FF\u0178\u2044\u20AC\u2039\u203A\uFB01\uFB02' +
  '\u2021\u00B7\u201A\u201E\u2030\u00C2\u00CA\u00C1\u00CB\u00C8\u00CD\u00CE\u00CF\u00CC\u00D3\u00D4' +
  '\uF8FF\u00D2\u00DA\u00DB\u00D9\u0131\u02C6\u02DC\u00AF\u02D8\u02D9\u02DA\u00B8\u02DD\u02DB\u02C7';

const OWN = new Map<number, string>([
  [874, CP874],
  [1255, CP1255],
  [1256, CP1256],
  [1257, CP1257],
  [1258, CP1258],
  [10000, CP10000],
]);

/** Code pages decodeBytes() of the xls reader has a table for. */
const SHARED = [1250, 1251, 1252, 1253, 1254, 437, 850];

const tables = new Map<number, string>();

/** The 128 characters of the upper half of a single-byte code page; those of Windows-1252 for one without a table. */
export function highTable(codepage: number): string {
  let table = tables.get(codepage);
  if (table !== undefined) return table;
  table = OWN.get(codepage);
  if (table === undefined) {
    const bytes = new Uint8Array(128);
    for (let i = 0; i < 128; i++) bytes[i] = 0x80 + i;
    table = decodeBytes(bytes, 0, 128, SHARED.indexOf(codepage) >= 0 ? codepage : 1252);
  }
  tables.set(codepage, table);
  return table;
}

/** Whether text in the code page is read with a table of its own (or as UTF-8); others are read as Windows-1252. */
export function isKnownRtfCodepage(codepage: number): boolean {
  return codepage === 65001 || OWN.has(codepage) || SHARED.indexOf(codepage) >= 0 || isDoubleByte(codepage);
}

/** The pseudo code page of a font of symbols (\fcharset2): a byte is a glyph of the font, kept in the private use area. */
export const SYMBOL_CODEPAGE = 42;

/**
 * Code page for the character set of a font (\fcharsetN); undefined for
 * "the default" (1) and for one that is not known, which both mean the code
 * page of the document.
 *
 * Character set 0 ("ANSI") is taken as the code page of the document as
 * well, not as Windows-1252: writers that never state another character set
 * put their text in the code page they declare with \ansicpg.
 */
export function charsetCodepage(charset: number): number | undefined {
  switch (charset) {
    case 2:
      return SYMBOL_CODEPAGE;
    case 77:
      return 10000;
    case 128:
      return 932;
    case 129:
      return 949;
    case 130:
      return 1361;
    case 134:
      return 936;
    case 136:
      return 950;
    case 161:
      return 1253;
    case 162:
      return 1254;
    case 163:
      return 1258;
    case 177:
      return 1255;
    case 178:
      return 1256;
    case 186:
      return 1257;
    case 204:
      return 1251;
    case 222:
      return 874;
    case 238:
      return 1250;
    case 254:
      return 437;
    case 255:
      return 850;
  }
  return undefined;
}

export function isDoubleByte(codepage: number): boolean {
  return codepage === 932 || codepage === 936 || codepage === 949 || codepage === 950 || codepage === 1361;
}

/** Whether a byte is the first of a two-byte character in a double-byte code page. */
export function isLeadByte(codepage: number, byte: number): boolean {
  switch (codepage) {
    case 932:
      return (byte >= 0x81 && byte <= 0x9f) || (byte >= 0xe0 && byte <= 0xfc);
    case 936:
    case 949:
    case 950:
      return byte >= 0x81 && byte <= 0xfe;
    case 1361:
      return (byte >= 0x84 && byte <= 0xd3) || (byte >= 0xd8 && byte <= 0xde) || (byte >= 0xe0 && byte <= 0xf9);
  }
  return false;
}

interface HostDecoder {
  decode(bytes: Uint8Array): string;
}

const LABELS: { [codepage: number]: string } = { 932: 'shift_jis', 936: 'gbk', 949: 'euc-kr', 950: 'big5', 1361: 'euc-kr' };
const hostDecoders = new Map<number, HostDecoder | null>();
const pair = new Uint8Array(2);

/** The decoder of the host for a double-byte code page; null where the host has none. */
function hostDecoder(codepage: number): HostDecoder | null {
  let decoder = hostDecoders.get(codepage);
  if (decoder !== undefined) return decoder;
  decoder = null;
  const label = LABELS[codepage];
  // Not every host has a TextDecoder, and not every TextDecoder knows these.
  const Ctor = typeof TextDecoder === 'undefined' ? undefined : TextDecoder;
  if (Ctor && label !== undefined && codepage !== 1361) {
    try {
      decoder = new Ctor(label);
    } catch (error) {
      decoder = null;
    }
  }
  hostDecoders.set(codepage, decoder);
  return decoder;
}

/** The character two bytes of a double-byte code page stand for; U+FFFD when the host cannot tell. */
export function decodePair(codepage: number, lead: number, trail: number): string {
  const decoder = hostDecoder(codepage);
  if (!decoder) return '\uFFFD';
  pair[0] = lead;
  pair[1] = trail;
  const text = decoder.decode(pair);
  // One character (or a surrogate pair); anything else means the bytes were not one.
  return text.length === 1 || (text.length === 2 && text.charCodeAt(0) >= 0xd800 && text.charCodeAt(0) < 0xdc00) ? text : '\uFFFD';
}

/** The character of one byte in a single-byte code page (the first half of a double-byte one included). */
export function decodeByte(codepage: number, byte: number): string {
  // A space is a space in every font: lines break at it.
  if (codepage === SYMBOL_CODEPAGE) return String.fromCharCode(byte <= 0x20 ? byte : 0xf000 + byte);
  if (byte < 0x80) return String.fromCharCode(byte);
  if (isDoubleByte(codepage)) {
    // Half-width katakana are single bytes in the Japanese code page.
    if (codepage === 932 && byte >= 0xa1 && byte <= 0xdf) return String.fromCharCode(0xff61 + byte - 0xa1);
    return '\uFFFD';
  }
  return highTable(codepage).charAt(byte - 0x80);
}

/** Text of bytes[start..end) in a code page, two-byte characters included. */
export function decodeRange(bytes: Uint8Array, start: number, end: number, codepage: number): string {
  if (codepage === 65001) return decodeBytes(bytes, start, end, 65001);
  let out = '';
  const double = isDoubleByte(codepage);
  for (let i = start; i < end; i++) {
    const b = bytes[i];
    if (double && i + 1 < end && isLeadByte(codepage, b)) out += decodePair(codepage, b, bytes[++i]);
    else out += decodeByte(codepage, b);
  }
  return out;
}

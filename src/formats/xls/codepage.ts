// Code pages for the byte strings of BIFF5 files.
//
// The tables are carried here so that reading does not depend on what a
// platform's TextDecoder happens to support. Each table is the upper half
// (0x80..0xFF) of a single-byte code page; the lower half is ASCII.

/** 0xA0..0xFF of ISO 8859-1, which Windows-1252 and -1254 share. */
function latin1(from: number): string {
  let out = '';
  for (let code = from; code <= 0xff; code++) out += String.fromCharCode(code);
  return out;
}

const CP1252_LOW =
  '€\u0081‚ƒ„…†‡ˆ‰Š‹Œ\u008DŽ\u008F' +
  '\u0090‘’“”•–—˜™š›œ\u009DžŸ';

const CP1252 = CP1252_LOW + latin1(0xa0);

/** Turkish: Windows-1252 without the caron letters, and six letters replaced. */
const CP1254 = (
  CP1252_LOW.replace('Ž', '\u008E').replace('ž', '\u009E') + latin1(0xa0)
)
  .replace('Ð', 'Ğ')
  .replace('Ý', 'İ')
  .replace('Þ', 'Ş')
  .replace('ð', 'ğ')
  .replace('ý', 'ı')
  .replace('þ', 'ş');

const CP1250 =
  '€\u0081‚\u0083„…†‡\u0088‰Š‹ŚŤŽŹ' +
  '\u0090‘’“”•–—\u0098™š›śťžź' +
  ' ˇ˘Ł¤Ą¦§¨©Ş«¬­®Ż' +
  '°±˛ł´µ¶·¸ąş»Ľ˝ľż' +
  'ŔÁÂĂÄĹĆÇČÉĘËĚÍÎĎ' +
  'ĐŃŇÓÔŐÖ×ŘŮÚŰÜÝŢß' +
  'ŕáâăäĺćçčéęëěíîď' +
  'đńňóôőö÷řůúűüýţ˙';

function cyrillic(): string {
  let out =
    'ЂЃ‚ѓ„…†‡€‰Љ‹ЊЌЋЏ' +
    'ђ‘’“”•–—\u0098™љ›њќћџ' +
    ' ЎўЈ¤Ґ¦§Ё©Є«¬­®Ї' +
    '°±Ііґµ¶·ё№є»јЅѕї';
  for (let code = 0x0410; code <= 0x044f; code++) out += String.fromCharCode(code);
  return out;
}

function greek(): string {
  let out =
    '€\u0081‚ƒ„…†‡\u0088‰\u008A‹\u008C\u008D\u008E\u008F' +
    '\u0090‘’“”•–—\u0098™\u009A›\u009C\u009D\u009E\u009F' +
    ' ΅Ά£¤¥¦§¨©�«¬­®―' +
    '°±²³΄µ¶·ΈΉΊ»Ό½ΎΏ';
  // 0xC0..0xFF follow the Greek block, which has no letter at U+03A2 and none after U+03CE.
  for (let byte = 0xc0; byte <= 0xff; byte++) {
    const code = 0x0390 + (byte - 0xc0);
    out += code === 0x03a2 || code > 0x03ce ? '�' : String.fromCharCode(code);
  }
  return out;
}

const CP437 =
  'ÇüéâäàåçêëèïîìÄÅ' +
  'ÉæÆôöòûùÿÖÜ¢£¥₧ƒ' +
  'áíóúñÑªº¿⌐¬½¼¡«»' +
  '░▒▓│┤╡╢╖╕╣║╗╝╜╛┐' +
  '└┴┬├─┼╞╟╚╔╩╦╠═╬╧' +
  '╨╤╥╙╘╒╓╫╪┘┌█▄▌▐▀' +
  'αßΓπΣσµτΦΘΩδ∞φε∩' +
  '≡±≥≤⌠⌡÷≈°∙·√ⁿ²■ ';

const CP850 =
  'ÇüéâäàåçêëèïîìÄÅ' +
  'ÉæÆôöòûùÿÖÜø£Ø×ƒ' +
  'áíóúñÑªº¿®¬½¼¡«»' +
  '░▒▓│┤ÁÂÀ©╣║╗╝¢¥┐' +
  '└┴┬├─┼ãÃ╚╔╩╦╠═╬¤' +
  'ðÐÊËÈıÍÎÏ┘┌█▄¦Ì▀' +
  'ÓßÔÒõÕµþÞÚÛÙýÝ¯´' +
  '­±‗¾¶§÷¸°¨·¹³²■ ';

const tables = new Map<number, string>();

function tableOf(codepage: number): string | undefined {
  let table = tables.get(codepage);
  if (table !== undefined) return table;
  switch (codepage) {
    case 1252:
    case 0x8001: // "ANSI" as BIFF2-5 files state it
      table = CP1252;
      break;
    case 1254:
      table = CP1254;
      break;
    case 1250:
      table = CP1250;
      break;
    case 1251:
      table = cyrillic();
      break;
    case 1253:
      table = greek();
      break;
    case 437:
      table = CP437;
      break;
    case 850:
      table = CP850;
      break;
    default:
      return undefined;
  }
  tables.set(codepage, table);
  return table;
}

/** Whether decodeBytes() has a table for the code page; others are read as Windows-1252. */
export function isKnownCodepage(codepage: number): boolean {
  return codepage === 65001 || codepage === 1200 || tableOf(codepage) !== undefined;
}

function decodeUtf8(bytes: Uint8Array, start: number, end: number): string {
  let out = '';
  for (let i = start; i < end; ) {
    const b = bytes[i++];
    let code: number;
    let more: number;
    if (b < 0x80) {
      out += String.fromCharCode(b);
      continue;
    } else if (b >= 0xc2 && b < 0xe0) (code = b & 0x1f), (more = 1);
    else if (b >= 0xe0 && b < 0xf0) (code = b & 0x0f), (more = 2);
    else if (b >= 0xf0 && b < 0xf5) (code = b & 0x07), (more = 3);
    else {
      out += '�';
      continue;
    }
    let ok = i + more <= end;
    for (let k = 0; ok && k < more; k++) {
      const next = bytes[i + k];
      if ((next & 0xc0) !== 0x80) ok = false;
      else code = (code << 6) | (next & 0x3f);
    }
    // Overlong forms, surrogates and values past U+10FFFF are not characters.
    if (ok && ((more === 2 && (code < 0x800 || (code >= 0xd800 && code < 0xe000))) || (more === 3 && (code < 0x10000 || code > 0x10ffff)))) ok = false;
    if (!ok) {
      out += '�';
      continue;
    }
    i += more;
    if (code >= 0x10000) {
      code -= 0x10000;
      out += String.fromCharCode(0xd800 + (code >> 10), 0xdc00 + (code & 0x3ff));
    } else out += String.fromCharCode(code);
  }
  return out;
}

/** Text of bytes[start..end) in a code page. A code page without a table reads as Windows-1252. */
export function decodeBytes(bytes: Uint8Array, start: number, end: number, codepage: number): string {
  if (codepage === 65001) return decodeUtf8(bytes, start, end);
  const table = tableOf(codepage) || CP1252;
  let out = '';
  for (let i = start; i < end; i++) {
    const b = bytes[i];
    out += b < 0x80 ? String.fromCharCode(b) : table.charAt(b - 0x80);
  }
  return out;
}

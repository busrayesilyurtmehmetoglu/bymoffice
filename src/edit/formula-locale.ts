// Formulas as they are typed in a locale, and as they are stored.
//
// A file always holds a formula in one notation: "," between arguments, "."
// in numbers, English function names. Excel shows and takes the same formula
// in the notation of its language: in Turkish (and wherever the decimal
// separator is a comma) arguments are separated by ";" and numbers are
// written "1,5". The two functions here translate between the two, token by
// token: text in quotes, sheet names in apostrophes and anything in brackets
// is never touched.
//
// Nothing here parses a formula or decides whether it is one: text that makes
// no sense comes out as it went in, and the parser has the last word.

import type { FormatLocale } from '../format/locale';

/** Longer text is not a formula (Excel's limit is 8,192 characters) and is left alone. */
const MAX_LENGTH = 32768;

/** Error constants, which hold characters ("/", "0", "?") that mean something else outside them. */
const ERRORS = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#SPILL!', '#CALC!', '#GETTING_DATA', '#FIELD!', '#BLOCKED!', '#CONNECT!', '#UNKNOWN!', '#BUSY!'];

/**
 * Turkish names of common functions, as the Turkish Excel has them, with the names
 * files store. Only read when a formula is typed or pasted; what is shown always has
 * the English names.
 *
 * UNSURE: taken from memory of the Turkish user interface, not checked against an
 * installation: BAĞ_DEĞ_SAY, BAĞ_DEĞ_DOLU_SAY, BOŞLUKSAY, ÇOKEĞERSAY, EĞERORTALAMA,
 * ÇOKEĞERORTALAMA, METNEÇEVİR, SAYIYAÇEVİR, YERİNEKOY, DEĞİŞTİR, MBUL, EHATALIYSA,
 * EBOŞSA, ESAYIYSA, EMETİNSE, EĞERYOKSA, HAFTANINGÜNÜ, SERİAY, SERİTARİH, İŞGÜNÜ,
 * TAMİŞGÜNÜ, NSAT, TAVANAYUVARLA, TABANAYUVARLA, S_SAYI_ÜRET, RASTGELEARADA, ALTTOPLAM,
 * SATIRSAY, SÜTUNSAY, KAYDIR, DOLAYLI, ÖZDEŞ, YİNELE, DAMGA.
 */
const TURKISH_FUNCTIONS: [string, string][] = [
  ['TOPLA', 'SUM'],
  ['EĞER', 'IF'],
  ['ORTALAMA', 'AVERAGE'],
  ['DÜŞEYARA', 'VLOOKUP'],
  ['YATAYARA', 'HLOOKUP'],
  ['ARA', 'LOOKUP'],
  ['EĞERSAY', 'COUNTIF'],
  ['ÇOKEĞERSAY', 'COUNTIFS'],
  ['ETOPLA', 'SUMIF'],
  ['ÇOKETOPLA', 'SUMIFS'],
  ['EĞERORTALAMA', 'AVERAGEIF'],
  ['ÇOKEĞERORTALAMA', 'AVERAGEIFS'],
  ['EĞERHATA', 'IFERROR'],
  ['EĞERYOKSA', 'IFNA'],
  ['BUGÜN', 'TODAY'],
  ['ŞİMDİ', 'NOW'],
  ['MAK', 'MAX'],
  ['MİN', 'MIN'],
  ['BÜYÜK', 'LARGE'],
  ['KÜÇÜK', 'SMALL'],
  ['ORTANCA', 'MEDIAN'],
  ['YUVARLA', 'ROUND'],
  ['YUKARIYUVARLA', 'ROUNDUP'],
  ['AŞAĞIYUVARLA', 'ROUNDDOWN'],
  ['TAVANAYUVARLA', 'CEILING'],
  ['TABANAYUVARLA', 'FLOOR'],
  ['NSAT', 'TRUNC'],
  ['TAMSAYI', 'INT'],
  ['MUTLAK', 'ABS'],
  ['MOD', 'MOD'],
  ['KUVVET', 'POWER'],
  ['KAREKÖK', 'SQRT'],
  ['ÇARPIM', 'PRODUCT'],
  ['TOPLA.ÇARPIM', 'SUMPRODUCT'],
  ['ALTTOPLAM', 'SUBTOTAL'],
  ['S_SAYI_ÜRET', 'RAND'],
  ['RASTGELEARADA', 'RANDBETWEEN'],
  ['BAĞ_DEĞ_SAY', 'COUNT'],
  ['BAĞ_DEĞ_DOLU_SAY', 'COUNTA'],
  ['BOŞLUKSAY', 'COUNTBLANK'],
  ['BİRLEŞTİR', 'CONCATENATE'],
  ['UZUNLUK', 'LEN'],
  ['SOLDAN', 'LEFT'],
  ['SAĞDAN', 'RIGHT'],
  ['PARÇAAL', 'MID'],
  ['BUL', 'FIND'],
  ['MBUL', 'SEARCH'],
  ['DEĞİŞTİR', 'REPLACE'],
  ['YERİNEKOY', 'SUBSTITUTE'],
  ['KIRP', 'TRIM'],
  ['BÜYÜKHARF', 'UPPER'],
  ['KÜÇÜKHARF', 'LOWER'],
  ['ÖZDEŞ', 'EXACT'],
  ['YİNELE', 'REPT'],
  ['DAMGA', 'CHAR'],
  ['METNEÇEVİR', 'TEXT'],
  ['SAYIYAÇEVİR', 'VALUE'],
  ['KAÇINCI', 'MATCH'],
  ['İNDİS', 'INDEX'],
  ['KAYDIR', 'OFFSET'],
  ['DOLAYLI', 'INDIRECT'],
  ['SATIR', 'ROW'],
  ['SÜTUN', 'COLUMN'],
  ['SATIRSAY', 'ROWS'],
  ['SÜTUNSAY', 'COLUMNS'],
  ['VE', 'AND'],
  ['YADA', 'OR'],
  ['DEĞİL', 'NOT'],
  ['DOĞRU', 'TRUE'],
  ['YANLIŞ', 'FALSE'],
  ['EHATALIYSA', 'ISERROR'],
  ['EBOŞSA', 'ISBLANK'],
  ['ESAYIYSA', 'ISNUMBER'],
  ['EMETİNSE', 'ISTEXT'],
  ['TARİH', 'DATE'],
  ['ZAMAN', 'TIME'],
  ['GÜN', 'DAY'],
  ['AY', 'MONTH'],
  ['YIL', 'YEAR'],
  ['SAAT', 'HOUR'],
  ['DAKİKA', 'MINUTE'],
  ['SANİYE', 'SECOND'],
  ['HAFTANINGÜNÜ', 'WEEKDAY'],
  ['SERİAY', 'EOMONTH'],
  ['SERİTARİH', 'EDATE'],
  ['İŞGÜNÜ', 'WORKDAY'],
  ['TAMİŞGÜNÜ', 'NETWORKDAYS'],
];

/** Upper case that treats the four Turkish "i" letters as one, so that "indis", "INDIS" and "İNDİS" are the same name. */
function nameKey(text: string): string {
  return text.replace(/[iıİ]/g, 'I').toUpperCase();
}

let turkishNames: Map<string, string> | undefined;

function turkishName(word: string): string | undefined {
  if (!turkishNames) {
    turkishNames = new Map();
    for (const [local, stored] of TURKISH_FUNCTIONS) turkishNames.set(nameKey(local), stored);
  }
  return turkishNames.get(nameKey(word));
}

const isDigit = (c: number) => c >= 48 && c <= 57;

/** Characters of names and references, as the formula parser has them. */
function isWordChar(c: number): boolean {
  if (c >= 128) return true;
  if (c >= 97) return c <= 122;
  if (c >= 65) return c <= 90 || c === 95 || c === 92;
  if (c >= 48) return c <= 57 || c === 63;
  return c === 36 || c === 46;
}

function isWordStart(c: number): boolean {
  if (c >= 128) return true;
  if (c >= 97) return c <= 122;
  return (c >= 65 && c <= 90) || c === 95 || c === 36;
}

/** Whether the locale writes formulas differently from the way they are stored. */
function usesComma(locale: FormatLocale | undefined): locale is FormatLocale {
  return !!locale && locale.decimal === ',';
}

function isTurkish(locale: FormatLocale): boolean {
  return locale.trueText === 'DOĞRU';
}

/** End of the quoted run that starts at `i` (a doubled quote stays inside); the length of the text when it never ends. */
function quotedEnd(s: string, i: number, quote: number): number {
  const n = s.length;
  for (let j = i + 1; j < n; j++) {
    if (s.charCodeAt(j) !== quote) continue;
    if (s.charCodeAt(j + 1) === quote) j++;
    else return j + 1;
  }
  return n;
}

/** End of the bracketed run that starts at `i`, brackets inside it included. */
function bracketEnd(s: string, i: number): number {
  const n = s.length;
  let depth = 0;
  for (let j = i; j < n; j++) {
    const c = s.charCodeAt(j);
    if (c === 91) depth++;
    else if (c === 93 && --depth === 0) return j + 1;
  }
  return n;
}

function errorAt(s: string, i: number): number {
  for (let k = 0; k < ERRORS.length; k++) {
    const code = ERRORS[k];
    if (s.length - i >= code.length && s.substr(i, code.length).toUpperCase() === code) return code.length;
  }
  return 0;
}

/** Length of an exponent ("E+12") at `i`, or 0. */
function exponentAt(s: string, i: number): number {
  const c = s.charCodeAt(i);
  if (c !== 69 && c !== 101) return 0;
  let j = i + 1;
  const sign = s.charCodeAt(j);
  if (sign === 43 || sign === 45) j++;
  const digits = j;
  while (isDigit(s.charCodeAt(j))) j++;
  return j > digits ? j - i : 0;
}

/**
 * One pass over the text. `toLocal` says which way: stored -> typed, or typed -> stored.
 * `names` (typed -> stored only) also turns Turkish function names into the stored ones.
 */
function translate(text: string, toLocal: boolean, names: boolean): string {
  const n = text.length;
  let out = '';
  let i = 0;
  // Inside an array constant the separators are those of columns and rows.
  let braces = 0;
  while (i < n) {
    const c = text.charCodeAt(i);
    if (c === 34 || c === 39) {
      const end = quotedEnd(text, i, c);
      out += text.slice(i, end);
      i = end;
    } else if (c === 91) {
      const end = bracketEnd(text, i);
      out += text.slice(i, end);
      i = end;
    } else if (c === 35) {
      const length = errorAt(text, i) || 1;
      out += text.substr(i, length);
      i += length;
    } else if (c === 123) {
      braces++;
      out += '{';
      i++;
    } else if (c === 125) {
      if (braces > 0) braces--;
      out += '}';
      i++;
    } else if (isDigit(c) || (toLocal && c === 46 && isDigit(text.charCodeAt(i + 1)))) {
      // A number: digits, a fraction, an exponent.
      let j = i;
      while (isDigit(text.charCodeAt(j))) j++;
      let whole = text.slice(i, j);
      let fraction = '';
      const separator = text.charCodeAt(j);
      if (separator === (toLocal ? 46 : 44) && isDigit(text.charCodeAt(j + 1))) {
        let k = j + 1;
        while (isDigit(text.charCodeAt(k))) k++;
        fraction = text.slice(j + 1, k);
        j = k;
      } else if (toLocal && separator === 46) {
        // "1." is the number 1; typed with a comma it would read as an argument separator.
        j++;
      }
      const exponent = exponentAt(text, j);
      // ".5" is written "0,5": a comma in front of digits would read as a separator.
      if (whole === '' && fraction !== '') whole = '0';
      out += whole + (fraction === '' ? '' : (toLocal ? ',' : '.') + fraction) + text.substr(j, exponent);
      i = j + exponent;
    } else if (braces > 0 && c === 92 && !toLocal) {
      // UNSURE: in an array constant Turkish Excel separates columns with ";" and rows with "\".
      out += ';';
      i++;
    } else if (isWordStart(c) || (c === 92 && braces === 0)) {
      let j = i + 1;
      // In an array constant a backslash ends the word: it may be the row separator.
      while (j < n && isWordChar(text.charCodeAt(j)) && !(braces > 0 && text.charCodeAt(j) === 92)) j++;
      let word = text.slice(i, j);
      if (names) {
        const stored = turkishName(word);
        // A function where it is called; TRUE and FALSE also where they stand as values.
        if (stored !== undefined && (text.charCodeAt(j) === 40 || stored === 'TRUE' || stored === 'FALSE')) word = stored;
      }
      out += word;
      i = j;
    } else if (c === 44) {
      out += toLocal ? ';' : ',';
      i++;
    } else if (c === 59) {
      // Stored: the row separator of an array constant. Typed: the separator of arguments, and of columns.
      out += toLocal ? (braces > 0 ? '\\' : ';') : ',';
      i++;
    } else {
      out += text.charAt(i);
      i++;
    }
  }
  return out;
}

/**
 * A stored formula (without "=") the way it is typed in the locale: ";" between
 * arguments and "," in numbers where the locale's decimal separator is a comma.
 * Function names stay English. Unchanged for every other locale.
 */
export function formulaToLocal(text: string, locale?: FormatLocale): string {
  if (typeof text !== 'string' || text.length > MAX_LENGTH || !usesComma(locale)) return text;
  return translate(text, true, false);
}

/**
 * A formula typed in the locale (without "=") the way it is stored. Where the decimal
 * separator is a comma: ";" separates arguments, a comma between two digits is a
 * decimal comma, any other comma is taken as the separator it is in a stored formula
 * (so a formula typed the English way is understood as long as it has no number list
 * like "1,5"). In Turkish, the Turkish names of common functions are understood too.
 */
export function formulaFromLocal(text: string, locale?: FormatLocale): string {
  if (typeof text !== 'string' || text.length > MAX_LENGTH || !usesComma(locale)) return text;
  return translate(text, false, isTurkish(locale));
}

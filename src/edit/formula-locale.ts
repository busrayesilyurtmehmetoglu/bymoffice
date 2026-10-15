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
// no sense comes out as it went in, and the parser has the last word. There is
// one exception. Where the decimal separator is a comma, a comma that does not
// stand between two digits means nothing, as in Excel, and such text is refused
// here: handed on, "IF(A1>5,1,0)" typed the English way would be read as another
// formula ("5,1" is a number), and the parser would have nothing to object to.

import type { FormatLocale } from '../format/locale';
import { withFuturePrefixes, withoutFuturePrefixes } from './future-names';

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
 * SATIRSAY, SÜTUNSAY, KAYDIR, DOLAYLI, ÖZDEŞ, YİNELE, DAMGA, and every name after
 * TAMİŞGÜNÜ in the list (the mathematical, statistical, text, date, lookup, information,
 * financial and database names added with the rest of the library).
 *
 * A function that has no line here is typed by its English name, in Turkish too.
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
  // Names that are the English ones but for a dotted İ: the engine knows no "SİN".
  ['SİN', 'SIN'],
  ['ASİN', 'ASIN'],
  ['SİNH', 'SINH'],
  ['ASİNH', 'ASINH'],
  ['Pİ', 'PI'],
  ['NOMİNAL', 'NOMINAL'],
  // The rest of the library, family by family. A function whose Turkish name is the English
  // one letter for letter (COS, LN, LOG, VAR, RANK, PEARSON, ATAN2 ...) needs no line.
  ['ÜS', 'EXP'],
  ['İŞARET', 'SIGN'],
  ['ÇARPINIM', 'FACT'],
  ['ÇİFTFAKTÖR', 'FACTDOUBLE'],
  ['KOMBİNASYON', 'COMBIN'],
  ['PERMÜTASYON', 'PERMUT'],
  ['OBEB', 'GCD'],
  ['OKEK', 'LCM'],
  ['KYUVARLA', 'MROUND'],
  ['BÖLÜM', 'QUOTIENT'],
  ['ÇİFT', 'EVEN'],
  ['TEK', 'ODD'],
  ['DERECE', 'DEGREES'],
  ['RADYAN', 'RADIANS'],
  ['TOPKARE', 'SUMSQ'],
  ['KAREKÖKPİ', 'SQRTPI'],
  ['ROMEN', 'ROMAN'],
  ['TAVANAYUVARLA.MATEMATİK', 'CEILING.MATH'],
  ['TABANAYUVARLA.MATEMATİK', 'FLOOR.MATH'],
  ['TAVANAYUVARLA.DUYARLI', 'CEILING.PRECISE'],
  ['TABANAYUVARLA.DUYARLI', 'FLOOR.PRECISE'],
  ['ISO.TAVAN', 'ISO.CEILING'],
  ['TOPLAMA', 'AGGREGATE'],
  ['DÇARP', 'MMULT'],
  ['TOPX2EY2', 'SUMX2MY2'],
  ['TOPX2AY2', 'SUMX2PY2'],
  ['TOPXEY2', 'SUMXMY2'],
  ['ORTALAMAA', 'AVERAGEA'],
  ['MAKA', 'MAXA'],
  ['MİNA', 'MINA'],
  ['ÇOKEĞERMAK', 'MAXIFS'],
  ['ÇOKEĞERMİN', 'MINIFS'],
  ['ENÇOK_OLAN', 'MODE'],
  ['ENÇOK_OLAN.TEK', 'MODE.SNGL'],
  ['STDSAPMA', 'STDEV'],
  ['STDSAPMA.S', 'STDEV.S'],
  ['STDSAPMA.P', 'STDEV.P'],
  ['STDSAPMAS', 'STDEVP'],
  ['STDSAPMAA', 'STDEVA'],
  ['STDSAPMASA', 'STDEVPA'],
  ['VARS', 'VARP'],
  ['VARSA', 'VARPA'],
  ['RANK.EŞİT', 'RANK.EQ'],
  ['RANK.ORT', 'RANK.AVG'],
  ['YÜZDEBİRLİK', 'PERCENTILE'],
  ['YÜZDEBİRLİK.DHL', 'PERCENTILE.INC'],
  ['YÜZDEBİRLİK.HRC', 'PERCENTILE.EXC'],
  ['DÖRTTEBİRLİK', 'QUARTILE'],
  ['DÖRTTEBİRLİK.DHL', 'QUARTILE.INC'],
  ['DÖRTTEBİRLİK.HRC', 'QUARTILE.EXC'],
  ['KORELASYON', 'CORREL'],
  ['EĞİM', 'SLOPE'],
  ['KESMENOKTASI', 'INTERCEPT'],
  ['TAHMİN', 'FORECAST'],
  ['TAHMİN.DOĞRUSAL', 'FORECAST.LINEAR'],
  ['GEOORT', 'GEOMEAN'],
  ['HARORT', 'HARMEAN'],
  ['KIRPORTALAMA', 'TRIMMEAN'],
  ['SIKLIK', 'FREQUENCY'],
  ['ORTSAP', 'AVEDEV'],
  ['SAPKARE', 'DEVSQ'],
  ['RKARE', 'RSQ'],
  ['KOVARYANS', 'COVAR'],
  ['KOVARYANS.P', 'COVARIANCE.P'],
  ['KOVARYANS.S', 'COVARIANCE.S'],
  ['NORM.DAĞ', 'NORM.DIST'],
  ['NORMDAĞ', 'NORMDIST'],
  ['NORM.TERS', 'NORM.INV'],
  ['NORMTERS', 'NORMINV'],
  ['NORM.S.DAĞ', 'NORM.S.DIST'],
  ['NORMSDAĞ', 'NORMSDIST'],
  ['NORM.S.TERS', 'NORM.S.INV'],
  ['NORMSTERS', 'NORMSINV'],
  ['YAZIM.DÜZENİ', 'PROPER'],
  ['TEMİZ', 'CLEAN'],
  ['KOD', 'CODE'],
  ['SAYIDÜZENLE', 'FIXED'],
  ['LİRA', 'DOLLAR'],
  ['SAYIDEĞERİ', 'NUMBERVALUE'],
  ['ARALIKBİRLEŞTİR', 'CONCAT'],
  ['METİNBİRLEŞTİR', 'TEXTJOIN'],
  ['M', 'T'],
  ['S', 'N'],
  ['TARİHSAYISI', 'DATEVALUE'],
  ['ZAMANSAYISI', 'TIMEVALUE'],
  ['GÜNSAY', 'DAYS'],
  ['GÜN360', 'DAYS360'],
  ['HAFTASAY', 'WEEKNUM'],
  ['ISOHAFTASAY', 'ISOWEEKNUM'],
  ['YILORAN', 'YEARFRAC'],
  ['ETARİHLİ', 'DATEDIF'],
  ['TAMİŞGÜNÜ.ULUSL', 'NETWORKDAYS.INTL'],
  ['İŞGÜNÜ.ULUSL', 'WORKDAY.INTL'],
  ['ELEMAN', 'CHOOSE'],
  ['ADRES', 'ADDRESS'],
  ['ALANSAY', 'AREAS'],
  ['DEVRİK_DÖNÜŞÜM', 'TRANSPOSE'],
  ['ÇAPRAZARA', 'XLOOKUP'],
  ['ÇAPRAZEŞLEŞTİR', 'XMATCH'],
  // The functions of dynamic arrays. UNSURE, like the rest of this part of the list.
  ['FİLTRE', 'FILTER'],
  ['SIRALA', 'SORT'],
  ['SIRALAÖLÇÜT', 'SORTBY'],
  ['BENZERSİZ', 'UNIQUE'],
  ['SIRALI', 'SEQUENCE'],
  ['RASTGELEDİZİ', 'RANDARRAY'],
  ['KÖPRÜ', 'HYPERLINK'],
  ['FORMÜLMETNİ', 'FORMULATEXT'],
  ['ÇOKEĞER', 'IFS'],
  ['İLKEŞLEŞEN', 'SWITCH'],
  ['ÖZELVEYA', 'XOR'],
  ['YOKSAY', 'NA'],
  ['EHATA', 'ISERR'],
  ['EYOKSA', 'ISNA'],
  ['EMETİNDEĞİLSE', 'ISNONTEXT'],
  ['EMANTIKSALSA', 'ISLOGICAL'],
  ['ÇİFTMİ', 'ISEVEN'],
  ['TEKMİ', 'ISODD'],
  ['EREFSE', 'ISREF'],
  ['EFORMÜLSE', 'ISFORMULA'],
  ['HATA.TİPİ', 'ERROR.TYPE'],
  ['TİP', 'TYPE'],
  ['SAYFA', 'SHEET'],
  ['SAYFALAR', 'SHEETS'],
  ['HÜCRE', 'CELL'],
  ['DEVRESEL_ÖDEME', 'PMT'],
  ['BD', 'PV'],
  ['GD', 'FV'],
  ['TAKSİT_SAYISI', 'NPER'],
  ['FAİZ_ORANI', 'RATE'],
  ['NBD', 'NPV'],
  ['İÇ_VERİM_ORANI', 'IRR'],
  ['ANBD', 'XNPV'],
  ['AİÇVERİMORANI', 'XIRR'],
  ['FAİZTUTARI', 'IPMT'],
  ['ANA_PARA_ÖDEMESİ', 'PPMT'],
  ['DA', 'SLN'],
  ['AZALANBAKİYE', 'DB'],
  ['ÇİFTAZALANBAKİYE', 'DDB'],
  ['YAT', 'SYD'],
  ['ETKİN', 'EFFECT'],
  ['VSEÇTOPLA', 'DSUM'],
  ['VSEÇSAY', 'DCOUNT'],
  ['VSEÇSAYDOLU', 'DCOUNTA'],
  ['VSEÇORT', 'DAVERAGE'],
  ['VSEÇMAK', 'DMAX'],
  ['VSEÇMİN', 'DMIN'],
  ['VAL', 'DGET'],
];

/**
 * Turkish names of the error constants, with the constants files store. Read when typed, in
 * a cell or in a formula; what is shown always has the English ones.
 *
 * UNSURE: taken from memory of the Turkish user interface, like the function names.
 * The longer of two names that begin alike comes first.
 */
const TURKISH_ERRORS: [string, string][] = [
  ['#SAYI/0!', '#DIV/0!'],
  ['#SAYI!', '#NUM!'],
  ['#DEĞER!', '#VALUE!'],
  ['#BAŞV!', '#REF!'],
  ['#BOŞ!', '#NULL!'],
  ['#YOK', '#N/A'],
  ['#AD?', '#NAME?'],
];

/** Upper case that treats the four Turkish "i" letters as one, so that "indis", "INDIS" and "İNDİS" are the same name. */
function nameKey(text: string): string {
  return text.replace(/[iıİ]/g, 'I').toUpperCase();
}

/**
 * Turkish names short enough to be the name of a function of a macro or an add-in as well
 * (as nameKey() has them). Where the workbook has a defined name of that spelling, a call
 * of it is left as it is typed.
 */
const SHORT_NAMES = new Set(['VAL', 'S', 'M', 'KOD', 'TEK', 'TIP', 'DA', 'BD', 'GD', 'YAT']);

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

/** The Turkish error constant at `i`: how long it is, and the constant that is stored for it. */
function turkishErrorAt(s: string, i: number): { length: number; stored: string } | undefined {
  for (let k = 0; k < TURKISH_ERRORS.length; k++) {
    const local = TURKISH_ERRORS[k][0];
    if (s.length - i >= local.length && nameKey(s.substr(i, local.length)) === nameKey(local)) return { length: local.length, stored: TURKISH_ERRORS[k][1] };
  }
  return undefined;
}

/**
 * The error constant that text typed into a cell stands for in the locale ("#YOK" is #N/A
 * in Turkish); undefined when the text is not the locale's own name of one.
 */
export function localErrorCode(text: string, locale?: FormatLocale): string | undefined {
  if (typeof text !== 'string' || text.length > 16 || !locale || !isTurkish(locale)) return undefined;
  const found = turkishErrorAt(text, 0);
  return found && found.length === text.length ? found.stored : undefined;
}

export interface FormulaLocaleOptions {
  /**
   * Whether the workbook has a defined name of this spelling. A name called DOĞRU (or
   * YANLIŞ) is then that name where it is typed, and not the constant.
   */
  isName?(name: string): boolean;
  /**
   * The formula the cell holds while it is edited, as it is stored. A name it calls is a
   * function of that file - of a macro, of an add-in, or one this table does not know - and
   * stays that name where it is typed, also when a Turkish function is called the same
   * ("VAL(" is DGET in Turkish): a formula typed again as it is shown is the formula it was.
   */
  stored?: string;
}

/** The names a stored formula calls, as nameKey() has them. Text in quotes and in brackets is passed over. */
function calledNames(stored: string): Set<string> {
  const called = new Set<string>();
  const n = stored.length;
  let i = 0;
  while (i < n) {
    const c = stored.charCodeAt(i);
    if (c === 34 || c === 39) i = quotedEnd(stored, i, c);
    else if (c === 91) i = bracketEnd(stored, i);
    else if (isWordStart(c)) {
      const j = wordEnd(stored, i + 1);
      if (stored.charCodeAt(j) === 40) called.add(nameKey(stored.slice(i, j)));
      i = j;
    } else i++;
  }
  return called;
}

/** What a pass over the text met that the caller has to know of. */
interface Findings {
  /** A comma that is neither inside a number nor a separator of the notation. */
  strayComma: boolean;
}

/** Offset just past the word that starts at `i`. */
function wordEnd(s: string, i: number): number {
  let j = i;
  while (j < s.length && isWordChar(s.charCodeAt(j))) j++;
  return j;
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
 * `names` (typed -> stored only) also turns Turkish function names, TRUE and FALSE and
 * error constants into the stored ones.
 */
function translate(text: string, toLocal: boolean, names: boolean, findings: Findings, options?: FormulaLocaleOptions): string {
  const n = text.length;
  let out = '';
  let i = 0;
  // Inside an array constant the separators are those of columns and rows.
  let braces = 0;
  // The names the formula of the cell calls already (FormulaLocaleOptions.stored), worked out when first asked for.
  let kept: Set<string> | undefined;
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
      const local = names ? turkishErrorAt(text, i) : undefined;
      if (local) {
        out += local.stored;
        i += local.length;
      } else {
        const length = errorAt(text, i) || 1;
        out += text.substr(i, length);
        i += length;
      }
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
        if (stored !== undefined) {
          const next = text.charCodeAt(j);
          // A function where it is called, and nowhere else: a sheet or a defined name may be called the same.
          if (next === 40) {
            // Not where the file has its own function of that spelling: one the formula of the cell
            // already calls, or - for the short names only - a defined name. Any other Turkish name
            // is the function whatever is defined, as in Excel. A name that is the stored one but for its dots
            // ("SİN") is no other function and is always stored the English way.
            if (nameKey(stored) !== nameKey(word)) {
              if (kept === undefined) kept = options !== undefined && typeof options.stored === 'string' && options.stored.length <= MAX_LENGTH ? calledNames(options.stored) : new Set<string>();
              const own = kept.has(nameKey(word)) || (SHORT_NAMES.has(nameKey(word)) && options !== undefined && typeof options.isName === 'function' && options.isName(word) === true);
              if (!own) word = stored;
            } else word = stored;
          }
          else if (stored === 'TRUE' || stored === 'FALSE') {
            // TRUE and FALSE also where they stand as values. Not where the word names a sheet
            // ("DOĞRU!A1", "DOĞRU:YANLIŞ!A1"), a name of a sheet ("Sayfa1!DOĞRU") or a defined name.
            const sheet = next === 33 || (next === 58 && text.charCodeAt(wordEnd(text, j + 1)) === 33 && wordEnd(text, j + 1) > j + 1);
            const scoped = out.charCodeAt(out.length - 1) === 33;
            const defined = options !== undefined && typeof options.isName === 'function' && options.isName(word) === true;
            if (!sheet && !scoped && !defined) word = stored;
          }
        }
      }
      out += word;
      i = j;
    } else if (c === 44) {
      // Stored: the separator of arguments. Typed: a comma is the decimal comma, and that was read
      // with its number; any other comma has no meaning.
      if (!toLocal) findings.strayComma = true;
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
  if (typeof text !== 'string' || text.length > MAX_LENGTH) return text;
  // What files add to newer functions ("_xlfn.") is not shown; "A1#" and "@" are.
  const shown = withoutFuturePrefixes(text);
  if (!usesComma(locale)) return shown;
  return translate(shown, true, false, { strayComma: false });
}

/**
 * A formula typed in the locale (without "=") the way it is stored. Where the decimal
 * separator is a comma: ";" separates arguments (and the areas of a union, and the columns
 * of an array constant, whose rows "\" separates), and a comma between two digits is the
 * decimal comma. In Turkish, the Turkish names of common functions, of TRUE and FALSE and
 * of the error constants are understood too.
 *
 * Throws RangeError("The formula cannot be read") for a comma anywhere else, as Excel refuses
 * it: "IF(A1>5,1,0)" typed the English way is not taken for "IF(A1>5.1,0)", and
 * "ROUND(A1,2)" is not understood either, so that the English notation does not work in one
 * formula and silently mean something else in the next. Unchanged for every other locale.
 */
export function formulaFromLocal(text: string, locale?: FormatLocale, options?: FormulaLocaleOptions): string {
  if (typeof text !== 'string' || text.length > MAX_LENGTH) return text;
  if (!usesComma(locale)) return withFuturePrefixes(text);
  const findings: Findings = { strayComma: false };
  const stored = translate(text, false, isTurkish(locale), findings, options);
  if (findings.strayComma) throw new RangeError('The formula cannot be read');
  // As files store it: "_xlfn." before the newer functions, "_xlpm." before the names of LET and LAMBDA.
  return withFuturePrefixes(stored);
}

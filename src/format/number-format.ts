// Excel number format codes: "#,##0.00", "0%", "d-mmm-yy", "[Red](0.00)", ...
//
// A code is compiled once into sections (positive;negative;zero;text) and
// cached; formatting a value then only walks the compiled tokens. The compiler
// never throws: a code it cannot make sense of shows values the "General" way.

import { formatGeneral } from '../render/cell-text';
import { LOCALE_EN, LOCALE_TR, normalizeLocale, type FormatLocale } from './locale';

export interface FormatOptions {
  locale?: FormatLocale;
  date1904?: boolean;
  /**
   * Built-in format id of the cell. Ids 14 and 22 mean "the system short date" (22: plus the
   * short time) whatever code is stored for them; without the id the code is applied as written.
   */
  builtinId?: number;
}

export interface FormattedValue {
  text: string;
  /** "#RRGGBB" when the section names a colour ([Red], [Color 3], ...). */
  color?: string;
  /** True when the value cannot be shown with this format (negative date): the grid draws "####". */
  overflow?: boolean;
}

/** Longer codes are not compiled; Excel itself stops at 255 characters. */
const MAX_CODE_LENGTH = 1024;
const MAX_CACHE_ENTRIES = 500;
/** Cap for text sections that repeat "@" many times over a long string. */
const MAX_TEXT_OUTPUT = 65536;
/** Last serial of each date system: 31 Dec 9999. */
const MAX_SERIAL_1900 = 2958465;
const MAX_SERIAL_1904 = 2957003;
/** Text shown when a value has no representation in the format; callers draw their own fill. */
const OVERFLOW_TEXT = '####';
/** Largest denominator of a "?/8" style fraction; longer ones are clamped. */
const MAX_FIXED_DENOMINATOR = 1e15;

const BUILTIN_SHORT_DATE = 14;
const BUILTIN_SHORT_DATE_TIME = 22;
const NUMBER_ERROR = '#NUM!';
const DEFAULT_SHORT_DATE = 'm/d/yyyy';
const DEFAULT_SHORT_TIME = 'h:mm';
const DEFAULT_LONG_DATE = 'dddd, mmmm d, yyyy';
const DEFAULT_LONG_TIME = 'h:mm:ss AM/PM';

// ---------------------------------------------------------------------------
// Compiled representation

// Tokens of number, text and General sections.
const T_LIT = 0;
const T_ZERO = 1;
const T_HASH = 2;
const T_QUESTION = 3;
/** Where the text ("@") or the General rendering goes. */
const T_SLOT = 4;

interface Tok {
  k: number;
  s: string;
}

// Tokens of date sections.
const D_LIT = 0;
const D_YEAR = 1;
const D_MONTH = 2;
const D_DAY = 3;
const D_HOUR = 4;
const D_MINUTE = 5;
const D_SECOND = 6;
const D_SUBSECOND = 7;
const D_AMPM = 8;
const D_ELAPSED_HOURS = 9;
const D_ELAPSED_MINUTES = 10;
const D_ELAPSED_SECONDS = 11;
const D_DATE_SEPARATOR = 12;
const D_TIME_SEPARATOR = 13;
const D_BUDDHIST_YEAR = 14;

interface DateTok {
  k: number;
  /** Letter count ("mmm" = 3); for D_AMPM the variant: 0 AM/PM in any case, 2 A/P, 3 a/p. */
  n: number;
  s: string;
}

const K_LITERAL = 0;
const K_NUMBER = 1;
const K_DATE = 2;
const K_GENERAL = 3;
const K_TEXT = 4;
/** "[$-F800]": the locale's long date. */
const K_SYSTEM_DATE = 5;
/** "[$-F400]": the locale's time. */
const K_SYSTEM_TIME = 6;

interface Condition {
  /** 0 "<", 1 "<=", 2 ">", 3 ">=", 4 "=", 5 "<>" */
  op: number;
  value: number;
}

interface Section {
  kind: number;
  color: string | undefined;
  cond: Condition | null;
  /** Month/day names forced by an LCID such as [$-409]; null means the caller's locale. */
  names: FormatLocale | null;
  /** K_LITERAL: the whole text. */
  literal: string;
  /** K_TEXT and K_GENERAL: literals around the slot(s). */
  slotToks: Tok[];

  // K_NUMBER
  intToks: Tok[];
  intCount: number;
  decToks: Tok[];
  decimals: number;
  hasPoint: boolean;
  grouped: boolean;
  /** Power of ten applied before rounding: +2 per "%", -3 per scaling comma. */
  scale: number;
  scientific: boolean;
  expPlus: boolean;
  expToks: Tok[];
  expCount: number;
  fraction: boolean;
  numToks: Tok[];
  /** Fixed denominator ("?/8"), or 0 when the denominator is searched. */
  fixedDen: number;
  /** Character width of the denominator in the code. */
  denWidth: number;
  /** Whether the denominator is padded with spaces ("?" placeholders). */
  denPadded: boolean;
  /** Minimum digits of a searched denominator ("0" placeholders). */
  denZeros: number;
  /** A "0" in the numerator: the fraction is written even when it is zero ("3 0/1"). */
  numForced: boolean;
  tailToks: Tok[];

  // K_DATE
  dateToks: DateTok[];
  subDigits: number;
  hour12: boolean;
}

interface Compiled {
  /** Sections for numbers, at most three. Empty: numbers are shown as General. */
  numeric: Section[];
  text: Section | null;
  conditional: boolean;
  date: boolean;
}

const GENERAL: Compiled = { numeric: [], text: null, conditional: false, date: false };

function newSection(): Section {
  return {
    kind: K_LITERAL,
    color: undefined,
    cond: null,
    names: null,
    literal: '',
    slotToks: [],
    intToks: [],
    intCount: 0,
    decToks: [],
    decimals: 0,
    hasPoint: false,
    grouped: false,
    scale: 0,
    scientific: false,
    expPlus: true,
    expToks: [],
    expCount: 0,
    fraction: false,
    numToks: [],
    fixedDen: 0,
    denWidth: 0,
    denPadded: false,
    denZeros: 0,
    numForced: false,
    tailToks: [],
    dateToks: [],
    subDigits: 0,
    hour12: false,
  };
}

// ---------------------------------------------------------------------------
// Colours

const NAMED_COLORS: { [name: string]: string } = {
  black: '#000000',
  blue: '#0000FF',
  cyan: '#00FFFF',
  green: '#00FF00',
  magenta: '#FF00FF',
  red: '#FF0000',
  white: '#FFFFFF',
  yellow: '#FFFF00',
};

/** [Color 1] .. [Color 56]: the default workbook palette (indexed colours 8..63). */
const INDEXED_COLORS =
  '000000FFFFFFFF000000FF000000FFFFFF00FF00FF00FFFF' +
  '800000008000000080808000800080008080C0C0C0808080' +
  '9999FF993366FFFFCCCCFFFF660066FF80800066CCCCCCFF' +
  '000080FF00FFFFFF0000FFFF8000808000000080800000FF' +
  '00CCFFCCFFFFCCFFCCFFFF9999CCFFFF99CCCC99FFFFCC99' +
  '3366FF33CCCC99CC00FFCC00FF9900FF6600666699969696' +
  '003366339966003300333300993300993366333399333333';

function parseColor(content: string): string | undefined {
  const lower = content.toLowerCase();
  if (Object.prototype.hasOwnProperty.call(NAMED_COLORS, lower)) return NAMED_COLORS[lower];
  const match = /^color\s*(\d{1,2})$/.exec(lower);
  if (!match) return undefined;
  const index = Number(match[1]);
  if (index < 1 || index > 56) return undefined;
  return '#' + INDEXED_COLORS.substr((index - 1) * 6, 6);
}

function parseCondition(content: string): Condition | null {
  // No condition is this long, and on a long run of digits the expression below backtracks for minutes.
  if (content.length > 64) return null;
  const match = /^(<=|>=|<>|<|>|=)\s*(-?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)$/.exec(content);
  if (!match) return null;
  const value = Number(match[2]);
  if (value !== value) return null;
  const ops = ['<', '<=', '>', '>=', '=', '<>'];
  return { op: ops.indexOf(match[1]), value };
}

function testCondition(cond: Condition, v: number): boolean {
  switch (cond.op) {
    case 0:
      return v < cond.value;
    case 1:
      return v <= cond.value;
    case 2:
      return v > cond.value;
    case 3:
      return v >= cond.value;
    case 4:
      return v === cond.value;
    default:
      return v !== cond.value;
  }
}

/** A condition only negative numbers can meet: its section shows them without the minus sign. */
function negativeOnly(cond: Condition): boolean {
  return (cond.op === 0 && cond.value <= 0) || (cond.op === 1 && cond.value < 0);
}

// ---------------------------------------------------------------------------
// Compiler

// Raw tokens: what the scanner sees before the section's kind is known.
const R_LIT = 0;
const R_PLACEHOLDER = 1;
const R_DOT = 2;
const R_COMMA = 3;
const R_PERCENT = 4;
const R_EXPONENT = 5;
const R_SLASH = 6;
const R_COLON = 7;
const R_AT = 8;
const R_GENERAL = 9;
const R_DATE = 10;
const R_AMPM = 11;
const R_ELAPSED = 12;

interface Raw {
  k: number;
  s: string;
  n: number;
  /** An unquoted single character (only those can be the digits of a fixed denominator). */
  bare: boolean;
  /** Date letters as written, for sections where they turn out to be plain text. */
  src?: string;
}

interface RawSection {
  raws: Raw[];
  section: Section;
  hasDate: boolean;
  hasPlaceholder: boolean;
  hasGeneral: boolean;
  hasAt: boolean;
  system: number;
}

function newRawSection(): RawSection {
  return {
    raws: [],
    section: newSection(),
    hasDate: false,
    hasPlaceholder: false,
    hasGeneral: false,
    hasAt: false,
    system: 0,
  };
}

/** Length in code units of the character at `i` (2 for a surrogate pair). */
function charLength(code: string, i: number): number {
  const unit = code.charCodeAt(i);
  if (unit >= 0xd800 && unit <= 0xdbff && i + 1 < code.length) {
    const next = code.charCodeAt(i + 1);
    if (next >= 0xdc00 && next <= 0xdfff) return 2;
  }
  return 1;
}

function matchesAt(code: string, i: number, lowerWord: string): boolean {
  if (i + lowerWord.length > code.length) return false;
  return code.substr(i, lowerWord.length).toLowerCase() === lowerWord;
}

function localeForLcid(lcid: number): FormatLocale | null {
  const language = lcid & 0x3ff;
  if (language === 0x1f) return LOCALE_TR;
  if (language === 0x09) return LOCALE_EN;
  return null;
}

/** Handles "[...]"; `content` is the text between the brackets. */
function applyBracket(content: string, current: RawSection): void {
  const section = current.section;
  if (content.length === 0) return;
  const first = content.charAt(0);
  if (first === '$') {
    // [$symbol-locale]: the symbol (up to the first "-") is literal text; the locale, a
    // hexadecimal LCID or a language tag such as "tr-TR", is not shown.
    const body = content.substring(1);
    const dash = body.indexOf('-');
    const symbol = dash < 0 ? body : body.substring(0, dash);
    const id = dash < 0 ? '' : body.substring(dash + 1).toLowerCase();
    if (symbol.length > 0) current.raws.push({ k: R_LIT, s: symbol, n: 0, bare: false });
    if (id === 'f800' || id === 'x-sysdate') {
      current.system = K_SYSTEM_DATE;
      return;
    }
    if (id === 'f400' || id === 'x-systime') {
      current.system = K_SYSTEM_TIME;
      return;
    }
    const next = id.indexOf('-');
    const primary = next < 0 ? id : id.substring(0, next);
    let names: FormatLocale | null = null;
    if (primary === 'tr') names = LOCALE_TR;
    else if (primary === 'en') names = LOCALE_EN;
    else if (/^[0-9a-f]{1,8}$/.test(primary)) names = localeForLcid(parseInt(primary, 16));
    if (names) section.names = names;
    return;
  }
  if (first === '<' || first === '>' || first === '=') {
    const cond = parseCondition(content);
    if (cond && !section.cond) section.cond = cond;
    return;
  }
  const lower = content.toLowerCase();
  if (/^(h+|m+|s+)$/.test(lower)) {
    current.raws.push({ k: R_ELAPSED, s: lower.charAt(0), n: lower.length, bare: false });
    current.hasDate = true;
    return;
  }
  const color = parseColor(content);
  if (color !== undefined) section.color = color;
  // Anything else ([DBNum1], [ENG], ...) is not understood and has no visible effect.
}

function scan(code: string): RawSection[] | null {
  const sections: RawSection[] = [];
  let current = newRawSection();
  const length = code.length;
  let i = 0;
  while (i < length) {
    const c = code.charAt(i);
    switch (c) {
      case ';':
        sections.push(current);
        if (sections.length >= 4) return null;
        current = newRawSection();
        i++;
        continue;
      case '"': {
        let end = code.indexOf('"', i + 1);
        if (end < 0) end = length;
        if (end > i + 1) current.raws.push({ k: R_LIT, s: code.substring(i + 1, end), n: 0, bare: false });
        i = end + 1;
        continue;
      }
      case '\\': {
        if (i + 1 < length) {
          const size = charLength(code, i + 1);
          current.raws.push({ k: R_LIT, s: code.substr(i + 1, size), n: 0, bare: false });
          i += 1 + size;
        } else {
          i++;
        }
        continue;
      }
      case '_':
        // "_x": a space as wide as x. Without font metrics that is one space.
        if (i + 1 < length) {
          current.raws.push({ k: R_LIT, s: ' ', n: 0, bare: false });
          i += 1 + charLength(code, i + 1);
        } else {
          i++;
        }
        continue;
      case '*':
        // "*x": repeat x to fill the cell. The width is unknown here, so nothing is emitted.
        i += i + 1 < length ? 1 + charLength(code, i + 1) : 1;
        continue;
      case '[': {
        const end = code.indexOf(']', i + 1);
        if (end < 0) return null;
        applyBracket(code.substring(i + 1, end), current);
        i = end + 1;
        continue;
      }
      case '0':
      case '#':
      case '?':
        current.raws.push({ k: R_PLACEHOLDER, s: c, n: 0, bare: true });
        current.hasPlaceholder = true;
        i++;
        continue;
      case '.':
        current.raws.push({ k: R_DOT, s: c, n: 0, bare: true });
        i++;
        continue;
      case ',':
        current.raws.push({ k: R_COMMA, s: c, n: 0, bare: true });
        i++;
        continue;
      case '%':
        current.raws.push({ k: R_PERCENT, s: c, n: 0, bare: true });
        i++;
        continue;
      case '/':
        current.raws.push({ k: R_SLASH, s: c, n: 0, bare: true });
        i++;
        continue;
      case ':':
        current.raws.push({ k: R_COLON, s: c, n: 0, bare: true });
        i++;
        continue;
      case '@':
        current.raws.push({ k: R_AT, s: c, n: 0, bare: true });
        current.hasAt = true;
        i++;
        continue;
      default:
        break;
    }

    const lower = c.toLowerCase();
    if (lower === 'g' && matchesAt(code, i, 'general')) {
      current.raws.push({ k: R_GENERAL, s: '', n: 0, bare: false });
      current.hasGeneral = true;
      i += 7;
      continue;
    }
    if (lower === 'a') {
      if (matchesAt(code, i, 'am/pm')) {
        current.raws.push({ k: R_AMPM, s: '', n: 0, bare: false, src: code.substr(i, 5) });
        current.hasDate = true;
        i += 5;
        continue;
      }
      if (matchesAt(code, i, 'a/p')) {
        current.raws.push({ k: R_AMPM, s: '', n: c === 'a' ? 3 : 2, bare: false, src: code.substr(i, 3) });
        current.hasDate = true;
        i += 3;
        continue;
      }
    }
    if (lower === 'e' && i + 1 < length && (code.charAt(i + 1) === '+' || code.charAt(i + 1) === '-')) {
      current.raws.push({ k: R_EXPONENT, s: code.charAt(i + 1), n: 0, bare: false });
      i += 2;
      continue;
    }
    if (
      lower === 'y' || lower === 'm' || lower === 'd' || lower === 'h' || lower === 's' ||
      lower === 'e' || lower === 'b' || lower === 'g'
    ) {
      let end = i + 1;
      while (end < length && code.charAt(end).toLowerCase() === lower) end++;
      current.raws.push({ k: R_DATE, s: lower, n: end - i, bare: false, src: code.substring(i, end) });
      // "g" (era name) shows nothing in the Gregorian calendar and does not make a date by itself.
      if (lower !== 'g') current.hasDate = true;
      i = end;
      continue;
    }
    const size = charLength(code, i);
    current.raws.push({ k: R_LIT, s: code.substr(i, size), n: 0, bare: size === 1 });
    i += size;
  }
  sections.push(current);
  return sections;
}

function pushLit(toks: Tok[], text: string): void {
  if (text.length === 0) return;
  const last = toks.length > 0 ? toks[toks.length - 1] : null;
  if (last && last.k === T_LIT) last.s += text;
  else toks.push({ k: T_LIT, s: text });
}

function pushDateLit(toks: DateTok[], text: string): void {
  if (text.length === 0) return;
  const last = toks.length > 0 ? toks[toks.length - 1] : null;
  if (last && last.k === D_LIT) last.s += text;
  else toks.push({ k: D_LIT, n: 0, s: text });
}

function placeholderKind(c: string): number {
  return c === '0' ? T_ZERO : c === '#' ? T_HASH : T_QUESTION;
}

function isBareDigit(raw: Raw): boolean {
  return raw.k === R_LIT && raw.bare && raw.s >= '1' && raw.s <= '9';
}

function buildNumber(raws: Raw[], section: Section): void {
  const count = raws.length;
  // 0 integer part, 1 decimals, 2 exponent, 3 after a fraction
  let phase = 0;
  let cur = section.intToks;
  let seenPlaceholder = false;
  let i = 0;
  while (i < count) {
    const raw = raws[i];
    switch (raw.k) {
      case R_PLACEHOLDER:
        if (phase === 3) break; // digits after the denominator have nowhere to come from
        cur.push({ k: placeholderKind(raw.s), s: '' });
        seenPlaceholder = true;
        if (phase === 0) section.intCount++;
        else if (phase === 1) section.decimals++;
        else section.expCount++;
        break;
      case R_COMMA: {
        let end = i;
        while (end < count && raws[end].k === R_COMMA) end++;
        const afterPlaceholder = i > 0 && raws[i - 1].k === R_PLACEHOLDER;
        const beforePlaceholder = end < count && raws[end].k === R_PLACEHOLDER;
        if (afterPlaceholder && beforePlaceholder) {
          // Between digits: thousands separators (no meaning among decimals).
          if (phase === 0) section.grouped = true;
        } else if (afterPlaceholder || (seenPlaceholder && i > 0 && raws[i - 1].k === R_DOT)) {
          // After the last digit: each comma divides by a thousand.
          section.scale -= 3 * (end - i);
        } else {
          pushLit(cur, repeat(',', end - i));
        }
        i = end;
        continue;
      }
      case R_PERCENT:
        section.scale += 2;
        pushLit(cur, '%');
        break;
      case R_DOT:
        if (phase === 0 && !section.hasPoint) {
          section.hasPoint = true;
          phase = 1;
          cur = section.decToks;
        } else {
          pushLit(cur, '.');
        }
        break;
      case R_EXPONENT:
        if (phase <= 1 && seenPlaceholder && i + 1 < count && raws[i + 1].k === R_PLACEHOLDER) {
          section.scientific = true;
          section.expPlus = raw.s === '+';
          phase = 2;
          cur = section.expToks;
        } else {
          pushLit(cur, 'E' + raw.s);
        }
        break;
      case R_SLASH: {
        const next = i + 1 < count ? raws[i + 1] : null;
        if (
          phase === 0 && i > 0 && raws[i - 1].k === R_PLACEHOLDER &&
          next !== null && (next.k === R_PLACEHOLDER || isBareDigit(next))
        ) {
          // The placeholders right before "/" are the numerator; earlier ones the whole part.
          const intToks = section.intToks;
          while (intToks.length > 0 && intToks[intToks.length - 1].k !== T_LIT) {
            section.numToks.unshift(intToks.pop() as Tok);
          }
          section.intCount -= section.numToks.length;
          let end = i + 1;
          let text = '';
          while (end < count && (raws[end].k === R_PLACEHOLDER || isBareDigit(raws[end]))) {
            text += raws[end].s;
            end++;
          }
          section.denWidth = text.length;
          if (/[1-9]/.test(text)) {
            const fixed = parseInt(text.replace(/[#?]/g, ''), 10);
            section.fixedDen = fixed >= 1 && fixed <= MAX_FIXED_DENOMINATOR ? fixed : MAX_FIXED_DENOMINATOR;
          } else {
            section.denPadded = text.indexOf('?') >= 0;
            for (let j = 0; j < text.length; j++) if (text.charAt(j) === '0') section.denZeros++;
          }
          section.fraction = true;
          phase = 3;
          cur = section.tailToks;
          i = end;
          continue;
        }
        pushLit(cur, '/');
        break;
      }
      case R_LIT:
      case R_COLON:
        pushLit(cur, raw.s);
        break;
      case R_DATE:
      case R_AMPM:
        // Unquoted letters in a number ("0 kg", "0.00 EUR") are plain text.
        pushLit(cur, raw.src || '');
        break;
      default:
        // "@" and "General" next to digit placeholders have no meaning.
        break;
    }
    i++;
  }
  for (let j = 0; j < section.numToks.length; j++) {
    if (section.numToks[j].k === T_ZERO) section.numForced = true;
  }
}

/** True when the next date part after `from` is seconds: then "m" means minutes. */
function secondsFollow(raws: Raw[], from: number): boolean {
  for (let i = from; i < raws.length; i++) {
    const k = raws[i].k;
    if (k === R_DATE && raws[i].s === 'g') continue;
    if (k === R_DATE || k === R_ELAPSED) return raws[i].s === 's';
  }
  return false;
}

function buildDate(raws: Raw[], section: Section): void {
  const toks = section.dateToks;
  const count = raws.length;
  // The previous date part: y, m (month), d, h, n (minutes), s.
  let previous = '';
  let i = 0;
  while (i < count) {
    const raw = raws[i];
    switch (raw.k) {
      case R_DATE:
        switch (raw.s) {
          case 'y':
            toks.push({ k: D_YEAR, n: raw.n, s: '' });
            previous = 'y';
            break;
          case 'e':
            toks.push({ k: D_YEAR, n: 4, s: '' });
            previous = 'y';
            break;
          case 'b':
            toks.push({ k: D_BUDDHIST_YEAR, n: raw.n, s: '' });
            previous = 'y';
            break;
          case 'd':
            toks.push({ k: D_DAY, n: raw.n, s: '' });
            previous = 'd';
            break;
          case 'h':
            toks.push({ k: D_HOUR, n: raw.n, s: '' });
            previous = 'h';
            break;
          case 's':
            toks.push({ k: D_SECOND, n: raw.n, s: '' });
            previous = 's';
            break;
          case 'g':
            break;
          default:
            // "m" and "mm" are minutes right after hours or right before seconds, otherwise the month.
            if (raw.n <= 2 && (previous === 'h' || secondsFollow(raws, i + 1))) {
              toks.push({ k: D_MINUTE, n: raw.n, s: '' });
              previous = 'n';
            } else {
              toks.push({ k: D_MONTH, n: raw.n, s: '' });
              previous = 'm';
            }
            break;
        }
        break;
      case R_ELAPSED:
        if (raw.s === 'h') {
          toks.push({ k: D_ELAPSED_HOURS, n: raw.n, s: '' });
          previous = 'h';
        } else if (raw.s === 'm') {
          toks.push({ k: D_ELAPSED_MINUTES, n: raw.n, s: '' });
          previous = 'n';
        } else {
          toks.push({ k: D_ELAPSED_SECONDS, n: raw.n, s: '' });
          previous = 's';
        }
        break;
      case R_AMPM:
        toks.push({ k: D_AMPM, n: raw.n, s: '' });
        section.hour12 = true;
        break;
      case R_DOT: {
        // ".0", ".00", ".000" directly after seconds: fractions of a second.
        const afterSeconds =
          i > 0 && (raws[i - 1].k === R_DATE || raws[i - 1].k === R_ELAPSED) && raws[i - 1].s === 's';
        let end = i + 1;
        while (end < count && raws[end].k === R_PLACEHOLDER && raws[end].s === '0') end++;
        if (afterSeconds && end > i + 1) {
          const digits = Math.min(end - i - 1, 3);
          toks.push({ k: D_SUBSECOND, n: digits, s: '' });
          if (digits > section.subDigits) section.subDigits = digits;
          i = end;
          continue;
        }
        pushDateLit(toks, '.');
        break;
      }
      case R_SLASH:
        toks.push({ k: D_DATE_SEPARATOR, n: 0, s: '' });
        break;
      case R_COLON:
        toks.push({ k: D_TIME_SEPARATOR, n: 0, s: '' });
        break;
      case R_EXPONENT:
        pushDateLit(toks, 'E' + raw.s);
        break;
      case R_AT:
      case R_GENERAL:
        break;
      default:
        pushDateLit(toks, raw.s);
        break;
    }
    i++;
  }
}

function buildSlots(raws: Raw[], section: Section): void {
  for (let i = 0; i < raws.length; i++) {
    const raw = raws[i];
    if (raw.k === R_DATE) continue;
    if (raw.k === R_AT || raw.k === R_GENERAL) section.slotToks.push({ k: T_SLOT, s: '' });
    else if (raw.k === R_EXPONENT) pushLit(section.slotToks, 'E' + raw.s);
    else pushLit(section.slotToks, raw.s);
  }
}

/**
 * Date letters make a date section unless digit placeholders say it is a number: then the
 * letters are text ("0 kg", "0 ms"). Zeros after seconds ("ss.00") are part of the time, and
 * an elapsed-time bracket ("[h]") always means time.
 */
function isDateSection(raw: RawSection): boolean {
  if (!raw.hasDate) return false;
  if (!raw.hasPlaceholder) return true;
  const raws = raw.raws;
  for (let i = 0; i < raws.length; i++) if (raws[i].k === R_ELAPSED) return true;
  for (let i = 0; i < raws.length; i++) {
    if (raws[i].k !== R_PLACEHOLDER) continue;
    if (raws[i].s !== '0') return false;
    let dot = i - 1;
    while (dot >= 0 && raws[dot].k === R_PLACEHOLDER && raws[dot].s === '0') dot--;
    if (dot < 1 || raws[dot].k !== R_DOT || raws[dot - 1].k !== R_DATE || raws[dot - 1].s !== 's') return false;
  }
  return true;
}

function buildSection(raw: RawSection): Section {
  const section = raw.section;
  if (raw.system !== 0) {
    section.kind = raw.system;
  } else if (isDateSection(raw)) {
    section.kind = K_DATE;
    buildDate(raw.raws, section);
  } else if (raw.hasPlaceholder) {
    section.kind = K_NUMBER;
    buildNumber(raw.raws, section);
  } else if (raw.hasGeneral || raw.hasAt) {
    section.kind = raw.hasGeneral ? K_GENERAL : K_TEXT;
    buildSlots(raw.raws, section);
  } else {
    section.kind = K_LITERAL;
    let text = '';
    for (let i = 0; i < raw.raws.length; i++) {
      const token = raw.raws[i];
      if (token.k !== R_DATE) text += token.k === R_EXPONENT ? 'E' + token.s : token.s;
    }
    section.literal = text;
  }
  return section;
}

function compile(code: string): Compiled {
  const raws = scan(code);
  if (raws === null) return GENERAL;
  const sections: Section[] = [];
  for (let i = 0; i < raws.length; i++) sections.push(buildSection(raws[i]));

  let text: Section | null = null;
  // The fourth section is for text; so is a last section that only holds "@".
  if (sections.length === 4 || sections[sections.length - 1].kind === K_TEXT) {
    text = sections.pop() as Section;
  }
  // A lone empty section is an empty code: General.
  if (sections.length === 1 && text === null && raws[0].raws.length === 0 && raws[0].system === 0) {
    if (sections[0].color === undefined && sections[0].cond === null) return GENERAL;
    // Only a colour or a condition: General with that colour.
    sections[0].kind = K_GENERAL;
    sections[0].slotToks.push({ k: T_SLOT, s: '' });
  }
  let conditional = false;
  let date = false;
  for (let i = 0; i < sections.length; i++) {
    const kind = sections[i].kind;
    if (sections[i].cond) conditional = true;
    if (kind === K_DATE || kind === K_SYSTEM_DATE || kind === K_SYSTEM_TIME) date = true;
  }
  return { numeric: sections, text, conditional, date };
}

const cache = new Map<string, Compiled>();
let lastCode = '';
let lastCompiled: Compiled = GENERAL;

function getCompiled(code: string): Compiled {
  if (code === lastCode) return lastCompiled;
  if (typeof code !== 'string' || code.length > MAX_CODE_LENGTH) return GENERAL;
  let compiled = cache.get(code);
  if (compiled === undefined) {
    try {
      compiled = compile(code);
    } catch (error) {
      // The compiler is not expected to throw; a format code must never break rendering.
      compiled = GENERAL;
    }
    if (cache.size >= MAX_CACHE_ENTRIES) {
      const oldest = cache.keys().next();
      if (!oldest.done) cache.delete(oldest.value);
    }
    cache.set(code, compiled);
  }
  lastCode = code;
  lastCompiled = compiled;
  return compiled;
}

// ---------------------------------------------------------------------------
// Decimal digits of a number

const ZEROS = '0000000000000000000000000000000000000000000000000000000000000000';

function repeat(text: string, n: number): string {
  let out = '';
  for (let i = 0; i < n; i++) out += text;
  return out;
}

function zeros(n: number): string {
  if (n <= 0) return '';
  return n <= 64 ? ZEROS.substring(0, n) : repeat('0', n);
}

// Results of roundFixed(), kept in module variables so the hot path allocates no tuple.
let rInt = '0';
let rFrac = '';

/**
 * Splits 0.d1d2...d15 * 10^(exp + 1) into integer and fraction digits, rounded half up
 * to `decimals` places. Working on the 15 significant digits Excel itself keeps makes
 * 1.005 round to 1.01 although the binary value is slightly below.
 */
function fixedFromDigits(digits: string, exp: number, decimals: number): void {
  const keep = exp + 1 + decimals;
  let all: string;
  if (keep < 0) {
    all = '';
  } else if (keep >= 15) {
    all = digits + zeros(keep - 15);
  } else {
    let n = keep > 0 ? Number(digits.substring(0, keep)) : 0;
    if (digits.charCodeAt(keep) >= 53) n++;
    all = n === 0 ? '' : String(n);
    if (all.length < keep) all = zeros(keep - all.length) + all;
  }
  const length = all.length;
  if (length > decimals) {
    rInt = all.substring(0, length - decimals);
    rFrac = all.substring(length - decimals);
  } else {
    rInt = '0';
    rFrac = length < decimals ? zeros(decimals - length) + all : all;
  }
}

/** Rounds v * 10^scale (v >= 0, finite) to `decimals` places into rInt / rFrac. */
function roundFixed(v: number, scale: number, decimals: number): void {
  if (v === 0 || (scale === 0 && v < 1e15 && v === Math.floor(v))) {
    rInt = String(v);
    rFrac = zeros(decimals);
    return;
  }
  // "d.dddddddddddddde+x"
  const text = v.toExponential(14);
  fixedFromDigits(text.charAt(0) + text.substring(2, 16), Number(text.substring(17)) + scale, decimals);
}

// ---------------------------------------------------------------------------
// Numbers

/**
 * Lays digits over the placeholders, right aligned: "0" pads with zeros, "?" with a
 * space, "#" with nothing; digits beyond the placeholders go in front of the first one.
 */
function renderDigits(toks: Tok[], count: number, digits: string, grouped: boolean, separator: string): string {
  let out = '';
  const length = digits.length;
  let position = count - 1;
  let first = true;
  for (let i = 0; i < toks.length; i++) {
    const tok = toks[i];
    if (tok.k === T_LIT) {
      out += tok.s;
      continue;
    }
    if (first) {
      first = false;
      for (let p = length - 1; p > position; p--) {
        out += digits.charAt(length - 1 - p);
        if (grouped && p % 3 === 0) out += separator;
      }
    }
    if (position < length) {
      out += digits.charAt(length - 1 - position);
      if (grouped && position > 0 && position % 3 === 0) out += separator;
    } else if (tok.k === T_ZERO) {
      out += '0';
      if (grouped && position > 0 && position % 3 === 0) out += separator;
    } else if (tok.k === T_QUESTION) {
      out += ' ';
    }
    position--;
  }
  return out;
}

/** Decimal placeholders: trailing zeros stay for "0", become a space for "?", vanish for "#". */
function renderDecimals(toks: Tok[], frac: string, lastNonZero: number): string {
  let out = '';
  let index = 0;
  for (let i = 0; i < toks.length; i++) {
    const tok = toks[i];
    if (tok.k === T_LIT) {
      out += tok.s;
      continue;
    }
    if (index <= lastNonZero || tok.k === T_ZERO) out += frac.charAt(index);
    else if (tok.k === T_QUESTION) out += ' ';
    index++;
  }
  return out;
}

function lastNonZeroIndex(digits: string): number {
  let last = digits.length - 1;
  while (last >= 0 && digits.charCodeAt(last) === 48) last--;
  return last;
}

function renderIntegerPart(section: Section, digits: string, locale: FormatLocale): string {
  if (section.intCount > 0) {
    return renderDigits(section.intToks, section.intCount, digits === '0' ? '' : digits, section.grouped, locale.group);
  }
  // ".00": no integer placeholder, but the integer digits are still shown.
  let out = '';
  for (let i = 0; i < section.intToks.length; i++) out += section.intToks[i].s;
  return digits === '0' ? out : out + digits;
}

function renderFixed(section: Section, v: number, minus: boolean, locale: FormatLocale): string {
  roundFixed(v, section.scale, section.decimals);
  const frac = rFrac;
  const last = lastNonZeroIndex(frac);
  // Excel keeps the sign of a negative number that rounds to zero: -0.004 shows as "-0.00".
  let out = minus ? '-' : '';
  out += renderIntegerPart(section, rInt, locale);
  if (section.hasPoint) out += locale.decimal + renderDecimals(section.decToks, frac, last);
  return out;
}

function renderScientific(section: Section, v: number, minus: boolean, locale: FormatLocale): string {
  let digits = ZEROS.substring(0, 15);
  let exponent = 0;
  if (v !== 0) {
    const text = v.toExponential(14);
    digits = text.charAt(0) + text.substring(2, 16);
    exponent = Number(text.substring(17)) + section.scale;
  }
  // With several integer placeholders the exponent moves in steps of that many ("##0.0E+0").
  const step = section.intCount > 1 ? section.intCount : 1;
  let shown = Math.floor(exponent / step) * step;
  fixedFromDigits(digits, exponent - shown, section.decimals);
  if (rInt.length > step) {
    // Rounding carried into a new digit: 999.96 -> 1.0E+3.
    shown += step;
    rInt = '1';
    rFrac = zeros(section.decimals);
  }
  const frac = rFrac;
  let out = minus && v !== 0 ? '-' : '';
  out += renderIntegerPart(section, rInt, locale);
  if (section.hasPoint) out += locale.decimal + renderDecimals(section.decToks, frac, lastNonZeroIndex(frac));
  out += 'E' + (shown < 0 ? '-' : section.expPlus ? '+' : '');
  const magnitude = String(shown < 0 ? -shown : shown);
  return out + renderDigits(section.expToks, section.expCount, magnitude, false, '');
}

// Result of bestFraction().
let fracNum = 0;
let fracDen = 1;

/** Closest num/den to x (0 <= x) with den <= maxDen; the smallest denominator wins ties. */
function bestFraction(x: number, maxDen: number): void {
  if (maxDen <= 999) {
    let bestError = Infinity;
    for (let den = 1; den <= maxDen; den++) {
      const num = Math.round(x * den);
      const error = Math.abs(x - num / den);
      if (error < bestError) {
        bestError = error;
        fracNum = num;
        fracDen = den;
        if (error < 1e-15) break;
      }
    }
    return;
  }
  // Large denominators: continued fraction convergents.
  let h0 = 0;
  let h1 = 1;
  let k0 = 1;
  let k1 = 0;
  let rest = x;
  for (let i = 0; i < 40; i++) {
    const a = Math.floor(rest);
    const h2 = a * h1 + h0;
    const k2 = a * k1 + k0;
    if (k2 > maxDen) break;
    h0 = h1;
    h1 = h2;
    k0 = k1;
    k1 = k2;
    const remainder = rest - a;
    if (remainder < 1e-12) break;
    rest = 1 / remainder;
  }
  fracNum = h1;
  fracDen = k1 === 0 ? 1 : k1;
  if (k1 === 0) fracNum = Math.round(x);
}

function renderFraction(section: Section, v: number, minus: boolean, locale: FormatLocale): string {
  const mixed = section.intCount > 0;
  let whole = mixed ? Math.floor(v) : 0;
  const part = v - whole;
  let num: number;
  let den: number;
  if (section.fixedDen > 0) {
    den = section.fixedDen;
    num = Math.round(Number((part * den).toPrecision(14)));
  } else {
    bestFraction(part, Math.pow(10, Math.min(section.denWidth, 7)) - 1);
    num = fracNum;
    den = fracDen;
  }
  if (mixed && num >= den) {
    whole += 1;
    num -= den;
  }
  let out = minus ? '-' : '';
  const numCount = section.numToks.length;
  if (mixed) {
    // "0 1/2" is written " 1/2": the whole part is only shown when there is one, or nothing else.
    const digits = whole === 0 && num !== 0 ? '' : String(whole);
    out += renderDigits(section.intToks, section.intCount, digits, section.grouped, locale.group);
  } else {
    for (let i = 0; i < section.intToks.length; i++) out += section.intToks[i].s;
  }
  if (mixed && num === 0 && !section.numForced) {
    // A whole number keeps its width: blanks where the fraction would be.
    out += repeat(' ', numCount + 1 + section.denWidth);
  } else {
    out += renderDigits(section.numToks, numCount, String(num), false, '') + '/';
    let text = String(den);
    if (text.length < section.denZeros) text = zeros(section.denZeros - text.length) + text;
    out += text;
    if (section.denPadded && text.length < section.denWidth) out += repeat(' ', section.denWidth - text.length);
  }
  for (let i = 0; i < section.tailToks.length; i++) out += section.tailToks[i].s;
  return out;
}

function generalText(v: number, locale: FormatLocale): string {
  const text = formatGeneral(v);
  return locale.decimal === '.' ? text : text.replace('.', locale.decimal);
}

function renderSlots(toks: Tok[], slot: string): string {
  let out = '';
  for (let i = 0; i < toks.length; i++) {
    const tok = toks[i];
    out += tok.k === T_SLOT ? slot : tok.s;
    if (out.length > MAX_TEXT_OUTPUT) break;
  }
  return out;
}

// ---------------------------------------------------------------------------
// Dates

// Result of computeDate().
let dYear = 0;
let dMonth = 0;
let dDay = 0;
let dHour = 0;
let dMinute = 0;
let dSecond = 0;
/** Fraction of the second in units of 10^-subDigits seconds. */
let dSub = 0;
let dWeekday = 0;
let dDays = 0;

const POWERS = [1, 10, 100, 1000];

/** Days since 1 Jan 1970 -> dYear / dMonth / dDay (proleptic Gregorian calendar). */
function civilFromDays(days: number): void {
  const z = days + 719468;
  const era = Math.floor(z / 146097);
  const dayOfEra = z - era * 146097;
  const yearOfEra = Math.floor(
    (dayOfEra - Math.floor(dayOfEra / 1460) + Math.floor(dayOfEra / 36524) - Math.floor(dayOfEra / 146096)) / 365,
  );
  const dayOfYear = dayOfEra - (365 * yearOfEra + Math.floor(yearOfEra / 4) - Math.floor(yearOfEra / 100));
  const shifted = Math.floor((5 * dayOfYear + 2) / 153);
  dDay = dayOfYear - Math.floor((153 * shifted + 2) / 5) + 1;
  dMonth = shifted < 10 ? shifted + 3 : shifted - 9;
  dYear = yearOfEra + era * 400 + (dMonth <= 2 ? 1 : 0);
}

/**
 * Fills the d* variables for a serial; the time is rounded to 10^-subDigits seconds,
 * which may carry into the next day. False when the serial is not a date.
 */
function computeDate(serial: number, date1904: boolean, subDigits: number): boolean {
  if (!(serial >= 0) || serial > MAX_SERIAL_1900 + 1) return false;
  let days = Math.floor(serial);
  const perSecond = POWERS[subDigits];
  const perDay = 86400 * perSecond;
  let time = Math.round((serial - days) * perDay);
  if (time >= perDay) {
    time = 0;
    days++;
  }
  if (days > (date1904 ? MAX_SERIAL_1904 : MAX_SERIAL_1900)) return false;
  dDays = days;
  if (date1904) {
    civilFromDays(days - 24107);
    dWeekday = (days + 5) % 7;
  } else {
    // The 1900 system counts a 29 Feb 1900 that never existed, and shows serial 0 as "0 January".
    if (days === 60) {
      dYear = 1900;
      dMonth = 2;
      dDay = 29;
    } else if (days === 0) {
      dYear = 1900;
      dMonth = 1;
      dDay = 0;
    } else {
      civilFromDays(days < 60 ? days - 25568 : days - 25569);
    }
    dWeekday = (days + 6) % 7;
  }
  const seconds = Math.floor(time / perSecond);
  dSub = time - seconds * perSecond;
  dHour = Math.floor(seconds / 3600);
  dMinute = Math.floor(seconds / 60) % 60;
  dSecond = seconds % 60;
  return true;
}

function pad(value: number, width: number): string {
  const text = String(value);
  return text.length >= width ? text : zeros(width - text.length) + text;
}

/** The d* variables through a date section; null when the serial is not a date. */
function renderDate(section: Section, serial: number, locale: FormatLocale, date1904: boolean): string | null {
  if (!computeDate(serial, date1904, section.subDigits)) return null;
  const names = section.names || locale;
  const toks = section.dateToks;
  let out = '';
  for (let i = 0; i < toks.length; i++) {
    const tok = toks[i];
    const n = tok.n;
    switch (tok.k) {
      case D_LIT:
        out += tok.s;
        break;
      case D_YEAR:
        out += n <= 2 ? pad(dYear % 100, 2) : String(dYear);
        break;
      case D_BUDDHIST_YEAR:
        out += n <= 2 ? pad((dYear + 543) % 100, 2) : String(dYear + 543);
        break;
      case D_MONTH:
        if (n === 1) out += String(dMonth);
        else if (n === 2) out += pad(dMonth, 2);
        else if (n === 3) out += names.shortMonthNames[dMonth - 1];
        else if (n === 5) out += names.monthNames[dMonth - 1].charAt(0);
        else out += names.monthNames[dMonth - 1];
        break;
      case D_DAY:
        if (n === 1) out += String(dDay);
        else if (n === 2) out += pad(dDay, 2);
        else if (n === 3) out += names.shortDayNames[dWeekday];
        else out += names.dayNames[dWeekday];
        break;
      case D_HOUR: {
        const hour = section.hour12 ? dHour % 12 || 12 : dHour;
        out += n === 1 ? String(hour) : pad(hour, 2);
        break;
      }
      case D_MINUTE:
        out += n === 1 ? String(dMinute) : pad(dMinute, 2);
        break;
      case D_SECOND:
        out += n === 1 ? String(dSecond) : pad(dSecond, 2);
        break;
      case D_SUBSECOND:
        out += locale.decimal + pad(dSub, section.subDigits).substring(0, n);
        break;
      case D_AMPM:
        if (n === 0) out += dHour < 12 ? names.am : names.pm;
        else if (n === 2) out += dHour < 12 ? 'A' : 'P';
        else out += dHour < 12 ? 'a' : 'p';
        break;
      case D_ELAPSED_HOURS:
        out += pad(dDays * 24 + dHour, n);
        break;
      case D_ELAPSED_MINUTES:
        out += pad((dDays * 24 + dHour) * 60 + dMinute, n);
        break;
      case D_ELAPSED_SECONDS:
        out += pad(((dDays * 24 + dHour) * 60 + dMinute) * 60 + dSecond, n);
        break;
      case D_DATE_SEPARATOR:
        out += locale.dateSeparator;
        break;
      case D_TIME_SEPARATOR:
        out += locale.timeSeparator;
        break;
      default:
        break;
    }
  }
  return out;
}

// ---------------------------------------------------------------------------
// Section choice

/** Set by pickSection(): the chosen section shows negative numbers without a minus sign. */
let dropSign = false;

/** The section for a number, or null when no section's condition accepts it. */
function pickSection(compiled: Compiled, v: number): Section | null {
  const list = compiled.numeric;
  const count = list.length;
  const first = list[0];
  dropSign = false;
  if (!compiled.conditional) {
    if (count === 1) return first;
    if (v < 0) {
      dropSign = true;
      return list[1];
    }
    return count === 2 || v > 0 ? first : list[2];
  }
  if (first.cond ? testCondition(first.cond, v) : count >= 3 ? v > 0 : v >= 0) {
    dropSign = first.cond !== null && negativeOnly(first.cond);
    return first;
  }
  if (count === 1) return null;
  const second = list[1];
  if (second.cond) {
    if (testCondition(second.cond, v)) {
      dropSign = negativeOnly(second.cond);
      return second;
    }
    return count === 3 ? list[2] : null;
  }
  // Two sections: the second takes everything the first refused, sign included.
  if (count === 2) return second;
  if (v < 0) {
    dropSign = true;
    return second;
  }
  return list[2];
}

function withColor(text: string, color: string | undefined): FormattedValue {
  return color === undefined ? { text } : { text, color };
}

function formatNumber(v: number, compiled: Compiled, locale: FormatLocale, date1904: boolean): FormattedValue {
  if (v !== v || v === Infinity || v === -Infinity) return { text: NUMBER_ERROR };
  if (compiled.numeric.length === 0) {
    // "@" alone, or a code that could not be compiled.
    return withColor(generalText(v, locale), compiled.text ? compiled.text.color : undefined);
  }
  let section = pickSection(compiled, v);
  if (section === null) {
    // A lone section whose condition fails leaves the value to General; with two
    // conditions and no section for the rest the value cannot be shown.
    if (compiled.numeric.length === 1) return { text: generalText(v, locale) };
    return { text: OVERFLOW_TEXT, overflow: true };
  }
  const color = section.color;
  const minus = v < 0 && !dropSign;
  const abs = v < 0 ? -v : v;

  if (section.kind === K_SYSTEM_DATE || section.kind === K_SYSTEM_TIME) {
    const code =
      section.kind === K_SYSTEM_DATE
        ? locale.longDate || DEFAULT_LONG_DATE
        : locale.longTime || DEFAULT_LONG_TIME;
    const system = getCompiled(code).numeric;
    if (system.length === 0 || system[0].kind !== K_DATE) return withColor(generalText(v, locale), color);
    section = system[0];
  }

  switch (section.kind) {
    case K_NUMBER: {
      const text = section.fraction
        ? abs < 1e15
          ? renderFraction(section, abs, minus, locale)
          : generalText(v, locale)
        : section.scientific
          ? renderScientific(section, abs, minus, locale)
          : renderFixed(section, abs, minus, locale);
      return color === undefined ? { text } : { text, color };
    }
    case K_DATE: {
      // The 1904 system shows negative dates and durations: the date of the absolute value, signed.
      const signed = v < 0 && date1904;
      let text = renderDate(section, signed ? abs : v, locale, date1904);
      if (text === null) return { text: OVERFLOW_TEXT, overflow: true };
      if (signed && minus) text = '-' + text;
      return color === undefined ? { text } : { text, color };
    }
    case K_GENERAL:
    case K_TEXT:
      return withColor((minus ? '-' : '') + renderSlots(section.slotToks, generalText(abs, locale)), color);
    default:
      return withColor(minus ? '-' + section.literal : section.literal, color);
  }
}

function formatText(value: string, compiled: Compiled): FormattedValue {
  const section = compiled.text;
  if (section === null) return { text: value };
  if (section.kind === K_TEXT || section.kind === K_GENERAL) {
    return withColor(renderSlots(section.slotToks, value), section.color);
  }
  if (section.kind === K_LITERAL) return withColor(section.literal, section.color);
  // A fourth section made of digit or date placeholders says nothing about text.
  return withColor(value, section.color);
}

// ---------------------------------------------------------------------------
// Public interface

/**
 * Text of a cell value under a number format code. Never throws: a code that cannot be
 * understood shows the value the "General" way.
 */
export function formatValue(
  value: number | string | boolean | null,
  code: string,
  options?: FormatOptions,
): FormattedValue {
  if (value === null || value === undefined) return { text: '' };
  let locale = options === undefined ? LOCALE_EN : (options.locale as FormatLocale);
  if (locale !== LOCALE_EN && locale !== LOCALE_TR) locale = normalizeLocale(locale);
  if (typeof value === 'boolean') return { text: value ? locale.trueText : locale.falseText };
  if (options !== undefined && options.builtinId !== undefined) {
    if (options.builtinId === BUILTIN_SHORT_DATE) {
      code = locale.shortDate || DEFAULT_SHORT_DATE;
    } else if (options.builtinId === BUILTIN_SHORT_DATE_TIME) {
      code = (locale.shortDate || DEFAULT_SHORT_DATE) + ' ' + (locale.shortTime || DEFAULT_SHORT_TIME);
    }
  }
  const compiled = getCompiled(code);
  if (typeof value === 'string') return formatText(value, compiled);
  if (typeof value !== 'number') return { text: String(value) };
  return formatNumber(value, compiled, locale, options !== undefined && options.date1904 === true);
}

/** Number of compiled codes held in the cache; exposed for tests of its size limit. */
export function compiledFormatCount(): number {
  return cache.size;
}

/** True when the code shows numbers as dates or times (any of its number sections). */
export function isDateFormat(code: string): boolean {
  return getCompiled(code).date;
}

/**
 * Serial -> calendar parts, honouring Excel's 1900 leap year bug (serial 60 = 29 Feb 1900)
 * and the 1904 system. The time is rounded to the millisecond. Serial 0 of the 1900 system
 * is day 0 of January 1900, as Excel shows it. `weekday` is 0 for Sunday. Every part is
 * NaN when the serial is negative, not a number or past 31 Dec 9999.
 */
export function serialToDate(
  serial: number,
  date1904?: boolean,
): {
  year: number;
  month: number;
  day: number;
  hour: number;
  minute: number;
  second: number;
  millisecond: number;
  weekday: number;
} {
  if (typeof serial !== 'number' || !computeDate(serial, date1904 === true, 3)) {
    return {
      year: NaN, month: NaN, day: NaN, hour: NaN, minute: NaN, second: NaN, millisecond: NaN, weekday: NaN,
    };
  }
  return {
    year: dYear,
    month: dMonth,
    day: dDay,
    hour: dHour,
    minute: dMinute,
    second: dSecond,
    millisecond: dSub,
    weekday: dWeekday,
  };
}

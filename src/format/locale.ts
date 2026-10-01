// Locale data for the number format engine. Kept as plain tables: `Intl` is
// not used because its output differs between engines and WebView versions.

export interface FormatLocale {
  decimal: string;
  group: string;
  /** Date separator that "/" in a format code stands for, and time separator for ":". */
  dateSeparator: string;
  timeSeparator: string;
  monthNames: string[];
  shortMonthNames: string[];
  dayNames: string[];
  shortDayNames: string[];
  am: string;
  pm: string;
  trueText: string;
  falseText: string;
  /** Format code of the system short date (what built-in format 14 shows). Default "m/d/yyyy". */
  shortDate?: string;
  /** Format code of the system long date (what "[$-F800]" shows). Default "dddd, mmmm d, yyyy". */
  longDate?: string;
  /** Format code of the system time (what "[$-F400]" shows). Default "h:mm:ss AM/PM". */
  longTime?: string;
  /** Format code of the system short time (the time part of built-in format 22). Default "h:mm". */
  shortTime?: string;
  /** Order of day, month and year in typed dates such as "3/4/2024". Default "mdy". */
  dateOrder?: 'dmy' | 'mdy' | 'ymd';
}

export const LOCALE_EN: FormatLocale = {
  decimal: '.',
  group: ',',
  dateSeparator: '/',
  timeSeparator: ':',
  monthNames: [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December',
  ],
  shortMonthNames: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
  dayNames: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
  shortDayNames: ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'],
  am: 'AM',
  pm: 'PM',
  trueText: 'TRUE',
  falseText: 'FALSE',
  shortDate: 'm/d/yyyy',
  longDate: 'dddd, mmmm d, yyyy',
  longTime: 'h:mm:ss AM/PM',
  shortTime: 'h:mm',
  dateOrder: 'mdy',
};

export const LOCALE_TR: FormatLocale = {
  decimal: ',',
  group: '.',
  dateSeparator: '.',
  timeSeparator: ':',
  monthNames: [
    'Ocak', 'Şubat', 'Mart', 'Nisan', 'Mayıs', 'Haziran',
    'Temmuz', 'Ağustos', 'Eylül', 'Ekim', 'Kasım', 'Aralık',
  ],
  shortMonthNames: ['Oca', 'Şub', 'Mar', 'Nis', 'May', 'Haz', 'Tem', 'Ağu', 'Eyl', 'Eki', 'Kas', 'Ara'],
  dayNames: ['Pazar', 'Pazartesi', 'Salı', 'Çarşamba', 'Perşembe', 'Cuma', 'Cumartesi'],
  shortDayNames: ['Paz', 'Pzt', 'Sal', 'Çar', 'Per', 'Cum', 'Cmt'],
  am: 'ÖÖ',
  pm: 'ÖS',
  trueText: 'DOĞRU',
  falseText: 'YANLIŞ',
  shortDate: 'dd/mm/yyyy',
  longDate: 'd mmmm yyyy dddd',
  longTime: 'hh:mm:ss',
  shortTime: 'hh:mm',
  dateOrder: 'dmy',
};

/** Locale for a BCP 47 tag: "tr", "tr-TR", "tr_TR" give Turkish; everything else English. */
export function getLocale(tag: string | undefined): FormatLocale {
  if (typeof tag !== 'string') return LOCALE_EN;
  const lower = tag.toLowerCase();
  if (lower === 'tr' || lower.indexOf('tr-') === 0 || lower.indexOf('tr_') === 0) return LOCALE_TR;
  return LOCALE_EN;
}

const STRING_FIELDS: (keyof FormatLocale)[] = [
  'decimal', 'group', 'dateSeparator', 'timeSeparator', 'am', 'pm', 'trueText', 'falseText',
  'shortDate', 'longDate', 'longTime', 'shortTime',
];
const LIST_FIELDS: [keyof FormatLocale, number][] = [
  ['monthNames', 12], ['shortMonthNames', 12], ['dayNames', 7], ['shortDayNames', 7],
];

const normalized = new WeakMap<object, FormatLocale>();

function isStringList(value: unknown, length: number): boolean {
  if (!Array.isArray(value) || value.length !== length) return false;
  for (let i = 0; i < length; i++) if (typeof value[i] !== 'string') return false;
  return true;
}

/**
 * A complete locale from a partial or untyped one: missing fields and fields of the wrong
 * type (name lists must have 12 / 7 strings) are taken from LOCALE_EN. The result is cached
 * per input object, so the input must not be mutated afterwards.
 */
export function normalizeLocale(locale: Partial<FormatLocale> | undefined | null): FormatLocale {
  if (locale === LOCALE_EN || locale === LOCALE_TR) return locale as FormatLocale;
  if (locale === undefined || locale === null || typeof locale !== 'object') return LOCALE_EN;
  const cached = normalized.get(locale);
  if (cached !== undefined) return cached;
  const source = locale as { [field: string]: unknown };
  const result = {} as { [field: string]: unknown };
  const fallback = LOCALE_EN as unknown as { [field: string]: unknown };
  for (let i = 0; i < STRING_FIELDS.length; i++) {
    const field = STRING_FIELDS[i];
    result[field] = typeof source[field] === 'string' ? source[field] : fallback[field];
  }
  for (let i = 0; i < LIST_FIELDS.length; i++) {
    const field = LIST_FIELDS[i][0];
    result[field] = isStringList(source[field], LIST_FIELDS[i][1])
      ? (source[field] as string[]).slice()
      : fallback[field];
  }
  const order = source.dateOrder;
  result.dateOrder = order === 'dmy' || order === 'mdy' || order === 'ymd' ? order : fallback.dateOrder;
  const complete = result as unknown as FormatLocale;
  normalized.set(locale, complete);
  return complete;
}

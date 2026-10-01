// Date and time functions. Dates are serial numbers: days since the start of
// the workbook's date system, with the time of day as the fraction. The 1900
// system keeps Excel's 29 February 1900; the 1904 system does not have it.

import {
  ERROR_NUM,
  ERROR_VALUE,
  isError,
  toBoolean,
  toText,
  type CalcContext,
  type CalcError,
  type CalcFunction,
  type FunctionSpec,
  type Scalar,
  type Value,
} from '../values';
import {
  dateParts,
  dateSerial,
  daysInMonth,
  eachNumber,
  isLeapYear,
  mapScalars,
  mapScalarsExcept,
  maxSerial,
  monthStart,
  parseDateText,
  scalarNumber,
  scalarSerial,
  spec,
  weekdayOf,
  type DateParts,
} from './helpers';

const SECONDS_PER_DAY = 86400;
/** Largest hour, minute or second TIME accepts. */
const MAX_TIME_PART = 32767;
/** Months and days beyond this cannot land inside the calendar whatever the year is. */
const MAX_DATE_PART = 10000000;

function whole(value: number): number {
  return value < 0 ? Math.ceil(value) : Math.floor(value);
}

function inCalendar(serial: number, context: CalcContext): number | CalcError {
  return serial >= 0 && serial <= maxSerial(context.date1904 === true) ? serial : ERROR_NUM;
}

/** A date argument as a whole serial with its calendar parts. */
function dateArg(value: Scalar, context: CalcContext): { serial: number; parts: DateParts } | CalcError {
  const serial = scalarSerial(value, context);
  if (typeof serial !== 'number') return serial;
  const parts = dateParts(serial, context.date1904 === true);
  if (parts === undefined) return ERROR_NUM;
  return { serial: Math.floor(serial), parts };
}

const date = mapScalars((args, context) => {
  const y = scalarNumber(args[0], context);
  if (typeof y !== 'number') return y;
  const m = scalarNumber(args[1], context);
  if (typeof m !== 'number') return m;
  const d = scalarNumber(args[2], context);
  if (typeof d !== 'number') return d;
  let year = Math.floor(y);
  const month = whole(m);
  const day = whole(d);
  if (!(year >= 0 && year < 10000) || !(Math.abs(month) <= MAX_DATE_PART) || !(Math.abs(day) <= MAX_DATE_PART)) {
    return ERROR_NUM;
  }
  if (year < 1900) year += 1900;
  return inCalendar(dateSerial(year, month, day, context.date1904 === true), context);
});

const time = mapScalars((args, context) => {
  let seconds = 0;
  const weights = [3600, 60, 1];
  for (let i = 0; i < 3; i++) {
    const part = scalarNumber(args[i], context);
    if (typeof part !== 'number') return part;
    if (!(Math.abs(part) <= MAX_TIME_PART)) return ERROR_NUM;
    seconds += whole(part) * weights[i];
  }
  if (seconds < 0) return ERROR_NUM;
  return (seconds % SECONDS_PER_DAY) / SECONDS_PER_DAY;
});

function datePart(pick: (parts: DateParts) => number): CalcFunction {
  return mapScalars((args, context) => {
    const arg = dateArg(args[0], context);
    return isError(arg) ? arg : pick(arg.parts);
  });
}

function timePart(pick: (seconds: number) => number): CalcFunction {
  return mapScalars((args, context) => {
    const serial = scalarNumber(args[0], context);
    if (typeof serial !== 'number') return serial;
    if (!(serial >= 0) || !isFinite(serial)) return ERROR_NUM;
    // Times are read to the nearest second; one that rounds up to midnight shows as 0:00:00.
    const seconds = Math.round((serial - Math.floor(serial)) * SECONDS_PER_DAY) % SECONDS_PER_DAY;
    return pick(seconds);
  });
}

/** An optional whole number: the fallback when it is not given at all; given but empty ("WEEKDAY(d,)") it is 0. */
function optionalInteger(args: Scalar[], index: number, fallback: number, context: CalcContext): number | CalcError {
  if (index >= args.length) return fallback;
  const value = scalarNumber(args[index], context);
  return typeof value === 'number' ? whole(value) : value;
}

/** First day of the week (0 for Sunday) for the "return type" numbers shared by WEEKDAY and WEEKNUM. */
function weekStart(type: number): number | undefined {
  if (type === 1 || type === 17) return 0;
  if (type === 2 || type === 11) return 1;
  if (type >= 12 && type <= 16) return type - 10;
  return undefined;
}

const weekday = mapScalars((args, context) => {
  const arg = dateArg(args[0], context);
  if (isError(arg)) return arg;
  const type = optionalInteger(args, 1, 1, context);
  if (typeof type !== 'number') return type;
  const day = weekdayOf(arg.serial, context.date1904 === true);
  // Type 3 numbers Monday to Sunday 0 to 6; every other type runs from 1 to 7.
  if (type === 3) return (day + 6) % 7;
  const start = weekStart(type);
  if (start === undefined) return ERROR_NUM;
  return ((day - start + 7) % 7) + 1;
});

function isoWeek(serial: number, date1904: boolean): number {
  // The week belongs to the year its Thursday is in.
  const thursday = serial - ((weekdayOf(serial, date1904) + 6) % 7) + 3;
  const parts = dateParts(Math.max(thursday, 0), date1904) || dateParts(serial, date1904);
  if (parts === undefined) return 1;
  return Math.floor((thursday - dateSerial(parts.year, 1, 1, date1904)) / 7) + 1;
}

const weekNumber = mapScalars((args, context) => {
  const arg = dateArg(args[0], context);
  if (isError(arg)) return arg;
  const type = optionalInteger(args, 1, 1, context);
  if (typeof type !== 'number') return type;
  const date1904 = context.date1904 === true;
  if (type === 21) return isoWeek(arg.serial, date1904);
  const start = weekStart(type);
  if (start === undefined) return ERROR_NUM;
  // Week 1 is the week that holds 1 January.
  const first = dateSerial(arg.parts.year, 1, 1, date1904);
  const offset = (weekdayOf(first, date1904) - start + 7) % 7;
  return Math.floor((arg.serial - first + offset) / 7) + 1;
});

const isoWeekNumber = mapScalars((args, context) => {
  const arg = dateArg(args[0], context);
  return isError(arg) ? arg : isoWeek(arg.serial, context.date1904 === true);
});

function monthShift(endOfMonth: boolean): CalcFunction {
  return mapScalars((args, context) => {
    const arg = dateArg(args[0], context);
    if (isError(arg)) return arg;
    const shift = scalarNumber(args[1], context);
    if (typeof shift !== 'number') return shift;
    const months = whole(shift);
    if (!(Math.abs(months) <= MAX_DATE_PART)) return ERROR_NUM;
    const date1904 = context.date1904 === true;
    const month = arg.parts.month + months;
    if (endOfMonth) return inCalendar(monthStart(arg.parts.year, month + 1, date1904) - 1, context);
    const day = Math.min(arg.parts.day, daysInMonth(arg.parts.year, month, date1904));
    return inCalendar(dateSerial(arg.parts.year, month, day, date1904), context);
  });
}

const dateDif = mapScalars((args, context) => {
  const start = dateArg(args[0], context);
  if (isError(start)) return start;
  const end = dateArg(args[1], context);
  if (isError(end)) return end;
  const unit = toText(args[2]);
  if (typeof unit !== 'string') return unit;
  if (start.serial > end.serial) return ERROR_NUM;
  const a = start.parts;
  const b = end.parts;
  const date1904 = context.date1904 === true;
  const dayShort = b.day < a.day ? 1 : 0;
  const yearShort = b.month < a.month || (b.month === a.month && b.day < a.day) ? 1 : 0;
  switch (unit.toUpperCase()) {
    case 'Y':
      return b.year - a.year - yearShort;
    case 'M':
      return (b.year - a.year) * 12 + b.month - a.month - dayShort;
    case 'D':
      return end.serial - start.serial;
    case 'MD':
      // Days left after whole months: counted from the start day in the month before the end.
      return dayShort === 0 ? b.day - a.day : end.serial - dateSerial(b.year, b.month - 1, a.day, date1904);
    case 'YM':
      return (((b.month - a.month - dayShort) % 12) + 12) % 12;
    case 'YD':
      return end.serial - dateSerial(b.year - yearShort, a.month, a.day, date1904);
    default:
      return ERROR_NUM;
  }
});

const days = mapScalars((args, context) => {
  const end = scalarSerial(args[0], context);
  if (typeof end !== 'number') return end;
  const start = scalarSerial(args[1], context);
  if (typeof start !== 'number') return start;
  return Math.floor(end) - Math.floor(start);
});

function isLastOfFebruary(parts: DateParts, date1904: boolean): boolean {
  return parts.month === 2 && parts.day === daysInMonth(parts.year, 2, date1904);
}

/** Days between two dates on the 30/360 calendar, US (NASD) or European rules. */
function days360(a: DateParts, b: DateParts, european: boolean, date1904: boolean): number {
  let d1 = a.day;
  let d2 = b.day;
  let m2 = b.month;
  if (european) {
    if (d1 > 30) d1 = 30;
    if (d2 > 30) d2 = 30;
  } else {
    if (d1 === 31 || isLastOfFebruary(a, date1904)) d1 = 30;
    if (d2 === 31) {
      if (d1 < 30) {
        d2 = 1;
        m2++;
      } else d2 = 30;
    }
  }
  return (b.year - a.year) * 360 + (m2 - a.month) * 30 + (d2 - d1);
}

const days360Function = mapScalars((args, context) => {
  const start = dateArg(args[0], context);
  if (isError(start)) return start;
  const end = dateArg(args[1], context);
  if (isError(end)) return end;
  const european = args.length > 2 ? toBoolean(args[2]) : false;
  if (typeof european !== 'boolean') return european;
  return days360(start.parts, end.parts, european, context.date1904 === true);
});

// ---------------------------------------------------------------------------
// Working days (Saturday and Sunday off)

function isWeekend(serial: number, date1904: boolean): boolean {
  const day = weekdayOf(serial, date1904);
  return day === 0 || day === 6;
}

/** The holidays that fall on working days, without repeats, in ascending order. */
function holidayList(arg: Value | undefined, context: CalcContext): number[] | CalcError {
  if (arg === undefined || arg === null) return [];
  const date1904 = context.date1904 === true;
  const seen = new Set<number>();
  const error = eachNumber([arg], context, undefined, (value) => {
    const day = Math.floor(value);
    if (day >= 0 && !isWeekend(day, date1904)) seen.add(day);
  });
  if (error !== undefined) return error;
  const list: number[] = [];
  seen.forEach((day) => {
    list.push(day);
  });
  return list.sort((x, y) => x - y);
}

/** Working days from `low` to `high`, both included. */
function workingDays(low: number, high: number, date1904: boolean): number {
  const span = high - low + 1;
  const weeks = Math.floor(span / 7);
  let count = weeks * 5;
  for (let day = low + weeks * 7; day <= high; day++) if (!isWeekend(day, date1904)) count++;
  return count;
}

const networkDays = mapScalarsExcept(
  (index) => index === 2,
  (args, context) => {
    const start = scalarSerial(args[0] as Scalar, context);
    if (typeof start !== 'number') return start;
    const end = scalarSerial(args[1] as Scalar, context);
    if (typeof end !== 'number') return end;
    const holidays = holidayList(args[2], context);
    if (isError(holidays)) return holidays;
    const low = Math.floor(Math.min(start, end));
    const high = Math.floor(Math.max(start, end));
    let count = workingDays(low, high, context.date1904 === true);
    for (let i = 0; i < holidays.length; i++) if (holidays[i] >= low && holidays[i] <= high) count--;
    return start > end ? -count : count;
  },
);

/** The date `count` working days after (or, negative, before) a date, weekends skipped. */
function shiftWorkingDays(start: number, count: number, date1904: boolean): number {
  const step = count < 0 ? -1 : 1;
  let day = start;
  // From a weekend the count starts at the working day just behind it.
  for (let i = 0; i < 2 && isWeekend(day, date1904); i++) day -= step;
  const total = Math.abs(count);
  day += step * Math.floor(total / 5) * 7;
  for (let left = total % 5; left > 0; left--) {
    day += step;
    for (let i = 0; i < 2 && isWeekend(day, date1904); i++) day += step;
  }
  return day;
}

const workDay = mapScalarsExcept(
  (index) => index === 2,
  (args, context) => {
    const from = scalarSerial(args[0] as Scalar, context);
    if (typeof from !== 'number') return from;
    const shift = scalarNumber(args[1] as Scalar, context);
    if (typeof shift !== 'number') return shift;
    const holidays = holidayList(args[2], context);
    if (isError(holidays)) return holidays;
    const date1904 = context.date1904 === true;
    const start = Math.floor(from);
    const count = whole(shift);
    const limit = maxSerial(date1904);
    if (!(Math.abs(count) <= limit)) return ERROR_NUM;
    if (count === 0) return start;
    let result = shiftWorkingDays(start, count, date1904);
    // Every holiday passed on the way costs one more day; passing it may reach further holidays.
    let passed = 0;
    if (count > 0) {
      let next = 0;
      while (next < holidays.length && holidays[next] <= start) next++;
      const base = next;
      for (;;) {
        while (next < holidays.length && holidays[next] <= result) next++;
        if (next - base === passed) break;
        passed = next - base;
        result = shiftWorkingDays(start, count + passed, date1904);
      }
    } else {
      let next = holidays.length - 1;
      while (next >= 0 && holidays[next] >= start) next--;
      const base = next;
      for (;;) {
        while (next >= 0 && holidays[next] >= result) next--;
        if (base - next === passed) break;
        passed = base - next;
        result = shiftWorkingDays(start, count - passed, date1904);
      }
    }
    return inCalendar(result, context);
  },
);

// ---------------------------------------------------------------------------

function yearLength(start: number, end: number, a: DateParts, b: DateParts, date1904: boolean): number {
  const withinYear = a.year === b.year || (b.year === a.year + 1 && (a.month > b.month || (a.month === b.month && a.day >= b.day)));
  if (!withinYear) {
    return (dateSerial(b.year + 1, 1, 1, date1904) - dateSerial(a.year, 1, 1, date1904)) / (b.year - a.year + 1);
  }
  if (a.year === b.year) return isLeapYear(a.year, date1904) ? 366 : 365;
  const leapDayInside =
    (isLeapYear(a.year, date1904) && start <= dateSerial(a.year, 2, 29, date1904)) ||
    (isLeapYear(b.year, date1904) && end >= dateSerial(b.year, 2, 29, date1904));
  return leapDayInside ? 366 : 365;
}

const yearFrac = mapScalars((args, context) => {
  let start = dateArg(args[0], context);
  if (isError(start)) return start;
  let end = dateArg(args[1], context);
  if (isError(end)) return end;
  const basis = optionalInteger(args, 2, 0, context);
  if (typeof basis !== 'number') return basis;
  if (basis < 0 || basis > 4) return ERROR_NUM;
  if (start.serial > end.serial) {
    const swap = start;
    start = end;
    end = swap;
  }
  const date1904 = context.date1904 === true;
  const span = end.serial - start.serial;
  switch (basis) {
    case 0: {
      // US 30/360 as Excel applies it here, which is not how DAYS360 does: the rules are
      // alternatives, the first that fits is the only one applied. So an end on the 31st
      // stays a 31st when the start is the last day of February (28 Feb to 31 Mar is 31 days).
      const a = start.parts;
      const b = end.parts;
      let d1 = a.day;
      let d2 = b.day;
      if (d1 === 31 && d2 === 31) {
        d1 = 30;
        d2 = 30;
      } else if (d1 === 31) d1 = 30;
      else if (d1 === 30 && d2 === 31) d2 = 30;
      else if (isLastOfFebruary(a, date1904) && isLastOfFebruary(b, date1904)) {
        d1 = 30;
        d2 = 30;
      } else if (isLastOfFebruary(a, date1904)) d1 = 30;
      return ((b.year - a.year) * 360 + (b.month - a.month) * 30 + (d2 - d1)) / 360;
    }
    case 1:
      return span / yearLength(start.serial, end.serial, start.parts, end.parts, date1904);
    case 2:
      return span / 360;
    case 3:
      return span / 365;
    default:
      return days360(start.parts, end.parts, true, date1904) / 360;
  }
});

function fromText(pick: (serial: number) => number): CalcFunction {
  return mapScalars((args, context) => {
    const arg = args[0];
    if (isError(arg)) return arg;
    if (typeof arg !== 'string') return ERROR_VALUE;
    const serial = parseDateText(arg, context);
    return serial === undefined ? ERROR_VALUE : pick(serial);
  });
}

function now(context: CalcContext): number | CalcError {
  const value = context.now();
  return value >= 0 && isFinite(value) ? value : ERROR_NUM;
}

export const DATE_FUNCTIONS: { [name: string]: FunctionSpec } = {
  DATE: spec(date, 3, 3),
  TIME: spec(time, 3, 3),
  NOW: spec((_args, context) => now(context), 0, 0),
  TODAY: spec((_args, context) => {
    const value = now(context);
    return typeof value === 'number' ? Math.floor(value) : value;
  }, 0, 0),
  YEAR: spec(datePart((parts) => parts.year), 1, 1),
  MONTH: spec(datePart((parts) => parts.month), 1, 1),
  DAY: spec(datePart((parts) => parts.day), 1, 1),
  HOUR: spec(timePart((seconds) => Math.floor(seconds / 3600)), 1, 1),
  MINUTE: spec(timePart((seconds) => Math.floor(seconds / 60) % 60), 1, 1),
  SECOND: spec(timePart((seconds) => seconds % 60), 1, 1),
  WEEKDAY: spec(weekday, 1, 2),
  WEEKNUM: spec(weekNumber, 1, 2),
  ISOWEEKNUM: spec(isoWeekNumber, 1, 1),
  EDATE: spec(monthShift(false), 2, 2),
  EOMONTH: spec(monthShift(true), 2, 2),
  DATEDIF: spec(dateDif, 3, 3),
  DAYS: spec(days, 2, 2),
  DAYS360: spec(days360Function, 2, 3),
  NETWORKDAYS: spec(networkDays, 2, 3),
  WORKDAY: spec(workDay, 2, 3),
  YEARFRAC: spec(yearFrac, 2, 3),
  DATEVALUE: spec(fromText((serial) => Math.floor(serial)), 1, 1),
  TIMEVALUE: spec(fromText((serial) => serial - Math.floor(serial)), 1, 1),
};

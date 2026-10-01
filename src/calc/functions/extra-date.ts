// Working days with a weekend of one's own choice: NETWORKDAYS.INTL and
// WORKDAY.INTL.

import {
  ERROR_NUM,
  ERROR_VALUE,
  isArray,
  isError,
  isRef,
  type CalcContext,
  type CalcError,
  type FunctionSpec,
  type Scalar,
  type Value,
} from '../values';
import { clean, mapScalarsExcept, maxSerial, scalarNumber, scalarSerial, spec, weekdayOf } from './helpers';

/** Which days of the week are off, by weekday with 0 for Sunday, and how many are worked. */
interface Week {
  off: boolean[];
  worked: number;
}

function week(off: boolean[]): Week {
  let worked = 0;
  for (let i = 0; i < 7; i++) if (!off[i]) worked++;
  return { off, worked };
}

/**
 * The weekend argument: 1 to 7 name two days in a row (1 is Saturday and
 * Sunday, 2 Sunday and Monday ...), 11 to 17 one day (11 is Sunday), and a text
 * of seven 0s and 1s marks the days off from Monday on.
 */
function weekend(arg: Scalar | undefined, context: CalcContext): Week | CalcError {
  const off = [false, false, false, false, false, false, false];
  if (arg === undefined || arg === null) {
    off[0] = off[6] = true;
    return week(off);
  }
  if (isError(arg)) return arg;
  if (typeof arg === 'boolean') return ERROR_VALUE;
  if (typeof arg === 'string') {
    if (arg.length !== 7) return ERROR_VALUE;
    for (let i = 0; i < 7; i++) {
      const c = arg.charCodeAt(i);
      if (c !== 48 && c !== 49) return ERROR_VALUE;
      if (c === 49) off[(i + 1) % 7] = true;
    }
    const result = week(off);
    // A week without a working day has no answer.
    return result.worked === 0 ? ERROR_VALUE : result;
  }
  const number = scalarNumber(arg, context);
  if (typeof number !== 'number') return number;
  const code = Math.trunc(number);
  if (code >= 1 && code <= 7) {
    off[(code + 5) % 7] = true;
    off[(code + 6) % 7] = true;
  } else if (code >= 11 && code <= 17) off[code - 11] = true;
  else return ERROR_NUM;
  return week(off);
}

function isOff(serial: number, days: Week, date1904: boolean): boolean {
  return days.off[weekdayOf(serial, date1904)];
}

/**
 * The holidays that fall on working days, without repeats, in ascending order.
 * Dates may be written as text; text that is no date is #VALUE!.
 */
function holidayList(arg: Value | undefined, days: Week, context: CalcContext): number[] | CalcError {
  if (arg === undefined || arg === null) return [];
  const date1904 = context.date1904 === true;
  const seen = new Set<number>();
  let error = undefined as CalcError | undefined;
  const visit = (value: Scalar): void => {
    if (error !== undefined || value === null || typeof value === 'boolean') return;
    if (isError(value)) {
      error = value;
      return;
    }
    const serial = scalarNumber(value, context);
    if (typeof serial !== 'number') {
      error = serial;
      return;
    }
    const day = Math.floor(serial);
    if (day >= 0 && !isOff(day, days, date1904)) seen.add(day);
  };
  if (isRef(arg)) {
    if (arg.range.bottom >= arg.range.top && arg.range.right >= arg.range.left) {
      context.forEachCell(arg.sheet, arg.range, (value) => visit(clean(value)));
    }
  } else if (isArray(arg)) {
    for (let i = 0; i < arg.data.length; i++) visit(clean(arg.data[i]));
  } else visit(arg);
  if (error !== undefined) return error;
  const list: number[] = [];
  seen.forEach((day) => {
    list.push(day);
  });
  return list.sort((x, y) => x - y);
}

/** Working days from `low` to `high`, both included. */
function workingDays(low: number, high: number, days: Week, date1904: boolean): number {
  const weeks = Math.floor((high - low + 1) / 7);
  let count = weeks * days.worked;
  for (let day = low + weeks * 7; day <= high; day++) if (!isOff(day, days, date1904)) count++;
  return count;
}

const networkDays = mapScalarsExcept(
  (index) => index === 3,
  (args, context) => {
    const start = scalarSerial(args[0] as Scalar, context);
    if (typeof start !== 'number') return start;
    const end = scalarSerial(args[1] as Scalar, context);
    if (typeof end !== 'number') return end;
    const days = weekend(args[2] as Scalar | undefined, context);
    if (isError(days)) return days;
    const holidays = holidayList(args[3], days, context);
    if (isError(holidays)) return holidays;
    const low = Math.floor(Math.min(start, end));
    const high = Math.floor(Math.max(start, end));
    let count = workingDays(low, high, days, context.date1904 === true);
    for (let i = 0; i < holidays.length; i++) if (holidays[i] >= low && holidays[i] <= high) count--;
    return start > end ? -count : count;
  },
);

/** The date `count` (not 0) working days after or, negative, before a date. */
function shiftWorkingDays(start: number, count: number, days: Week, date1904: boolean): number {
  const step = count < 0 ? -1 : 1;
  const total = Math.abs(count);
  // Seven days on are always a full week of working days, wherever one starts.
  let weeks = Math.floor(total / days.worked);
  let left = total - weeks * days.worked;
  if (left === 0) {
    weeks--;
    left = days.worked;
  }
  let day = start + step * weeks * 7;
  for (; left > 0; left--) {
    day += step;
    for (let i = 0; i < 7 && isOff(day, days, date1904); i++) day += step;
  }
  return day;
}

const workDay = mapScalarsExcept(
  (index) => index === 3,
  (args, context) => {
    const from = scalarSerial(args[0] as Scalar, context);
    if (typeof from !== 'number') return from;
    const shift = scalarNumber(args[1] as Scalar, context);
    if (typeof shift !== 'number') return shift;
    const days = weekend(args[2] as Scalar | undefined, context);
    if (isError(days)) return days;
    const holidays = holidayList(args[3], days, context);
    if (isError(holidays)) return holidays;
    const date1904 = context.date1904 === true;
    const start = Math.floor(from);
    const count = Math.trunc(shift);
    const limit = maxSerial(date1904);
    if (!(Math.abs(count) <= limit)) return ERROR_NUM;
    if (count === 0) return start;
    let result = shiftWorkingDays(start, count, days, date1904);
    // Every holiday passed on the way costs one more day; passing it may reach further holidays.
    let passed = 0;
    if (count > 0) {
      let next = 0;
      while (next < holidays.length && holidays[next] <= start) next++;
      const base = next;
      for (let round = 0; round <= holidays.length; round++) {
        while (next < holidays.length && holidays[next] <= result) next++;
        if (next - base === passed) break;
        passed = next - base;
        result = shiftWorkingDays(start, count + passed, days, date1904);
      }
    } else {
      let next = holidays.length - 1;
      while (next >= 0 && holidays[next] >= start) next--;
      const base = next;
      for (let round = 0; round <= holidays.length; round++) {
        while (next >= 0 && holidays[next] >= result) next--;
        if (base - next === passed) break;
        passed = base - next;
        result = shiftWorkingDays(start, count - passed, days, date1904);
      }
    }
    return result >= 0 && result <= limit ? result : ERROR_NUM;
  },
);

export const EXTRA_DATE_FUNCTIONS: { [name: string]: FunctionSpec } = {
  'NETWORKDAYS.INTL': spec(networkDays, 2, 4),
  'WORKDAY.INTL': spec(workDay, 2, 4),
};

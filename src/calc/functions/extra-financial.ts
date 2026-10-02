// More financial functions: cash flows on irregular dates (XNPV, XIRR),
// interest rate conversion (EFFECT, NOMINAL), sums over the periods of a loan
// (CUMIPMT, CUMPRINC) and depreciation (DB, DDB, SYD).
//
// Signs follow Excel: money paid out is negative, money received positive.

import {
  checkNumber,
  ERROR_NUM,
  ERROR_VALUE,
  isError,
  type CalcContext,
  type CalcError,
  type CalcFunction,
  type FunctionSpec,
  type Scalar,
  type Value,
} from '../values';
import { clean, mapScalarsExcept, maxSerial, numeric, roundTo, scalarNumber, spec, toArray } from './helpers';

/** Iterations XIRR may take before giving up with #NUM!. */
const MAX_ITERATIONS = 100;
const TOLERANCE = 1e-12;

// ---------------------------------------------------------------------------
// Cash flows on dates

interface DatedFlows {
  values: number[];
  /** Years (of 365 days) from the first date. */
  times: number[];
}

/** The flows and their dates, side by side. Every flow must be a number and no date may precede the first. */
function datedFlows(valuesArg: Value, datesArg: Value, context: CalcContext): DatedFlows | CalcError {
  const values = toArray(valuesArg, context);
  if (isError(values)) return values;
  const dates = toArray(datesArg, context);
  if (isError(dates)) return dates;
  const count = values.rows * values.cols;
  if (count !== dates.rows * dates.cols) return ERROR_NUM;
  const limit = maxSerial(context.date1904 === true);
  const result: DatedFlows = { values: new Array(count), times: new Array(count) };
  let origin = 0;
  for (let i = 0; i < count; i++) {
    const value = clean(values.data[i]);
    if (isError(value)) return value;
    if (typeof value !== 'number') return ERROR_VALUE;
    const date = clean(dates.data[i]);
    if (isError(date)) return date;
    // A date may be written as text; anything else that is not a number is no date.
    const serial = typeof date === 'string' ? scalarNumber(date, context) : date;
    if (typeof serial !== 'number') return ERROR_VALUE;
    const day = Math.floor(serial);
    if (!(day >= 0 && day <= limit)) return ERROR_NUM;
    if (i === 0) origin = day;
    else if (day < origin) return ERROR_NUM;
    result.values[i] = value;
    result.times[i] = (day - origin) / 365;
  }
  return result;
}

/** Present value of dated flows at a yearly rate above -100%, and its slope in the rate. */
function datedValue(rate: number, flows: DatedFlows): { value: number; slope: number } {
  const base = 1 + rate;
  let value = 0;
  let slope = 0;
  for (let i = 0; i < flows.values.length; i++) {
    const time = flows.times[i];
    const term = time === 0 ? flows.values[i] : flows.values[i] * Math.pow(base, -time);
    value += term;
    slope -= (time * term) / base;
  }
  return { value, slope };
}

const xnpv: CalcFunction = (args, context) => {
  // With an array of rates the flows are the same for every one of them: they are read once.
  let flows: DatedFlows | CalcError | undefined;
  return mapScalarsExcept(
    (index) => index > 0,
    (parts) => {
      const rate = scalarNumber(parts[0] as Scalar, context);
      if (typeof rate !== 'number') return rate;
      if (flows === undefined) flows = datedFlows(parts[1], parts[2], context);
      if (isError(flows)) return flows;
      // Excel's XNPV refuses a negative rate (XIRR, which searches, does not).
      if (!(rate >= 0)) return ERROR_NUM;
      return checkNumber(datedValue(rate, flows).value);
    },
  )(args, context);
};

function datedRate(flows: DatedFlows, guess: number): number | CalcError {
  let positive = false;
  let negative = false;
  for (let i = 0; i < flows.values.length; i++) {
    if (flows.values[i] > 0) positive = true;
    if (flows.values[i] < 0) negative = true;
  }
  if (!positive || !negative || !(guess > -1)) return ERROR_NUM;
  let rate = guess;
  for (let i = 0; i < MAX_ITERATIONS; i++) {
    const at = datedValue(rate, flows);
    if (!isFinite(at.value) || !isFinite(at.slope) || at.slope === 0) break;
    let next = rate - at.value / at.slope;
    if (!isFinite(next)) break;
    // The rate stays above -100%.
    if (!(next > -1)) next = (rate - 1) / 2;
    if (Math.abs(next - rate) <= TOLERANCE * Math.max(1, Math.abs(next))) return next;
    rate = next;
  }
  // Newton left the road: look for a sign change on a grid of rates and halve the interval.
  let low = -0.999999;
  let lowValue = datedValue(low, flows).value;
  for (let step = 0; step < 80; step++) {
    const high = step < 40 ? -0.975 + step * 0.05 : Math.pow(10, (step - 40) / 8);
    const highValue = datedValue(high, flows).value;
    if (isFinite(lowValue) && isFinite(highValue) && lowValue * highValue <= 0) {
      let a = low;
      let b = high;
      let fa = lowValue;
      for (let i = 0; i < 100 && b - a > TOLERANCE * Math.max(1, Math.abs(a)); i++) {
        const middle = (a + b) / 2;
        const fm = datedValue(middle, flows).value;
        if (fa * fm <= 0) b = middle;
        else {
          a = middle;
          fa = fm;
        }
      }
      return (a + b) / 2;
    }
    low = high;
    lowValue = highValue;
  }
  return ERROR_NUM;
}

const xirr: CalcFunction = (args, context) => {
  // With an array of guesses the flows are the same for every one of them: they are read once.
  let flows: DatedFlows | CalcError | undefined;
  return mapScalarsExcept(
    (index) => index < 2,
    (parts) => {
      if (flows === undefined) flows = datedFlows(parts[0], parts[1], context);
      if (isError(flows)) return flows;
      let guess = 0.1;
      if (parts.length > 2 && parts[2] !== null) {
        const given = scalarNumber(parts[2] as Scalar, context);
        if (typeof given !== 'number') return given;
        guess = given;
      }
      const result = datedRate(flows, guess);
      return typeof result === 'number' ? checkNumber(result) : result;
    },
  )(args, context);
};

// ---------------------------------------------------------------------------
// Loans

/** The payment of a loan, as PMT gives it (the formula of the one in financial.ts). */
function payment(rate: number, nper: number, pv: number, fv: number, type: number): number {
  if (rate === 0) return -(pv + fv) / nper;
  // expm1 and log1p keep the digits that (1 + rate)^nper - 1 loses for a small rate.
  const gain = Math.expm1(nper * Math.log1p(rate));
  return (-rate * (pv * (gain + 1) + fv)) / ((1 + rate * type) * gain);
}

/**
 * What is still owed once `paid` payments have been made. With payments at the
 * start of a period the first one meets the whole loan before any interest.
 */
function owed(rate: number, pmt: number, pv: number, paid: number, type: number): number {
  if (type === 1) {
    if (paid <= 0) return pv;
    const gain = Math.expm1((paid - 1) * Math.log1p(rate));
    return (pv + pmt) * (gain + 1) + (pmt * gain) / rate;
  }
  const gain = Math.expm1(paid * Math.log1p(rate));
  return pv * (gain + 1) + (pmt * gain) / rate;
}

function cumulative(principal: boolean): CalcFunction {
  return numeric((rate, nper, pv, start, end, type) => {
    if (!(rate > 0 && nper > 0 && pv > 0)) return ERROR_NUM;
    const from = Math.trunc(start);
    const to = Math.trunc(end);
    const when = Math.trunc(type);
    // The periods must lie inside the loan, in order.
    if (!(from >= 1 && to >= 1 && from <= to && to <= nper) || (when !== 0 && when !== 1)) return ERROR_NUM;
    const pmt = payment(rate, nper, pv, 0, when);
    // Each payment lowers the debt by its principal part, so the parts of a run of
    // periods add up to the change of the debt; the rest of the payments is interest.
    const repaid = owed(rate, pmt, pv, to, when) - owed(rate, pmt, pv, from - 1, when);
    return principal ? repaid : pmt * (to - from + 1) - repaid;
  });
}

// ---------------------------------------------------------------------------
// Depreciation

function fixedDeclining(cost: number, salvage: number, life: number, period: number, month?: number): number | CalcError {
  const months = month === undefined ? 12 : Math.trunc(month);
  const per = Math.trunc(period);
  if (!(cost > 0 && salvage >= 0 && salvage <= cost && life > 0 && per >= 1 && months >= 1 && months <= 12)) {
    return ERROR_NUM;
  }
  if (per > life + 1 || (per > life && months === 12)) return ERROR_NUM;
  // The rate is rounded to three decimals.
  const rate = roundTo(1 - Math.pow(salvage / cost, 1 / life), 3, 0);
  const firstYear = (cost * rate * months) / 12;
  if (per === 1) return firstYear;
  const share = (cost - firstYear) * Math.pow(1 - rate, per - 2) * rate;
  // The year after the last one takes the months the first year left out.
  return per > life ? (share * (12 - months)) / 12 : share;
}

function doubleDeclining(cost: number, salvage: number, life: number, period: number, factor?: number): number | CalcError {
  const speed = factor === undefined ? 2 : factor;
  if (!(cost >= 0 && salvage >= 0 && life > 0 && period > 0 && period <= life && speed > 0)) return ERROR_NUM;
  let rate = speed / life;
  let before: number;
  if (rate >= 1) {
    rate = 1;
    before = period === 1 ? cost : 0;
  } else before = cost * Math.pow(1 - rate, period - 1);
  const after = cost * Math.pow(1 - rate, period);
  // The book value never falls below the salvage value.
  const share = after < salvage ? before - salvage : before - after;
  return share < 0 ? 0 : share;
}

export const EXTRA_FINANCIAL_FUNCTIONS: { [name: string]: FunctionSpec } = {
  XNPV: spec(xnpv, 3, 3, (index) => index > 0),
  XIRR: spec(xirr, 2, 3, (index) => index < 2),
  EFFECT: spec(
    numeric((nominal, npery) => {
      const periods = Math.trunc(npery);
      if (!(nominal > 0 && periods >= 1)) return ERROR_NUM;
      return Math.pow(1 + nominal / periods, periods) - 1;
    }),
    2,
    2,
  ),
  NOMINAL: spec(
    numeric((effect, npery) => {
      const periods = Math.trunc(npery);
      if (!(effect > 0 && periods >= 1)) return ERROR_NUM;
      return periods * (Math.pow(1 + effect, 1 / periods) - 1);
    }),
    2,
    2,
  ),
  CUMIPMT: spec(cumulative(false), 6, 6),
  CUMPRINC: spec(cumulative(true), 6, 6),
  DB: spec(numeric(fixedDeclining), 4, 5),
  DDB: spec(numeric(doubleDeclining), 4, 5),
  SYD: spec(
    numeric((cost, salvage, life, per) => {
      if (!(salvage >= 0 && life > 0 && per > 0 && per <= life)) return ERROR_NUM;
      return ((cost - salvage) * (life - per + 1) * 2) / (life * (life + 1));
    }),
    4,
    4,
  ),
};

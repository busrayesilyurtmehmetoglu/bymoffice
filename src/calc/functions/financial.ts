// Financial functions: the time value of money (PMT, FV, PV, NPER, RATE and
// the split of a payment into interest and principal), NPV, IRR and SLN.
//
// Signs follow Excel: money paid out is negative, money received positive.

import {
  checkNumber,
  ERROR_DIV0,
  ERROR_NUM,
  isError,
  type CalcError,
  type CalcFunction,
  type FunctionSpec,
  type Scalar,
} from '../values';
import { collectNumbers, mapScalarsExcept, numeric, scalarNumber, spec } from './helpers';

/** Iterations RATE and IRR may take before giving up with #NUM!. */
const MAX_ITERATIONS = 100;
const TOLERANCE = 1e-10;

function due(type: number | undefined): number {
  return type === undefined || type === 0 ? 0 : 1;
}

function payment(rate: number, nper: number, pv: number, fv: number, type: number): number {
  if (rate === 0) return -(pv + fv) / nper;
  const growth = Math.pow(1 + rate, nper);
  return (-rate * (pv * growth + fv)) / ((1 + rate * type) * (growth - 1));
}

function futureValue(rate: number, nper: number, pmt: number, pv: number, type: number): number {
  if (rate === 0) return -(pv + pmt * nper);
  const growth = Math.pow(1 + rate, nper);
  return -(pv * growth + (pmt * (1 + rate * type) * (growth - 1)) / rate);
}

function presentValue(rate: number, nper: number, pmt: number, fv: number, type: number): number {
  if (rate === 0) return -(fv + pmt * nper);
  const growth = Math.pow(1 + rate, nper);
  return -(fv + (pmt * (1 + rate * type) * (growth - 1)) / rate) / growth;
}

function periods(rate: number, pmt: number, pv: number, fv: number, type: number): number | CalcError {
  if (rate === 0) return pmt === 0 ? ERROR_DIV0 : -(pv + fv) / pmt;
  if (rate <= -1) return ERROR_NUM;
  const adjusted = (pmt * (1 + rate * type)) / rate;
  const ratio = (adjusted - fv) / (adjusted + pv);
  return ratio > 0 ? Math.log(ratio) / Math.log(1 + rate) : ERROR_NUM;
}

/** The balance equation RATE solves for its root. */
function balance(rate: number, nper: number, pmt: number, pv: number, fv: number, type: number): number {
  if (rate === 0) return pv + pmt * nper + fv;
  // expm1 and log1p keep the digits that (1 + rate)^nper - 1 loses for a small rate.
  const gain = Math.expm1(nper * Math.log1p(rate));
  return pv * (gain + 1) + (pmt * (1 + rate * type) * gain) / rate + fv;
}

function interestRate(nper: number, pmt: number, pv: number, fv: number, type: number, guess: number): number | CalcError {
  if (!(nper > 0) || !(guess > -1)) return ERROR_NUM;
  let rate = guess;
  // Newton's method with a numeric slope; the rate is kept above -100%.
  for (let i = 0; i < MAX_ITERATIONS; i++) {
    const y = balance(rate, nper, pmt, pv, fv, type);
    if (!isFinite(y)) return ERROR_NUM;
    const h = Math.max(1e-7, Math.abs(rate) * 1e-7);
    const slope = (balance(rate + h, nper, pmt, pv, fv, type) - balance(rate - h, nper, pmt, pv, fv, type)) / (2 * h);
    if (!isFinite(slope) || slope === 0) return ERROR_NUM;
    let next = rate - y / slope;
    if (!(next > -1)) next = (rate - 1) / 2;
    if (Math.abs(next - rate) <= TOLERANCE * Math.max(1, Math.abs(next))) {
      const scale = Math.abs(pv) + Math.abs(fv) + Math.abs(pmt * nper);
      return Math.abs(balance(next, nper, pmt, pv, fv, type)) <= 1e-6 * Math.max(1, scale) ? next : ERROR_NUM;
    }
    rate = next;
  }
  return ERROR_NUM;
}

function interestPart(rate: number, per: number, nper: number, pv: number, fv: number, type: number): number {
  const pmt = payment(rate, nper, pv, fv, type);
  if (per === 1) return type === 1 ? 0 : -pv * rate;
  // Interest is charged on the balance left after the payments so far.
  return type === 1
    ? (futureValue(rate, per - 2, pmt, pv, 1) - pmt) * rate
    : futureValue(rate, per - 1, pmt, pv, 0) * rate;
}

function netPresentValue(rate: number, flows: number[]): number {
  let total = 0;
  let factor = 1;
  for (let i = 0; i < flows.length; i++) {
    factor *= 1 + rate;
    total += flows[i] / factor;
  }
  return total;
}

/** Value at time 0 of flows whose first one is at time 0, and its slope in the rate. */
function flowValue(rate: number, flows: number[]): { value: number; slope: number } {
  let value = 0;
  let slope = 0;
  let factor = 1;
  for (let i = 0; i < flows.length; i++) {
    value += flows[i] / factor;
    slope -= (i * flows[i]) / (factor * (1 + rate));
    factor *= 1 + rate;
  }
  return { value, slope };
}

function internalRate(flows: number[], guess: number): number | CalcError {
  let positive = false;
  let negative = false;
  for (let i = 0; i < flows.length; i++) {
    if (flows[i] > 0) positive = true;
    if (flows[i] < 0) negative = true;
  }
  if (!positive || !negative || !(guess > -1)) return ERROR_NUM;
  let rate = guess;
  for (let i = 0; i < MAX_ITERATIONS; i++) {
    const at = flowValue(rate, flows);
    if (!isFinite(at.value) || !isFinite(at.slope) || at.slope === 0) break;
    const next = rate - at.value / at.slope;
    if (!(next > -1)) break;
    if (Math.abs(next - rate) <= TOLERANCE * Math.max(1, Math.abs(next))) return next;
    rate = next;
  }
  // Newton left the road: look for a sign change on a grid of rates and halve the interval.
  let low = -0.999999;
  let lowValue = flowValue(low, flows).value;
  for (let step = 0; step < 200; step++) {
    const high = step < 100 ? -0.99 + step * 0.02 : Math.pow(10, (step - 100) / 16);
    const highValue = flowValue(high, flows).value;
    if (isFinite(lowValue) && isFinite(highValue) && lowValue * highValue <= 0) {
      let a = low;
      let b = high;
      let fa = lowValue;
      for (let i = 0; i < 200 && b - a > TOLERANCE * Math.max(1, Math.abs(a)); i++) {
        const middle = (a + b) / 2;
        const fm = flowValue(middle, flows).value;
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

const npv = mapScalarsExcept(
  (index) => index > 0,
  (args, context) => {
    const rate = scalarNumber(args[0] as Scalar, context);
    if (typeof rate !== 'number') return rate;
    const flows = collectNumbers(args.slice(1), context);
    if (isError(flows)) return flows;
    if (rate === -1) return ERROR_DIV0;
    return checkNumber(netPresentValue(rate, flows));
  },
);

const irr = mapScalarsExcept(
  (index) => index === 0,
  (args, context) => {
    const flows = collectNumbers([args[0]], context);
    if (isError(flows)) return flows;
    let guess = 0.1;
    // Left out, the guess is 10%; left empty ("IRR(values,)") it is 0.
    if (args.length > 1) {
      const given = scalarNumber(args[1] as Scalar, context);
      if (typeof given !== 'number') return given;
      guess = given;
    }
    const result = internalRate(flows, guess);
    return typeof result === 'number' ? checkNumber(result) : result;
  },
);

function part(principal: boolean): CalcFunction {
  return numeric((rate, per, nper, pv, fv, type) => {
    if (!(per >= 1 && per <= nper)) return ERROR_NUM;
    const when = due(type);
    const future = fv === undefined ? 0 : fv;
    const interest = interestPart(rate, per, nper, pv, future, when);
    return principal ? payment(rate, nper, pv, future, when) - interest : interest;
  });
}

export const FINANCIAL_FUNCTIONS: { [name: string]: FunctionSpec } = {
  PMT: spec(
    numeric((rate, nper, pv, fv, type) =>
      nper === 0 ? ERROR_NUM : payment(rate, nper, pv, fv === undefined ? 0 : fv, due(type)),
    ),
    3,
    5,
  ),
  FV: spec(
    numeric((rate, nper, pmt, pv, type) => futureValue(rate, nper, pmt, pv === undefined ? 0 : pv, due(type))),
    3,
    5,
  ),
  PV: spec(
    numeric((rate, nper, pmt, fv, type) => presentValue(rate, nper, pmt, fv === undefined ? 0 : fv, due(type))),
    3,
    5,
  ),
  NPER: spec(
    numeric((rate, pmt, pv, fv, type) => periods(rate, pmt, pv, fv === undefined ? 0 : fv, due(type))),
    3,
    5,
  ),
  RATE: spec(
    numeric((nper, pmt, pv, fv, type, guess) =>
      interestRate(nper, pmt, pv, fv === undefined ? 0 : fv, due(type), guess === undefined ? 0.1 : guess),
    ),
    3,
    6,
  ),
  NPV: spec(npv, 2),
  IRR: spec(irr, 1, 2, (index) => index === 0),
  IPMT: spec(part(false), 4, 6),
  PPMT: spec(part(true), 4, 6),
  SLN: spec(numeric((cost, salvage, life) => (life === 0 ? ERROR_DIV0 : (cost - salvage) / life)), 3, 3),
};

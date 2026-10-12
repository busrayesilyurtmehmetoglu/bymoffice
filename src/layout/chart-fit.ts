// The arithmetic behind the lines a chart adds to its series: trend lines
// (least squares, worked out by hand), the spread of a series for its error
// bars, curves through the points of a smoothed line, and cutting a line to
// the plot area. Nothing here knows of charts or of text: numbers in, numbers out.
//
// Excel is not at hand to compare with: rules that follow what Excel is known
// to do but could not be checked against it are marked UNSURE.

export type TrendType = 'linear' | 'exp' | 'log' | 'poly' | 'power' | 'movingAvg';

export interface TrendOptions {
  /** Polynomial: the degree, 2..6. Default 2. */
  order?: number;
  /** Moving average: how many points are averaged. Default 2. */
  period?: number;
}

export interface TrendFit {
  /** The value of the trend at `x`; undefined where it has none (a logarithm of nothing above zero). */
  at(x: number): number | undefined;
  /** Moving average: the points of the line, x and y in turn. The line has no values between or beyond them. */
  points?: number[];
}

/** A straight line through (u, v) by least squares, around the means so that large numbers lose nothing. */
function line(us: ArrayLike<number>, vs: ArrayLike<number>, count: number): { slope: number; meanU: number; meanV: number } | undefined {
  if (count < 2) return undefined;
  let meanU = 0;
  let meanV = 0;
  for (let i = 0; i < count; i++) {
    meanU += us[i] / count;
    meanV += vs[i] / count;
  }
  let suu = 0;
  let suv = 0;
  for (let i = 0; i < count; i++) {
    const du = us[i] - meanU;
    suu += du * du;
    suv += du * (vs[i] - meanV);
  }
  if (!(suu > 0) || !Number.isFinite(suu) || !Number.isFinite(suv)) return undefined;
  return { slope: suv / suu, meanU, meanV };
}

/**
 * A polynomial of the given degree through (x, y) by least squares. The x values are moved and
 * scaled to -1..1 and the polynomial is written in Chebyshev's polynomials of that: the sums
 * that have to be solved then stay small and far from singular, where sums of powers of the
 * x values themselves (x = 45000, 45001, ...) would lose every digit.
 */
function polynomial(xs: ArrayLike<number>, ys: ArrayLike<number>, count: number, degree: number): ((x: number) => number) | undefined {
  const terms = degree + 1;
  if (count < terms) return undefined;
  let low = Infinity;
  let high = -Infinity;
  let meanY = 0;
  for (let i = 0; i < count; i++) {
    if (xs[i] < low) low = xs[i];
    if (xs[i] > high) high = xs[i];
    meanY += ys[i] / count;
  }
  const centre = low / 2 + high / 2;
  const half = high / 2 - low / 2;
  if (!(half > 0) || !Number.isFinite(half) || !Number.isFinite(meanY)) return undefined;
  // The sums of the products of the basis with itself and with y.
  const matrix: number[][] = [];
  const right: number[] = [];
  for (let r = 0; r < terms; r++) {
    matrix.push(new Array<number>(terms).fill(0));
    right.push(0);
  }
  const basis = new Array<number>(terms).fill(0);
  for (let i = 0; i < count; i++) {
    const t = (xs[i] - centre) / half;
    basis[0] = 1;
    if (terms > 1) basis[1] = t;
    for (let k = 2; k < terms; k++) basis[k] = 2 * t * basis[k - 1] - basis[k - 2];
    const y = ys[i] - meanY;
    for (let r = 0; r < terms; r++) {
      right[r] += basis[r] * y;
      for (let c = r; c < terms; c++) matrix[r][c] += basis[r] * basis[c];
    }
  }
  for (let r = 0; r < terms; r++) for (let c = 0; c < r; c++) matrix[r][c] = matrix[c][r];
  // Elimination with the largest pivot of each column.
  for (let k = 0; k < terms; k++) {
    let pivot = k;
    for (let r = k + 1; r < terms; r++) if (Math.abs(matrix[r][k]) > Math.abs(matrix[pivot][k])) pivot = r;
    // Fewer different x values than the degree needs: no such polynomial.
    if (!(Math.abs(matrix[pivot][k]) > count * 1e-11)) return undefined;
    if (pivot !== k) {
      const row = matrix[k];
      matrix[k] = matrix[pivot];
      matrix[pivot] = row;
      const value = right[k];
      right[k] = right[pivot];
      right[pivot] = value;
    }
    for (let r = k + 1; r < terms; r++) {
      const factor = matrix[r][k] / matrix[k][k];
      if (factor === 0) continue;
      for (let c = k; c < terms; c++) matrix[r][c] -= factor * matrix[k][c];
      right[r] -= factor * right[k];
    }
  }
  const weights = new Array<number>(terms).fill(0);
  for (let r = terms - 1; r >= 0; r--) {
    let sum = right[r];
    for (let c = r + 1; c < terms; c++) sum -= matrix[r][c] * weights[c];
    weights[r] = sum / matrix[r][r];
  }
  for (let r = 0; r < terms; r++) if (!Number.isFinite(weights[r])) return undefined;
  return (x) => {
    const t = (x - centre) / half;
    let before = 1;
    let now = t;
    let sum = weights[0] + (terms > 1 ? weights[1] * t : 0);
    for (let k = 2; k < terms; k++) {
      const next = 2 * t * now - before;
      sum += weights[k] * next;
      before = now;
      now = next;
    }
    return sum + meanY;
  };
}

/**
 * The trend of the points (xs[i], ys[i]), every one of them a finite number. Undefined when
 * the points give none: too few of them, all at one x, or values a type cannot take.
 *
 * UNSURE (what Excel is known to do, not checked against it): an exponential or a power trend
 * of a series with a value that is not above zero, and a logarithmic or a power trend of
 * points with an x that is not above zero, is not drawn at all.
 */
export function fitTrend(type: TrendType, xs: ArrayLike<number>, ys: ArrayLike<number>, options: TrendOptions = {}): TrendFit | undefined {
  const count = Math.min(xs.length, ys.length);
  if (count < 2) return undefined;
  if (type === 'movingAvg') {
    const period = options.period !== undefined && options.period >= 2 ? Math.floor(options.period) : 2;
    if (count < period) return undefined;
    const points: number[] = [];
    let sum = 0;
    for (let i = 0; i < count; i++) {
      sum += ys[i];
      if (i >= period) sum -= ys[i - period];
      if (i >= period - 1) points.push(xs[i], sum / period);
    }
    return { at: () => undefined, points };
  }
  if (type === 'poly') {
    const degree = options.order !== undefined && options.order >= 2 ? Math.min(Math.floor(options.order), 6) : 2;
    const curve = polynomial(xs, ys, count, degree);
    return curve ? { at: curve } : undefined;
  }
  const logX = type === 'log' || type === 'power';
  const logY = type === 'exp' || type === 'power';
  const us = new Float64Array(count);
  const vs = new Float64Array(count);
  for (let i = 0; i < count; i++) {
    if ((logX && !(xs[i] > 0)) || (logY && !(ys[i] > 0))) return undefined;
    us[i] = logX ? Math.log(xs[i]) : xs[i];
    vs[i] = logY ? Math.log(ys[i]) : ys[i];
  }
  const fit = line(us, vs, count);
  if (!fit) return undefined;
  return {
    at: (x) => {
      if (logX && !(x > 0)) return undefined;
      const v = fit.meanV + fit.slope * ((logX ? Math.log(x) : x) - fit.meanU);
      const y = logY ? Math.exp(v) : v;
      return Number.isFinite(y) ? y : undefined;
    },
  };
}

/** The mean of the values, their standard deviation (of a sample: n - 1) and the standard error of the mean. */
export function seriesSpread(values: ArrayLike<number>): { mean: number; deviation: number; error: number } {
  const count = values.length;
  if (count === 0) return { mean: 0, deviation: 0, error: 0 };
  let mean = 0;
  for (let i = 0; i < count; i++) mean += values[i] / count;
  let squares = 0;
  for (let i = 0; i < count; i++) squares += (values[i] - mean) * (values[i] - mean);
  const deviation = count > 1 ? Math.sqrt(squares / (count - 1)) : 0;
  const safe = Number.isFinite(deviation) ? deviation : 0;
  return { mean: Number.isFinite(mean) ? mean : 0, deviation: safe, error: safe / Math.sqrt(count) };
}

/**
 * A curve through the points (x and y in turn) as cubic Bézier pieces: the first point, then
 * two control points and an end point for every further one. The curve goes through every
 * point (a Catmull-Rom spline); between points it may swing beyond them.
 */
export function smoothPath(points: ArrayLike<number>): number[] {
  const count = Math.floor(points.length / 2);
  const out: number[] = [];
  if (count === 0) return out;
  out.push(points[0], points[1]);
  const x = (i: number): number => points[2 * (i < 0 ? 0 : i >= count ? count - 1 : i)];
  const y = (i: number): number => points[2 * (i < 0 ? 0 : i >= count ? count - 1 : i) + 1];
  for (let i = 0; i + 1 < count; i++) {
    out.push(x(i) + (x(i + 1) - x(i - 1)) / 6, y(i) + (y(i + 1) - y(i - 1)) / 6, x(i + 1) - (x(i + 2) - x(i)) / 6, y(i + 1) - (y(i + 2) - y(i)) / 6, x(i + 1), y(i + 1));
  }
  return out;
}

/**
 * The parts of a line (x and y in turn) that lie inside a box, as lines of their own: where
 * the line leaves the box it ends on the edge. Points that are no numbers break the line.
 */
export function clipLine(points: ArrayLike<number>, box: { x: number; y: number; width: number; height: number }): number[][] {
  const runs: number[][] = [];
  let run: number[] = [];
  const left = box.x;
  const right = box.x + box.width;
  const top = box.y;
  const bottom = box.y + box.height;
  const close = (): void => {
    if (run.length >= 4) runs.push(run);
    run = [];
  };
  for (let i = 0; i + 3 < points.length; i += 2) {
    const x0 = points[i];
    const y0 = points[i + 1];
    const x1 = points[i + 2];
    const y1 = points[i + 3];
    if (!Number.isFinite(x0) || !Number.isFinite(y0) || !Number.isFinite(x1) || !Number.isFinite(y1)) {
      close();
      continue;
    }
    // Liang-Barsky: the share of the segment that is inside all four edges.
    let from = 0;
    let to = 1;
    const dx = x1 - x0;
    const dy = y1 - y0;
    const edges: [number, number][] = [
      [-dx, x0 - left],
      [dx, right - x0],
      [-dy, y0 - top],
      [dy, bottom - y0],
    ];
    let outside = false;
    for (const [p, q] of edges) {
      if (p === 0) {
        if (q < 0) outside = true;
        continue;
      }
      const t = q / p;
      if (p < 0) {
        if (t > from) from = t;
      } else if (t < to) to = t;
    }
    if (outside || from > to) {
      close();
      continue;
    }
    const ax = x0 + dx * from;
    const ay = y0 + dy * from;
    if (from > 0 || run.length === 0 || run[run.length - 2] !== ax || run[run.length - 1] !== ay) {
      close();
      run.push(ax, ay);
    }
    run.push(x0 + dx * to, y0 + dy * to);
    if (to < 1) close();
  }
  close();
  return runs;
}

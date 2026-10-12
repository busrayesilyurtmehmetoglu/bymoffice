// The arithmetic of trend lines, of the spread of a series, of smoothed
// lines and of cutting a line to a box. The expected numbers are worked out
// by hand (see the comments): no library is at hand to take them from.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { clipLine, fitTrend, seriesSpread, smoothPath } from '../src/layout/chart-fit';

const near = (actual: number | undefined, expected: number, message: string, tolerance = 1e-9) =>
  assert.ok(actual !== undefined && Math.abs(actual - expected) <= tolerance * Math.max(1, Math.abs(expected)), `${message}: ${actual} is not ${expected}`);

test('a linear trend: the line of least squares', () => {
  // x = 1..5, y = 2, 4, 5, 4, 5: mean x 3, mean y 4, Sxy = 6, Sxx = 10: slope 0.6, intercept 2.2.
  const fit = fitTrend('linear', [1, 2, 3, 4, 5], [2, 4, 5, 4, 5]);
  near(fit?.at(0), 2.2, 'the intercept');
  near(fit?.at(6), 5.8, 'one step beyond the last point');
  // The same line far from zero, where sums of squares of x would lose the digits that matter.
  const far = fitTrend('linear', [1e9 + 1, 1e9 + 2, 1e9 + 3, 1e9 + 4, 1e9 + 5], [2, 4, 5, 4, 5]);
  near(far?.at(1e9 + 6), 5.8, 'far from zero');
  assert.equal(fitTrend('linear', [1], [1]), undefined, 'one point has no trend');
  assert.equal(fitTrend('linear', [2, 2, 2], [1, 2, 3]), undefined, 'points at one x have none');
});

test('a polynomial trend: exact where the points are a polynomial, least squares where they are not', () => {
  // y = 2x² - 3x + 1.
  const xs = [1, 2, 3, 4, 5, 6];
  const exact = fitTrend('poly', xs, xs.map((x) => 2 * x * x - 3 * x + 1), { order: 2 });
  near(exact?.at(7), 78, 'beyond the points');
  near(exact?.at(2.5), 6, 'between two points');
  // x = -1, 0, 1, 2 and y = 1, 0, 0, 3. The normal equations 4a + 2b + 6c = 4, 2a + 6b + 8c = 5,
  // 6a + 8b + 18c = 13 give a = -0.3, b = -0.4, c = 1.
  const least = fitTrend('poly', [-1, 0, 1, 2], [1, 0, 0, 3], { order: 2 });
  near(least?.at(0), -0.3, 'a');
  near(least?.at(3), 7.5, 'a + 3b + 9c');
  near(least?.at(1), 0.3, 'a + b + c');
  assert.equal(fitTrend('poly', [1, 2], [1, 2], { order: 2 }), undefined, 'two points have no parabola');
  assert.equal(fitTrend('poly', [1, 1, 2, 2], [1, 2, 3, 4], { order: 2 }), undefined, 'nor have points at two x values');
});

test('a polynomial of degree six through dates: stable where powers of x would not be', () => {
  // Serial numbers of dates: x⁶ is near 1e28 and its sums hold no digit of the differences.
  const p = (u: number) => 0.5 * u * u * u - 4 * u * u + 3 * u - 7;
  const xs: number[] = [];
  const ys: number[] = [];
  for (let i = 0; i < 40; i++) {
    xs.push(45000 + i);
    ys.push(p(i));
  }
  for (const order of [3, 4, 5, 6]) {
    const fit = fitTrend('poly', xs, ys, { order });
    near(fit?.at(45000 + 12.5), p(12.5), `degree ${order}, between two points`, 1e-7);
    near(fit?.at(45000 + 41), p(41), `degree ${order}, beyond the last`, 1e-6);
  }
  // A polynomial of the sixth degree is not one of the fifth: u⁶ through thirteen points.
  const us = Array.from({ length: 13 }, (_value, i) => i - 6);
  const sixth = fitTrend('poly', us, us.map((u) => Math.pow(u, 6)), { order: 6 });
  const fifth = fitTrend('poly', us, us.map((u) => Math.pow(u, 6)), { order: 5 });
  near(sixth?.at(2.5), Math.pow(2.5, 6), 'degree 6 gives u⁶ back', 1e-7);
  assert.ok(Math.abs((fifth?.at(2.5) as number) - Math.pow(2.5, 6)) > 1, 'degree 5 cannot');
  near(fitTrend('poly', us, us.map((u) => Math.pow(u, 6)), { order: 99 })?.at(2.5), Math.pow(2.5, 6), 'a degree beyond six is six, not five', 1e-7);
  const clipped = fitTrend('poly', xs, ys, { order: 99 });
  near(clipped?.at(45000 + 3), p(3), 'a degree beyond six is six', 1e-7);
});

test('exponential, logarithmic and power trends are straight lines of logarithms', () => {
  const xs = [1, 2, 3, 4, 5];
  // y = 3 e^(x/2): at 4 it is 3e² = 22.16716829...
  near(fitTrend('exp', xs, xs.map((x) => 3 * Math.exp(0.5 * x)))?.at(4), 22.16716829679195, 'exponential', 1e-12);
  // y = 2 + 3 ln x: at 10 it is 2 + 3 × 2.302585092994046.
  near(fitTrend('log', xs, xs.map((x) => 2 + 3 * Math.log(x)))?.at(10), 8.907755278982137, 'logarithmic', 1e-12);
  // y = 2 x^1.5: at 9 it is 54.
  near(fitTrend('power', xs, xs.map((x) => 2 * Math.pow(x, 1.5)))?.at(9), 54, 'power', 1e-12);
  // Least squares of the logarithms: x = 1, 2, 3 and y = e, e², e⁴. ln y = 1, 2, 4: slope 1.5, mean 7/3.
  near(fitTrend('exp', [1, 2, 3], [Math.E, Math.E * Math.E, Math.pow(Math.E, 4)])?.at(2), Math.exp(7 / 3), 'not exact: the line through the logarithms', 1e-12);
  assert.equal(fitTrend('exp', xs, [1, 2, 0, 4, 5]), undefined, 'no exponential trend with a value that is not above zero');
  assert.equal(fitTrend('power', xs, [1, 2, -3, 4, 5]), undefined, 'nor a power trend');
  assert.equal(fitTrend('log', [0, 1, 2], [1, 2, 3]), undefined, 'no logarithmic trend with an x that is not above zero');
  assert.equal(fitTrend('log', xs, xs)?.at(0), undefined, 'and no value where the logarithm has none');
  assert.equal(fitTrend('exp', xs, xs.map((x) => Math.exp(100 * x)))?.at(1e6), undefined, 'a value beyond what a number can be is none');
});

test('a moving average: the mean of the last points, from the first point that has enough before it', () => {
  const fit = fitTrend('movingAvg', [1, 2, 3, 4, 5, 6], [1, 2, 3, 4, 5, 9], { period: 3 });
  assert.deepEqual(fit?.points, [3, 2, 4, 3, 5, 4, 6, 6]);
  assert.deepEqual(fitTrend('movingAvg', [1, 2, 3], [4, 6, 10])?.points, [2, 5, 3, 8], 'two points by default');
  assert.equal(fitTrend('movingAvg', [1, 2], [1, 2], { period: 3 }), undefined, 'fewer points than the period');
});

test('the spread of a series: mean, standard deviation of a sample, standard error', () => {
  // 2, 4, 4, 4, 5, 5, 7, 9: mean 5, squares of the differences 9+1+1+1+0+0+4+16 = 32, variance 32/7.
  const spread = seriesSpread([2, 4, 4, 4, 5, 5, 7, 9]);
  near(spread.mean, 5, 'mean');
  near(spread.deviation, Math.sqrt(32 / 7), 'deviation');
  near(spread.deviation, 2.138089935299395, 'deviation, as a number', 1e-12);
  near(spread.error, 0.7559289460184544, 'deviation over the root of eight', 1e-12);
  assert.deepEqual(seriesSpread([]), { mean: 0, deviation: 0, error: 0 });
  assert.deepEqual(seriesSpread([3]), { mean: 3, deviation: 0, error: 0 });
  const huge = seriesSpread([1.7e308, -1.7e308]);
  assert.ok(Number.isFinite(huge.mean) && Number.isFinite(huge.deviation) && Number.isFinite(huge.error), 'numbers too large to square give no spread, not a number that is none');
});

test('a smoothed line goes through its points; points on a straight line stay on it', () => {
  const curve = smoothPath([0, 0, 10, 10, 20, 20, 30, 30]);
  assert.equal(curve.length, 2 + 3 * 6);
  for (let i = 0; i < curve.length; i += 2) near(curve[i + 1], curve[i], `number ${i} lies on the line`);
  assert.deepEqual([curve[6], curve[7], curve[12], curve[13], curve[18], curve[19]], [10, 10, 20, 20, 30, 30], 'the end of each piece is a point');
  // A peak: the control points beside it swing level, not through it.
  const peak = smoothPath([0, 0, 10, 10, 20, 0]);
  near(peak[4], 10 - 20 / 6, 'the control point before the peak, x');
  near(peak[5], 10, 'and y: level with the peak');
  assert.deepEqual(smoothPath([5, 6]), [5, 6], 'one point is one point');
  assert.deepEqual(smoothPath([]), []);
});

test('a line is cut to a box: what is outside is left out, and the ends lie on the edges', () => {
  const box = { x: 0, y: 0, width: 10, height: 10 };
  assert.deepEqual(clipLine([2, 2, 8, 8], box), [[2, 2, 8, 8]], 'inside: as it is');
  assert.deepEqual(clipLine([-5, 5, 15, 5], box), [[0, 5, 10, 5]], 'through: from edge to edge');
  assert.deepEqual(clipLine([5, 5, 5, 20, 8, 20, 8, 5], box), [[5, 5, 5, 10], [8, 10, 8, 5]], 'out and back in: two lines');
  assert.deepEqual(clipLine([20, 20, 30, 30], box), [], 'outside: nothing');
  assert.deepEqual(clipLine([1, 1, 2, 2, NaN, 3, 4, 4, 5, 5], box), [[1, 1, 2, 2], [4, 4, 5, 5]], 'a point that is none breaks the line');
  assert.deepEqual(clipLine([1, 1, 2, 2, 3, 3], box), [[1, 1, 2, 2, 3, 3]], 'points inside stay one line');
});

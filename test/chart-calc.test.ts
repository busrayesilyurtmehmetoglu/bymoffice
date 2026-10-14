// The arithmetic of the newer chart types and of the surface chart (layout/chart-calc.ts):
// bins of a histogram, quartiles of a box, running totals of a waterfall, the hierarchy of a
// treemap and the rectangles it is divided into, the bands of a surface.
//
// The quartiles are compared with Python's statistics.quantiles(), computed outside this test
// (method "inclusive" is QUARTILE.INC, "exclusive" QUARTILE.EXC); the bins by hand.
// UNSURE: that Excel's charts use these rules is not checked against Excel.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { bandPolygon, boxStats, buildTree, histogramBins, levelSegment, MAX_HISTOGRAM_BINS, quartile, runningShares, scottWidth, squarify, treeDepth, waterfallSteps } from '../src/layout/chart-calc';

const near = (actual: number, expected: number, message: string, tolerance = 1e-9) => assert.ok(Math.abs(actual - expected) <= tolerance, `${message}: ${actual} is not ${expected}`);

// ------------------------------------------------------------------ quartiles

test('quartiles: inclusive and exclusive as statistics.quantiles() of Python gives them', () => {
  // [data, inclusive Q1 Q2 Q3, exclusive Q1 Q2 Q3], from Python 3: statistics.quantiles(data, n=4, method=...).
  const cases: [number[], number[], number[]][] = [
    [[1, 2, 3, 4, 5, 6, 7, 8], [2.75, 4.5, 6.25], [2.25, 4.5, 6.75]],
    [[1, 2, 3, 4, 5, 6, 7, 8, 9], [3, 5, 7], [2.5, 5, 7.5]],
    [[2, 4, 4, 5, 7, 9, 10, 12, 30], [4, 7, 10], [4, 7, 11]],
    [[3.5, 1.25, 8, 2, 2, 40, 7.75, 6, 5, 4.5, -20], [2, 4.5, 6.875], [2, 4.5, 7.75]],
  ];
  for (const [data, inclusive, exclusive] of cases) {
    const sorted = data.slice().sort((a, b) => a - b);
    [0.25, 0.5, 0.75].forEach((p, i) => {
      near(quartile(sorted, p, 'inclusive'), inclusive[i], `${data} inclusive ${p}`);
      near(quartile(sorted, p, 'exclusive'), exclusive[i], `${data} exclusive ${p}`);
    });
  }
  // Too few values for the exclusive place: held to the ends.
  assert.equal(quartile([4, 9], 0.25, 'exclusive'), 4);
  assert.equal(quartile([4, 9], 0.75, 'exclusive'), 9);
  assert.equal(quartile([7], 0.25, 'inclusive'), 7);
  assert.ok(Number.isNaN(quartile([], 0.5, 'inclusive')));
});

test('box: whiskers reach the last values within 1.5 times the box, outliers lie beyond, the mean of all', () => {
  const stats = boxStats([2, 4, 4, 5, 7, 9, 10, 12, 30], 'inclusive');
  assert.ok(stats);
  assert.deepEqual([stats.q1, stats.median, stats.q3], [4, 7, 10]);
  // The box is 6 high: whiskers reach no further than 4 - 9 and 10 + 9.
  assert.deepEqual([stats.low, stats.high], [2, 12]);
  assert.deepEqual(stats.outliers, [30]);
  assert.deepEqual(stats.inner, [2, 4, 4, 5, 7, 9, 10, 12]);
  near(stats.mean, 83 / 9, 'mean (Python: 9.222222222222221)');
  const excl = boxStats([3.5, 1.25, 8, 2, 2, 40, 7.75, 6, 5, 4.5, -20], 'exclusive');
  assert.ok(excl);
  assert.deepEqual([excl.q1, excl.median, excl.q3], [2, 4.5, 7.75]);
  assert.deepEqual(excl.outliers, [-20, 40], 'below 2 - 8.625 and above 7.75 + 8.625');
  assert.deepEqual([excl.low, excl.high], [1.25, 8]);
  assert.equal(boxStats([NaN, Infinity], 'exclusive'), undefined, 'nothing that is a number: no box');
  const single = boxStats([5], 'exclusive');
  assert.deepEqual(single && [single.q1, single.median, single.q3, single.low, single.high, single.outliers.length], [5, 5, 5, 5, 5, 0]);
});

// ------------------------------------------------------------------ histogram

test('histogram: bins of a stated width start at the lowest value; a value on an edge goes to the bin on its left', () => {
  const result = histogramBins([1, 2, 3, 4, 5, 6, 7, 8, 9, 10], { closed: 'r', size: 3 });
  assert.deepEqual(
    result.bins.map((bin) => [bin.from, bin.to, bin.count]),
    [
      [1, 4, 4], // [1, 4]: 1 2 3 4
      [4, 7, 3], // (4, 7]: 5 6 7
      [7, 10, 3], // (7, 10]: 8 9 10
    ],
  );
  const left = histogramBins([1, 2, 3, 4, 5, 6, 7, 8, 9, 10], { closed: 'l', size: 3 });
  assert.deepEqual(
    left.bins.map((bin) => [bin.from, bin.to, bin.count]),
    [
      [1, 4, 3], // [1, 4): 1 2 3
      [4, 7, 3], // [4, 7): 4 5 6
      [7, 10, 4], // [7, 10]: the last also holds its end
    ],
  );
});

test('histogram: a number of bins divides the span; underflow and overflow make bins of their own', () => {
  const counted = histogramBins([0, 10, 20, 30, 40], { closed: 'r', count: 4 });
  assert.deepEqual(counted.bins.map((bin) => [bin.from, bin.to, bin.count]), [
    [0, 10, 2],
    [10, 20, 1],
    [20, 30, 1],
    [30, 40, 1],
  ]);
  const ends = histogramBins([-5, 0, 1, 2, 3, 4, 5, 6, 50], { closed: 'r', underflow: 1, overflow: 5, size: 2 });
  assert.deepEqual(
    ends.bins.map((bin) => [bin.kind, bin.from, bin.to, bin.count]),
    [
      ['under', -Infinity, 1, 3], // -5 0 1
      ['bin', 1, 3, 2], // (1, 3]: 2 3
      ['bin', 3, 5, 2], // (3, 5]: 4 5
      ['over', 5, Infinity, 2], // 6 50
    ],
  );
  // An overflow below the underflow is none.
  assert.ok(!histogramBins([1, 2, 3], { closed: 'r', underflow: 2, overflow: 1 }).bins.some((bin) => bin.kind === 'over'));
});

test('histogram: automatic bins are as wide as Scott\'s rule says; no spread is one bin; what is no number is left out', () => {
  const data = [2, 4, 4, 5, 7, 9, 10, 12, 30];
  // Python: 3.5 * statistics.stdev(data) / len(data) ** (1/3) = 14.197447146135074
  near(scottWidth(data), 14.197447146135074, 'Scott', 1e-9);
  const auto = histogramBins(data, { closed: 'r' });
  assert.equal(auto.bins.length, 2, '28 / 14.2 rounded up');
  assert.deepEqual(auto.bins.map((bin) => bin.count), [8, 1]);
  near(auto.bins[1].to, 2 + 2 * 14.197447146135074, 'the last bin is as wide as the others');
  const flat = histogramBins([5, 5, 5, NaN, Infinity], { closed: 'r' });
  assert.deepEqual(flat.bins.map((bin) => bin.count), [3]);
  assert.deepEqual(histogramBins([], { closed: 'r' }).bins, []);
});

test('histogram: no more bins than the limit, whatever the file asks for', () => {
  const values = Array.from({ length: 100 }, (_v, i) => i);
  const tiny = histogramBins(values, { closed: 'r', size: 1e-9 });
  assert.equal(tiny.capped, true);
  assert.equal(tiny.bins.length, MAX_HISTOGRAM_BINS);
  assert.equal(tiny.bins.reduce((sum, bin) => sum + bin.count, 0), 100, 'every value is in one bin');
  const many = histogramBins(values, { closed: 'r', count: 1e9 });
  assert.equal(many.bins.length, MAX_HISTOGRAM_BINS);
  // Large numbers far apart: nothing loops for long, every value is counted.
  const wide = histogramBins([-1e300, 1e300, 0], { closed: 'l', size: 1 });
  assert.equal(wide.bins.reduce((sum, bin) => sum + bin.count, 0), 3);
});

test('pareto: the running share of the counts', () => {
  assert.deepEqual(runningShares([5, 3, 2]), [0.5, 0.8, 1]);
  assert.deepEqual(runningShares([0, 0]), [0, 0]);
});

// ------------------------------------------------------------------ waterfall

test('waterfall: bars from the running total; a total stands on zero and sets the total', () => {
  const steps = waterfallSteps([100, 20, -50, null, 30, 100], new Set([0, 5]));
  assert.deepEqual(steps, [
    { from: 0, to: 100, kind: 'total' },
    { from: 100, to: 120, kind: 'increase' },
    { from: 120, to: 70, kind: 'decrease' },
    null,
    { from: 70, to: 100, kind: 'increase' },
    { from: 0, to: 100, kind: 'total' },
  ]);
  // A total in the middle: the step after it goes on from the total, not from what came before it.
  assert.deepEqual(waterfallSteps([10, 5, 50, 3, -8], new Set([2])), [
    { from: 0, to: 10, kind: 'increase' },
    { from: 10, to: 15, kind: 'increase' },
    { from: 0, to: 50, kind: 'total' },
    { from: 50, to: 53, kind: 'increase' },
    { from: 53, to: 45, kind: 'decrease' },
  ]);
  assert.deepEqual(waterfallSteps([5, 5], new Set()), [
    { from: 0, to: 5, kind: 'increase' },
    { from: 5, to: 10, kind: 'increase' },
  ]);
});

// ------------------------------------------------------------------ hierarchy

test('tree: paths from the outermost level to the leaves, values summed, largest first, branches numbered', () => {
  const names = [
    ['a1', 'a2', 'b1', 'b2', 'c1'],
    ['A', 'A', 'B', 'B', ''],
  ];
  const tree = buildTree(5, 2, (level, i) => names[level][i], (i) => [10, 20, 30, 5, 15][i]);
  assert.equal(tree.value, 80);
  assert.deepEqual(
    tree.children.map((node) => [node.name, node.value, node.branch]),
    [
      ['B', 35, 0],
      ['A', 30, 1],
      ['c1', 15, 2],
    ],
    'a point without a name at the top is a leaf further up',
  );
  assert.deepEqual(tree.children[0].children.map((node) => [node.name, node.value, node.point, node.branch, node.depth]), [
    ['b1', 30, 2, 0, 1],
    ['b2', 5, 3, 0, 1],
  ]);
  assert.equal(treeDepth(tree), 2);
  // Values that are not above zero do not count.
  const none = buildTree(3, 1, (_l, i) => `x${i}`, (i) => [0, -4, null][i]);
  assert.equal(none.children.length, 0);
  // Two points with one path are one leaf.
  const same = buildTree(2, 1, () => 'x', () => 2);
  assert.deepEqual(same.children.map((node) => [node.name, node.value, node.point]), [['x', 4, 0]]);
});

test('squarify: the pieces fill the box, each as large as its share, as near to squares as the rule makes them', () => {
  // The example of Bruls, Huizing and van Wijk: 6, 6, 4, 3, 2, 2, 1 in a box of 6 by 4.
  const values = [6, 6, 4, 3, 2, 2, 1];
  const pieces = squarify(values, { x: 0, y: 0, width: 6, height: 4 });
  let area = 0;
  pieces.forEach((piece, i) => {
    assert.ok(piece, `piece ${i}`);
    near(piece.width * piece.height, values[i], `area of ${i}`, 1e-9);
    assert.ok(piece.x >= -1e-9 && piece.y >= -1e-9 && piece.x + piece.width <= 6 + 1e-9 && piece.y + piece.height <= 4 + 1e-9, `piece ${i} inside`);
    area += piece.width * piece.height;
  });
  near(area, 24, 'the box is filled');
  // The first row is the two 6s side by side along the shorter side: 3 wide, 2 high each.
  assert.deepEqual([pieces[0], pieces[1]].map((piece) => piece && [piece.x, piece.y, piece.width, piece.height]), [
    [0, 0, 3, 2],
    [0, 2, 3, 2],
  ]);
  // No two pieces overlap.
  for (let i = 0; i < pieces.length; i++) {
    for (let j = i + 1; j < pieces.length; j++) {
      const a = pieces[i] as NonNullable<(typeof pieces)[number]>;
      const b = pieces[j] as NonNullable<(typeof pieces)[number]>;
      const overlap = Math.max(0, Math.min(a.x + a.width, b.x + b.width) - Math.max(a.x, b.x)) * Math.max(0, Math.min(a.y + a.height, b.y + b.height) - Math.max(a.y, b.y));
      assert.ok(overlap < 1e-9, `${i} and ${j} overlap`);
    }
  }
  assert.deepEqual(squarify([0, -1, NaN], { x: 0, y: 0, width: 10, height: 10 }), [undefined, undefined, undefined]);
  assert.deepEqual(squarify([1], { x: 0, y: 0, width: 0, height: 10 }), [undefined]);
});

// -------------------------------------------------------------------- surface

test('surface: the part of a triangle within a band, and the line of one value through it', () => {
  // Values 0, 10 and 10 at (0,0), (10,0) and (0,10): the band 0..5 is the corner at the origin.
  const corners = [0, 0, 0, 10, 0, 10, 0, 10, 10];
  const low = bandPolygon(corners, 0, 5);
  assert.deepEqual(low, [0, 0, 5, 0, 0, 5]);
  const high = bandPolygon(corners, 5, 10);
  assert.equal(high.length, 8, 'the rest is a four-cornered piece');
  assert.deepEqual(bandPolygon(corners, 20, 30), [], 'no part lies in a band above all values');
  assert.deepEqual(levelSegment(corners, 5), [5, 0, 0, 5]);
  assert.deepEqual(levelSegment(corners, 50), []);
});

// The arithmetic behind the newer chart types and the surface chart: the bins
// of a histogram, the quartiles of a box and whisker chart, the running totals
// of a waterfall, the hierarchy of a treemap or a sunburst and the rectangles
// a treemap divides its area into (squarified), and the pieces of a surface
// seen from above that lie in one band of values. Nothing here knows of charts
// or of text: numbers in, numbers out. Every loop is bounded by what is handed
// in or by a stated limit.
//
// Excel is not at hand to compare with: rules that follow what Excel is known
// to do but could not be checked against it are marked UNSURE.

// --------------------------------------------------------------- histogram

/** Bins of one histogram, at most. */
export const MAX_HISTOGRAM_BINS = 1000;

export interface BinRule {
  /** Which end of a bin belongs to it: "r" (a, b] (the first bin also holds its start), or "l" [a, b) (the last also holds its end). */
  closed: 'l' | 'r';
  /** Values at or below it (below it, closed on the left) make one bin of their own. */
  underflow?: number;
  /** Values above it (at or above it, closed on the left) make one bin of their own. */
  overflow?: number;
  /** The width of a bin. */
  size?: number;
  /** How many bins lie between the start and the end. */
  count?: number;
}

export interface HistogramBin {
  /** The ends of the bin; -Infinity / Infinity for the bins of values below and above all others. */
  from: number;
  to: number;
  /** How many values lie in it. */
  count: number;
  /** "under" and "over": the bins of `underflow` and `overflow`. */
  kind: 'bin' | 'under' | 'over';
}

export interface Histogram {
  bins: HistogramBin[];
  /** The width the bins were given. */
  width: number;
  /** More bins were asked for (or worked out) than are made: the width was widened. */
  capped: boolean;
}

/**
 * The width Scott's normal reference rule gives a bin of `values`: 3.5 standard deviations (of
 * the sample) over the cube root of their number. 0 for fewer than two values or no spread.
 * UNSURE: Excel names this rule for its automatic bins; whether it rounds the width is not known.
 */
export function scottWidth(values: ArrayLike<number>): number {
  const n = values.length;
  if (n < 2) return 0;
  let mean = 0;
  for (let i = 0; i < n; i++) mean += values[i];
  mean /= n;
  let sum = 0;
  for (let i = 0; i < n; i++) sum += (values[i] - mean) * (values[i] - mean);
  const deviation = Math.sqrt(sum / (n - 1));
  return Number.isFinite(deviation) && deviation > 0 ? (3.5 * deviation) / Math.cbrt(n) : 0;
}

/**
 * The bins of a histogram of `values` (finite numbers; anything else is left out). Regular bins
 * run from the underflow (or the lowest value) to the overflow (or the highest), `size` wide, or
 * `count` of them, or as wide as scottWidth() says. A bin that would hold nothing is still a bin.
 */
export function histogramBins(values: ArrayLike<number>, rule: BinRule): Histogram {
  const list: number[] = [];
  for (let i = 0; i < values.length; i++) {
    const value = values[i];
    if (typeof value === 'number' && Number.isFinite(value)) list.push(value);
  }
  const bins: HistogramBin[] = [];
  if (list.length === 0) return { bins, width: 0, capped: false };
  let min = Infinity;
  let max = -Infinity;
  for (const value of list) {
    if (value < min) min = value;
    if (value > max) max = value;
  }
  const left = rule.closed === 'l';
  const under = typeof rule.underflow === 'number' && Number.isFinite(rule.underflow) ? rule.underflow : undefined;
  let over = typeof rule.overflow === 'number' && Number.isFinite(rule.overflow) ? rule.overflow : undefined;
  if (under !== undefined && over !== undefined && over < under) over = undefined;
  const start = under !== undefined ? under : min;
  const end = over !== undefined ? over : max;
  const span = Math.max(end - start, 0);
  let width: number;
  let count: number;
  if (rule.count !== undefined && rule.count >= 1) {
    count = Math.floor(rule.count);
    width = span / count;
  } else {
    width = rule.size !== undefined && rule.size > 0 ? rule.size : scottWidth(list);
    count = width > 0 && span > 0 ? Math.ceil(span / width - 1e-9) : 1;
  }
  if (!(width > 0)) {
    // No spread: one bin holds everything between the ends.
    width = span > 0 ? span : 1;
    count = 1;
  }
  let capped = false;
  if (count > MAX_HISTOGRAM_BINS) {
    capped = true;
    count = MAX_HISTOGRAM_BINS;
    width = span / count;
  }
  count = Math.max(count, 1);
  // Bins of a stated or worked out width are all that wide: the last may end beyond the highest value. A number of bins ends at it.
  const even = rule.count === undefined && !capped;
  const edge = (k: number): number => (k >= count && !even ? end : start + k * width);
  const counts = new Array<number>(count).fill(0);
  let below = 0;
  let above = 0;
  for (const value of list) {
    if (under !== undefined && (left ? value < under : value <= under)) {
      below++;
      continue;
    }
    if (over !== undefined && (left ? value >= over : value > over)) {
      above++;
      continue;
    }
    // The bin by arithmetic, then moved by the edges themselves, so that a value on an edge goes where the rule says.
    let k = Math.floor((value - start) / width);
    if (!Number.isFinite(k)) k = 0;
    k = Math.min(Math.max(k, 0), count - 1);
    if (left) {
      while (k > 0 && value < edge(k)) k--;
      while (k < count - 1 && value >= edge(k + 1)) k++;
    } else {
      while (k > 0 && value <= edge(k)) k--;
      while (k < count - 1 && value > edge(k + 1)) k++;
    }
    counts[k]++;
  }
  if (under !== undefined) bins.push({ from: -Infinity, to: under, count: below, kind: 'under' });
  for (let k = 0; k < count; k++) bins.push({ from: edge(k), to: edge(k + 1), count: counts[k], kind: 'bin' });
  if (over !== undefined) bins.push({ from: over, to: Infinity, count: above, kind: 'over' });
  return { bins, width, capped };
}

/** The share of the total each of `counts` adds up to, in turn: the line of a pareto chart. */
export function runningShares(counts: ArrayLike<number>): number[] {
  let total = 0;
  for (let i = 0; i < counts.length; i++) if (counts[i] > 0) total += counts[i];
  const out: number[] = [];
  let sum = 0;
  for (let i = 0; i < counts.length; i++) {
    if (counts[i] > 0) sum += counts[i];
    out.push(total > 0 ? sum / total : 0);
  }
  return out;
}

// ------------------------------------------------------------ box and whisker

/**
 * The quartile `p` (0.25, 0.5, 0.75) of sorted values. "inclusive" is QUARTILE.INC (the place
 * p(n-1) counted from 0); "exclusive" is QUARTILE.EXC (the place p(n+1) counted from 1), held to
 * the first and the last value where that place lies outside them. UNSURE: these are taken to be
 * what the quartile calculation of Excel's box and whisker chart means by its two methods.
 */
export function quartile(sorted: ArrayLike<number>, p: number, method: 'inclusive' | 'exclusive'): number {
  const n = sorted.length;
  if (n === 0) return NaN;
  if (n === 1) return sorted[0];
  let place = method === 'inclusive' ? p * (n - 1) : p * (n + 1) - 1;
  if (place <= 0) return sorted[0];
  if (place >= n - 1) return sorted[n - 1];
  const below = Math.floor(place);
  place -= below;
  return sorted[below] + (sorted[below + 1] - sorted[below]) * place;
}

export interface BoxStats {
  q1: number;
  median: number;
  q3: number;
  mean: number;
  /** The ends of the whiskers: the lowest and the highest value within 1.5 times the height of the box beyond it. */
  low: number;
  high: number;
  /** The values beyond the whiskers. */
  outliers: number[];
  /** The values between the whiskers. */
  inner: number[];
  count: number;
}

/** What a box of a box and whisker chart shows of `values` (finite numbers; anything else is left out). Undefined for none. */
export function boxStats(values: ArrayLike<number>, method: 'inclusive' | 'exclusive'): BoxStats | undefined {
  const sorted: number[] = [];
  for (let i = 0; i < values.length; i++) {
    const value = values[i];
    if (typeof value === 'number' && Number.isFinite(value)) sorted.push(value);
  }
  if (sorted.length === 0) return undefined;
  sorted.sort((a, b) => a - b);
  const q1 = quartile(sorted, 0.25, method);
  const median = quartile(sorted, 0.5, method);
  const q3 = quartile(sorted, 0.75, method);
  const reach = (q3 - q1) * 1.5;
  let sum = 0;
  let low = Infinity;
  let high = -Infinity;
  const outliers: number[] = [];
  const inner: number[] = [];
  for (const value of sorted) {
    sum += value;
    if (value < q1 - reach || value > q3 + reach) outliers.push(value);
    else {
      inner.push(value);
      if (value < low) low = value;
      if (value > high) high = value;
    }
  }
  if (low === Infinity) {
    low = q1;
    high = q3;
  }
  return { q1, median, q3, mean: sum / sorted.length, low, high, outliers, inner, count: sorted.length };
}

// ------------------------------------------------------------------ waterfall

export interface WaterfallStep {
  /** Where the bar starts and ends. */
  from: number;
  to: number;
  kind: 'increase' | 'decrease' | 'total';
}

/**
 * The bars of a waterfall: every value is added to the running total, and a bar goes from the
 * total before it to the total after it. A point that is a total stands on zero and is the
 * running total from there on. A point without a number has no bar and changes nothing.
 * UNSURE: a total is its own value (Excel takes the value of the cell as the total).
 */
export function waterfallSteps(values: ArrayLike<number | null>, totals: ReadonlySet<number>): (WaterfallStep | null)[] {
  const out: (WaterfallStep | null)[] = [];
  let running = 0;
  for (let i = 0; i < values.length; i++) {
    const value = values[i];
    if (typeof value !== 'number' || !Number.isFinite(value)) {
      out.push(null);
      continue;
    }
    if (totals.has(i)) {
      out.push({ from: 0, to: value, kind: 'total' });
      running = value;
      continue;
    }
    out.push({ from: running, to: running + value, kind: value < 0 ? 'decrease' : 'increase' });
    running += value;
  }
  return out;
}

// ------------------------------------------------------------------ hierarchy

export interface TreeNode {
  name: string;
  /** The sum of the values of the points under it. */
  value: number;
  /** How far down it is: 0 for the children of the root. */
  depth: number;
  children: TreeNode[];
  /** A leaf: the point it is made of (the first, when several points give one path). */
  point?: number;
  /** The place among the children of the root of the branch it is in: what colours it. */
  branch: number;
}

/** Points a hierarchy is built of, at most; further points are left out. */
export const MAX_TREE_POINTS = 100000;

/**
 * The hierarchy of a treemap or a sunburst. Point `i` is the path of the names `name(level, i)`
 * from the outermost level (`depth` - 1) to the leaves (0); an empty name ends a path early
 * (that point is a leaf further up). Only values above zero count. Children are listed by their
 * value, largest first; UNSURE: Excel is taken to sort them that way.
 */
export function buildTree(count: number, depth: number, name: (level: number, index: number) => string, value: (index: number) => number | null): TreeNode {
  const root: TreeNode = { name: '', value: 0, depth: -1, children: [], branch: -1 };
  const levels = Math.max(Math.min(Math.floor(depth), 64), 1);
  const lookup = new Map<TreeNode, Map<string, TreeNode>>();
  const points = Math.min(Math.max(Math.floor(count), 0), MAX_TREE_POINTS);
  for (let i = 0; i < points; i++) {
    const size = value(i);
    if (typeof size !== 'number' || !Number.isFinite(size) || !(size > 0)) continue;
    let node = root;
    root.value += size;
    let named = false;
    for (let level = levels - 1; level >= 0; level--) {
      const text = name(level, i);
      if (text === '') {
        // A gap above the leaves is passed over; a gap below what was named ends the path.
        if (named) break;
        continue;
      }
      named = true;
      let map = lookup.get(node);
      if (!map) lookup.set(node, (map = new Map()));
      let next = map.get(text);
      if (!next) {
        next = { name: text, value: 0, depth: node.depth + 1, children: [], branch: node.depth < 0 ? node.children.length : node.branch };
        map.set(text, next);
        node.children.push(next);
      }
      next.value += size;
      node = next;
    }
    if (!named) {
      const leaf: TreeNode = { name: '', value: size, depth: 0, children: [], branch: root.children.length, point: i };
      root.children.push(leaf);
      continue;
    }
    if (node.point === undefined) node.point = i;
  }
  // Largest first; the branch of a child of the root is its place once sorted.
  const stack: TreeNode[] = [root];
  while (stack.length > 0) {
    const node = stack.pop() as TreeNode;
    if (node.children.length > 1) node.children.sort((a, b) => b.value - a.value);
    for (const child of node.children) stack.push(child);
  }
  const relabel = (node: TreeNode, branch: number): void => {
    const nodes: TreeNode[] = [node];
    while (nodes.length > 0) {
      const at = nodes.pop() as TreeNode;
      at.branch = branch;
      for (const child of at.children) nodes.push(child);
    }
  };
  root.children.forEach((child, i) => relabel(child, i));
  return root;
}

/** The deepest level any leaf of a tree lies at, plus one: how many rings a sunburst has. */
export function treeDepth(root: TreeNode): number {
  let deepest = 0;
  const stack: TreeNode[] = root.children.slice();
  while (stack.length > 0) {
    const node = stack.pop() as TreeNode;
    if (node.depth + 1 > deepest) deepest = node.depth + 1;
    for (const child of node.children) stack.push(child);
  }
  return deepest;
}

export interface Rect {
  x: number;
  y: number;
  width: number;
  height: number;
}

/**
 * Divides a rectangle among `values` (largest first; nothing that is not above zero gets room)
 * so that the pieces are as near to squares as can be: the squarified layout of Bruls, Huizing
 * and van Wijk. The pieces are in the order of the values; one that gets no room is undefined.
 */
export function squarify(values: readonly number[], box: Rect): (Rect | undefined)[] {
  const out: (Rect | undefined)[] = values.map(() => undefined);
  let total = 0;
  for (const value of values) if (value > 0 && Number.isFinite(value)) total += value;
  if (!(total > 0) || !(box.width > 0) || !(box.height > 0)) return out;
  const scale = (box.width * box.height) / total;
  let { x, y, width, height } = box;
  let i = 0;
  const n = values.length;
  const area = (k: number): number => (values[k] > 0 && Number.isFinite(values[k]) ? values[k] * scale : 0);
  /** The worst ratio of the sides of the pieces of a row of areas `sum` (largest `big`, smallest `small`) along a side `side`. */
  const worst = (sum: number, big: number, small: number, side: number): number => {
    const s2 = sum * sum;
    const w2 = side * side;
    return Math.max((w2 * big) / s2, s2 / (w2 * small));
  };
  while (i < n) {
    if (!(area(i) > 0)) {
      i++;
      continue;
    }
    const side = Math.min(width, height);
    let end = i + 1;
    let sum = area(i);
    let big = sum;
    let small = sum;
    let ratio = worst(sum, big, small, side);
    while (end < n) {
      const next = area(end);
      if (!(next > 0)) {
        end++;
        continue;
      }
      const after = worst(sum + next, Math.max(big, next), Math.min(small, next), side);
      if (after > ratio) break;
      sum += next;
      big = Math.max(big, next);
      small = Math.min(small, next);
      ratio = after;
      end++;
    }
    // The row lies along the shorter side.
    if (width >= height) {
      const rowWidth = height > 0 ? sum / height : 0;
      let at = y;
      for (let k = i; k < end; k++) {
        const piece = area(k);
        if (!(piece > 0)) continue;
        const h = rowWidth > 0 ? piece / rowWidth : 0;
        out[k] = { x, y: at, width: rowWidth, height: h };
        at += h;
      }
      x += rowWidth;
      width = Math.max(width - rowWidth, 0);
    } else {
      const rowHeight = width > 0 ? sum / width : 0;
      let at = x;
      for (let k = i; k < end; k++) {
        const piece = area(k);
        if (!(piece > 0)) continue;
        const w = rowHeight > 0 ? piece / rowHeight : 0;
        out[k] = { x: at, y, width: w, height: rowHeight };
        at += w;
      }
      y += rowHeight;
      height = Math.max(height - rowHeight, 0);
    }
    i = end;
  }
  return out;
}

// -------------------------------------------------------------------- surface

/**
 * The part of a triangle (x, y and value at each corner) whose values lie between `low` and
 * `high`, values in between taken linearly: the corners of a polygon, x and y in turn; empty
 * when no part does.
 */
export function bandPolygon(corners: readonly number[], low: number, high: number): number[] {
  // Each corner: x, y, v.
  let poly: number[] = corners.slice(0, 9);
  const clip = (input: number[], keep: (v: number) => boolean, level: number): number[] => {
    const out: number[] = [];
    const count = input.length / 3;
    for (let i = 0; i < count; i++) {
      const a = i * 3;
      const b = ((i + 1) % count) * 3;
      const ina = keep(input[a + 2]);
      const inb = keep(input[b + 2]);
      if (ina) out.push(input[a], input[a + 1], input[a + 2]);
      if (ina !== inb) {
        const dv = input[b + 2] - input[a + 2];
        const t = dv === 0 ? 0 : (level - input[a + 2]) / dv;
        out.push(input[a] + (input[b] - input[a]) * t, input[a + 1] + (input[b + 1] - input[a + 1]) * t, level);
      }
    }
    return out;
  };
  poly = clip(poly, (v) => v >= low, low);
  if (poly.length < 9) return [];
  poly = clip(poly, (v) => v <= high, high);
  if (poly.length < 9) return [];
  const out: number[] = [];
  for (let i = 0; i < poly.length; i += 3) out.push(poly[i], poly[i + 1]);
  return out;
}

/** Where a triangle (x, y and value at each corner) is crossed by the value `level`: a line, x and y of its ends in turn; empty where it is not. */
export function levelSegment(corners: readonly number[], level: number): number[] {
  const out: number[] = [];
  for (let i = 0; i < 3; i++) {
    const a = i * 3;
    const b = ((i + 1) % 3) * 3;
    const va = corners[a + 2];
    const vb = corners[b + 2];
    if ((va < level && vb >= level) || (vb < level && va >= level)) {
      const t = (level - va) / (vb - va);
      out.push(corners[a] + (corners[b] - corners[a]) * t, corners[a + 1] + (corners[b + 1] - corners[a + 1]) * t);
    }
  }
  return out.length === 4 ? out : [];
}

// Lays out the charts that are not drawn along axes of series: the newer chart
// types (histogram and pareto, waterfall, funnel, box and whisker, treemap,
// sunburst) and the surface chart, seen from above as bands of colour. Like
// layout/chart.ts, of which it is a part: pure calculation, a flat list of
// items for the painter. What it needs of the layout of the other charts
// (fonts, colours, fitting text, data labels) it is handed as a kit, so that
// it imports nothing of that module but its types.
//
// Excel is not at hand to compare with: rules that follow what Excel is known
// to do but could not be checked against it are marked UNSURE.

import type { ChartAxis, ChartDataLabels, ChartGroup, ChartLine, ChartSeries, ChartTextStyle, Chart } from '../model/sheet/chart';
import type { ChartFont, ChartItem, ChartLabelAnchor, ChartLayout, ChartLayoutOptions, ChartMeasure, ChartScale, ChartSeriesValues, ChartValues } from './chart';
import { bandPolygon, boxStats, buildTree, histogramBins, levelSegment, runningShares, squarify, treeDepth, waterfallSteps, type BinRule, type BoxStats, type Rect, type TreeNode } from './chart-calc';

/** Bars, boxes or categories drawn of one series of the newer chart types; what lies beyond is left out. */
export const MAX_EX_CATEGORIES = 2000;
/** Rings of a sunburst, and levels of a treemap, that are drawn. */
export const MAX_EX_DEPTH = 8;

/** The labels of one series, as the kit's placer knows them. */
export type ExSeriesLabels = { base: ChartDataLabels | undefined; own: Map<number, ChartDataLabels> | undefined };

/** Places data labels so that none lies over another (the LabelPlacer of layout/chart.ts). */
export interface ExPlacer {
  readonly items: ChartItem[];
  of(entry: ChartSeriesValues): ExSeriesLabels | undefined;
  add(entry: ChartSeriesValues, labels: ExSeriesLabels, index: number, value: number, anchor: ChartLabelAnchor): void;
}

/** What the layout of the other charts lends to this one, for one chart in one frame. */
export interface ExKit {
  chart: Chart;
  values: ChartValues;
  measure: ChartMeasure;
  options: ChartLayoutOptions;
  items: ChartItem[];
  detail: ChartLayout['detail'];
  width: number;
  height: number;
  /** What may be drawn, all series together: MAX_DRAWN_ITEMS or less. */
  budget: number;
  /** A font of the chart: what `style` states, else the text of the chart, else `points` points. */
  font(style: ChartTextStyle | undefined, points: number, bold?: boolean): ChartFont;
  color(style: ChartTextStyle | undefined): string;
  clean(text: string): string;
  fit(text: string, font: ChartFont, room: number): string;
  /** A number as an axis or a label shows it. */
  tick(value: number, format?: { code: string; builtinId?: number }): string;
  line(line: ChartLine | undefined, automatic: string | undefined): string | undefined;
  /** The colour the chart gives its n-th series or point that states none. */
  auto(index: number): string;
  /** What a series is drawn with. */
  look(group: ChartGroup, series: ChartSeries): { color: string; fill: string | undefined; stroke: string | undefined; lineWidth: number };
  /** A scale for values from `low` to `high`, as the axis states it, in at most `steps` steps. */
  scale(low: number, high: number, axis: ChartAxis | undefined, steps: number): ChartScale;
  placer(): ExPlacer;
}

interface Box {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface ExLegendEntry {
  label: string;
  color: string;
  kind: 'box' | 'line';
}

const GRID_COLOR = '#D9D9D9';
const AXIS_COLOR = '#BFBFBF';
const CONNECTOR_COLOR = '#A6A6A6';
const HEX_COLOR = /^#[0-9A-Fa-f]{6}$/;
const clamp = (value: number, low: number, high: number): number => (value < low ? low : value > high ? high : value);
const safe = (color: string | null | undefined): string | undefined => (typeof color === 'string' && HEX_COLOR.test(color) ? color : undefined);

/** A colour made darker (`factor` below 1) or lighter (above 1, towards white). */
function shade(color: string, factor: number): string {
  const base = safe(color) || '#808080';
  let out = '#';
  for (let i = 0; i < 3; i++) {
    const channel = parseInt(base.slice(1 + i * 2, 3 + i * 2), 16);
    const value = factor <= 1 ? channel * factor : channel + (255 - channel) * (factor - 1);
    const byte = Math.round(clamp(value, 0, 255));
    out += (byte < 16 ? '0' : '') + byte.toString(16).toUpperCase();
  }
  return out;
}

/** Black or white, whichever reads better on `color`. */
function textOn(color: string | undefined): string {
  const base = safe(color);
  if (!base) return '#000000';
  const [r, g, b] = [1, 3, 5].map((at) => parseInt(base.slice(at, at + 2), 16) / 255);
  return 0.299 * r + 0.587 * g + 0.114 * b > 0.6 ? '#000000' : '#FFFFFF';
}

/** The fill a point of a series states for itself (c:dPt / cx:dataPt); undefined when it states none. */
function pointFills(series: ChartSeries): Map<number, string | null> {
  const out = new Map<number, string | null>();
  for (const point of series.points || []) {
    if (!point || !point.shape || point.shape.fill === undefined || out.has(point.index)) continue;
    out.set(point.index, point.shape.fill === null ? null : safe(point.shape.fill) || null);
  }
  return out;
}

function axisOf(chart: Chart, group: ChartGroup, kind: 'cat' | 'val'): ChartAxis | undefined {
  for (const id of group.axisIds || []) {
    const axis = chart.axes.find((candidate) => candidate.id === id && candidate.kind === kind);
    if (axis) return axis;
  }
  return chart.axes.find((axis) => axis.kind === kind);
}

const labelsShown = (axis: ChartAxis | undefined): boolean => !(axis && (axis.deleted || axis.tickLabels === 'none'));

// ----------------------------------------------------------- columns on axes

interface ColumnFrame {
  plot: Box;
  scale: ChartScale;
  /** The place of a value on the axis of the values. */
  y(value: number): number;
  /** The width of one category. */
  slot: number;
  /** The place of a share on the second axis (0..1), when there is one. */
  y2?(share: number): number;
}

interface ColumnSpec {
  count: number;
  low: number;
  high: number;
  label(index: number): string;
  catAxis: ChartAxis | undefined;
  valAxis: ChartAxis | undefined;
  format?: { code: string; builtinId?: number };
  /** A second axis on the right, from 0% to 100%. */
  shares?: boolean;
}

/**
 * The frame of a chart of columns: the labels of the values on the left (and of the shares on the
 * right), those of the categories below, the gridlines; what is left is the plot area.
 */
function columnFrame(kit: ExKit, layout: ChartLayout, box: Box, spec: ColumnSpec): ColumnFrame {
  const { chart, items, measure } = kit;
  const words = kit.detail === 'full';
  const catFont = kit.font(spec.catAxis ? spec.catAxis.labelStyle : undefined, 9);
  const catColor = kit.color(spec.catAxis ? spec.catAxis.labelStyle : undefined);
  const valFont = kit.font(spec.valAxis ? spec.valAxis.labelStyle : undefined, 9);
  const valColor = kit.color(spec.valAxis ? spec.valAxis.labelStyle : undefined);
  const catLabels = words && labelsShown(spec.catAxis) && spec.count > 0;
  const valLabels = words && labelsShown(spec.valAxis);
  const plot: Box = { ...box };
  if (catLabels) plot.height -= catFont.size * 1.4 + 4;
  const steps = clamp(Math.floor(plot.height / Math.max(valFont.size * 2.2, 8)), 1, 10);
  const scale = kit.scale(spec.low, spec.high, spec.valAxis, steps);
  layout.scale = scale;
  const format = spec.valAxis && spec.valAxis.formatCode && spec.valAxis.sourceLinked === false ? { code: spec.valAxis.formatCode } : spec.format;
  const ticks = valLabels ? scale.ticks.map((tick) => kit.tick(tick, format)) : [];
  if (valLabels) {
    let widest = 0;
    for (const text of ticks) widest = Math.max(widest, measure(text, valFont));
    const taken = Math.min(widest, box.width * 0.3) + 6;
    plot.x += taken;
    plot.width -= taken;
  }
  const shares = spec.shares ? [0, 0.2, 0.4, 0.6, 0.8, 1] : [];
  if (spec.shares && words) {
    let widest = 0;
    for (const share of shares) widest = Math.max(widest, measure(kit.tick(share, { code: '0%' }), valFont));
    plot.width -= Math.min(widest, box.width * 0.2) + 6;
  }
  if (plot.width < 1) plot.width = 1;
  if (plot.height < 1) plot.height = 1;
  layout.plot = { ...plot };
  const span = scale.max - scale.min;
  const y = (value: number): number => plot.y + plot.height - clamp(span > 0 ? (value - scale.min) / span : 0, 0, 1) * plot.height;
  const y2 = (share: number): number => plot.y + plot.height - clamp(share, 0, 1) * plot.height;
  if (kit.detail !== 'bare') {
    // Gridlines where the axis of the values asks for them.
    const grid = spec.valAxis ? spec.valAxis.majorGridlines : {};
    const stroke = grid ? kit.line(grid, GRID_COLOR) : undefined;
    if (grid && stroke && !(spec.valAxis && spec.valAxis.deleted)) {
      const points: number[] = [];
      for (const tick of scale.ticks) points.push(plot.x, y(tick), plot.x + plot.width, y(tick));
      if (points.length > 0) items.push({ kind: 'path', role: 'grid', points, segments: true, stroke, lineWidth: clamp(grid.width || 1, 0.5, 8), dash: grid.dash === true });
    }
    const axisLine = spec.catAxis ? kit.line(spec.catAxis.line, AXIS_COLOR) : AXIS_COLOR;
    if (axisLine && !(spec.catAxis && spec.catAxis.deleted)) {
      const base = y(clamp(0, scale.min, scale.max));
      items.push({ kind: 'path', role: 'axis', points: [plot.x, base, plot.x + plot.width, base], stroke: axisLine, lineWidth: 1 });
    }
  }
  if (valLabels) {
    for (let i = 0; i < scale.ticks.length; i++) {
      const text = kit.fit(ticks[i], valFont, Math.max(plot.x - box.x - 4, 0));
      if (text !== '') items.push({ kind: 'text', role: 'tick', text, x: plot.x - 4, y: y(scale.ticks[i]) + valFont.size * 0.35, font: valFont, color: valColor, align: 'right' });
    }
  }
  if (spec.shares && words) {
    for (const share of shares) {
      const text = kit.tick(share, { code: '0%' });
      items.push({ kind: 'text', role: 'tick', text, x: plot.x + plot.width + 4, y: y2(share) + valFont.size * 0.35, font: valFont, color: valColor, align: 'left' });
    }
  }
  const slot = plot.width / Math.max(spec.count, 1);
  if (catLabels) {
    // As many labels as there is room for: every n-th one where they would run into each other.
    let widest = 0;
    const stride = Math.max(Math.floor(spec.count / 200), 1);
    for (let i = 0; i < spec.count; i += stride) widest = Math.max(widest, measure(kit.clean(spec.label(i)), catFont));
    const every = Math.max(Math.ceil((Math.min(widest, plot.width * 0.5) + 6) / Math.max(slot, 1)), 1);
    const room = slot * every - 4;
    const top = plot.y + plot.height + catFont.size + 3;
    for (let i = 0; i < spec.count; i += every) {
      const text = kit.fit(kit.clean(spec.label(i)), catFont, room);
      if (text !== '') items.push({ kind: 'text', role: 'category', text, x: plot.x + slot * (i + 0.5), y: top, font: catFont, color: catColor, align: 'center' });
    }
  }
  void chart;
  return { plot, scale, y, slot, y2: spec.shares ? y2 : undefined };
}

// ------------------------------------------------------------------- waterfall

/** The colours of the bars of a waterfall: increase, decrease, total. UNSURE: the first three colours of the chart, as Excel gives a new one. */
function waterfallColors(kit: ExKit): [string, string, string] {
  return [kit.auto(0), kit.auto(1), kit.auto(2)];
}

function layoutWaterfall(kit: ExKit, layout: ChartLayout, box: Box, entry: ChartSeriesValues): void {
  const { chart, values, items } = kit;
  const group = chart.groups[entry.group];
  const series = entry.series;
  const count = Math.min(entry.values.length, MAX_EX_CATEGORIES, Math.max(kit.budget, 1));
  const list = entry.values.toArray(count);
  const totals = new Set<number>((series.ex && series.ex.subtotals) || []);
  const steps = waterfallSteps(list, totals);
  let low = 0;
  let high = 0;
  for (const step of steps) {
    if (!step) continue;
    low = Math.min(low, step.from, step.to);
    high = Math.max(high, step.from, step.to);
  }
  const frame = columnFrame(kit, layout, box, { count, low, high, label: (i) => values.label(i), catAxis: axisOf(chart, group, 'cat'), valAxis: axisOf(chart, group, 'val'), format: entry.format });
  const colors = waterfallColors(kit);
  const fills = pointFills(series);
  const look = kit.look(group, series);
  // UNSURE: half a bar between two bars where the file states no gap.
  const gap = group.gapWidth !== undefined ? group.gapWidth : 50;
  const width = frame.slot / (1 + gap / 100);
  const placer = kit.detail === 'full' ? kit.placer() : undefined;
  const labels = placer ? placer.of(entry) : undefined;
  const index = values.series.indexOf(entry);
  const lines: number[] = [];
  let before: { x: number; level: number } | undefined;
  for (let i = 0; i < steps.length; i++) {
    const step = steps[i];
    if (!step) continue;
    const x = frame.plot.x + frame.slot * i + (frame.slot - width) / 2;
    const top = frame.y(Math.max(step.from, step.to));
    const bottom = frame.y(Math.min(step.from, step.to));
    const own = fills.get(i);
    const fill = own !== undefined ? own || undefined : step.kind === 'increase' ? colors[0] : step.kind === 'decrease' ? colors[1] : colors[2];
    items.push({ kind: 'rect', role: 'bar', x, y: top, width, height: Math.max(bottom - top, 0), fill, stroke: look.stroke, lineWidth: look.lineWidth, series: index, point: i });
    if (before && !(series.ex && series.ex.connectorLines === false)) {
      const level = frame.y(before.level);
      lines.push(before.x, level, x, level);
    }
    before = { x: x + width, level: step.to };
    const value = list[i];
    if (placer && labels && typeof value === 'number') placer.add(entry, labels, i, value, { kind: 'bar', x, y: top, width, height: Math.max(bottom - top, 0), lying: false, low: step.to >= step.from, stacked: false });
  }
  if (lines.length > 0) items.push({ kind: 'path', role: 'line', points: lines, segments: true, stroke: CONNECTOR_COLOR, lineWidth: 1 });
  if (placer) for (const item of placer.items) items.push(item);
}

// ---------------------------------------------------------------------- funnel

function layoutFunnel(kit: ExKit, layout: ChartLayout, box: Box, entry: ChartSeriesValues): void {
  const { chart, values, items, measure } = kit;
  const group = chart.groups[entry.group];
  const count = Math.min(entry.values.length, MAX_EX_CATEGORIES, Math.max(kit.budget, 1));
  const list = entry.values.toArray(count);
  let most = 0;
  for (const value of list) if (typeof value === 'number' && value > most) most = value;
  const catAxis = axisOf(chart, group, 'cat');
  const font = kit.font(catAxis ? catAxis.labelStyle : undefined, 9);
  const color = kit.color(catAxis ? catAxis.labelStyle : undefined);
  const words = kit.detail === 'full' && labelsShown(catAxis) && count > 0;
  const plot: Box = { ...box };
  const row = box.height / Math.max(count, 1);
  const shown = words && row >= font.size * 0.9;
  if (shown) {
    let widest = 0;
    for (let i = 0; i < count; i++) widest = Math.max(widest, measure(kit.clean(values.label(i)), font));
    const taken = Math.min(widest, box.width * 0.35) + 8;
    plot.x += taken;
    plot.width = Math.max(plot.width - taken, 1);
  }
  layout.plot = { ...plot };
  // UNSURE: a tenth of a bar between two bars where the file states no gap.
  const gap = group.gapWidth !== undefined ? group.gapWidth : 10;
  const height = row / (1 + gap / 100);
  const look = kit.look(group, entry.series);
  const fills = pointFills(entry.series);
  const placer = kit.detail === 'full' ? kit.placer() : undefined;
  const labels = placer ? placer.of(entry) : undefined;
  const index = values.series.indexOf(entry);
  const middle = plot.x + plot.width / 2;
  for (let i = 0; i < count; i++) {
    const top = plot.y + row * i + (row - height) / 2;
    if (shown) {
      const text = kit.fit(kit.clean(values.label(i)), font, plot.x - box.x - 8);
      if (text !== '') items.push({ kind: 'text', role: 'category', text, x: plot.x - 8, y: top + height / 2 + font.size * 0.35, font, color, align: 'right' });
    }
    const value = list[i];
    if (typeof value !== 'number' || !(value > 0) || !(most > 0)) continue;
    const width = (value / most) * plot.width;
    const own = fills.get(i);
    const fill = own !== undefined ? own || undefined : look.fill || look.color;
    items.push({ kind: 'rect', role: 'bar', x: middle - width / 2, y: top, width, height, fill, stroke: look.stroke, lineWidth: look.lineWidth, series: index, point: i });
    // The labels stand in the middle of their bars unless the file says otherwise.
    if (placer && labels) placer.add(entry, labels, i, value, { kind: 'bar', x: middle - width / 2, y: top, width, height, lying: true, low: false, stacked: true });
  }
  if (placer) for (const item of placer.items) items.push(item);
}

// ------------------------------------------------------------------- histogram

interface Bars {
  labels: string[];
  counts: number[];
}

/** The text of the edges of a bin. */
function binLabel(kit: ExKit, bin: { from: number; to: number; kind: string }, first: boolean, last: boolean, closed: 'l' | 'r', format: { code: string } | undefined): string {
  const text = (value: number) => kit.tick(value, format);
  if (bin.kind === 'under') return `${closed === 'r' ? '≤' : '<'}${text(bin.to)}`;
  if (bin.kind === 'over') return `${closed === 'r' ? '>' : '≥'}${text(bin.from)}`;
  // The first bin holds its start, and the last its end, unless a bin of their own holds what lies beyond.
  const open = closed === 'r' && !first ? '(' : '[';
  const shut = closed === 'l' && !last ? ')' : ']';
  return `${open}${text(bin.from)}, ${text(bin.to)}${shut}`;
}

/** The bars of a histogram: its bins, or the sums of its categories. */
function histogramBars(kit: ExKit, entry: ChartSeriesValues): Bars {
  const ex = entry.series.ex || {};
  const values = kit.values;
  if (ex.aggregate || (!ex.binning && entry.series.categories !== undefined && entry.series.categories.kind === 'str')) {
    // By category: the values of equal categories added up. UNSURE: that is what Excel does with a histogram "by category".
    const order = new Map<string, number>();
    const labels: string[] = [];
    const counts: number[] = [];
    entry.values.each((value, index) => {
      const label = values.label(index);
      let at = order.get(label);
      if (at === undefined) {
        if (labels.length >= MAX_EX_CATEGORIES) return;
        at = labels.length;
        order.set(label, at);
        labels.push(label);
        counts.push(0);
      }
      counts[at] += value;
    });
    return { labels, counts };
  }
  const numbers: number[] = [];
  entry.values.each((value) => void numbers.push(value));
  const binning = ex.binning;
  const rule: BinRule = { closed: binning ? binning.closed : 'r' };
  if (binning) {
    if (typeof binning.underflow === 'number') rule.underflow = binning.underflow;
    if (typeof binning.overflow === 'number') rule.overflow = binning.overflow;
    if (binning.size !== undefined) rule.size = binning.size;
    if (binning.count !== undefined) rule.count = binning.count;
  }
  const result = histogramBins(numbers, rule);
  const format = entry.format && entry.format.code !== '' && !/^general$/i.test(entry.format.code) ? { code: entry.format.code } : undefined;
  const bins = result.bins.slice(0, MAX_EX_CATEGORIES);
  const regular = bins.filter((bin) => bin.kind === 'bin');
  return {
    labels: bins.map((bin) => binLabel(kit, bin, bin === regular[0] && rule.underflow === undefined, bin === regular[regular.length - 1] && rule.overflow === undefined, rule.closed, format)),
    counts: bins.map((bin) => bin.count),
  };
}

function layoutHistogram(kit: ExKit, layout: ChartLayout, box: Box, entry: ChartSeriesValues): void {
  const { chart, values, items } = kit;
  const group = chart.groups[entry.group];
  const series = entry.series;
  const bars = histogramBars(kit, entry);
  const pareto = series.ex && series.ex.pareto;
  if (pareto) {
    // A pareto chart: the bars from the largest down, and the share they add up to.
    const order = bars.counts.map((_count, i) => i).sort((a, b) => bars.counts[b] - bars.counts[a] || a - b);
    bars.labels = order.map((i) => bars.labels[i]);
    bars.counts = order.map((i) => bars.counts[i]);
  }
  const count = Math.min(bars.counts.length, Math.max(kit.budget, 1));
  let low = 0;
  let high = 0;
  for (let i = 0; i < count; i++) {
    low = Math.min(low, bars.counts[i]);
    high = Math.max(high, bars.counts[i]);
  }
  const frame = columnFrame(kit, layout, box, { count, low, high, label: (i) => bars.labels[i] || '', catAxis: axisOf(chart, group, 'cat'), valAxis: axisOf(chart, group, 'val'), shares: pareto !== undefined });
  const look = kit.look(group, series);
  const gap = group.gapWidth !== undefined ? group.gapWidth : 0;
  const width = frame.slot / (1 + gap / 100);
  const index = values.series.indexOf(entry);
  const placer = kit.detail === 'full' ? kit.placer() : undefined;
  const labels = placer ? placer.of(entry) : undefined;
  const base = frame.y(clamp(0, frame.scale.min, frame.scale.max));
  for (let i = 0; i < count; i++) {
    const x = frame.plot.x + frame.slot * i + (frame.slot - width) / 2;
    const top = frame.y(bars.counts[i]);
    const y = Math.min(top, base);
    const height = Math.abs(base - top);
    // UNSURE: a histogram is outlined in white between its bars where it states no outline (Excel's look).
    items.push({ kind: 'rect', role: 'bar', x, y, width, height, fill: look.fill || look.color, stroke: look.stroke || (gap === 0 ? '#FFFFFF' : undefined), lineWidth: look.lineWidth, series: index, point: i });
    if (placer && labels) placer.add(entry, labels, i, bars.counts[i], { kind: 'bar', x, y, width, height, lying: false, low: bars.counts[i] >= 0, stacked: false });
  }
  if (pareto && frame.y2 && count > 0) {
    const shares = runningShares(bars.counts.slice(0, count));
    const points: number[] = [];
    for (let i = 0; i < count; i++) points.push(frame.plot.x + frame.slot * (i + 0.5), frame.y2(shares[i]));
    const stroke = kit.line(pareto.line, kit.auto(series.index + 1)) || kit.auto(series.index + 1);
    if (points.length >= 4) items.push({ kind: 'path', role: 'line', points, stroke, lineWidth: clamp((pareto.line && pareto.line.width) || 2, 0.5, 32), dash: pareto.line !== undefined && pareto.line.dash === true, series: index });
    else items.push({ kind: 'markers', role: 'marker', symbol: 'circle', size: 5, points, fill: stroke, series: index });
  }
  if (placer) for (const item of placer.items) items.push(item);
}

// -------------------------------------------------------------- box and whisker

function layoutBoxes(kit: ExKit, layout: ChartLayout, box: Box, shown: ChartSeriesValues[]): void {
  const { chart, values, items } = kit;
  const group = chart.groups[shown[0].group];
  // The categories: the distinct labels of the points in the order they first come, or one without a label.
  const order = new Map<string, number>();
  const names: string[] = [];
  const labelled = values.count > 0 && shown.some((entry) => entry.series.categories !== undefined);
  const categoryOf = (index: number): number => {
    const label = labelled ? values.label(index) : '';
    let at = order.get(label);
    if (at === undefined) {
      if (names.length >= MAX_EX_CATEGORIES) return -1;
      at = names.length;
      order.set(label, at);
      names.push(label);
    }
    return at;
  };
  const drawnSeries = shown.slice(0, Math.max(Math.min(shown.length, 64), 1));
  const stats: (BoxStats | undefined)[][] = [];
  let low = Infinity;
  let high = -Infinity;
  for (const entry of drawnSeries) {
    const buckets: number[][] = [];
    entry.values.each((value, index) => {
      const at = categoryOf(index);
      if (at < 0) return;
      while (buckets.length <= at) buckets.push([]);
      buckets[at].push(value);
    });
    const method = entry.series.ex && entry.series.ex.quartiles === 'inclusive' ? 'inclusive' : 'exclusive';
    const row = buckets.map((bucket) => boxStats(bucket, method));
    for (const item of row) {
      if (!item) continue;
      low = Math.min(low, item.low, item.mean, item.outliers.length > 0 ? item.outliers[0] : item.low);
      high = Math.max(high, item.high, item.mean, item.outliers.length > 0 ? item.outliers[item.outliers.length - 1] : item.high);
    }
    stats.push(row);
  }
  if (low === Infinity) {
    low = 0;
    high = 0;
  }
  const count = Math.max(names.length, 1);
  const frame = columnFrame(kit, layout, box, { count, low, high, label: (i) => names[i] || '', catAxis: axisOf(chart, group, 'cat'), valAxis: axisOf(chart, group, 'val'), format: shown[0].format });
  // UNSURE: a box and its neighbours of other series fill the category but for a gap as wide as one box.
  const gap = group.gapWidth !== undefined ? group.gapWidth : 100;
  const cluster = frame.slot / (1 + gap / 100);
  const width = cluster / drawnSeries.length;
  let left = kit.budget;
  drawnSeries.forEach((entry, s) => {
    const series = entry.series;
    const look = kit.look(group, series);
    const fill = look.fill || look.color;
    // UNSURE: the lines of a box in a darker form of its colour where the file states none.
    const stroke = look.stroke || shade(look.color, 0.6);
    const ex = series.ex || {};
    const index = values.series.indexOf(entry);
    const lines: number[] = [];
    const means: number[] = [];
    const outliers: number[] = [];
    const inner: number[] = [];
    const meanPath: number[] = [];
    stats[s].forEach((item, c) => {
      if (!item || left <= 0) return;
      left -= 6 + item.outliers.length + (ex.nonoutliers ? item.inner.length : 0);
      const x = frame.plot.x + frame.slot * c + (frame.slot - cluster) / 2 + width * s;
      const middle = x + width / 2;
      const w = width * 0.8;
      const bx = middle - w / 2;
      const q1 = frame.y(item.q1);
      const q3 = frame.y(item.q3);
      items.push({ kind: 'rect', role: 'bar', x: bx, y: Math.min(q1, q3), width: w, height: Math.abs(q1 - q3), fill, stroke, lineWidth: look.lineWidth, series: index, point: c });
      const lo = frame.y(item.low);
      const hi = frame.y(item.high);
      const cap = w / 2;
      // The median, the whiskers and their ends.
      lines.push(bx, frame.y(item.median), bx + w, frame.y(item.median));
      lines.push(middle, Math.max(q1, q3), middle, lo, middle, Math.min(q1, q3), middle, hi);
      lines.push(middle - cap / 2, lo, middle + cap / 2, lo, middle - cap / 2, hi, middle + cap / 2, hi);
      if (ex.meanMarker !== false) means.push(middle, frame.y(item.mean));
      if (ex.meanLine === true) meanPath.push(middle, frame.y(item.mean));
      if (ex.outliers !== false) for (const value of item.outliers) outliers.push(middle, frame.y(value));
      if (ex.nonoutliers === true) for (const value of item.inner) inner.push(middle, frame.y(value));
    });
    if (lines.length > 0) items.push({ kind: 'path', role: 'line', points: lines, segments: true, stroke, lineWidth: look.lineWidth, series: index });
    if (meanPath.length >= 4) items.push({ kind: 'path', role: 'line', points: meanPath, stroke, lineWidth: look.lineWidth, series: index });
    if (inner.length > 0) items.push({ kind: 'markers', role: 'marker', symbol: 'circle', size: 5, points: inner, stroke, series: index });
    if (outliers.length > 0) items.push({ kind: 'markers', role: 'marker', symbol: 'circle', size: 5, points: outliers, stroke, series: index });
    if (means.length > 0) items.push({ kind: 'markers', role: 'marker', symbol: 'x', size: 7, points: means, stroke, series: index });
  });
}

// ------------------------------------------------------------ treemap, sunburst

const TREES = new WeakMap<ChartValues, Map<ChartSeriesValues, TreeNode>>();

/** The hierarchy of a series of a treemap or a sunburst, worked out once for the values of a chart. */
export function treeOf(values: ChartValues, entry: ChartSeriesValues): TreeNode {
  let kept = TREES.get(values);
  if (!kept) TREES.set(values, (kept = new Map()));
  let tree = kept.get(entry);
  if (!tree) {
    const levels = values.levels;
    tree = levels
      ? buildTree(entry.values.length, Math.min(levels.depth, MAX_EX_DEPTH), (level, index) => levels.name(level, index), (index) => entry.values.at(index))
      : buildTree(entry.values.length, 1, (_level, index) => (entry.series.categories ? values.label(index) : ''), (index) => entry.values.at(index));
    kept.set(entry, tree);
  }
  return tree;
}

/** The colour of a node: that of its branch, or of its point where the point states one. */
function nodeColor(kit: ExKit, node: TreeNode, fills: Map<number, string | null>): string | undefined {
  if (node.point !== undefined && node.children.length === 0) {
    const own = fills.get(node.point);
    if (own !== undefined) return own || undefined;
  }
  return kit.auto(Math.max(node.branch, 0));
}

/** What the labels of the points of a treemap or a sunburst say: the name by itself unless the file says otherwise. */
function nodeText(kit: ExKit, labels: ChartDataLabels | undefined, node: TreeNode, format: { code: string; builtinId?: number } | undefined): string {
  if (labels && labels.deleted) return '';
  const parts: string[] = [];
  if (!labels || labels.showCategory === true) parts.push(kit.clean(node.name));
  if (labels && labels.showValue === true) parts.push(kit.tick(node.value, labels.formatCode !== undefined && labels.sourceLinked !== true ? { code: labels.formatCode } : format));
  return parts.filter((part) => part !== '').join(typeof labels?.separator === 'string' ? kit.clean(labels.separator) : ', ');
}

function layoutTreemap(kit: ExKit, layout: ChartLayout, box: Box, entry: ChartSeriesValues): void {
  const { chart, values, items } = kit;
  const tree = treeOf(values, entry);
  layout.plot = { ...box };
  const fills = pointFills(entry.series);
  const index = values.series.indexOf(entry);
  const words = kit.detail === 'full';
  const labels = entry.series.dataLabels;
  const font = kit.font(labels ? labels.style : undefined, 9);
  const parentFont = kit.font(labels ? labels.style : undefined, 9, true);
  const parents = (entry.series.ex && entry.series.ex.parentLabels) || 'overlapping';
  const texts: ChartItem[] = [];
  let left = kit.budget;
  // Breadth first, so that a budget that runs out leaves the deeper levels out.
  const queue: { node: TreeNode; rect: Rect }[] = [{ node: tree, rect: box }];
  for (let head = 0; head < queue.length && left > 0; head++) {
    const { node, rect } = queue[head];
    let inner = rect;
    const own = node.depth >= 0;
    if (own) {
      const fill = nodeColor(kit, node, fills);
      items.push({ kind: 'rect', role: 'bar', x: rect.x, y: rect.y, width: rect.width, height: rect.height, fill, stroke: '#FFFFFF', lineWidth: 1, series: index, point: node.point });
      left--;
      const leaf = node.children.length === 0 || node.depth + 1 >= MAX_EX_DEPTH;
      const fits = (f: ChartFont) => rect.height >= f.size + 4 && rect.width >= f.size * 2;
      if (words && leaf && fits(font)) {
        const text = kit.fit(nodeText(kit, labels, node, entry.format), font, rect.width - 8);
        if (text !== '') texts.push({ kind: 'text', role: 'label', text, x: rect.x + 4, y: rect.y + font.size + 2, font, color: (labels && labels.style && safe(labels.style.color)) || textOn(fill), align: 'left' });
      }
      if (!leaf && parents === 'banner' && rect.height > parentFont.size * 3) {
        // A band at the top of the parent holds its name; the children share what is left.
        const band = parentFont.size * 1.4 + 2;
        inner = { x: rect.x, y: rect.y + band, width: rect.width, height: rect.height - band };
        if (words) {
          const text = kit.fit(kit.clean(node.name), parentFont, rect.width - 8);
          if (text !== '') texts.push({ kind: 'text', role: 'label', text, x: rect.x + 4, y: rect.y + parentFont.size + 1, font: parentFont, color: textOn(fill), align: 'left' });
        }
      } else if (!leaf && parents === 'overlapping' && words && fits(parentFont)) {
        const text = kit.fit(kit.clean(node.name), parentFont, rect.width - 8);
        // Drawn after the children, over them.
        if (text !== '') texts.push({ kind: 'text', role: 'label', text, x: rect.x + 4, y: rect.y + parentFont.size + 2, font: parentFont, color: textOn(fill), align: 'left' });
      }
      if (leaf) continue;
    }
    const pieces = squarify(
      node.children.map((child) => child.value),
      inner,
    );
    for (let i = 0; i < node.children.length && queue.length < kit.budget * 2; i++) {
      const piece = pieces[i];
      if (!piece || piece.width < 0.5 || piece.height < 0.5) continue;
      queue.push({ node: node.children[i], rect: piece });
    }
  }
  // Labels of parents that overlap their children are drawn over the labels of the leaves: they come last.
  const ordered = texts.filter((item) => item.kind === 'text' && !item.font.bold).concat(texts.filter((item) => item.kind === 'text' && item.font.bold));
  for (const item of ordered) items.push(item);
  void chart;
}

function layoutSunburst(kit: ExKit, layout: ChartLayout, box: Box, entry: ChartSeriesValues): void {
  const { values, items, measure } = kit;
  const tree = treeOf(values, entry);
  const depth = Math.max(Math.min(treeDepth(tree), MAX_EX_DEPTH), 1);
  const fills = pointFills(entry.series);
  const index = values.series.indexOf(entry);
  const radius = Math.max(Math.min(box.width, box.height) / 2 - (kit.detail === 'full' ? 4 : 0), 1);
  const cx = box.x + box.width / 2;
  const cy = box.y + box.height / 2;
  layout.plot = { x: cx - radius, y: cy - radius, width: radius * 2, height: radius * 2 };
  // UNSURE: a hole in the middle half a ring wide.
  const ring = radius / (depth + 0.5);
  const hole = ring / 2;
  const labels = entry.series.dataLabels;
  const font = kit.font(labels ? labels.style : undefined, 9);
  const words = kit.detail === 'full';
  const texts: ChartItem[] = [];
  let left = kit.budget;
  const queue: { node: TreeNode; start: number; end: number }[] = [{ node: tree, start: -Math.PI / 2, end: (Math.PI * 3) / 2 }];
  for (let head = 0; head < queue.length && left > 0; head++) {
    const { node, start, end } = queue[head];
    if (node.depth >= 0) {
      const inner = hole + node.depth * ring;
      const outer = inner + ring;
      const fill = nodeColor(kit, node, fills) || '#FFFFFF';
      items.push({ kind: 'slice', cx, cy, radius: outer, inner, start, end, fill, stroke: '#FFFFFF', lineWidth: 1, series: index, point: node.point === undefined ? -1 : node.point });
      left--;
      if (words) {
        const text = nodeText(kit, labels, node, entry.format);
        const middle = (start + end) / 2;
        const along = (inner + outer) / 2;
        const width = measure(text, font);
        // A label stands in its piece only where the piece has room for it.
        if (text !== '' && width <= ring - 4 && font.size <= (end - start) * inner - 2) {
          texts.push({ kind: 'text', role: 'label', text, x: cx + Math.cos(middle) * along, y: cy + Math.sin(middle) * along + font.size * 0.35, font, color: (labels && labels.style && safe(labels.style.color)) || textOn(fill), align: 'center' });
        }
      }
      if (node.depth + 1 >= MAX_EX_DEPTH) continue;
    }
    const total = node.value;
    if (!(total > 0)) continue;
    let angle = start;
    for (const child of node.children) {
      const span = ((end - start) * child.value) / total;
      const next = angle + span;
      // A piece too thin to be seen is left out, with what lies under it.
      if (span * radius >= 0.3 && queue.length < kit.budget * 2) queue.push({ node: child, start: angle, end: next });
      angle = next;
    }
  }
  for (const item of texts) items.push(item);
}

// --------------------------------------------------------------------- surface

/** The scale of the values of a surface chart and the colour of each band between two of its ticks. */
export function surfaceBands(kit: Pick<ExKit, 'chart' | 'auto' | 'scale'>, shown: ChartSeriesValues[]): { scale: ChartScale; colors: string[] } {
  const { chart } = kit;
  let low = Infinity;
  let high = -Infinity;
  for (const entry of shown) {
    entry.values.each((value) => {
      if (value < low) low = value;
      if (value > high) high = value;
    });
  }
  if (low === Infinity) {
    low = 0;
    high = 0;
  }
  const group = shown.length > 0 ? chart.groups[shown[0].group] : chart.groups[0];
  const scale = kit.scale(low, high, group ? axisOf(chart, group, 'val') : undefined, 10);
  const colors: string[] = [];
  const bands = (group && group.bands) || [];
  for (let k = 0; k + 1 < scale.ticks.length; k++) {
    const own = bands.find((band) => band.index === k);
    // UNSURE: the bands take the colours of the chart in turn, from the lowest, as the series of other charts do.
    colors.push((own && own.shape.fill !== undefined && own.shape.fill !== null && safe(own.shape.fill)) || kit.auto(k));
  }
  return { scale, colors };
}

function layoutSurface(kit: ExKit, layout: ChartLayout, box: Box, shown: ChartSeriesValues[]): void {
  const { chart, values, items, measure } = kit;
  const words = kit.detail === 'full';
  const group = chart.groups[shown[0].group];
  const rows = shown.slice(0, 255);
  const { scale, colors } = surfaceBands(kit, rows);
  layout.scale = scale;
  const catAxis = axisOf(chart, group, 'cat');
  const catFont = kit.font(catAxis ? catAxis.labelStyle : undefined, 9);
  const catColor = kit.color(catAxis ? catAxis.labelStyle : undefined);
  const plot: Box = { ...box };
  const columns = Math.max(values.count, 1);
  if (words && labelsShown(catAxis)) plot.height -= catFont.size * 1.4 + 4;
  // The names of the series stand beside their rows, as the series axis of the file names them.
  const seriesAxis = chart.axes.find((axis) => axis.kind === 'cat' && axis !== catAxis && group.axisIds.indexOf(axis.id) >= 0);
  const rowNames = words && rows.length > 1 && labelsShown(seriesAxis);
  if (rowNames) {
    let widest = 0;
    for (const entry of rows) widest = Math.max(widest, measure(kit.clean(entry.name), catFont));
    const taken = Math.min(widest, box.width * 0.3) + 6;
    plot.x += taken;
    plot.width -= taken;
  }
  plot.width = Math.max(plot.width, 1);
  plot.height = Math.max(plot.height, 1);
  layout.plot = { ...plot };
  // Fewer columns and rows than the chart has where drawing all of them would cost more than a frame is worth.
  const pieces = Math.max(Math.floor(kit.budget / Math.max(Math.min(colors.length, 6), 1) / 2), 4);
  let across = Math.min(columns, Math.max(Math.floor(plot.width), 2), MAX_EX_CATEGORIES);
  let down = Math.min(rows.length, Math.max(Math.floor(plot.height), 2));
  while (across > 2 && Math.max(across - 1, 1) * Math.max(down - 1, 1) > pieces) across = Math.max(Math.floor(across / 2), 2);
  while (down > 2 && Math.max(across - 1, 1) * Math.max(down - 1, 1) > pieces) down = Math.max(Math.floor(down / 2), 2);
  const column = (i: number): number => (across >= columns ? i : Math.round((i * (columns - 1)) / Math.max(across - 1, 1)));
  const row = (j: number): number => (down >= rows.length ? j : Math.round((j * (rows.length - 1)) / Math.max(down - 1, 1)));
  // One column or one row is drawn as a strip: its values stand at both edges.
  const gridX = Math.max(across, 2);
  const gridY = Math.max(down, 2);
  const xAt = (i: number): number => plot.x + (i * plot.width) / (gridX - 1);
  const yAt = (j: number): number => plot.y + plot.height - (j * plot.height) / (gridY - 1);
  const grid: (number | null)[][] = [];
  for (let j = 0; j < gridY; j++) {
    const entry = rows[row(Math.min(j, down - 1))];
    const line: (number | null)[] = [];
    for (let i = 0; i < gridX; i++) line.push(entry ? entry.values.at(column(Math.min(i, across - 1))) : null);
    grid.push(line);
  }
  const index = values.series.indexOf(rows[0]);
  const wire = group.wireframe === true;
  const segments: number[][] = colors.map(() => []);
  let left = kit.budget;
  for (let j = 0; j + 1 < gridY && left > 0; j++) {
    for (let i = 0; i + 1 < gridX && left > 0; i++) {
      const a = grid[j][i];
      const b = grid[j][i + 1];
      const c = grid[j + 1][i + 1];
      const d = grid[j + 1][i];
      if (a === null || b === null || c === null || d === null) continue;
      const triangles = [
        [xAt(i), yAt(j), a, xAt(i + 1), yAt(j), b, xAt(i + 1), yAt(j + 1), c],
        [xAt(i), yAt(j), a, xAt(i + 1), yAt(j + 1), c, xAt(i), yAt(j + 1), d],
      ];
      for (const corners of triangles) {
        const lo = Math.min(corners[2], corners[5], corners[8]);
        const hi = Math.max(corners[2], corners[5], corners[8]);
        for (let k = 0; k < colors.length; k++) {
          const from = scale.ticks[k];
          const to = scale.ticks[k + 1];
          if (hi < from || lo > to) continue;
          if (wire) {
            const line = levelSegment(corners, to);
            if (line.length === 4 && k + 1 < colors.length) segments[k].push(...line);
            continue;
          }
          const polygon = bandPolygon(corners, k === 0 ? -Infinity : from, k === colors.length - 1 ? Infinity : to);
          if (polygon.length < 6) continue;
          // The outline in the colour of the piece hides the seams between pieces.
          items.push({ kind: 'path', role: 'area', points: polygon, closed: true, fill: colors[k], stroke: colors[k], lineWidth: 0.5, series: index });
          left--;
        }
      }
    }
  }
  if (wire) {
    for (let k = 0; k < colors.length; k++) if (segments[k].length > 0) items.push({ kind: 'path', role: 'line', points: segments[k], segments: true, stroke: colors[k], lineWidth: 1.5, series: index });
  }
  if (!words) return;
  if (labelsShown(catAxis)) {
    const slot = plot.width / Math.max(columns - 1, 1);
    let widest = 0;
    const stride = Math.max(Math.floor(columns / 200), 1);
    for (let i = 0; i < columns; i += stride) widest = Math.max(widest, measure(kit.clean(values.label(i)), catFont));
    const every = Math.max(Math.ceil((Math.min(widest, plot.width * 0.5) + 6) / Math.max(slot, 1)), 1);
    for (let i = 0; i < columns; i += every) {
      const text = kit.fit(kit.clean(values.label(i)), catFont, slot * every - 4);
      const x = columns > 1 ? plot.x + slot * i : plot.x + plot.width / 2;
      if (text !== '') items.push({ kind: 'text', role: 'category', text, x, y: plot.y + plot.height + catFont.size + 3, font: catFont, color: catColor, align: columns > 1 && i === 0 ? 'left' : columns > 1 && i === columns - 1 ? 'right' : 'center' });
    }
  }
  if (rowNames) {
    const step = plot.height / Math.max(rows.length - 1, 1);
    const every = Math.max(Math.ceil((catFont.size * 1.2) / Math.max(step, 1)), 1);
    for (let j = 0; j < rows.length; j += every) {
      const text = kit.fit(kit.clean(rows[j].name), catFont, plot.x - box.x - 6);
      if (text !== '') items.push({ kind: 'text', role: 'category', text, x: plot.x - 6, y: plot.y + plot.height - step * j + catFont.size * 0.35, font: catFont, color: catColor, align: 'right' });
    }
  }
}

// ---------------------------------------------------------------------- legend

/** The entries of the legend of a chart of a newer type or of a surface chart. */
export function extendedLegend(kit: ExKit, shown: ChartSeriesValues[]): ExLegendEntry[] {
  const { chart, values } = kit;
  const first = shown[0];
  if (!first) return [];
  const group = chart.groups[first.group];
  switch (first.kind) {
    case 'waterfall': {
      const [up, down, total] = waterfallColors(kit);
      return [
        { label: kit.clean(values.word ? values.word('increase') : 'Increase'), color: up, kind: 'box' },
        { label: kit.clean(values.word ? values.word('decrease') : 'Decrease'), color: down, kind: 'box' },
        { label: kit.clean(values.word ? values.word('total') : 'Total'), color: total, kind: 'box' },
      ];
    }
    case 'treemap':
    case 'sunburst': {
      // One entry for each branch at the top.
      const tree = treeOf(values, first);
      return tree.children.slice(0, 100).map((node) => ({ label: kit.clean(node.name), color: kit.auto(Math.max(node.branch, 0)), kind: 'box' as const }));
    }
    case 'surface': {
      const { scale, colors } = surfaceBands(kit, shown);
      const format = first.format;
      return colors.slice(0, 100).map((color, k) => ({ label: `${kit.tick(scale.ticks[k], format)}-${kit.tick(scale.ticks[k + 1], format)}`, color, kind: 'box' as const }));
    }
    default: {
      const out: ExLegendEntry[] = [];
      const list = first.kind === 'boxWhisker' ? shown : [first];
      for (const entry of list.slice(0, 100)) {
        const look = kit.look(group, entry.series);
        out.push({ label: kit.clean(entry.name), color: look.fill || look.color, kind: 'box' });
      }
      return out;
    }
  }
}

// ---------------------------------------------------------------------- entry

/** Lays out a chart of a newer type, or a surface chart, in `box`: the series `shown` are those of its first chart type. */
export function layoutExtended(kit: ExKit, layout: ChartLayout, box: Box, shown: ChartSeriesValues[]): void {
  const first = shown[0];
  if (!first || !(box.width > 0) || !(box.height > 0)) return;
  switch (first.kind) {
    case 'waterfall':
      return layoutWaterfall(kit, layout, box, first);
    case 'funnel':
      return layoutFunnel(kit, layout, box, first);
    case 'histogram':
      return layoutHistogram(kit, layout, box, first);
    case 'boxWhisker':
      return layoutBoxes(kit, layout, box, shown);
    case 'treemap':
      return layoutTreemap(kit, layout, box, first);
    case 'sunburst':
      return layoutSunburst(kit, layout, box, first);
    case 'surface':
      return layoutSurface(kit, layout, box, shown);
    default:
      return;
  }
}

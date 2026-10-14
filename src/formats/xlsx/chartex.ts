// Reads a part of the newer chart types (chartex: xl/charts/chartExN.xml or
// word/charts/chartExN.xml, <cx:chartSpace>, [MS-ODRAWXML] 2.24) into the
// model the charts of the older parts are read into: histogram (and pareto),
// waterfall, funnel, box and whisker, treemap and sunburst. Each series of
// the part names its data (cx:chartData) by number and states its layout
// (cx:layoutPr); series of one layout that follow one another make one chart
// type of the model.
//
// Nothing here is ever serialized, and a chart read here is only shown
// (Chart.shownOnly): the part stays in the package and is written back from
// there. A part that holds a region map (it needs map data the engine does
// not have), a layout that is not known, or nothing that can be drawn is no
// chart of the model: parseChartEx() answers undefined and its frame shows a
// placeholder, or the fallback the file gives beside it.

import {
  MAX_CHART_AXES,
  MAX_CHART_POINTS,
  MAX_CHART_SERIES,
  type Chart,
  type ChartAxis,
  type ChartBinning,
  type ChartCache,
  type ChartData,
  type ChartDataLabels,
  type ChartGroup,
  type ChartKind,
  type ChartLegend,
  type ChartPoint,
  type ChartPointLabel,
  type ChartSeries,
  type ChartSeriesEx,
  type ChartTextStyle,
  type ChartTitle,
} from '../../model/sheet/chart';
import { localName, type XmlElement } from '../../xml';
import {
  bodyStyle,
  ChartBudgetError,
  child,
  children,
  clamp,
  HEX6,
  int,
  lineOf,
  MAX_CHART_CACHE,
  MAX_CHART_DATA_POINTS,
  MAX_TEXT,
  note,
  number,
  OFFICE_ACCENTS,
  readLayout,
  richText,
  shapeOf,
  sortCache,
  textOf,
  type ChartReadBudget,
  type ReadState,
} from './chart';

/** Data sets (cx:data) read from one part. */
const MAX_EX_DATA = 255;
/** Levels of a dimension (cx:lvl) read: a hierarchy of a treemap or a sunburst is no deeper. */
export const MAX_EX_LEVELS = 8;
/** Subtotals of a waterfall series that are read. */
const MAX_SUBTOTALS = 4096;

/** A namespace of the newer chart parts: 2014/chartex, and the later ones that mark a version (2015/9/8/chartex, ...). */
export const CHARTEX_NAMESPACE = /^http:\/\/schemas\.microsoft\.com\/office\/drawing\/20\d\d(\/\d{1,2}){0,3}\/chartex$/;

/** What a cx:series layout is drawn as; a layout that is not here is not drawn. */
const LAYOUTS: { [layout: string]: ChartKind } = {
  clusteredColumn: 'histogram',
  waterfall: 'waterfall',
  funnel: 'funnel',
  boxWhisker: 'boxWhisker',
  treemap: 'treemap',
  sunburst: 'sunburst',
};

const TRUE = (value: string | undefined): boolean | undefined => (value === undefined ? undefined : value === '1' || value === 'true');

/** The namespace an element's prefix stands for, as far as the element itself declares it. */
function namespaceOf(el: XmlElement): string | undefined {
  const colon = el.name.indexOf(':');
  return el.attrs.get(colon < 0 ? 'xmlns' : `xmlns:${el.name.slice(0, colon)}`);
}

/** A number or "auto" (ST_DoubleOrAutomatic); undefined for anything else. */
function numberOrAuto(text: string | undefined): number | 'auto' | undefined {
  if (text === undefined) return undefined;
  if (text.trim() === 'auto') return 'auto';
  return number(text);
}

// --------------------------------------------------------------------- data

interface Dimension {
  type: string;
  data: ChartData;
}

/** The points of one level (cx:lvl) into `cache`; the text of a point is its content. */
function readLevel(state: ReadState, level: XmlElement | undefined, kind: 'num' | 'str', cache: ChartCache): number {
  if (!level) return 0;
  const index = cache.index;
  const values = cache.values;
  let ordered = true;
  let truncated = false;
  for (const node of level.children) {
    if (node.type !== 'element' || localName(node.name) !== 'pt') continue;
    const at = int(node.attrs.get('idx'));
    if (at === undefined || at < 0) continue;
    if (at >= MAX_CHART_POINTS) {
      truncated = true;
      continue;
    }
    if (state.points >= state.pointLimit) {
      state.cut = true;
      break;
    }
    state.points++;
    if (index.length > 0 && at <= index[index.length - 1]) ordered = false;
    const text = textOf(node);
    index.push(at);
    if (kind === 'num') {
      const value = number(text);
      values.push(value === undefined ? null : value);
    } else values.push(text.slice(0, MAX_TEXT));
  }
  if (!ordered) {
    const holder: ChartData = { kind, cache: { index: index.slice(), values: values.slice() }, count: 0 };
    sortCache(holder);
    index.length = 0;
    values.length = 0;
    for (let i = 0; i < holder.cache.index.length; i++) {
      index.push(holder.cache.index[i]);
      values.push(holder.cache.values[i]);
    }
  }
  if (truncated) note(state, `Only the first ${MAX_CHART_POINTS} points of a series are read`);
  const stated = int(level.attrs.get('ptCount'));
  const last = index.length > 0 ? index[index.length - 1] + 1 : 0;
  return Math.max(last, stated === undefined || stated < 0 ? 0 : Math.min(stated, MAX_CHART_POINTS));
}

/** A dimension (cx:numDim, cx:strDim): its formula and its levels, the first of them next to the axis. */
function readDimension(state: ReadState, el: XmlElement): Dimension {
  const kind = localName(el.name) === 'numDim' ? 'num' : 'str';
  const data: ChartData = { kind, cache: { index: [], values: [] }, count: 0 };
  const formula = child(el, 'f');
  const text = textOf(formula).trim();
  if (text !== '' && text.length <= 8192) data.ref = text.charAt(0) === '=' ? text.slice(1) : text;
  if (formula && formula.attrs.get('dir') === 'row') data.byRow = true;
  const levels = children(el, 'lvl');
  if (levels.length > 0) {
    if (kind === 'num') {
      const code = levels[0].attrs.get('formatCode');
      if (code !== undefined && code !== '' && code.length <= 255) data.formatCode = code;
    }
    data.count = readLevel(state, levels[0], kind, data.cache);
    if (kind === 'str' && levels.length > 1) {
      const further: ChartCache[] = [];
      for (let i = 1; i < levels.length && i < MAX_EX_LEVELS; i++) {
        const cache: ChartCache = { index: [], values: [] };
        data.count = Math.max(data.count, readLevel(state, levels[i], 'str', cache));
        further.push(cache);
      }
      data.multiLevel = true;
      data.levels = further;
      if (levels.length > MAX_EX_LEVELS) note(state, `Of more than ${MAX_EX_LEVELS} levels of categories the ${MAX_EX_LEVELS} next to the leaves are read`);
    }
    if (kind === 'num' && levels.length > 1) note(state, 'Of a dimension of numbers with several levels the first is read');
  }
  return { type: el.attrs.get('type') || (kind === 'num' ? 'val' : 'cat'), data };
}

function readChartData(state: ReadState, el: XmlElement | undefined): Map<number, Dimension[]> {
  const out = new Map<number, Dimension[]>();
  for (const data of children(el, 'data')) {
    if (out.size >= MAX_EX_DATA) break;
    const id = int(data.attrs.get('id'));
    if (id === undefined || id < 0 || out.has(id)) continue;
    const dimensions: Dimension[] = [];
    for (const node of data.children) {
      if (node.type !== 'element') continue;
      const local = localName(node.name);
      if ((local === 'numDim' || local === 'strDim') && dimensions.length < 8) dimensions.push(readDimension(state, node));
    }
    out.set(id, dimensions);
  }
  return out;
}

// --------------------------------------------------------------------- text

/** A text (cx:tx): its rich text, or the text of the data (cx:txData) with the formula it is read from. */
function readText(state: ReadState, tx: XmlElement | undefined, style: ChartTextStyle): { text?: string; ref?: string } {
  const out: { text?: string; ref?: string } = {};
  if (!tx) return out;
  const rich = child(tx, 'rich');
  if (rich) {
    const text = richText(state, rich, style);
    if (text !== '') out.text = text;
    return out;
  }
  const data = child(tx, 'txData');
  const formula = textOf(child(data, 'f')).trim();
  if (formula !== '' && formula.length <= 8192) out.ref = formula.charAt(0) === '=' ? formula.slice(1) : formula;
  const value = child(data, 'v');
  if (value) out.text = textOf(value).slice(0, MAX_TEXT);
  return out;
}

function readTitle(state: ReadState, el: XmlElement | undefined): ChartTitle | undefined {
  if (!el) return undefined;
  const style: ChartTextStyle = { ...(bodyStyle(state, child(el, 'txPr')) || {}) };
  const { text, ref } = readText(state, child(el, 'tx'), style);
  const title: ChartTitle = {};
  if (text !== undefined && text !== '') title.text = text;
  if (ref !== undefined) title.ref = ref;
  if (Object.keys(style).length > 0) title.style = style;
  const position = el.attrs.get('pos');
  if (position !== undefined && position !== 't') note(state, 'A title that does not stand at the top is drawn at the top');
  const layout = readLayout(el);
  if (layout) title.layout = layout;
  return title;
}

const LABEL_POSITION = /^(outEnd|inEnd|ctr|inBase|bestFit|t|b|l|r)$/;

/** What a cx:dataLabels or cx:dataLabel says. A set of labels without a cx:visibility shows the values. */
function readLabels(state: ReadState, el: XmlElement, whole: boolean): ChartDataLabels {
  const labels: ChartDataLabels = {};
  const seen = child(el, 'visibility');
  if (seen) {
    const value = TRUE(seen.attrs.get('value'));
    const category = TRUE(seen.attrs.get('categoryName'));
    const series = TRUE(seen.attrs.get('seriesName'));
    if (value !== undefined) labels.showValue = value;
    if (category !== undefined) labels.showCategory = category;
    if (series !== undefined) labels.showSeries = series;
  } else if (whole) labels.showValue = true;
  const position = el.attrs.get('pos');
  if (position !== undefined && LABEL_POSITION.test(position)) labels.position = position as ChartDataLabels['position'];
  const format = child(el, 'numFmt');
  if (format) {
    const code = format.attrs.get('formatCode');
    if (code !== undefined && code !== '' && code.length <= 255) labels.formatCode = code;
    const linked = TRUE(format.attrs.get('sourceLinked'));
    if (linked !== undefined) labels.sourceLinked = linked;
  }
  const separator = textOf(child(el, 'separator')).slice(0, 16);
  if (separator !== '') labels.separator = separator;
  const style = bodyStyle(state, child(el, 'txPr'));
  if (style) labels.style = style;
  return labels;
}

// ------------------------------------------------------------------- series

/** What cx:layoutPr states, for the kind of the series. */
function readLayoutProps(state: ReadState, el: XmlElement | undefined, kind: ChartKind): ChartSeriesEx | undefined {
  if (!el) return undefined;
  const ex: ChartSeriesEx = {};
  const shown = child(el, 'visibility');
  if (kind === 'waterfall') {
    const totals: number[] = [];
    for (const node of children(child(el, 'subtotals'), 'idx')) {
      const at = int(node.attrs.get('val'));
      if (at !== undefined && at >= 0 && at < MAX_CHART_POINTS && totals.length < MAX_SUBTOTALS && totals.indexOf(at) < 0) totals.push(at);
    }
    if (totals.length > 0) ex.subtotals = totals.sort((a, b) => a - b);
    const lines = TRUE(shown ? shown.attrs.get('connectorLines') : undefined);
    if (lines !== undefined) ex.connectorLines = lines;
  }
  if (kind === 'histogram') {
    const binning = child(el, 'binning');
    if (binning) {
      const bins: ChartBinning = { closed: binning.attrs.get('intervalClosed') === 'l' ? 'l' : 'r' };
      const under = numberOrAuto(binning.attrs.get('underflow'));
      if (under !== undefined) bins.underflow = under;
      const over = numberOrAuto(binning.attrs.get('overflow'));
      if (over !== undefined) bins.overflow = over;
      const size = number(child(binning, 'binSize')?.attrs.get('val'));
      if (size !== undefined && size > 0) bins.size = size;
      const count = int(child(binning, 'binCount')?.attrs.get('val'));
      if (count !== undefined && count > 0) bins.count = Math.min(count, 10000);
      ex.binning = bins;
    } else if (child(el, 'aggregation')) ex.aggregate = true;
  }
  if (kind === 'boxWhisker') {
    const method = child(el, 'statistics')?.attrs.get('quartileMethod');
    if (method === 'inclusive' || method === 'exclusive') ex.quartiles = method;
    if (shown) {
      for (const key of ['meanLine', 'meanMarker', 'nonoutliers', 'outliers'] as const) {
        const value = TRUE(shown.attrs.get(key));
        if (value !== undefined) ex[key] = value;
      }
    }
  }
  if (kind === 'treemap') {
    const layout = child(el, 'parentLabelLayout')?.attrs.get('val');
    if (layout === 'none' || layout === 'banner' || layout === 'overlapping') ex.parentLabels = layout;
  }
  if (child(el, 'geography') || child(el, 'regionLabelLayout')) note(state, 'Map data of a series is not drawn');
  return Object.keys(ex).length > 0 ? ex : undefined;
}

interface ReadSeries {
  kind: ChartKind;
  series: ChartSeries;
  axisIds: number[];
}

function readSeries(state: ReadState, el: XmlElement, kind: ChartKind, position: number, data: Map<number, Dimension[]>): ReadSeries {
  const format = int(el.attrs.get('formatIdx'));
  const series: ChartSeries = { index: format !== undefined && format >= 0 && format < 10000 ? format : position, order: position };
  const style: ChartTextStyle = {};
  const name = readText(state, child(el, 'tx'), style);
  if (name.text !== undefined || name.ref !== undefined) series.name = name;
  const shape = shapeOf(state, child(el, 'spPr'));
  if (shape) series.shape = shape;
  const points: ChartPoint[] = [];
  for (const node of children(el, 'dataPt')) {
    if (points.length >= MAX_CHART_DATA_POINTS) break;
    const at = int(node.attrs.get('idx'));
    if (at === undefined || at < 0 || at >= MAX_CHART_POINTS) continue;
    const own = shapeOf(state, child(node, 'spPr'));
    if (own) points.push({ index: at, shape: own });
  }
  if (points.length > 0) series.points = points;
  const labels = child(el, 'dataLabels');
  if (labels) {
    series.dataLabels = readLabels(state, labels, true);
    const own: ChartPointLabel[] = [];
    for (const node of labels.children) {
      if (node.type !== 'element' || own.length >= MAX_CHART_DATA_POINTS) continue;
      const local = localName(node.name);
      if (local !== 'dataLabel' && local !== 'dataLabelHidden') continue;
      const at = int(node.attrs.get('idx'));
      if (at === undefined || at < 0 || at >= MAX_CHART_POINTS || own.some((label) => label.index === at)) continue;
      own.push(local === 'dataLabelHidden' ? { index: at, deleted: true } : { index: at, ...readLabels(state, node, false) });
    }
    if (own.length > 0) series.pointLabels = own.sort((a, b) => a.index - b.index);
    if (kind === 'boxWhisker') note(state, 'The data labels of a box and whisker chart are not drawn');
    if ((kind === 'treemap' || kind === 'sunburst') && (series.pointLabels || []).some((label) => !label.deleted)) note(state, 'Data labels of single points of a treemap or a sunburst are not drawn');
  }
  const id = int(child(el, 'dataId')?.attrs.get('val'));
  const dimensions = id === undefined ? undefined : data.get(id);
  if (dimensions) {
    const numbers = dimensions.filter((dimension) => dimension.data.kind === 'num');
    const values = numbers.find((dimension) => dimension.type === 'val') || numbers.find((dimension) => dimension.type === 'size') || numbers[0];
    if (values) series.values = values.data;
    const categories = dimensions.find((dimension) => dimension.data.kind === 'str' && dimension.type === 'cat') || dimensions.find((dimension) => dimension !== values && dimension.type === 'cat');
    if (categories) series.categories = categories.data;
    if (dimensions.some((dimension) => dimension.type === 'colorVal' || dimension.type === 'colorStr' || dimension.type === 'x' || dimension.type === 'y')) note(state, 'Some data of a series is not drawn');
  }
  const ex = readLayoutProps(state, child(el, 'layoutPr'), kind);
  if (ex) series.ex = ex;
  if (child(el, 'valueColors') || child(el, 'valueColorPositions')) note(state, 'The colour scale of a series is not drawn');
  const axisIds: number[] = [];
  for (const node of children(el, 'axisId')) {
    const axis = int(node.attrs.get('val'));
    if (axis !== undefined && axisIds.length < 4) axisIds.push(axis);
  }
  return { kind, series, axisIds };
}

// --------------------------------------------------------------------- axes

function readAxis(state: ReadState, el: XmlElement): ChartAxis | undefined {
  const id = int(el.attrs.get('id'));
  if (id === undefined) return undefined;
  const values = child(el, 'valScaling');
  const axis: ChartAxis = { id, kind: values ? 'val' : 'cat', position: values ? 'l' : 'b', tickLabels: child(el, 'tickLabels') ? 'nextTo' : 'none' };
  if (TRUE(el.attrs.get('hidden'))) axis.deleted = true;
  if (values) {
    const min = numberOrAuto(values.attrs.get('min'));
    if (typeof min === 'number') axis.min = min;
    const max = numberOrAuto(values.attrs.get('max'));
    if (typeof max === 'number') axis.max = max;
    const unit = numberOrAuto(values.attrs.get('majorUnit'));
    if (typeof unit === 'number' && unit > 0) axis.majorUnit = unit;
  }
  const format = child(el, 'numFmt');
  if (format) {
    const code = format.attrs.get('formatCode');
    if (code !== undefined && code !== '' && code.length <= 255) axis.formatCode = code;
    const linked = TRUE(format.attrs.get('sourceLinked'));
    if (linked !== undefined) axis.sourceLinked = linked;
  }
  const major = child(el, 'majorGridlines');
  if (major) axis.majorGridlines = lineOf(state, child(major, 'spPr')) || {};
  const minor = child(el, 'minorGridlines');
  if (minor) axis.minorGridlines = lineOf(state, child(minor, 'spPr')) || {};
  const title = readTitle(state, child(el, 'title'));
  if (title) axis.title = title;
  const line = lineOf(state, child(el, 'spPr'));
  if (line) axis.line = line;
  const style = bodyStyle(state, child(el, 'txPr'));
  if (style) axis.labelStyle = style;
  if (child(el, 'units')) note(state, 'Display units of an axis are not applied');
  return axis;
}

/** The gap between the columns a category axis states (cx:catScaling gapWidth), as a percentage of a column. */
function gapOf(el: XmlElement): number | undefined {
  const gap = number(child(el, 'catScaling')?.attrs.get('gapWidth'));
  // UNSURE: the gap is a ratio of the width of a column (0.5 is half of it), as Excel writes it.
  return gap === undefined || gap < 0 ? undefined : clamp(Math.round(gap * 100), 0, 500);
}

// -------------------------------------------------------------------- chart

/** Whether a root element is that of a part of the newer chart types. */
export function isChartExRoot(root: XmlElement | undefined): boolean {
  if (!root || localName(root.name) !== 'chartSpace') return false;
  const space = namespaceOf(root);
  return space !== undefined && CHARTEX_NAMESPACE.test(space);
}

/**
 * The chart of a part of the newer chart types, from its root element
 * (<cx:chartSpace>). Undefined for a part that is none, that holds a layout
 * that is not drawn (a region map) or nothing that is drawn. `themeColors`
 * and `budget` as parseChart() takes them; a chart that holds more cached
 * values than is left of the budget throws a ChartBudgetError (the only thing
 * this throws) and takes nothing.
 */
export function parseChartEx(root: XmlElement, themeColors: string[], budget?: ChartReadBudget): Chart | undefined {
  const limit = budget ? Math.max(Math.min(budget.points, MAX_CHART_CACHE), 0) : MAX_CHART_CACHE;
  const state: ReadState = { themeColors: Array.isArray(themeColors) ? themeColors : [], notes: [], points: 0, pointLimit: limit, series: 0 };
  let chart: Chart | undefined;
  try {
    chart = readChartEx(root, state);
  } catch (error) {
    chart = undefined;
  }
  if (budget) {
    if (chart && state.cut && limit < MAX_CHART_CACHE) throw new ChartBudgetError('the charts of the workbook hold too many cached values');
    if (chart) budget.points -= state.points;
  }
  if (chart && state.cut) note(state, `Only the first ${MAX_CHART_CACHE} cached values of a chart are read`);
  return chart;
}

function readChartEx(root: XmlElement, state: ReadState): Chart | undefined {
  if (!isChartExRoot(root)) return undefined;
  const chartEl = child(root, 'chart');
  const plot = child(chartEl, 'plotArea');
  const region = child(plot, 'plotAreaRegion');
  if (!chartEl || !plot || !region) return undefined;
  const data = readChartData(state, child(root, 'chartData'));

  const read: (ReadSeries | undefined)[] = [];
  const owners: { owner: number; el: XmlElement }[] = [];
  const all = children(region, 'series');
  for (let i = 0; i < all.length; i++) {
    const el = all[i];
    const layout = el.attrs.get('layoutId') || '';
    if (layout === 'paretoLine') {
      const owner = int(el.attrs.get('ownerIdx'));
      if (owner !== undefined && owner >= 0) owners.push({ owner, el });
      read.push(undefined);
      continue;
    }
    // A layout that is not drawn (a region map, one of a later version) hides the whole chart.
    if (!Object.prototype.hasOwnProperty.call(LAYOUTS, layout)) return undefined;
    if (state.series >= MAX_CHART_SERIES) {
      note(state, `Only the first ${MAX_CHART_SERIES} series are read`);
      break;
    }
    if (TRUE(el.attrs.get('hidden'))) {
      read.push(undefined);
      continue;
    }
    read.push(readSeries(state, el, LAYOUTS[layout], state.series, data));
    state.series++;
  }
  // The line of a pareto chart belongs to the series it names by its place among the series of the part.
  for (const { owner, el } of owners) {
    const target = read[owner];
    if (!target || target.kind !== 'histogram') continue;
    const line = lineOf(state, child(el, 'spPr'));
    const axis = int(child(el, 'axisId')?.attrs.get('val'));
    const pareto: NonNullable<ChartSeriesEx['pareto']> = {};
    if (line) pareto.line = line;
    if (axis !== undefined) pareto.axisId = axis;
    target.series.ex = { ...(target.series.ex || {}), pareto };
  }

  const axes: ChartAxis[] = [];
  const gaps = new Map<number, number>();
  for (const el of children(plot, 'axis')) {
    if (axes.length >= MAX_CHART_AXES) break;
    const axis = readAxis(state, el);
    if (!axis || axes.some((other) => other.id === axis.id)) continue;
    axes.push(axis);
    const gap = gapOf(el);
    if (gap !== undefined) gaps.set(axis.id, gap);
  }

  // Series of one layout that follow one another are one chart type.
  const groups: ChartGroup[] = [];
  for (const item of read) {
    if (!item) continue;
    const last = groups[groups.length - 1];
    if (last && last.kind === item.kind) {
      last.series.push(item.series);
      continue;
    }
    const group: ChartGroup = { kind: item.kind, grouping: 'standard', axisIds: item.axisIds, series: [item.series] };
    for (const id of item.axisIds) {
      const gap = gaps.get(id);
      if (gap !== undefined) group.gapWidth = gap;
    }
    if (item.kind === 'treemap' || item.kind === 'sunburst' || item.kind === 'funnel') group.varyColors = item.kind !== 'funnel';
    groups.push(group);
  }
  if (groups.length === 0 || !groups.some((group) => group.series.some((series) => series.values !== undefined))) return undefined;
  if (groups.length > 1) note(state, 'Of a chart of several layouts the first is drawn');
  for (const group of groups) {
    if (group.series.length > 1 && group.kind !== 'boxWhisker') note(state, 'Of several series of one layout the first is drawn');
  }

  const palette: string[] = [];
  for (let i = 0; i < 6; i++) {
    const theme = state.themeColors[4 + i];
    palette.push('#' + (typeof theme === 'string' && HEX6.test(theme) ? theme.toUpperCase() : OFFICE_ACCENTS[i]));
  }
  const chart: Chart = { extended: true, shownOnly: true, groups, axes, plotVisOnly: true, blanks: 'gap', palette, notes: state.notes };
  const title = readTitle(state, child(chartEl, 'title'));
  if (title) chart.title = title;
  const legendEl = child(chartEl, 'legend');
  if (legendEl) {
    const position = legendEl.attrs.get('pos');
    const legend: ChartLegend = { position: position === 'l' || position === 't' || position === 'b' ? position : 'r' };
    const style = bodyStyle(state, child(legendEl, 'txPr'));
    if (style) legend.style = style;
    chart.legend = legend;
  }
  const area = shapeOf(state, child(root, 'spPr'));
  if (area) chart.area = area;
  const surface = shapeOf(state, child(child(region, 'plotSurface'), 'spPr'));
  if (surface) chart.plotArea = surface;
  const text = bodyStyle(state, child(root, 'txPr'));
  if (text) chart.textStyle = text;
  return chart;
}

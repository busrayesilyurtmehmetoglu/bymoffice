// Reads a chart part (xl/charts/chartN.xml, DrawingML <c:chartSpace>) into the
// read-only model the grid draws a chart from: the chart types of the plot
// area with their series, the axes, the title and the legend.
//
// Nothing here is ever serialized: the part stays in the package and is
// written back from there. A chart that holds a type that is not drawn (radar,
// bubble, stock, surface, of-pie) or that cannot be read is not a chart of the
// model: parseChart() answers undefined and the drawing shows a placeholder.

import {
  MAX_CHART_AXES,
  MAX_CHART_POINTS,
  MAX_CHART_SERIES,
  type Chart,
  type ChartAxis,
  type ChartData,
  type ChartGroup,
  type ChartKind,
  type ChartLegend,
  type ChartLine,
  type ChartMarker,
  type ChartPoint,
  type ChartSeries,
  type ChartShape,
  type ChartTextStyle,
  type ChartTitle,
} from '../../model/sheet/chart';
import { localName, type XmlElement } from '../../xml';
import { drawingColor, EMU_PER_PIXEL } from './drawings';

/** Cached values (<c:pt>) read from one chart; what lies beyond is left out. */
export const MAX_CHART_CACHE = 1000000;
/** Points drawn their own way (c:dPt) read from one series. */
export const MAX_CHART_DATA_POINTS = 4096;
const MAX_TEXT = 4096;
const MAX_NOTES = 32;

/** The accent colours of the Office theme, for a workbook without a theme part. */
const OFFICE_ACCENTS = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'];
const HEX6 = /^[0-9A-Fa-f]{6}$/;
const NUMBER = /^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/;

const GROUPS: { [local: string]: { kind: ChartKind; threeD?: true } } = {
  barChart: { kind: 'bar' },
  bar3DChart: { kind: 'bar', threeD: true },
  lineChart: { kind: 'line' },
  line3DChart: { kind: 'line', threeD: true },
  pieChart: { kind: 'pie' },
  pie3DChart: { kind: 'pie', threeD: true },
  doughnutChart: { kind: 'doughnut' },
  areaChart: { kind: 'area' },
  area3DChart: { kind: 'area', threeD: true },
  scatterChart: { kind: 'scatter' },
};

const AXES: { [local: string]: 'cat' | 'val' } = { catAx: 'cat', dateAx: 'cat', serAx: 'cat', valAx: 'val' };

interface ReadState {
  themeColors: string[];
  notes: string[];
  /** Cached values (<c:pt>) read so far. */
  points: number;
  /** The most of them this chart may read. */
  pointLimit: number;
  /** A cache was cut short at that limit. */
  cut?: boolean;
  series: number;
}

function note(state: ReadState, text: string): void {
  if (state.notes.length < MAX_NOTES && state.notes.indexOf(text) < 0) state.notes.push(text);
}

function child(parent: XmlElement | undefined, local: string): XmlElement | undefined {
  if (!parent) return undefined;
  for (const node of parent.children) if (node.type === 'element' && localName(node.name) === local) return node;
  return undefined;
}

function children(parent: XmlElement | undefined, local: string): XmlElement[] {
  const found: XmlElement[] = [];
  if (!parent) return found;
  for (const node of parent.children) if (node.type === 'element' && localName(node.name) === local) found.push(node);
  return found;
}

function textOf(el: XmlElement | undefined): string {
  if (!el) return '';
  let text = '';
  for (const node of el.children) if (node.type === 'text') text += node.value;
  return text;
}

/** The `val` attribute of a child element; undefined when the child or the attribute is missing. */
function val(parent: XmlElement | undefined, local: string): string | undefined {
  const el = child(parent, local);
  return el ? el.attrs.get('val') : undefined;
}

/** A flag element: missing -> `missing`; present without a value -> true (the default of the schema). */
function flag(parent: XmlElement | undefined, local: string, missing: boolean): boolean {
  const el = child(parent, local);
  if (!el) return missing;
  const value = el.attrs.get('val');
  return value === undefined ? true : value === '1' || value === 'true';
}

function number(text: string | undefined): number | undefined {
  if (text === undefined) return undefined;
  const trimmed = text.trim();
  if (!NUMBER.test(trimmed)) return undefined;
  const value = Number(trimmed);
  return Number.isFinite(value) ? value : undefined;
}

function int(text: string | undefined): number | undefined {
  const value = number(text);
  return value !== undefined && Math.floor(value) === value && Math.abs(value) <= 2147483647 ? value : undefined;
}

const clamp = (value: number, low: number, high: number): number => (value < low ? low : value > high ? high : value);

// ------------------------------------------------------------------ colours

/**
 * The colour of a fill among the children of `holder`: "#RRGGBB", null for
 * one that shows nothing, undefined when none is stated. Transparency is
 * drawn against white; a gradient is its first colour, a pattern its
 * foreground colour.
 */
function fillOf(state: ReadState, holder: XmlElement | undefined): string | null | undefined {
  if (!holder) return undefined;
  for (const node of holder.children) {
    if (node.type !== 'element') continue;
    const local = localName(node.name);
    if (local === 'noFill') return null;
    if (local === 'solidFill') return colorOf(state, node);
    if (local === 'gradFill') {
      note(state, 'A gradient fill is drawn as its first colour');
      return colorOf(state, child(child(node, 'gsLst'), 'gs'));
    }
    if (local === 'pattFill') {
      note(state, 'A pattern fill is drawn as its foreground colour');
      return colorOf(state, child(node, 'fgClr'));
    }
    if (local === 'blipFill') {
      note(state, 'A picture fill is not drawn');
      return undefined;
    }
  }
  return undefined;
}

/** The first colour element inside `holder`; null when it is fully transparent, undefined when there is none that can be read. */
function colorOf(state: ReadState, holder: XmlElement | undefined): string | null | undefined {
  if (!holder) return undefined;
  for (const node of holder.children) {
    if (node.type !== 'element') continue;
    const alpha = number(val(node, 'alpha'));
    if (alpha !== undefined && alpha <= 0) return null;
    const color = drawingColor(state.themeColors, node);
    if (color === undefined) return undefined;
    if (alpha === undefined || alpha >= 100000) return color;
    // No blending on the canvas: the colour as it shows over white.
    note(state, 'Transparent colours are drawn as they show on white');
    const share = alpha / 100000;
    let out = '#';
    for (let i = 0; i < 3; i++) {
      const channel = Math.round(parseInt(color.slice(1 + i * 2, 3 + i * 2), 16) * share + 255 * (1 - share));
      out += (channel < 16 ? '0' : '') + channel.toString(16).toUpperCase();
    }
    return out;
  }
  return undefined;
}

function lineOf(state: ReadState, props: XmlElement | undefined): ChartLine | undefined {
  const outline = child(props, 'ln');
  if (!outline) return undefined;
  const line: ChartLine = {};
  const color = fillOf(state, outline);
  if (color !== undefined) line.color = color;
  const width = int(outline.attrs.get('w'));
  if (width !== undefined && width >= 0) line.width = clamp(width / EMU_PER_PIXEL, 0, 32);
  const dash = val(outline, 'prstDash');
  if ((dash !== undefined && dash !== 'solid') || child(outline, 'custDash')) line.dash = true;
  return line.color !== undefined || line.width !== undefined || line.dash ? line : undefined;
}

function shapeOf(state: ReadState, props: XmlElement | undefined): ChartShape | undefined {
  if (!props) return undefined;
  const shape: ChartShape = {};
  const fill = fillOf(state, props);
  if (fill !== undefined) shape.fill = fill;
  const line = lineOf(state, props);
  if (line) shape.line = line;
  return shape.fill !== undefined || shape.line ? shape : undefined;
}

// --------------------------------------------------------------------- text

function runStyle(state: ReadState, props: XmlElement | undefined, into: ChartTextStyle): void {
  if (!props) return;
  const size = int(props.attrs.get('sz'));
  if (size !== undefined && size >= 100 && size <= 400000) into.size = size / 100;
  const bold = props.attrs.get('b');
  if (bold !== undefined) into.bold = bold === '1' || bold === 'true';
  const italic = props.attrs.get('i');
  if (italic !== undefined) into.italic = italic === '1' || italic === 'true';
  const color = fillOf(state, props);
  if (color) into.color = color;
  const latin = child(props, 'latin');
  const face = latin ? latin.attrs.get('typeface') : undefined;
  // "+mn-lt" and the like name a font of the theme, not a font.
  if (face !== undefined && face !== '' && face.charAt(0) !== '+') into.font = face.slice(0, 100);
}

/** What a <c:txPr> (or the body of a rich text) says about the text it does not set itself. */
function bodyStyle(state: ReadState, body: XmlElement | undefined): ChartTextStyle | undefined {
  if (!body) return undefined;
  const style: ChartTextStyle = {};
  runStyle(state, child(child(child(body, 'lstStyle'), 'lvl1pPr'), 'defRPr'), style);
  runStyle(state, child(child(child(body, 'p'), 'pPr'), 'defRPr'), style);
  return Object.keys(style).length > 0 ? style : undefined;
}

/** A title: the plain text of its rich text with the look of the first run, or the text a reference cached. */
function readTitle(state: ReadState, title: XmlElement | undefined): ChartTitle | undefined {
  if (!title) return undefined;
  const out: ChartTitle = {};
  const style: ChartTextStyle = { ...(bodyStyle(state, child(title, 'txPr')) || {}) };
  const tx = child(title, 'tx');
  const rich = child(tx, 'rich');
  if (rich) {
    Object.assign(style, bodyStyle(state, rich) || {});
    let text = '';
    let first = true;
    let lines = 0;
    for (const paragraph of children(rich, 'p')) {
      if (text.length >= MAX_TEXT) break;
      if (lines++ > 0) text += '\n';
      for (const node of paragraph.children) {
        if (node.type !== 'element') continue;
        const kind = localName(node.name);
        if (kind === 'br') text += '\n';
        if (kind !== 'r' && kind !== 'fld') continue;
        const piece = textOf(child(node, 't'));
        if (piece === '') continue;
        if (first) {
          runStyle(state, child(node, 'rPr'), style);
          first = false;
        }
        text += piece;
      }
    }
    text = text.slice(0, MAX_TEXT).replace(/^\n+|\n+$/g, '');
    if (text !== '') out.text = text;
  } else {
    const ref = child(tx, 'strRef');
    if (ref) {
      const formula = textOf(child(ref, 'f')).trim();
      if (formula !== '') out.ref = formula.slice(0, 8192);
      const cached = textOf(child(child(child(ref, 'strCache'), 'pt'), 'v')).slice(0, MAX_TEXT);
      if (cached !== '') out.text = cached;
    }
  }
  if (Object.keys(style).length > 0) out.style = style;
  return out;
}

// --------------------------------------------------------------------- data

function readCache(state: ReadState, holder: XmlElement | undefined, kind: 'num' | 'str', data: ChartData): void {
  if (!holder) return;
  if (kind === 'num') {
    const code = textOf(child(holder, 'formatCode'));
    if (code !== '' && code.length <= 255) data.formatCode = code;
  }
  const stated = int(val(holder, 'ptCount'));
  // One entry per point the file has, whatever its index: what the cache costs is what the file holds.
  const index = data.cache.index;
  const values = data.cache.values;
  let truncated = false;
  let ordered = true;
  for (const node of holder.children) {
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
    const text = textOf(child(node, 'v'));
    index.push(at);
    if (kind === 'num') {
      const value = number(text);
      values.push(value === undefined ? null : value);
    } else values.push(text.slice(0, MAX_TEXT));
  }
  if (!ordered) sortCache(data);
  if (stated !== undefined && stated > MAX_CHART_POINTS) truncated = true;
  if (truncated) note(state, `Only the first ${MAX_CHART_POINTS} points of a series are read`);
  const sorted = data.cache.index;
  const last = sorted.length > 0 ? sorted[sorted.length - 1] + 1 : 0;
  data.count = Math.max(last, stated === undefined || stated < 0 ? 0 : Math.min(stated, MAX_CHART_POINTS));
}

/** Puts the points of a cache in the order of their indexes; of two points with one index the later one of the file stays. */
function sortCache(data: ChartData): void {
  const { index, values } = data.cache;
  const order: number[] = [];
  for (let i = 0; i < index.length; i++) order.push(i);
  order.sort((a, b) => index[a] - index[b] || a - b);
  const sortedIndex: number[] = [];
  const sortedValues: (number | string | null)[] = [];
  for (const place of order) {
    if (sortedIndex.length > 0 && sortedIndex[sortedIndex.length - 1] === index[place]) sortedValues[sortedValues.length - 1] = values[place];
    else {
      sortedIndex.push(index[place]);
      sortedValues.push(values[place]);
    }
  }
  data.cache = { index: sortedIndex, values: sortedValues };
}

/** A data source: <c:cat>, <c:val>, <c:xVal>, <c:yVal>. Undefined when it holds none of the forms that are read. */
function readData(state: ReadState, holder: XmlElement | undefined): ChartData | undefined {
  if (!holder) return undefined;
  for (const node of holder.children) {
    if (node.type !== 'element') continue;
    const local = localName(node.name);
    if (local === 'numRef' || local === 'strRef' || local === 'multiLvlStrRef') {
      const kind = local === 'numRef' ? 'num' : 'str';
      const data: ChartData = { kind, cache: { index: [], values: [] }, count: 0 };
      const formula = textOf(child(node, 'f')).trim();
      if (formula !== '' && formula.length <= 8192) data.ref = formula.charAt(0) === '=' ? formula.slice(1) : formula;
      if (local === 'multiLvlStrRef') {
        // The first level is the one next to the axis.
        data.multiLevel = true;
        note(state, 'Of several levels of categories the one next to the axis is shown');
        readCache(state, child(child(node, 'multiLvlStrCache'), 'lvl'), 'str', data);
        const stated = int(val(child(node, 'multiLvlStrCache'), 'ptCount'));
        if (stated !== undefined && stated > data.count) data.count = Math.min(stated, MAX_CHART_POINTS);
      } else readCache(state, child(node, kind === 'num' ? 'numCache' : 'strCache'), kind, data);
      return data;
    }
    if (local === 'numLit' || local === 'strLit') {
      const kind = local === 'numLit' ? 'num' : 'str';
      const data: ChartData = { kind, cache: { index: [], values: [] }, count: 0 };
      readCache(state, node, kind, data);
      return data;
    }
  }
  return undefined;
}

// ------------------------------------------------------------------- series

const SYMBOLS: { [name: string]: string } = {
  none: 'none',
  auto: 'auto',
  circle: 'circle',
  square: 'square',
  diamond: 'diamond',
  triangle: 'triangle',
  x: 'x',
  plus: 'plus',
  star: 'star',
  dash: 'dash',
  dot: 'dot',
  picture: 'square',
};

function readMarker(state: ReadState, el: XmlElement | undefined): ChartMarker | undefined {
  if (!el) return undefined;
  const stated = val(el, 'symbol');
  const marker: ChartMarker = { symbol: stated !== undefined && Object.prototype.hasOwnProperty.call(SYMBOLS, stated) ? SYMBOLS[stated] : 'auto' };
  const size = int(val(el, 'size'));
  if (size !== undefined) marker.size = clamp(size, 2, 72);
  const props = child(el, 'spPr');
  const fill = fillOf(state, props);
  if (fill !== undefined) marker.fill = fill;
  const line = lineOf(state, props);
  if (line) marker.line = line;
  return marker;
}

function readSeries(state: ReadState, el: XmlElement, scatter: boolean, position: number): ChartSeries {
  const index = int(val(el, 'idx'));
  const order = int(val(el, 'order'));
  const series: ChartSeries = { index: index !== undefined && index >= 0 ? index : position, order: order !== undefined && order >= 0 ? order : position };
  const tx = child(el, 'tx');
  if (tx) {
    const ref = child(tx, 'strRef');
    const name: { ref?: string; text?: string } = {};
    if (ref) {
      const formula = textOf(child(ref, 'f')).trim();
      if (formula !== '' && formula.length <= 8192) name.ref = formula.charAt(0) === '=' ? formula.slice(1) : formula;
      const cached = child(child(child(ref, 'strCache'), 'pt'), 'v');
      if (cached) name.text = textOf(cached).slice(0, MAX_TEXT);
    } else {
      const literal = child(tx, 'v');
      if (literal) name.text = textOf(literal).slice(0, MAX_TEXT);
    }
    if (name.ref !== undefined || name.text !== undefined) series.name = name;
  }
  const categories = readData(state, child(el, scatter ? 'xVal' : 'cat'));
  if (categories) series.categories = categories;
  const values = readData(state, child(el, scatter ? 'yVal' : 'val'));
  if (values) series.values = values;
  const shape = shapeOf(state, child(el, 'spPr'));
  if (shape) series.shape = shape;
  const marker = readMarker(state, child(el, 'marker'));
  if (marker) series.marker = marker;
  const points: ChartPoint[] = [];
  for (const node of children(el, 'dPt')) {
    if (points.length >= MAX_CHART_DATA_POINTS) break;
    const at = int(val(node, 'idx'));
    if (at === undefined || at < 0 || at >= MAX_CHART_POINTS) continue;
    const point: ChartPoint = { index: at };
    const own = shapeOf(state, child(node, 'spPr'));
    if (own) point.shape = own;
    const mark = readMarker(state, child(node, 'marker'));
    if (mark) point.marker = mark;
    if (child(node, 'explosion')) note(state, 'Slices that are pulled out of a pie are drawn in place');
    if (point.shape || point.marker) points.push(point);
  }
  if (points.length > 0) series.points = points;
  if (child(el, 'invertIfNegative') && flag(el, 'invertIfNegative', false)) series.invertIfNegative = true;
  if (child(el, 'smooth') && flag(el, 'smooth', false)) {
    series.smooth = true;
    note(state, 'Smoothed lines are drawn straight');
  }
  if (hasLabels(child(el, 'dLbls'))) series.dataLabels = true;
  if (child(el, 'trendline')) note(state, 'Trend lines are not drawn');
  if (child(el, 'errBars')) note(state, 'Error bars are not drawn');
  if (child(el, 'explosion')) note(state, 'Slices that are pulled out of a pie are drawn in place');
  return series;
}

/** Whether a <c:dLbls> shows anything. */
function hasLabels(el: XmlElement | undefined): boolean {
  if (!el) return false;
  if (child(el, 'delete') && flag(el, 'delete', false)) return false;
  for (const name of ['showVal', 'showCatName', 'showSerName', 'showPercent', 'showBubbleSize', 'showLegendKey']) {
    if (child(el, name) && flag(el, name, false)) return true;
  }
  // Labels of single points count too.
  return children(el, 'dLbl').length > 0;
}

function readGroup(state: ReadState, el: XmlElement, kind: ChartKind, threeD: boolean): ChartGroup {
  const grouping = val(el, 'grouping');
  const group: ChartGroup = {
    kind,
    grouping: grouping === 'stacked' ? 'stacked' : grouping === 'percentStacked' ? 'percentStacked' : 'standard',
    axisIds: [],
    series: [],
  };
  if (threeD) {
    group.threeD = true;
    note(state, 'A 3-D chart is drawn flat');
  }
  if (kind === 'bar') {
    group.barDir = val(el, 'barDir') === 'bar' ? 'bar' : 'col';
    const gap = int(val(el, 'gapWidth'));
    if (gap !== undefined) group.gapWidth = clamp(gap, 0, 500);
    const overlap = int(val(el, 'overlap'));
    if (overlap !== undefined) group.overlap = clamp(overlap, -100, 100);
  }
  if (kind === 'pie' || kind === 'doughnut') {
    const angle = int(val(el, 'firstSliceAng'));
    if (angle !== undefined) group.firstSliceAngle = clamp(angle, 0, 360);
    const hole = int((val(el, 'holeSize') || '').replace(/%$/, ''));
    if (kind === 'doughnut') group.holeSize = hole === undefined ? 50 : clamp(hole, 1, 90);
  }
  if (kind === 'line' && child(el, 'marker')) group.markers = flag(el, 'marker', true);
  if (kind === 'scatter') {
    const style = val(el, 'scatterStyle');
    group.scatterStyle = style !== undefined && /^(lineMarker|line|marker|smooth|smoothMarker|none)$/.test(style) ? style : 'marker';
  }
  // A pie is the type whose points differ in colour when the file says nothing.
  if (child(el, 'varyColors')) group.varyColors = flag(el, 'varyColors', true);
  else if (kind === 'pie' || kind === 'doughnut') group.varyColors = true;
  if (hasLabels(child(el, 'dLbls'))) group.dataLabels = true;
  for (const node of children(el, 'axId')) {
    const id = int(node.attrs.get('val'));
    if (id !== undefined && group.axisIds.length < 4) group.axisIds.push(id);
  }
  const series = children(el, 'ser');
  for (let i = 0; i < series.length; i++) {
    if (state.series >= MAX_CHART_SERIES) {
      note(state, `Only the first ${MAX_CHART_SERIES} series are read`);
      break;
    }
    const read = readSeries(state, series[i], kind === 'scatter', state.series);
    state.series++;
    group.series.push(read);
  }
  group.series.sort((a, b) => a.order - b.order);
  if (child(el, 'dropLines') || child(el, 'hiLowLines') || child(el, 'upDownBars') || child(el, 'serLines')) note(state, 'Drop lines, high-low lines, up-down bars and series lines are not drawn');
  return group;
}

// --------------------------------------------------------------------- axes

function readAxis(state: ReadState, el: XmlElement, kind: 'cat' | 'val'): ChartAxis | undefined {
  const id = int(val(el, 'axId'));
  if (id === undefined) return undefined;
  const position = val(el, 'axPos');
  const labels = val(el, 'tickLblPos');
  const axis: ChartAxis = {
    id,
    kind,
    position: position === 'l' || position === 't' || position === 'r' || position === 'b' ? position : kind === 'cat' ? 'b' : 'l',
    tickLabels: labels === 'low' || labels === 'high' || labels === 'none' ? labels : 'nextTo',
  };
  const local = localName(el.name);
  if (local === 'dateAx') axis.date = true;
  if (local === 'serAx') note(state, 'A series axis is not drawn');
  // UNSURE: an axis without the element is drawn. (The default of the element in the schema is "deleted"; Excel always writes it.)
  if (child(el, 'delete') && flag(el, 'delete', false)) axis.deleted = true;
  const scaling = child(el, 'scaling');
  const min = number(val(scaling, 'min'));
  if (min !== undefined) axis.min = min;
  const max = number(val(scaling, 'max'));
  if (max !== undefined) axis.max = max;
  if (val(scaling, 'orientation') === 'maxMin') axis.reversed = true;
  const logBase = number(val(scaling, 'logBase'));
  if (logBase !== undefined && logBase >= 2) {
    axis.logBase = logBase;
    note(state, 'A logarithmic axis is drawn as a linear one');
  }
  const unit = number(val(el, 'majorUnit'));
  if (unit !== undefined && unit > 0) axis.majorUnit = unit;
  const format = child(el, 'numFmt');
  if (format) {
    const code = format.attrs.get('formatCode');
    if (code !== undefined && code !== '' && code.length <= 255) axis.formatCode = code;
    const linked = format.attrs.get('sourceLinked');
    if (linked === '1' || linked === 'true') axis.sourceLinked = true;
    else if (linked === '0' || linked === 'false') axis.sourceLinked = false;
  }
  const major = child(el, 'majorGridlines');
  if (major) axis.majorGridlines = lineOf(state, child(major, 'spPr')) || {};
  const minor = child(el, 'minorGridlines');
  if (minor) axis.minorGridlines = lineOf(state, child(minor, 'spPr')) || {};
  const title = readTitle(state, child(el, 'title'));
  if (title) axis.title = title;
  const crosses = val(el, 'crosses');
  if (crosses === 'autoZero' || crosses === 'min' || crosses === 'max') axis.crosses = crosses;
  const crossesAt = number(val(el, 'crossesAt'));
  if (crossesAt !== undefined) axis.crossesAt = crossesAt;
  const crossAxis = int(val(el, 'crossAx'));
  if (crossAxis !== undefined) axis.crossAxis = crossAxis;
  const between = val(el, 'crossBetween');
  if (between === 'between' || between === 'midCat') axis.crossBetween = between;
  const line = lineOf(state, child(el, 'spPr'));
  if (line) axis.line = line;
  const style = bodyStyle(state, child(el, 'txPr'));
  if (style) axis.labelStyle = style;
  if (child(el, 'dispUnits')) note(state, 'Display units of an axis are not applied');
  return axis;
}

function readLegend(state: ReadState, el: XmlElement | undefined): ChartLegend | undefined {
  if (!el) return undefined;
  const position = val(el, 'legendPos');
  const legend: ChartLegend = { position: position === 'l' || position === 't' || position === 'b' || position === 'tr' ? position : 'r' };
  const deleted: number[] = [];
  for (const entry of children(el, 'legendEntry')) {
    const index = int(val(entry, 'idx'));
    if (index !== undefined && index >= 0 && child(entry, 'delete') && flag(entry, 'delete', false) && deleted.length < MAX_CHART_SERIES) deleted.push(index);
  }
  if (deleted.length > 0) legend.deleted = deleted;
  const style = bodyStyle(state, child(el, 'txPr'));
  if (style) legend.style = style;
  return legend;
}

// -------------------------------------------------------------------- chart

/** What reading the charts of a workbook may still cost; shared by all of them. */
export interface ChartReadBudget {
  /** Cached values (<c:pt>) that may still be read. */
  points: number;
}

/** Thrown for a chart whose cached values do not fit what is left of the workbook's budget. */
export class ChartBudgetError extends Error {}

/** The cached values (<c:pt>) a chart holds, all series together. */
export function chartCachedPoints(chart: Chart): number {
  let total = 0;
  for (const group of chart.groups) {
    for (const series of group.series) {
      for (const data of [series.categories, series.values]) if (data && data.cache && Array.isArray(data.cache.index)) total += data.cache.index.length;
    }
  }
  return total;
}

/**
 * The chart of a chart part, from its root element (<c:chartSpace>). Undefined
 * for a part that is no chart, that holds no series of a type that is drawn,
 * or that holds a type that is not drawn at all: the drawing then shows a
 * placeholder. `themeColors` are those of the style sheet ("RRGGBB": lt1, dk1,
 * lt2, dk2, accent1..6, ...). Never throws for what the part states.
 *
 * With a `budget` the cached values of the chart are taken from it; a chart
 * that holds more than is left of it throws a ChartBudgetError (the only
 * thing this throws) and takes nothing.
 */
export function parseChart(root: XmlElement, themeColors: string[], budget?: ChartReadBudget): Chart | undefined {
  const limit = budget ? Math.max(Math.min(budget.points, MAX_CHART_CACHE), 0) : MAX_CHART_CACHE;
  const state: ReadState = { themeColors: Array.isArray(themeColors) ? themeColors : [], notes: [], points: 0, pointLimit: limit, series: 0 };
  let chart: Chart | undefined;
  try {
    chart = readChart(root, state);
  } catch (error) {
    // Whatever a part holds, it costs no more than its chart.
    chart = undefined;
  }
  if (budget) {
    // A chart cut short by what other charts took is not half a chart: it is not shown.
    if (chart && state.cut && limit < MAX_CHART_CACHE) throw new ChartBudgetError('the charts of the workbook hold too many cached values');
    if (chart) budget.points -= state.points;
  }
  if (chart && state.cut) note(state, `Only the first ${MAX_CHART_CACHE} cached values of a chart are read`);
  return chart;
}

function readChart(root: XmlElement, state: ReadState): Chart | undefined {
  if (!root || localName(root.name) !== 'chartSpace') return undefined;
  const chartEl = child(root, 'chart');
  const plot = child(chartEl, 'plotArea');
  if (!chartEl || !plot) return undefined;
  const groups: ChartGroup[] = [];
  const axes: ChartAxis[] = [];
  for (const node of plot.children) {
    if (node.type !== 'element') continue;
    const local = localName(node.name);
    if (Object.prototype.hasOwnProperty.call(GROUPS, local)) {
      const type = GROUPS[local];
      if (groups.length < 16) groups.push(readGroup(state, node, type.kind, type.threeD === true));
    } else if (Object.prototype.hasOwnProperty.call(AXES, local)) {
      const axis = axes.length < MAX_CHART_AXES ? readAxis(state, node, AXES[local]) : undefined;
      if (axis) axes.push(axis);
    } else if (/Chart$/.test(local)) {
      // radarChart, bubbleChart, stockChart, surfaceChart, ofPieChart, ...: one type that is not drawn hides the whole chart.
      return undefined;
    } else if (local === 'dTable') note(state, 'The data table of a chart is not drawn');
  }
  if (groups.length === 0 || !groups.some((group) => group.series.length > 0)) return undefined;

  const palette: string[] = [];
  for (let i = 0; i < 6; i++) {
    const theme = state.themeColors[4 + i];
    palette.push('#' + (typeof theme === 'string' && HEX6.test(theme) ? theme.toUpperCase() : OFFICE_ACCENTS[i]));
  }
  const blanks = child(chartEl, 'dispBlanksAs');
  const stated = blanks ? blanks.attrs.get('val') : undefined;
  const chart: Chart = {
    groups,
    axes,
    plotVisOnly: flag(chartEl, 'plotVisOnly', true),
    // UNSURE: a chart without the element is drawn with gaps, as Excel writes it for a new chart; an element without a value means "zero" (the schema).
    blanks: !blanks ? 'gap' : stated === 'gap' || stated === 'span' ? stated : 'zero',
    palette,
    notes: state.notes,
  };
  const title = readTitle(state, child(chartEl, 'title'));
  if (title) chart.title = title;
  if (child(chartEl, 'autoTitleDeleted') && flag(chartEl, 'autoTitleDeleted', false)) chart.autoTitleDeleted = true;
  const legend = readLegend(state, child(chartEl, 'legend'));
  if (legend) chart.legend = legend;
  if (child(root, 'roundedCorners') && flag(root, 'roundedCorners', false)) chart.roundedCorners = true;
  if (child(root, 'date1904') && flag(root, 'date1904', false)) chart.date1904 = true;
  const area = shapeOf(state, child(root, 'spPr'));
  if (area) chart.area = area;
  const plotShape = shapeOf(state, child(plot, 'spPr'));
  if (plotShape) chart.plotArea = plotShape;
  const text = bodyStyle(state, child(root, 'txPr'));
  if (text) chart.textStyle = text;

  if (child(root, 'style')) note(state, 'The style number of a chart is not applied; series without a colour get the accent colours of the theme');
  const layouts = [child(plot, 'layout'), child(child(chartEl, 'title'), 'layout'), child(child(chartEl, 'legend'), 'layout')];
  if (layouts.some((layout) => child(layout, 'manualLayout') !== undefined)) note(state, 'Places set by hand (of the plot area, the title or the legend) are not applied');
  let labels = 0;
  for (const group of groups) {
    if (group.dataLabels) labels++;
    for (const series of group.series) if (series.dataLabels) labels++;
  }
  if (labels > 0) note(state, `Data labels are not drawn (${labels} series or chart types have them)`);
  if (groups.some((group) => group.kind === 'pie' || group.kind === 'doughnut') && groups.some((group) => group.kind !== 'pie' && group.kind !== 'doughnut')) {
    note(state, 'A pie combined with another chart type: only the pie is drawn');
  }
  if (secondaryAxes(groups, axes)) note(state, 'Series on a secondary axis are drawn against the primary axis');
  return chart;
}

/** Whether the chart types of the plot area are drawn against more than one pair of axes. */
function secondaryAxes(groups: ChartGroup[], axes: ChartAxis[]): boolean {
  let first: string | undefined;
  for (const group of groups) {
    if (group.kind === 'pie' || group.kind === 'doughnut' || group.axisIds.length === 0) continue;
    const key = group.axisIds.slice().sort().join(',');
    if (first === undefined) first = key;
    else if (key !== first) return true;
  }
  const valueAxes = axes.filter((axis) => axis.kind === 'val').length;
  return valueAxes > (groups.some((group) => group.kind === 'scatter') ? 2 : 1);
}

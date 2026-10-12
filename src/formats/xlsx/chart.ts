// Reads a chart part (xl/charts/chartN.xml, DrawingML <c:chartSpace>) into the
// read-only model the grid draws a chart from: the chart types of the plot
// area with their series, the axes, the title and the legend.
//
// Nothing here is ever serialized: the part stays in the package and is
// written back from there. A chart that holds a type that is not drawn (a
// surface) or that cannot be read is not a chart of the model: parseChart()
// answers undefined and the drawing shows a placeholder.

import {
  MAX_CHART_AXES,
  MAX_CHART_POINTS,
  MAX_CHART_SERIES,
  MAX_CHART_TRENDLINES,
  type Chart,
  type ChartAxis,
  type ChartCache,
  type ChartData,
  type ChartDataLabels,
  type ChartErrorBars,
  type ChartGroup,
  type ChartKind,
  type ChartLegend,
  type ChartLine,
  type ChartManualLayout,
  type ChartMarker,
  type ChartPoint,
  type ChartPointLabel,
  type ChartSeries,
  type ChartShape,
  type ChartTextStyle,
  type ChartTitle,
  type ChartTrendline,
  type ChartUpDownBars,
} from '../../model/sheet/chart';
import { localName, type XmlElement } from '../../xml';
import { drawingColor, EMU_PER_PIXEL } from './drawings';

/** Cached values (<c:pt>) read from one chart; what lies beyond is left out. */
export const MAX_CHART_CACHE = 1000000;
/** Points drawn their own way (c:dPt) read from one series. */
export const MAX_CHART_DATA_POINTS = 4096;
/** Levels of categories (c:lvl) read from one reference. */
const MAX_CHART_LEVELS = 8;
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
  radarChart: { kind: 'radar' },
  bubbleChart: { kind: 'bubble' },
  stockChart: { kind: 'stock' },
  // Only the pie of a pie with a second plot is drawn.
  ofPieChart: { kind: 'pie' },
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

/** The plain text of a rich text (c:rich), its paragraphs as lines; the look of the first run goes into `style`. */
function richText(state: ReadState, rich: XmlElement, style: ChartTextStyle): string {
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
  return text.slice(0, MAX_TEXT).replace(/^\n+|\n+$/g, '');
}

/** A place set by hand (c:layout with c:manualLayout), as the file states it. */
function readLayout(holder: XmlElement | undefined): ChartManualLayout | undefined {
  const manual = child(child(holder, 'layout'), 'manualLayout');
  if (!manual) return undefined;
  const layout: ChartManualLayout = {};
  const target = val(manual, 'layoutTarget');
  if (target === 'inner' || target === 'outer') layout.target = target;
  const xMode = val(manual, 'xMode');
  if (xMode === 'edge' || xMode === 'factor') layout.xMode = xMode;
  const yMode = val(manual, 'yMode');
  if (yMode === 'edge' || yMode === 'factor') layout.yMode = yMode;
  // What the width and the height mean is only kept where it is not the usual (a size).
  if (val(manual, 'wMode') === 'edge') layout.wMode = 'edge';
  if (val(manual, 'hMode') === 'edge') layout.hMode = 'edge';
  for (const key of ['x', 'y', 'w', 'h'] as const) {
    const value = number(val(manual, key));
    if (value !== undefined) layout[key] = value;
  }
  return Object.keys(layout).length > 0 ? layout : undefined;
}

/** A title: the plain text of its rich text with the look of the first run, or the text a reference cached. */
function readTitle(state: ReadState, title: XmlElement | undefined): ChartTitle | undefined {
  if (!title) return undefined;
  const out: ChartTitle = {};
  const style: ChartTextStyle = { ...(bodyStyle(state, child(title, 'txPr')) || {}) };
  const tx = child(title, 'tx');
  const rich = child(tx, 'rich');
  const layout = readLayout(title);
  if (layout) out.layout = layout;
  if (rich) {
    const text = richText(state, rich, style);
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
        const cache = child(node, 'multiLvlStrCache');
        const levels = children(cache, 'lvl');
        readCache(state, levels[0], 'str', data);
        const further: ChartCache[] = [];
        for (let i = 1; i < levels.length && i < MAX_CHART_LEVELS; i++) {
          const level: ChartData = { kind: 'str', cache: { index: [], values: [] }, count: 0 };
          readCache(state, levels[i], 'str', level);
          further.push(level.cache);
        }
        if (further.length > 0) data.levels = further;
        if (levels.length > 2) note(state, 'Of more than two levels of categories the two next to the axis are shown');
        const stated = int(val(cache, 'ptCount'));
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

const LABEL_FLAGS: [string, 'showValue' | 'showCategory' | 'showSeries' | 'showPercent' | 'showBubbleSize' | 'showLegendKey'][] = [
  ['showVal', 'showValue'],
  ['showCatName', 'showCategory'],
  ['showSerName', 'showSeries'],
  ['showPercent', 'showPercent'],
  ['showBubbleSize', 'showBubbleSize'],
  ['showLegendKey', 'showLegendKey'],
];
const LABEL_POSITION = /^(outEnd|inEnd|ctr|inBase|bestFit|t|b|l|r)$/;

/** What a <c:dLbls> or a <c:dLbl> says; undefined when it says nothing. */
function readLabels(state: ReadState, el: XmlElement | undefined): ChartDataLabels | undefined {
  if (!el) return undefined;
  if (child(el, 'delete') && flag(el, 'delete', false)) return { deleted: true };
  const labels: ChartDataLabels = {};
  for (const [name, key] of LABEL_FLAGS) if (child(el, name)) labels[key] = flag(el, name, false);
  if (labels.showLegendKey) note(state, 'The legend key in a data label is not drawn');
  if (child(el, 'showLeaderLines') && flag(el, 'showLeaderLines', false)) note(state, 'Leader lines of data labels are not drawn');
  const separator = child(el, 'separator');
  if (separator) {
    // The text of the element; files not written by Excel state it as an attribute instead.
    const text = (textOf(separator) || separator.attrs.get('val') || '').slice(0, 16);
    if (text !== '') labels.separator = text;
  }
  const format = child(el, 'numFmt');
  if (format) {
    const code = format.attrs.get('formatCode');
    if (code !== undefined && code !== '' && code.length <= 255) labels.formatCode = code;
    const linked = format.attrs.get('sourceLinked');
    if (linked === '1' || linked === 'true') labels.sourceLinked = true;
    else if (linked === '0' || linked === 'false') labels.sourceLinked = false;
  }
  const position = val(el, 'dLblPos');
  if (position !== undefined && LABEL_POSITION.test(position)) labels.position = position as ChartDataLabels['position'];
  const style: ChartTextStyle = { ...(bodyStyle(state, child(el, 'txPr')) || {}) };
  const rich = child(child(el, 'tx'), 'rich');
  if (rich) {
    const text = richText(state, rich, style).replace(/\n/g, ' ');
    if (text !== '') labels.text = text;
  }
  if (Object.keys(style).length > 0) labels.style = style;
  return Object.keys(labels).length > 0 ? labels : undefined;
}

const TREND_TYPES = /^(linear|exp|log|poly|power|movingAvg)$/;

function readTrendline(state: ReadState, el: XmlElement): ChartTrendline | undefined {
  const type = val(el, 'trendlineType');
  if (type === undefined || !TREND_TYPES.test(type)) return undefined;
  const line: ChartTrendline = { type: type as ChartTrendline['type'] };
  if (type === 'poly') {
    const order = int(val(el, 'order'));
    line.order = order === undefined ? 2 : clamp(order, 2, 6);
  }
  if (type === 'movingAvg') {
    const period = int(val(el, 'period'));
    line.period = period === undefined ? 2 : clamp(period, 2, MAX_CHART_POINTS);
  }
  const forward = number(val(el, 'forward'));
  if (forward !== undefined && forward > 0) line.forward = forward;
  const backward = number(val(el, 'backward'));
  if (backward !== undefined && backward > 0) line.backward = backward;
  const name = child(el, 'name');
  if (name && textOf(name) !== '') line.name = textOf(name).slice(0, MAX_TEXT);
  const look = lineOf(state, child(el, 'spPr'));
  if (look) line.line = look;
  if ((child(el, 'dispRSqr') && flag(el, 'dispRSqr', false)) || (child(el, 'dispEq') && flag(el, 'dispEq', false))) note(state, 'The equation and R² of a trend line are not shown');
  if (child(el, 'intercept')) note(state, 'A set intercept of a trend line is not applied');
  return line;
}

function readErrorBars(state: ReadState, el: XmlElement, standing: boolean): ChartErrorBars {
  const direction = val(el, 'errDir');
  const type = val(el, 'errBarType');
  const kind = val(el, 'errValType');
  const bars: ChartErrorBars = {
    // A chart type without horizontal values has bars along its values only.
    direction: direction === 'x' && !standing ? 'x' : 'y',
    type: type === 'plus' || type === 'minus' ? type : 'both',
    valueType: kind === 'percentage' || kind === 'stdDev' || kind === 'stdErr' || kind === 'cust' ? kind : 'fixedVal',
  };
  const amount = number(val(el, 'val'));
  if (amount !== undefined) bars.value = amount;
  if (bars.valueType === 'cust') {
    const plus = readData(state, child(el, 'plus'));
    if (plus && plus.kind === 'num') bars.plus = plus;
    const minus = readData(state, child(el, 'minus'));
    if (minus && minus.kind === 'num') bars.minus = minus;
  }
  if (child(el, 'noEndCap') && flag(el, 'noEndCap', false)) bars.noEndCap = true;
  const look = lineOf(state, child(el, 'spPr'));
  if (look) bars.line = look;
  return bars;
}

/** How far a slice is pulled out, as a percentage of the radius; undefined when nothing is stated. */
function explosionOf(holder: XmlElement): number | undefined {
  const value = int(val(holder, 'explosion'));
  return value === undefined ? undefined : clamp(value, 0, 400);
}

function readSeries(state: ReadState, el: XmlElement, kind: ChartKind, position: number): ChartSeries {
  const scatter = kind === 'scatter' || kind === 'bubble';
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
    const pulled = explosionOf(node);
    if (pulled !== undefined) point.explosion = pulled;
    if (point.shape || point.marker || pulled !== undefined) points.push(point);
  }
  if (points.length > 0) series.points = points;
  if (child(el, 'invertIfNegative') && flag(el, 'invertIfNegative', false)) series.invertIfNegative = true;
  if (child(el, 'smooth') && flag(el, 'smooth', false)) series.smooth = true;
  const labelsEl = child(el, 'dLbls');
  const labels = readLabels(state, labelsEl);
  if (labels) series.dataLabels = labels;
  const own: ChartPointLabel[] = [];
  for (const node of children(labelsEl, 'dLbl')) {
    if (own.length >= MAX_CHART_DATA_POINTS) break;
    const at = int(val(node, 'idx'));
    if (at === undefined || at < 0 || at >= MAX_CHART_POINTS) continue;
    const label = readLabels(state, node);
    if (label) own.push({ index: at, ...label });
  }
  if (own.length > 0) series.pointLabels = own.sort((a, b) => a.index - b.index);
  const trends: ChartTrendline[] = [];
  for (const node of children(el, 'trendline')) {
    const trend = trends.length < MAX_CHART_TRENDLINES ? readTrendline(state, node) : undefined;
    if (trend) trends.push(trend);
  }
  if (trends.length > 0) series.trendlines = trends;
  const errors: ChartErrorBars[] = [];
  for (const node of children(el, 'errBars')) {
    const bars = readErrorBars(state, node, !scatter);
    // One set along each axis: the first the file has.
    if (!errors.some((other) => other.direction === bars.direction)) errors.push(bars);
  }
  if (errors.length > 0) series.errorBars = errors;
  const pulled = explosionOf(el);
  if (pulled !== undefined) series.explosion = pulled;
  if (kind === 'bubble') {
    const sizes = readData(state, child(el, 'bubbleSize'));
    if (sizes && sizes.kind === 'num') series.bubbleSizes = sizes;
    if (child(el, 'bubble3D') && flag(el, 'bubble3D', false)) note(state, 'A 3-D chart is drawn flat');
  }
  return series;
}

/** The look of a line an element of a chart type holds (c:dropLines, c:hiLowLines); undefined without the element. */
function groupLine(state: ReadState, el: XmlElement, name: string): ChartLine | undefined {
  const holder = child(el, name);
  return holder ? lineOf(state, child(holder, 'spPr')) || {} : undefined;
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
  if (kind === 'radar') {
    const style = val(el, 'radarStyle');
    group.radarStyle = style === 'marker' || style === 'filled' ? style : 'standard';
  }
  if (kind === 'bubble') {
    const scale = int((val(el, 'bubbleScale') || '').replace(/%$/, ''));
    if (scale !== undefined) group.bubbleScale = clamp(scale, 0, 300);
    if (child(el, 'showNegBubbles') && flag(el, 'showNegBubbles', false)) group.showNegBubbles = true;
    const represents = val(el, 'sizeRepresents');
    if (represents === 'area' || represents === 'w') group.sizeRepresents = represents;
  }
  if (kind === 'line' || kind === 'area' || kind === 'stock') {
    const drop = groupLine(state, el, 'dropLines');
    if (drop) group.dropLines = drop;
  }
  if (kind === 'line' || kind === 'stock') {
    const hiLow = groupLine(state, el, 'hiLowLines');
    if (hiLow) group.hiLowLines = hiLow;
    const upDown = child(el, 'upDownBars');
    if (upDown) {
      const boxes: ChartUpDownBars = {};
      const gap = int((val(upDown, 'gapWidth') || '').replace(/%$/, ''));
      if (gap !== undefined) boxes.gapWidth = clamp(gap, 0, 500);
      const up = shapeOf(state, child(child(upDown, 'upBars'), 'spPr'));
      if (up) boxes.up = up;
      const down = shapeOf(state, child(child(upDown, 'downBars'), 'spPr'));
      if (down) boxes.down = down;
      group.upDownBars = boxes;
    }
  }
  if (localName(el.name) === 'ofPieChart') {
    const split = val(el, 'splitType');
    const second: NonNullable<ChartGroup['ofPie']> = { split: split === 'pos' || split === 'val' || split === 'percent' || split === 'cust' ? split : 'auto' };
    const at = number(val(el, 'splitPos'));
    if (at !== undefined && second.split !== 'cust') second.position = at;
    if (second.split === 'cust') {
      const listed: number[] = [];
      for (const node of children(child(el, 'custSplit'), 'secondPiePt')) {
        const index = int(node.attrs.get('val'));
        if (index !== undefined && index >= 0 && index < MAX_CHART_POINTS && listed.length < MAX_CHART_DATA_POINTS) listed.push(index);
      }
      second.points = listed;
    }
    group.ofPie = second;
    note(state, 'Of a pie with a second plot the pie alone is drawn: the points of the second plot are one slice of it');
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
  const labels = readLabels(state, child(el, 'dLbls'));
  if (labels) group.dataLabels = labels;
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
    const read = readSeries(state, series[i], kind, state.series);
    state.series++;
    group.series.push(read);
  }
  group.series.sort((a, b) => a.order - b.order);
  if (child(el, 'serLines')) note(state, 'Series lines between stacked bars are not drawn');
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
  if (logBase !== undefined && logBase >= 2 && logBase <= 1000 && kind === 'val') {
    axis.logBase = logBase;
    note(state, 'On a logarithmic axis values that are not above zero are left out');
  }
  if (local === 'dateAx') {
    const base = val(el, 'baseTimeUnit');
    if (base === 'days' || base === 'months' || base === 'years') axis.baseTimeUnit = base;
    const major = val(el, 'majorTimeUnit');
    if (major === 'days' || major === 'months' || major === 'years') axis.majorTimeUnit = major;
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
  if (crossesAt !== undefined && kind === 'val') note(state, 'A value axis that crosses the other axis at a stated place stands at the start or the end of the plot area');
  const crossAxis = int(val(el, 'crossAx'));
  if (crossAxis !== undefined) axis.crossAxis = crossAxis;
  const between = val(el, 'crossBetween');
  if (between === 'between' || between === 'midCat') axis.crossBetween = between;
  const line = lineOf(state, child(el, 'spPr'));
  if (line) axis.line = line;
  const style = bodyStyle(state, child(el, 'txPr'));
  if (style) axis.labelStyle = style;
  // Sixty-thousandths of a degree, clockwise.
  const body = child(child(el, 'txPr'), 'bodyPr');
  const turned = int(body ? body.attrs.get('rot') : undefined);
  if (turned !== undefined) axis.labelRotation = Math.round(clamp(turned / 60000, -90, 90));
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
  const layout = readLayout(el);
  if (layout) legend.layout = layout;
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
      const sources = [series.categories, series.values, series.bubbleSizes];
      for (const bars of series.errorBars || []) sources.push(bars.plus, bars.minus);
      for (const data of sources) {
        if (!data || !data.cache || !Array.isArray(data.cache.index)) continue;
        total += data.cache.index.length;
        for (const level of data.levels || []) if (level && Array.isArray(level.index)) total += level.index.length;
      }
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
      // surfaceChart, surface3DChart, ...: one type that is not drawn hides the whole chart.
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
  const plotLayout = readLayout(plot);
  if (plotLayout) chart.plotLayout = plotLayout;
  const text = bodyStyle(state, child(root, 'txPr'));
  if (text) chart.textStyle = text;

  if (child(root, 'style')) note(state, 'The style number of a chart is not applied; series without a colour get the accent colours of the theme');
  if (groups.some((group) => group.kind === 'pie' || group.kind === 'doughnut') && groups.some((group) => group.kind !== 'pie' && group.kind !== 'doughnut')) {
    note(state, 'A pie combined with another chart type: only the pie is drawn');
  } else if (groups.some((group) => group.kind === 'radar') && groups.some((group) => group.kind !== 'radar')) {
    note(state, 'A radar combined with another chart type: only the radar is drawn');
  }
  return chart;
}

// The charts of a BIFF8 workbook: the substream of a chart (BOF with
// dt = 0x0020 ... EOF), which stands after the OBJ record of an embedded chart
// or is the substream of a chart sheet, read into the model the charts of an
// xlsx file are read into (model/sheet/chart.ts). A view for showing: an xls
// file is never written, and a chart read here is not written from its model
// either (Chart.shownOnly).
//
// Inside the substream the Chart record holds everything in blocks: a record,
// Begin, the records that belong to it, End. The series (Series) name their
// data in AI records - formula tokens that are turned into the text of a
// reference, so that the cells are read where the workbook has them - and the
// chart group they are drawn in (SerToCrt); the chart groups (ChartFormat with
// one record of a chart type) stand in the axis groups (AxisParent) beside
// their axes. After the block the values the file cached follow as cell
// records behind SIIndex records.
//
// UNSURE, all of it: no file written by Excel was at hand and no tool on the
// machine writes these records. The layouts are those of the specification of
// the format ([MS-XLS] 2.4.x) as far as they are remembered; single doubts are
// marked where they matter. For that reason the reader is strict: a record
// that is too short for its fields, blocks that do not close, a series that
// names no chart group, a chart type that is not drawn - any of them leaves
// the chart a placeholder, with a warning on its sheet, rather than a chart
// that may be wrong. Nothing is trusted: every read is checked against the
// record it is made in, and nesting, records, series, points and cached
// values are bounded.

import {
  MAX_CHART_AXES,
  MAX_CHART_POINTS,
  MAX_CHART_SERIES,
  type Chart,
  type ChartAxis,
  type ChartData,
  type ChartDataLabels,
  type ChartGroup,
  type ChartLabelPosition,
  type ChartLegend,
  type ChartLine,
  type ChartMarker,
  type ChartPoint,
  type ChartPointLabel,
  type ChartSeries,
  type ChartShape,
  type ChartTextStyle,
  type ChartTitle,
} from '../../model/sheet/chart';
import { formatCodeOf, resolveColor } from '../../model/sheet/color';
import type { Font, StyleSheet } from '../../model/sheet/types';
import { MAX_CHART_CACHE, MAX_CHART_DATA_POINTS } from '../xlsx/chart';
import { formulaText, type FormulaContext } from './formula';
import { REC_BLANK, REC_BOF, REC_BOOLERR, REC_EOF, REC_LABEL, REC_NUMBER, readerAt, type RecordStream } from './records';
import { readString16 } from './strings';

const REC_CHART = 0x1002;
const REC_SERIES = 0x1003;
const REC_DATAFORMAT = 0x1006;
const REC_LINEFORMAT = 0x1007;
const REC_MARKERFORMAT = 0x1009;
const REC_AREAFORMAT = 0x100a;
const REC_PIEFORMAT = 0x100b;
const REC_ATTACHEDLABEL = 0x100c;
const REC_SERIESTEXT = 0x100d;
const REC_CHARTFORMAT = 0x1014;
const REC_LEGEND = 0x1015;
const REC_BAR = 0x1017;
const REC_LINE = 0x1018;
const REC_PIE = 0x1019;
const REC_AREA = 0x101a;
const REC_SCATTER = 0x101b;
const REC_CRTLINE = 0x101c;
const REC_AXIS = 0x101d;
const REC_TICK = 0x101e;
const REC_VALUERANGE = 0x101f;
const REC_CATSERRANGE = 0x1020;
const REC_AXISLINE = 0x1021;
const REC_DEFAULTTEXT = 0x1024;
const REC_TEXT = 0x1025;
const REC_FONTX = 0x1026;
const REC_OBJECTLINK = 0x1027;
const REC_FRAME = 0x1032;
const REC_BEGIN = 0x1033;
const REC_END = 0x1034;
const REC_PLOTAREA = 0x1035;
const REC_CHART3D = 0x103a;
const REC_DROPBAR = 0x103d;
const REC_RADAR = 0x103e;
const REC_SURF = 0x103f;
const REC_RADARAREA = 0x1040;
const REC_AXISPARENT = 0x1041;
const REC_SHTPROPS = 0x1044;
const REC_SERTOCRT = 0x1045;
const REC_SERPARENT = 0x104a;
const REC_SERAUXTREND = 0x104b;
const REC_IFMT = 0x104e;
const REC_AI = 0x1051;
const REC_SERAUXERRBAR = 0x105b;
const REC_SERFMT = 0x105d;
const REC_BOPPOP = 0x1061;
const REC_AXCEXT = 0x1062;
const REC_DAT = 0x1063;
const REC_SIINDEX = 0x1065;
const REC_GELFRAME = 0x1066;

/** Records that are known and say nothing that is drawn: units, positions, links, the fonts a chart was scaled by, text runs, pictures of fills. */
const QUIET: { [id: number]: true } = {
  0x003c: true, // Continue
  0x00a0: true, // Scl
  0x1001: true, // Units
  0x1016: true, // SeriesList
  0x1022: true, // CrtLink
  0x103c: true, // PicF
  0x1043: true, // LegendException
  0x1046: true, // AxesUsed
  0x1048: true, // SBaseRef
  0x104f: true, // Pos
  0x1050: true, // AlRuns
  0x105c: true, // ClrtClient
  0x105f: true, // Chart3DBarShape
  0x1060: true, // Fbi
  0x1064: true, // PlotGrowth
  0x1067: true, // BopPopCustom
};
const KNOWN: { [id: number]: true } = {};
for (const id of [
  REC_CHART, REC_SERIES, REC_DATAFORMAT, REC_LINEFORMAT, REC_MARKERFORMAT, REC_AREAFORMAT, REC_PIEFORMAT, REC_ATTACHEDLABEL, REC_SERIESTEXT, REC_CHARTFORMAT, REC_LEGEND,
  REC_BAR, REC_LINE, REC_PIE, REC_AREA, REC_SCATTER, REC_CRTLINE, REC_AXIS, REC_TICK, REC_VALUERANGE, REC_CATSERRANGE, REC_AXISLINE, REC_DEFAULTTEXT, REC_TEXT, REC_FONTX,
  REC_OBJECTLINK, REC_FRAME, REC_PLOTAREA, REC_CHART3D, REC_DROPBAR, REC_RADAR, REC_SURF, REC_RADARAREA, REC_AXISPARENT, REC_SHTPROPS, REC_SERTOCRT, REC_SERPARENT,
  REC_SERAUXTREND, REC_IFMT, REC_AI, REC_SERAUXERRBAR, REC_SERFMT, REC_BOPPOP, REC_AXCEXT, REC_DAT, REC_GELFRAME,
]) KNOWN[id] = true;

/** Deepest nesting of blocks (Begin ... End) that is followed; a chart of Excel nests five deep. */
export const XLS_MAX_CHART_DEPTH = 16;
/** Records inside the Chart record of one chart; a chart with more is not read. */
export const XLS_MAX_CHART_RECORDS = 200000;
/**
 * Bytes the cached values of the charts of one workbook may take, all charts together: sixteen for a
 * number, sixteen and two for each of its characters for a text.
 */
export const XLS_MAX_CHART_CACHE_BYTES = 64 * 1024 * 1024;
/** Notes about one chart. */
const MAX_NOTES = 32;

/** What reading a chart needs of its workbook. */
export interface XlsChartContext {
  styles: StyleSheet;
  /** The model's font for a font index of the file. */
  fontOf(index: number): Font | undefined;
  /** The sheets and names the formula tokens of the references are turned into text with. */
  formula: FormulaContext;
  date1904: boolean;
  /** What the charts of the workbook may still cost, all sheets together: charts that are read, cached values, and the bytes those take. */
  budget: { charts: number; points: number; bytes: number };
}

/**
 * Where a substream may reach to, and how far it was read. No two substreams of a workbook share
 * records: one that runs into the records of another is not read (a file of a few bytes that
 * names the same records many times would cost each time what they cost once).
 */
export interface XlsSpan {
  /** The offset no record of the substream may reach beyond: the start of the next substream that was read, or the end of the stream. */
  limit: number;
  /** Set by the reader: the offset behind the last record it looked at, whether it read the substream or not. */
  end: number;
}

export interface XlsChartRead {
  /** The chart; missing for one that is not read: `problem` then says why. */
  chart?: Chart;
  problem?: string;
  /** The size the chart states for itself, in points. */
  width?: number;
  height?: number;
}

const OVERLAP = 'its records are those of another sheet';

/** Why a chart is left a placeholder. */
class Unreadable extends Error {}

interface Node {
  id: number;
  /** Offset of the content of the record in the stream, and its length. */
  d: number;
  length: number;
  children?: Node[];
}

const NONE: Node[] = [];
const kids = (node: Node | undefined, id: number): Node[] => (node && node.children ? node.children.filter((child) => child.id === id) : NONE);
const kid = (node: Node | undefined, id: number): Node | undefined => {
  const list = node && node.children ? node.children : NONE;
  for (let i = 0; i < list.length; i++) if (list[i].id === id) return list[i];
  return undefined;
};

const hex2 = (value: number): string => (value < 16 ? '0' : '') + value.toString(16).toUpperCase();
const clamp = (value: number, low: number, high: number): number => Math.min(Math.max(value, low), high);

/** imk of MarkerFormat. 6 and 7 are the short and the long bar ("Dow Jones", "standard deviation"). */
const MARKERS = ['none', 'square', 'diamond', 'triangle', 'x', 'star', 'dot', 'dash', 'circle', 'plus'];
/** dlp of Text: where a data label stands. */
const LABEL_PLACES: (ChartLabelPosition | undefined)[] = [undefined, 'outEnd', 'inEnd', 'ctr', 'inBase', 't', 'b', 'l', 'r', 'bestFit'];
const TIME_UNITS: ('days' | 'months' | 'years')[] = ['days', 'months', 'years'];

interface State {
  stream: RecordStream;
  view: DataView;
  context: XlsChartContext;
  notes: string[];
  /** Records of the tree so far. */
  records: number;
  span: XlsSpan;
  /** Ids of records nothing is known of, with how often each stands in the chart. */
  unknown: Map<number, number>;
  later: number;
}

function note(state: State, text: string): void {
  if (state.notes.length < MAX_NOTES && state.notes.indexOf(text) < 0) state.notes.push(text);
}

// ---------------------------------------------------------------------- tree

/** The records of a block, from the record after its Begin to its End; `depth` counts the blocks around it. */
function readBlock(state: State, start: number, depth: number): { nodes: Node[]; next: number } {
  if (depth > XLS_MAX_CHART_DEPTH) throw new Unreadable('its records are nested too deeply');
  const { view } = state;
  const size = state.stream.bytes.length;
  const nodes: Node[] = [];
  let p = start;
  for (;;) {
    if (p + 4 > size) throw new Unreadable('a block of its records is not closed');
    const id = view.getUint16(p, true);
    const length = view.getUint16(p + 2, true);
    const d = p + 4;
    const next = d + length;
    if (next > size) throw new Unreadable('a record runs past the end of the stream');
    if (next > state.span.limit) throw new Unreadable(OVERLAP);
    p = next;
    state.span.end = p;
    if (id === REC_END) return { nodes, next: p };
    // The substream ends, or another one starts, inside a block: the blocks do not close.
    if (id === REC_EOF || id === REC_BOF) throw new Unreadable('a block of its records is not closed');
    if (id === REC_BEGIN) {
      const owner = nodes[nodes.length - 1];
      if (!owner || owner.children) throw new Unreadable('a block of its records belongs to no record');
      const inner = readBlock(state, p, depth + 1);
      owner.children = inner.nodes;
      p = inner.next;
      continue;
    }
    if (++state.records > XLS_MAX_CHART_RECORDS) throw new Unreadable('it holds too many records');
    nodes.push({ id, d, length });
    if (id >= 0x0850 && id <= 0x08ff) state.later++;
    else if (!KNOWN[id] && !QUIET[id]) state.unknown.set(id, (state.unknown.get(id) || 0) + 1);
  }
}

// -------------------------------------------------------------------- fields

function need(node: Node, size: number, what: string): void {
  if (node.length < size) throw new Unreadable(`a ${what} record is too short`);
}

const u8 = (state: State, node: Node, at: number): number => state.view.getUint8(node.d + at);
const u16 = (state: State, node: Node, at: number): number => state.view.getUint16(node.d + at, true);
const i16 = (state: State, node: Node, at: number): number => state.view.getInt16(node.d + at, true);
const f64 = (state: State, node: Node, at: number): number => state.view.getFloat64(node.d + at, true);

/** LongRGB: red, green, blue and a byte that is not used. */
function rgb(state: State, node: Node, at: number): string {
  const b = state.stream.bytes;
  return '#' + hex2(b[node.d + at]) + hex2(b[node.d + at + 1]) + hex2(b[node.d + at + 2]);
}

/**
 * LineFormat: rgb (4), lns (2), we (2), flags (2: bit 0 fAuto, bit 2 fAxisOn - read where the line is
 * that of an axis), icv (2). Undefined for a line the chart gives by itself.
 */
function lineOf(state: State, node: Node | undefined): ChartLine | undefined {
  if (!node) return undefined;
  need(node, 10, 'LineFormat');
  if (u16(state, node, 8) & 1) return undefined;
  const style = u16(state, node, 4);
  // lns 5: no line.
  if (style === 5) return { color: null };
  const weight = i16(state, node, 6);
  // we: -1 hairline, 0 narrow, 1 medium, 2 wide.
  const line: ChartLine = { color: rgb(state, node, 0), width: weight < 0 ? 0.5 : weight === 0 ? 1 : weight === 1 ? 2 : 3 };
  if (style >= 1 && style <= 4) line.dash = true;
  return line;
}

/**
 * AreaFormat: rgbFore (4), rgbBack (4), fls (2), flags (2: bit 0 fAuto, bit 1 fInvertNeg). The
 * fill is the foreground colour, whatever the pattern (a pattern is drawn as a plain fill).
 */
function areaOf(state: State, node: Node | undefined): { fill?: string | null; invert?: boolean } {
  if (!node) return {};
  need(node, 12, 'AreaFormat');
  const flags = u16(state, node, 10);
  const out: { fill?: string | null; invert?: boolean } = {};
  if (flags & 2) out.invert = true;
  if (flags & 1) return out;
  out.fill = u16(state, node, 8) === 0 ? null : rgb(state, node, 0);
  return out;
}

/** The look of a frame (Frame with LineFormat and AreaFormat inside); undefined when it leaves everything to the chart. */
function frameShape(state: State, frame: Node | undefined): ChartShape | undefined {
  if (!frame) return undefined;
  const shape: ChartShape = {};
  const area = areaOf(state, kid(frame, REC_AREAFORMAT));
  if (area.fill !== undefined) shape.fill = area.fill;
  const line = lineOf(state, kid(frame, REC_LINEFORMAT));
  if (line) shape.line = line;
  if (kid(frame, REC_GELFRAME)) note(state, 'A gradient or picture fill of a chart is not drawn');
  return shape.fill !== undefined || shape.line ? shape : undefined;
}

function fontStyle(state: State, holder: Node | undefined): ChartTextStyle | undefined {
  const fontX = kid(holder, REC_FONTX);
  if (!fontX || fontX.length < 2) return undefined;
  const font = state.context.fontOf(u16(state, fontX, 0));
  if (!font) return undefined;
  const style: ChartTextStyle = {};
  if (typeof font.size === 'number' && font.size > 0) style.size = font.size;
  if (font.bold) style.bold = true;
  if (font.italic) style.italic = true;
  if (typeof font.name === 'string' && font.name !== '') style.font = font.name.slice(0, 255);
  const color = resolveColor(font.color, state.context.styles);
  if (color) style.color = color;
  return style;
}

/**
 * The style of a Text record: the font it names (FontX) and its colour.
 * Text: at (1), vat (1), wBkgMode (2), rgbText (4), x, y, dx, dy (4 each), flags (2: bit 0
 * fAutoColor, 1 fShowKey, 2 fShowValue, 4 fAutoText, 6 fDeleted, 11 fShowLabelAndPerc, 12
 * fShowPercent, 13 fShowBubbleSizes, 14 fShowLabel), icvText (2), dlp and more (2), trot (2).
 */
function textStyle(state: State, text: Node): ChartTextStyle | undefined {
  let style = fontStyle(state, text);
  if (text.length >= 26 && !(u16(state, text, 24) & 1)) {
    style = style || {};
    style.color = rgb(state, text, 4);
  }
  return style && Object.keys(style).length > 0 ? style : undefined;
}

/** SeriesText: two bytes that are not used, the count of characters (1), a byte that says how wide they are, the characters. */
function seriesText(state: State, node: Node | undefined): string | undefined {
  if (!node || node.length < 3) return undefined;
  const count = u8(state, node, 2);
  if (count === 0) return '';
  if (node.length < 4) return undefined;
  const wide = (u8(state, node, 3) & 1) !== 0;
  const b = state.stream.bytes;
  const end = node.d + node.length;
  let text = '';
  for (let i = 0, p = node.d + 4; i < count; i++) {
    if (p + (wide ? 2 : 1) > end) break;
    text += String.fromCharCode(wide ? b[p] | (b[p + 1] << 8) : b[p]);
    p += wide ? 2 : 1;
  }
  return text;
}

/**
 * AI: id (1: 0 name or title, 1 values, 2 categories, 3 bubble sizes), rt (1: 0 none, 1 literal,
 * 2 reference), flags (2), ifmt (2), cce (2), then cce bytes of formula tokens. The reference as
 * text, as a chart part of an xlsx file would state it; undefined for data that names no cells
 * and for tokens that cannot be turned into text.
 */
function reference(state: State, ai: Node | undefined): string | undefined {
  if (!ai) return undefined;
  need(ai, 8, 'AI');
  if (u8(state, ai, 1) !== 2) return undefined;
  const size = u16(state, ai, 6);
  if (size === 0) return undefined;
  const start = ai.d + 8;
  if (start + size > ai.d + ai.length) throw new Unreadable('the formula of an AI record runs past its record');
  const text = formulaText(state.context.formula, state.stream.bytes, start, start + size, start + size, ai.d + ai.length, 0, 0);
  if (text === undefined || text === '' || text.length > 8192) {
    note(state, 'A reference of the chart cannot be read; the values the file cached for it are shown');
    return undefined;
  }
  // Several areas are written in parentheses, as a chart part writes them.
  return text.indexOf(',') >= 0 && text.charCodeAt(0) !== 40 ? `(${text})` : text;
}

/** The AI record of a block with the given id; the first one when several state it. */
function aiOf(state: State, holder: Node, id: number): Node | undefined {
  for (const ai of kids(holder, REC_AI)) if (ai.length >= 1 && u8(state, ai, 0) === id) return ai;
  return undefined;
}

// --------------------------------------------------------------------- cache

/** The cached values by what they are of (1 values, 2 categories, 3 bubble sizes) and by series: point -> value. */
type Cache = Map<number, number | string>[][];

function dataOf(held: Map<number, number | string> | undefined, ref: string | undefined, stated: number, kind: 'num' | 'str' | undefined): ChartData {
  const index: number[] = [];
  if (held) held.forEach((_value, at) => index.push(at));
  index.sort((a, b) => a - b);
  let texts = false;
  const values: (number | string | null)[] = index.map((at) => {
    const value = (held as Map<number, number | string>).get(at) as number | string;
    if (typeof value === 'string') texts = true;
    return value;
  });
  const as = kind || (texts ? 'str' : 'num');
  for (let i = 0; i < values.length; i++) {
    const value = values[i];
    if (as === 'str' && typeof value === 'number') values[i] = String(value);
    else if (as === 'num' && typeof value !== 'number') values[i] = null;
  }
  const last = index.length > 0 ? index[index.length - 1] + 1 : 0;
  const data: ChartData = { kind: as, cache: { index, values }, count: clamp(Math.max(stated, last), 0, MAX_CHART_POINTS) };
  if (ref !== undefined) data.ref = ref;
  return data;
}

// -------------------------------------------------------------------- series

interface Formats {
  shape?: ChartShape;
  marker?: ChartMarker;
  explosion?: number;
  smooth?: boolean;
  invert?: boolean;
  labels?: ChartDataLabels;
}

/**
 * What a DataFormat block states: LineFormat, AreaFormat, PieFormat (pcExplode, 2), MarkerFormat
 * (rgbFore 4, rgbBack 4, imk 2, flags 2: bit 0 fAuto, 4 fNotShowInt, 5 fNotShowBrd; icvFore 2,
 * icvBack 2, miSize 4 in twips), SerFmt (flags 2: bit 0 fSmoothedLine), AttachedLabel (flags 2).
 * `filled`: the chart type fills its points (bars, areas, slices, bubbles).
 */
function formatsOf(state: State, block: Node, filled: boolean): Formats {
  const out: Formats = {};
  const shape: ChartShape = {};
  const area = areaOf(state, kid(block, REC_AREAFORMAT));
  if (filled && area.fill !== undefined) shape.fill = area.fill;
  if (area.invert) out.invert = true;
  const line = lineOf(state, kid(block, REC_LINEFORMAT));
  if (line) shape.line = line;
  if (shape.fill !== undefined || shape.line) out.shape = shape;
  if (kid(block, REC_GELFRAME)) note(state, 'A gradient or picture fill of a chart is not drawn');

  const pie = kid(block, REC_PIEFORMAT);
  if (pie) {
    need(pie, 2, 'PieFormat');
    const explosion = u16(state, pie, 0);
    if (explosion > 0) out.explosion = clamp(explosion, 0, 400);
  }
  const marker = kid(block, REC_MARKERFORMAT);
  if (marker) {
    need(marker, 12, 'MarkerFormat');
    const flags = u16(state, marker, 10);
    if (!(flags & 1)) {
      const type = u16(state, marker, 8);
      const made: ChartMarker = { symbol: type < MARKERS.length ? MARKERS[type] : 'square' };
      if (made.symbol !== 'none') {
        made.fill = flags & 0x10 ? null : rgb(state, marker, 4);
        made.line = { color: flags & 0x20 ? null : rgb(state, marker, 0) };
        if (marker.length >= 20) {
          const points = state.view.getUint32(marker.d + 16, true) / 20;
          if (points > 0) made.size = clamp(Math.round(points), 2, 72);
        }
      }
      out.marker = made;
    }
  }
  const format = kid(block, REC_SERFMT);
  if (format && format.length >= 2 && u16(state, format, 0) & 1) out.smooth = true;
  const attached = kid(block, REC_ATTACHEDLABEL);
  if (attached && attached.length >= 2) {
    // Bit 0 value, 1 per cent, 2 category and per cent, 4 category, 5 bubble size, 6 series name.
    const flags = u16(state, attached, 0);
    if (flags & 0x77) {
      const labels: ChartDataLabels = {};
      if (flags & 1) labels.showValue = true;
      if (flags & 6) labels.showPercent = true;
      if (flags & 0x14) labels.showCategory = true;
      if (flags & 0x20) labels.showBubbleSize = true;
      if (flags & 0x40) labels.showSeries = true;
      out.labels = labels;
    }
  }
  return out;
}

/** The data labels a Text record of a series or of a point states. */
function labelsOfText(state: State, text: Node, point: boolean): ChartDataLabels {
  need(text, 26, 'Text');
  const flags = u16(state, text, 24);
  if (flags & 0x40) return { deleted: true };
  const labels: ChartDataLabels = {};
  if (flags & 0x0004) labels.showValue = true;
  if (flags & 0x1800) labels.showPercent = true;
  if (flags & 0x4800) labels.showCategory = true;
  if (flags & 0x2000) labels.showBubbleSize = true;
  if (flags & 0x0002) labels.showLegendKey = true;
  const place = text.length >= 30 ? LABEL_PLACES[u16(state, text, 28) & 0xf] : undefined;
  if (place) labels.position = place;
  const style = textStyle(state, text);
  if (style) labels.style = style;
  // A label with words of its own (fAutoText is off): the words.
  if (point && !(flags & 0x10)) {
    const words = seriesText(state, kid(text, REC_SERIESTEXT));
    if (words !== undefined && words !== '') labels.text = words.slice(0, 1024);
  }
  return labels;
}

interface ReadSeries {
  series: ChartSeries;
  /** The number of the chart group the series names; undefined when it names none. */
  group: number | undefined;
  node: Node;
  /** Its place among all Series records of the chart: what the cache and the links of labels count by. */
  place: number;
}

// --------------------------------------------------------------------- chart

interface ReadGroup {
  group: ChartGroup;
  /** icrt of its ChartFormat. */
  number: number;
  node: Node;
}

/**
 * A chart group: ChartFormat (16 reserved bytes, flags 2: bit 0 fVaried, icrt 2) with one record
 * of a chart type inside.
 */
function readGroup(state: State, node: Node, axisIds: number[]): ReadGroup {
  need(node, 20, 'ChartFormat');
  const varied = (u16(state, node, 16) & 1) !== 0;
  const number = u16(state, node, 18);
  let type: Node | undefined;
  for (const child of node.children || NONE) {
    switch (child.id) {
      case REC_BAR:
      case REC_LINE:
      case REC_PIE:
      case REC_AREA:
      case REC_SCATTER:
      case REC_RADAR:
      case REC_RADARAREA:
        if (type) throw new Unreadable('a chart group states two chart types');
        type = child;
        break;
      case REC_SURF:
        throw new Unreadable('a surface chart is not drawn');
      case REC_BOPPOP:
        throw new Unreadable('a pie with a second plot (pie of pie, bar of pie) is not drawn');
      default:
    }
  }
  if (!type) throw new Unreadable('a chart group states no chart type');
  const group: ChartGroup = { kind: 'bar', grouping: 'standard', axisIds, series: [] };
  const grouping = (flags: number): ChartGroup['grouping'] => (flags & 2 ? 'percentStacked' : flags & 1 ? 'stacked' : 'standard');
  switch (type.id) {
    case REC_BAR: {
      // Bar: pcOverlap (2, signed), pcGap (2), flags (2: bit 0 fTranspose, 1 fStacked, 2 f100).
      need(type, 6, 'Bar');
      const flags = u16(state, type, 4);
      group.barDir = flags & 1 ? 'bar' : 'col';
      group.grouping = flags & 4 ? 'percentStacked' : flags & 2 ? 'stacked' : 'standard';
      group.gapWidth = clamp(u16(state, type, 2), 0, 500);
      // UNSURE: the record counts the other way round than a chart part does - a stacked chart of Excel states -100, where a chart part states 100.
      group.overlap = clamp(0 - i16(state, type, 0), -100, 100);
      break;
    }
    case REC_LINE:
      // Line: flags (2: bit 0 fStacked, 1 f100).
      need(type, 2, 'Line');
      group.kind = 'line';
      group.grouping = grouping(u16(state, type, 0));
      // UNSURE: a series that states no marker of its own (or an automatic one) has markers, as a line chart of Excel 2003 has by itself.
      group.markers = true;
      break;
    case REC_AREA:
      // Area: flags (2: bit 0 fStacked, 1 f100).
      need(type, 2, 'Area');
      group.kind = 'area';
      group.grouping = grouping(u16(state, type, 0));
      break;
    case REC_PIE: {
      // Pie: anStart (2), pcDonut (2), flags (2).
      need(type, 4, 'Pie');
      const hole = u16(state, type, 2);
      group.kind = hole > 0 ? 'doughnut' : 'pie';
      if (hole > 0) group.holeSize = clamp(hole, 10, 90);
      group.firstSliceAngle = u16(state, type, 0) % 360;
      break;
    }
    case REC_SCATTER: {
      // Scatter: pcBubbleSizeRatio (2), wBubbleSize (2: 1 area, 2 width), flags (2: bit 0 fBubbles, 1 fShowNegBubbles).
      const flags = type.length >= 6 ? u16(state, type, 4) : 0;
      if (flags & 1) {
        group.kind = 'bubble';
        group.bubbleScale = clamp(u16(state, type, 0), 0, 300);
        group.sizeRepresents = u16(state, type, 2) === 2 ? 'w' : 'area';
        if (flags & 2) group.showNegBubbles = true;
      } else {
        group.kind = 'scatter';
        group.scatterStyle = 'lineMarker';
      }
      break;
    }
    case REC_RADAR:
      group.kind = 'radar';
      group.radarStyle = 'marker';
      break;
    default:
      group.kind = 'radar';
      group.radarStyle = 'filled';
  }
  if (varied) group.varyColors = true;
  if (kid(node, REC_CHART3D)) {
    group.threeD = true;
    note(state, 'A 3-D chart is drawn flat');
  }
  // CrtLine: id (2: 0 drop lines, 1 high-low lines); the LineFormat after it is its look.
  const children = node.children || NONE;
  for (let i = 0; i < children.length; i++) {
    if (children[i].id !== REC_CRTLINE || children[i].length < 2) continue;
    const id = u16(state, children[i], 0);
    const line = (children[i + 1] && children[i + 1].id === REC_LINEFORMAT ? lineOf(state, children[i + 1]) : undefined) || {};
    if (id === 0) group.dropLines = line;
    else if (id === 1) group.hiLowLines = line;
  }
  // DropBar: pcGap (2); the first is the one of rising, the second the one of falling values.
  const bars = kids(node, REC_DROPBAR);
  if (bars.length > 0) {
    group.upDownBars = {};
    if (bars[0].length >= 2) group.upDownBars.gapWidth = clamp(u16(state, bars[0], 0), 0, 500);
    const up = frameShape(state, bars[0]);
    if (up) group.upDownBars.up = up;
    const down = frameShape(state, bars[1]);
    if (down) group.upDownBars.down = down;
  }
  return { group, number, node };
}

/** What an axis says about where the other axis of its pair crosses it. */
interface Crossing {
  max: boolean;
  at?: number;
  between?: boolean;
}

/**
 * An axis: Axis (wType 2: 0 category or horizontal, 1 value, 2 series; 16 reserved bytes) with
 * its range, its tick marks, its number format and its lines inside. `lying`: the bars of the
 * axis group lie, so that the categories stand at the side.
 */
function readAxis(state: State, node: Node, parent: number, lying: boolean): { axis: ChartAxis; crossing: Crossing; type: number } | undefined {
  need(node, 2, 'Axis');
  const type = u16(state, node, 0);
  if (type === 2) {
    note(state, 'The series axis of a chart is not drawn');
    return undefined;
  }
  if (type > 2) return undefined;
  const range = kid(node, REC_VALUERANGE);
  const upright = type === 1;
  const position: ChartAxis['position'] = upright !== lying ? (parent ? 'r' : 'l') : parent ? 't' : 'b';
  const axis: ChartAxis = { id: parent * 3 + type + 1, kind: upright || range ? 'val' : 'cat', position, tickLabels: 'nextTo' };
  const crossing: Crossing = { max: false };

  if (range) {
    // ValueRange: numMin, numMax, numMajor, numMinor, numCross (8 each), flags (2: bit 0 fAutoMin, 1 fAutoMax, 2 fAutoMajor,
    // 3 fAutoMinor, 4 fAutoCross, 5 fLog, 6 fReversed, 7 fMaxCross).
    need(range, 42, 'ValueRange');
    const flags = u16(state, range, 40);
    const log = (flags & 0x20) !== 0;
    // UNSURE: on a logarithmic axis the record holds the exponents of ten, not the numbers.
    const stated = (at: number): number | undefined => {
      const value = log ? Math.pow(10, f64(state, range, at)) : f64(state, range, at);
      return Number.isFinite(value) ? value : undefined;
    };
    if (log) axis.logBase = 10;
    if (!(flags & 1)) axis.min = stated(0);
    if (!(flags & 2)) axis.max = stated(8);
    if (axis.min === undefined) delete axis.min;
    if (axis.max === undefined) delete axis.max;
    if (!(flags & 4) && !log) {
      const major = f64(state, range, 16);
      if (Number.isFinite(major) && major > 0) axis.majorUnit = major;
    }
    if (flags & 0x40) axis.reversed = true;
    crossing.max = (flags & 0x80) !== 0;
    if (!(flags & 0x10)) crossing.at = stated(32);
  }
  const categories = kid(node, REC_CATSERRANGE);
  if (categories && axis.kind === 'cat') {
    // CatSerRange: catCross (2), catLabel (2), catMark (2), flags (2: bit 0 fBetween, 1 fMaxCross, 2 fReverse).
    need(categories, 8, 'CatSerRange');
    const flags = u16(state, categories, 6);
    if (flags & 4) axis.reversed = true;
    crossing.max = (flags & 2) !== 0;
    crossing.between = (flags & 1) !== 0;
    const at = u16(state, categories, 0);
    if (at > 1) crossing.at = at;
  }
  const dates = kid(node, REC_AXCEXT);
  if (dates && dates.length >= 18 && axis.kind === 'cat') {
    // AxcExt: catMin, catMax, catMajor, duMajor, catMinor, duMinor, duBase, catCrossDate (2 each), flags (2: bit 4 fDateAxis, 5 fAutoBase).
    const flags = u16(state, dates, 16);
    if (flags & 0x10) {
      axis.date = true;
      const base = u16(state, dates, 12);
      if (!(flags & 0x20) && base < TIME_UNITS.length) axis.baseTimeUnit = TIME_UNITS[base];
    }
  }
  const tick = kid(node, REC_TICK);
  if (tick) {
    // Tick: tktMajor (1), tktMinor (1), tlt (1: 0 none, 1 low, 2 high, 3 next to the axis), wBkgMode (1), rgb (4), 16 reserved
    // bytes, flags (2: bit 0 fAutoCo, bits 2..4 rot, bit 5 fAutoRot), icv (2), trot (2).
    need(tick, 3, 'Tick');
    const labels = u8(state, tick, 2);
    axis.tickLabels = labels === 0 ? 'none' : labels === 1 ? 'low' : labels === 2 ? 'high' : 'nextTo';
    if (tick.length >= 26) {
      const flags = u16(state, tick, 24);
      if (!(flags & 1)) axis.labelStyle = { color: rgb(state, tick, 4) };
      // UNSURE: rot 2 turns the labels a quarter counter-clockwise, 3 a quarter clockwise; other angles (trot) are not applied.
      const turn = (flags >> 2) & 7;
      if (!(flags & 0x20) && (turn === 2 || turn === 3)) axis.labelRotation = turn === 2 ? -90 : 90;
    }
  }
  const font = fontStyle(state, node);
  if (font) axis.labelStyle = { ...font, ...(axis.labelStyle || {}) };
  const format = kid(node, REC_IFMT);
  if (format && format.length >= 2) {
    // UNSURE: an axis that has the record states a format of its own; one without it takes that of the cells.
    axis.formatCode = formatCodeOf(state.context.styles, u16(state, format, 0)).slice(0, 255);
    axis.sourceLinked = false;
  }
  // AxisLine: id (2: 0 the axis itself, 1 major grid lines, 2 minor grid lines); the LineFormat after it is its look.
  const children = node.children || NONE;
  for (let i = 0; i < children.length; i++) {
    if (children[i].id !== REC_AXISLINE || children[i].length < 2) continue;
    const id = u16(state, children[i], 0);
    const format = children[i + 1] && children[i + 1].id === REC_LINEFORMAT ? children[i + 1] : undefined;
    const line = format ? lineOf(state, format) : undefined;
    if (id === 0) {
      if (line) axis.line = line;
      // Bit 2 of the flags of the line of the axis itself (fAxisOn): the axis is shown. Excel writes the second category
      // axis of a chart with two scales switched off this way. UNSURE: an axis without such a line is shown.
      if (format && !(u16(state, format, 8) & 4)) axis.deleted = true;
    } else if (id === 1) axis.majorGridlines = line || {};
    else if (id === 2) axis.minorGridlines = line || {};
  }
  return { axis, crossing, type };
}

/** A title: the words of a Text record (SeriesText), or the cell they are read from (AI); undefined for one that was taken away. */
function titleOf(state: State, text: Node): ChartTitle | undefined {
  need(text, 26, 'Text');
  if (u16(state, text, 24) & 0x40) return undefined;
  const title: ChartTitle = {};
  const words = seriesText(state, kid(text, REC_SERIESTEXT));
  if (words !== undefined && words !== '') title.text = words;
  const ref = reference(state, aiOf(state, text, 0));
  if (ref !== undefined) title.ref = ref;
  if (title.text === undefined && title.ref === undefined) return undefined;
  const style = textStyle(state, text);
  if (style) title.style = style;
  return title;
}

/** ObjectLink: wLinkObj (2: 1 chart title, 2 value axis title, 3 category axis title, 4 series or point), wLinkVar1 (2: series), wLinkVar2 (2: point). */
function linkOf(state: State, text: Node): { object: number; series: number; point: number } | undefined {
  const link = kid(text, REC_OBJECTLINK);
  if (!link || link.length < 6) return undefined;
  return { object: u16(state, link, 0), series: u16(state, link, 2), point: u16(state, link, 4) };
}

function interpret(state: State, chartNode: Node, cache: Cache): Chart {
  const top = chartNode.children || NONE;
  const { context } = state;

  // The axis groups with their axes and their chart groups.
  const groups: ReadGroup[] = [];
  const axes: ChartAxis[] = [];
  let plotArea: ChartShape | undefined;
  let legendNode: Node | undefined;
  const seenParents = new Set<number>();
  for (const parent of kids(chartNode, REC_AXISPARENT)) {
    // AxisParent: iax (2: 0 primary, 1 secondary), 16 bytes that are not used.
    need(parent, 2, 'AxisParent');
    const index = u16(state, parent, 0);
    if (index > 1 || seenParents.has(index)) throw new Unreadable('its axis groups are not a primary and a secondary one');
    seenParents.add(index);
    const formats = kids(parent, REC_CHARTFORMAT);
    if (groups.length + formats.length > 16) throw new Unreadable('it holds too many chart groups');
    // Bars that lie turn the axes of their group.
    const lying = formats.some((format) => {
      const bar = kid(format, REC_BAR);
      return bar !== undefined && bar.length >= 6 && (u16(state, bar, 4) & 1) !== 0;
    });
    const pair: ({ axis: ChartAxis; crossing: Crossing } | undefined)[] = [undefined, undefined];
    for (const node of kids(parent, REC_AXIS)) {
      const read = axes.length < MAX_CHART_AXES ? readAxis(state, node, index, lying) : undefined;
      if (!read) continue;
      if (pair[read.type]) throw new Unreadable('an axis group holds two axes of one kind');
      pair[read.type] = read;
      axes.push(read.axis);
    }
    const [across, up] = pair;
    if (across && up) {
      // Each axis says where the other one crosses it.
      across.axis.crossAxis = up.axis.id;
      up.axis.crossAxis = across.axis.id;
      if (across.crossing.max) up.axis.crosses = 'max';
      else if (across.crossing.at !== undefined) up.axis.crossesAt = across.crossing.at;
      if (across.crossing.between !== undefined) up.axis.crossBetween = across.crossing.between ? 'between' : 'midCat';
      if (up.crossing.max) across.axis.crosses = 'max';
      else if (up.crossing.at !== undefined) across.axis.crossesAt = up.crossing.at;
    }
    // The titles of the axes: Text records that say which axis they belong to.
    for (const text of kids(parent, REC_TEXT)) {
      const link = linkOf(state, text);
      const holder = link && link.object === 3 ? across : link && link.object === 2 ? up : undefined;
      if (!holder) continue;
      const title = titleOf(state, text);
      if (title) holder.axis.title = title;
    }
    // PlotArea is followed by the frame of the area the series are drawn in.
    const children = parent.children || NONE;
    for (let i = 0; i + 1 < children.length; i++) {
      if (children[i].id === REC_PLOTAREA && children[i + 1].id === REC_FRAME && index === 0) plotArea = frameShape(state, children[i + 1]);
    }
    const ids = [across, up].filter((entry) => entry !== undefined).map((entry) => (entry as { axis: ChartAxis }).axis.id);
    for (const format of formats) {
      groups.push(readGroup(state, format, ids));
      if (!legendNode) legendNode = kid(format, REC_LEGEND);
    }
  }
  if (groups.length === 0) throw new Unreadable('it holds no chart group');
  // UNSURE: a series names its chart group by the number the group states (icrt); in the files Excel writes that is also its place among the groups.
  const byNumber = new Map<number, ReadGroup>();
  let numbered = true;
  groups.forEach((entry) => {
    if (byNumber.has(entry.number)) numbered = false;
    byNumber.set(entry.number, entry);
  });

  // The series.
  const all: ReadSeries[] = [];
  const nodes = kids(chartNode, REC_SERIES);
  let auxiliary = 0;
  let beyond = 0;
  for (let place = 0; place < nodes.length; place++) {
    const node = nodes[place];
    // A series that is the trend line or the error bars of another one is no series of the chart.
    if (kid(node, REC_SERPARENT) || kid(node, REC_SERAUXTREND) || kid(node, REC_SERAUXERRBAR)) {
      auxiliary++;
      continue;
    }
    if (all.length >= MAX_CHART_SERIES) {
      beyond++;
      continue;
    }
    // Series: sdtX (2: 1 numbers, 3 texts), sdtY (2), cValx (2), cValy (2), sdtBSize (2), cValBSize (2).
    need(node, 8, 'Series');
    const link = kid(node, REC_SERTOCRT);
    let target: ReadGroup | undefined;
    let number: number | undefined;
    if (link) {
      need(link, 2, 'SerToCrt');
      number = u16(state, link, 0);
      target = numbered ? byNumber.get(number) : groups[number];
    } else if (groups.length === 1) target = groups[0];
    if (!target) throw new Unreadable('a series names a chart group the chart does not have');
    const kind = target.group.kind;
    const xy = kind === 'scatter' || kind === 'bubble';
    const filled = kind === 'bar' || kind === 'area' || kind === 'pie' || kind === 'doughnut' || kind === 'bubble' || (kind === 'radar' && target.group.radarStyle === 'filled');

    const series: ChartSeries = { index: all.length, order: all.length };
    const name: { ref?: string; text?: string } = {};
    const nameRef = reference(state, aiOf(state, node, 0));
    if (nameRef !== undefined) name.ref = nameRef;
    const words = seriesText(state, kid(node, REC_SERIESTEXT));
    if (words !== undefined && words !== '') name.text = words;
    if (name.ref !== undefined || name.text !== undefined) series.name = name;

    const held = (what: number): Map<number, number | string> | undefined => (cache[what] ? cache[what][place] : undefined);
    series.values = dataOf(held(1), reference(state, aiOf(state, node, 1)), u16(state, node, 6), 'num');
    const categories = reference(state, aiOf(state, node, 2));
    if (categories !== undefined || held(2)) {
      // The horizontal values of a scatter chart are numbers; categories are what the file says they are.
      series.categories = dataOf(held(2), categories, u16(state, node, 4), xy ? 'num' : held(2) ? undefined : u16(state, node, 0) === 3 ? 'str' : 'num');
    }
    if (kind === 'bubble') {
      const sizes = reference(state, aiOf(state, node, 3));
      if (sizes !== undefined || held(3)) series.bubbleSizes = dataOf(held(3), sizes, node.length >= 12 ? u16(state, node, 10) : 0, 'num');
    }

    // DataFormat: xi (2: the point, 0xFFFF for the whole series), yi (2: the series), iss (2: its number among the colours), flags (2).
    const points: ChartPoint[] = [];
    let whole = false;
    for (const block of kids(node, REC_DATAFORMAT)) {
      need(block, 6, 'DataFormat');
      const point = u16(state, block, 0);
      const formats = formatsOf(state, block, filled);
      if (point === 0xffff) {
        if (whole) continue;
        whole = true;
        series.index = u16(state, block, 4);
        if (formats.shape) series.shape = formats.shape;
        if (formats.marker) series.marker = formats.marker;
        if (formats.explosion !== undefined) series.explosion = formats.explosion;
        if (formats.smooth) series.smooth = true;
        if (formats.invert) series.invertIfNegative = true;
        if (formats.labels) series.dataLabels = formats.labels;
      } else if (points.length < MAX_CHART_DATA_POINTS && (formats.shape || formats.marker || formats.explosion !== undefined)) {
        const made: ChartPoint = { index: point };
        if (formats.shape) made.shape = formats.shape;
        if (formats.marker) made.marker = formats.marker;
        if (formats.explosion !== undefined) made.explosion = formats.explosion;
        points.push(made);
      }
    }
    if (points.length > 0) series.points = points.sort((a, b) => a.index - b.index);
    target.group.series.push(series);
    all.push({ series, group: number, node, place });
  }
  if (auxiliary > 0) note(state, 'Trend lines and error bars of a chart of this format are not drawn');
  if (beyond > 0) note(state, `Only the first ${MAX_CHART_SERIES} series of a chart are read`);
  const kept = groups.filter((entry) => entry.group.series.length > 0).map((entry) => entry.group);
  if (kept.length === 0) throw new Unreadable('it holds no series');

  // The title of the chart and the data labels: Text records that say what they belong to.
  let title: ChartTitle | undefined;
  const byPlace = new Map<number, ChartSeries>();
  for (const entry of all) byPlace.set(entry.place, entry.series);
  for (const text of kids(chartNode, REC_TEXT)) {
    const link = linkOf(state, text);
    if (!link) continue;
    if (link.object === 1) title = title || titleOf(state, text);
    else if (link.object === 4) {
      const series = byPlace.get(link.series);
      if (!series) continue;
      if (link.point === 0xffff) series.dataLabels = labelsOfText(state, text, false);
      else {
        const labels = (series.pointLabels = series.pointLabels || []);
        if (labels.length < MAX_CHART_DATA_POINTS) labels.push({ index: link.point, ...labelsOfText(state, text, true) } as ChartPointLabel);
      }
    }
  }
  for (const entry of all) if (entry.series.pointLabels) entry.series.pointLabels.sort((a, b) => a.index - b.index);

  // The colours series get that state none. UNSURE: Excel 2003 takes them from the palette of the workbook - the entries 24
  // and up for what is filled, 32 and up for lines; of the eight of each row the first six are taken, and one row for the whole chart.
  const first = kept[0].kind;
  const row = first === 'line' || first === 'scatter' || (first === 'radar' && kept[0].radarStyle !== 'filled') ? 32 : 24;
  const palette: string[] = [];
  for (let i = 0; i < 6; i++) palette.push(resolveColor({ indexed: row + i }, context.styles) || '#808080');

  const chart: Chart = { groups: kept, axes, plotVisOnly: true, blanks: 'gap', palette, notes: state.notes, shownOnly: true };
  // ShtProps: flags (2: bit 1 fPlotVisOnly), mdBlank (1: 0 gap, 1 zero, 2 span).
  const sheet = kid(chartNode, REC_SHTPROPS);
  if (sheet && sheet.length >= 3) {
    chart.plotVisOnly = (u16(state, sheet, 0) & 2) !== 0;
    const blanks = u8(state, sheet, 2);
    chart.blanks = blanks === 1 ? 'zero' : blanks === 2 ? 'span' : 'gap';
  }
  if (title) chart.title = title;
  // A chart of this format has the title it states and makes none up.
  else chart.autoTitleDeleted = true;
  if (context.date1904) chart.date1904 = true;
  // The first frame of the chart is that of the whole chart.
  const area = frameShape(state, kid(chartNode, REC_FRAME));
  if (area) chart.area = area;
  if (plotArea) chart.plotArea = plotArea;
  if (legendNode) {
    // Legend: x, y, dx, dy (4 each), wType (1: 0 bottom, 1 corner, 2 top, 3 right, 4 left, 7 not docked), wSpace (1), flags (2).
    need(legendNode, 17, 'Legend');
    const type = u8(state, legendNode, 16);
    const legend: ChartLegend = { position: type === 0 ? 'b' : type === 1 ? 'tr' : type === 2 ? 't' : type === 4 ? 'l' : 'r' };
    if (type > 4) note(state, 'A legend that was moved by hand stands at the right');
    const style = fontStyle(state, kid(legendNode, REC_TEXT));
    if (style) legend.style = style;
    chart.legend = legend;
  }
  // DefaultText: id (2); the Text after it is the default. UNSURE: 2 is the one of all text of the chart.
  for (let i = 0; i + 1 < top.length; i++) {
    if (top[i].id !== REC_DEFAULTTEXT || top[i].length < 2 || top[i + 1].id !== REC_TEXT || u16(state, top[i], 0) !== 2) continue;
    const style = fontStyle(state, top[i + 1]);
    if (style) chart.textStyle = style;
    break;
  }
  if (kid(chartNode, REC_DAT)) note(state, 'The data table of a chart is not drawn');
  if (state.later > 0) note(state, 'Formatting that later versions of Excel added to the chart is not applied');
  if (state.unknown.size > 0) {
    const ids: string[] = [];
    state.unknown.forEach((count, id) => {
      if (ids.length < 8) ids.push(`0x${(id < 0x1000 ? '0' : '') + id.toString(16).toUpperCase()}${count > 1 ? ` (${count})` : ''}`);
    });
    note(state, `Records of the chart that are not read: ${ids.join(', ')}${state.unknown.size > ids.length ? ', ...' : ''}`);
  }
  return chart;
}

function read(stream: RecordStream, start: number, context: XlsChartContext, span: XlsSpan): XlsChartRead {
  const view = stream.view;
  const size = stream.bytes.length;
  if (start < 0 || start + 8 > size || view.getUint16(start, true) !== REC_BOF || view.getUint16(start + 2, true) < 4 || view.getUint16(start + 6, true) !== 0x0020) {
    throw new Unreadable('it does not start as a chart does');
  }
  if (context.budget.charts <= 0) throw new Unreadable('the workbook holds too many charts');
  context.budget.charts--;
  const state: State = { stream, view, context, notes: [], records: 0, span, unknown: new Map(), later: 0 };
  if (start + 4 + view.getUint16(start + 2, true) > span.limit) throw new Unreadable(OVERLAP);
  const cache: Cache = [];
  let chartNode: Node | undefined;
  /** What the cell records that follow are of: see SIIndex. */
  let what = 0;
  let cached = 0;
  let cut = false;
  let depth = 0;
  let p = start + 4 + view.getUint16(start + 2, true);
  for (;;) {
    if (p + 4 > size) throw new Unreadable('its records are not terminated');
    const id = view.getUint16(p, true);
    const length = view.getUint16(p + 2, true);
    const d = p + 4;
    const next = d + length;
    if (next > size) throw new Unreadable('a record runs past the end of the stream');
    if (next > span.limit) throw new Unreadable(OVERLAP);
    const offset = p;
    p = next;
    span.end = p;
    // A substream inside the chart (there should be none) is passed over.
    if (depth > 0) {
      if (id === REC_BOF) depth++;
      else if (id === REC_EOF) depth--;
      continue;
    }
    if (id === REC_EOF) break;
    if (id === REC_BOF) {
      depth++;
      continue;
    }
    if (id === REC_BEGIN || id === REC_END) throw new Unreadable('a block of its records belongs to no record');
    if (id === REC_CHART) {
      if (chartNode) throw new Unreadable('it holds two Chart records');
      chartNode = { id, d, length };
      if (p + 4 > size || view.getUint16(p, true) !== REC_BEGIN) throw new Unreadable('its Chart record holds nothing');
      const block = readBlock(state, p + 4 + view.getUint16(p + 2, true), 1);
      chartNode.children = block.nodes;
      p = block.next;
      continue;
    }
    if (!chartNode) continue;
    if (id === REC_SIINDEX) {
      // SIIndex: numIndex (2: 1 values, 2 categories, 3 bubble sizes).
      what = length >= 2 ? view.getUint16(d, true) : 0;
      continue;
    }
    if (id !== REC_NUMBER && id !== REC_LABEL && id !== REC_BOOLERR && id !== REC_BLANK) continue;
    // UNSURE: of a cell record behind SIIndex the row is the point and the column the series.
    if (what < 1 || what > 3 || length < 6 || id === REC_BLANK || id === REC_BOOLERR) continue;
    const point = view.getUint16(d, true);
    const series = view.getUint16(d + 2, true);
    if (point >= MAX_CHART_POINTS) continue;
    let value: number | string | undefined;
    if (id === REC_NUMBER) {
      if (length < 14) throw new Unreadable('a cached number is too short');
      const number = view.getFloat64(d + 6, true);
      if (Number.isFinite(number)) value = number;
    } else {
      const reader = readerAt(stream, offset);
      reader.skip(6);
      value = readString16(reader).text.slice(0, 32767);
    }
    if (value === undefined) continue;
    if (cached >= MAX_CHART_CACHE) {
      cut = true;
      continue;
    }
    // A chart that does not fit what the other charts of the workbook left is not half a chart: it is not shown.
    if (context.budget.points <= 0) throw new Unreadable('the charts of the workbook hold too many cached values');
    const cost = typeof value === 'string' ? 16 + 2 * value.length : 16;
    if (context.budget.bytes < cost) throw new Unreadable('the charts of the workbook hold too many bytes of cached values');
    context.budget.bytes -= cost;
    context.budget.points--;
    cached++;
    const row = (cache[what] = cache[what] || []);
    (row[series] = row[series] || new Map()).set(point, value);
  }
  if (!chartNode) throw new Unreadable('it holds no Chart record');
  // Chart: x, y, dx, dy, 16.16 fixed point numbers of points.
  need(chartNode, 16, 'Chart');
  const chart = interpret(state, chartNode, cache);
  if (cut) note(state, `Only the first ${MAX_CHART_CACHE} cached values of a chart are read`);
  const out: XlsChartRead = { chart };
  const width = view.getInt32(chartNode.d + 8, true) / 65536;
  const height = view.getInt32(chartNode.d + 12, true) / 65536;
  if (width > 0 && height > 0 && width < 100000 && height < 100000) {
    out.width = width;
    out.height = height;
  }
  return out;
}

/**
 * Reads the chart whose substream starts at `start` (its BOF record). Never throws: a chart that
 * cannot be read for certain comes back as a `problem`, and is shown as a placeholder. `span`
 * says how far the substream may reach and is told how far it was read; without it the
 * substream may reach to the end of the stream.
 */
export function readXlsChart(stream: RecordStream, start: number, context: XlsChartContext, span?: XlsSpan): XlsChartRead {
  const within = span || { limit: stream.bytes.length, end: 0 };
  within.end = Math.max(start, 0);
  try {
    return read(stream, start, context, within);
  } catch (error) {
    return { problem: error instanceof Unreadable ? error.message : 'its records cannot be read' };
  }
}

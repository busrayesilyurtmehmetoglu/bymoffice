// The records of a BIFF8 chart, built field by field: the substream of a chart
// (BOF with dt = 0x0020 ... EOF) as it stands after the OBJ record of an
// embedded chart, or as the substream of a chart sheet.
//
// NOT CHECKED AGAINST A FILE OF EXCEL: no tool on the machine writes these
// records, so every layout below is the one of the specification of the
// format ([MS-XLS] 2.4.x) as far as it is remembered. Each builder states the
// fields in the order they are written. Where the reader depends on a field
// whose place or meaning is not certain, the reader says UNSURE.

import { cat, f64, rec, str8, u16, u32, type Bytes } from './xls-fixtures';

export const REC = {
  units: 0x1001,
  chart: 0x1002,
  series: 0x1003,
  dataFormat: 0x1006,
  lineFormat: 0x1007,
  markerFormat: 0x1009,
  areaFormat: 0x100a,
  pieFormat: 0x100b,
  attachedLabel: 0x100c,
  seriesText: 0x100d,
  chartFormat: 0x1014,
  legend: 0x1015,
  seriesList: 0x1016,
  bar: 0x1017,
  line: 0x1018,
  pie: 0x1019,
  area: 0x101a,
  scatter: 0x101b,
  crtLine: 0x101c,
  axis: 0x101d,
  tick: 0x101e,
  valueRange: 0x101f,
  catSerRange: 0x1020,
  axisLine: 0x1021,
  defaultText: 0x1024,
  text: 0x1025,
  fontX: 0x1026,
  objectLink: 0x1027,
  frame: 0x1032,
  begin: 0x1033,
  end: 0x1034,
  plotArea: 0x1035,
  chart3d: 0x103a,
  dropBar: 0x103d,
  radar: 0x103e,
  surf: 0x103f,
  radarArea: 0x1040,
  axisParent: 0x1041,
  shtProps: 0x1044,
  serToCrt: 0x1045,
  axesUsed: 0x1046,
  serParent: 0x104a,
  serAuxTrend: 0x104b,
  ifmt: 0x104e,
  pos: 0x104f,
  ai: 0x1051,
  serFmt: 0x105d,
  bopPop: 0x1061,
  axcExt: 0x1062,
  dat: 0x1063,
  siIndex: 0x1065,
};

export const begin = (): Uint8Array => rec(REC.begin);
export const end = (): Uint8Array => rec(REC.end);
/** A record with the records that belong to it: record, Begin, children, End. */
export const block = (head: Bytes, ...children: Bytes[]): Uint8Array => cat(head, begin(), ...children, end());

/** LongRGB: red, green, blue, one reserved byte. */
export const rgb = (hex: string): number[] => [parseInt(hex.slice(0, 2), 16), parseInt(hex.slice(2, 4), 16), parseInt(hex.slice(4, 6), 16), 0];
/** A 16.16 fixed point number. */
const fixed = (value: number): number[] => u32(Math.round(value * 65536));
const i32 = (value: number): number[] => u32(value >>> 0);

/** Chart: x, y, dx, dy of the chart in points, 16.16 fixed point each. */
export const chartRec = (width = 400, height = 300): Uint8Array => rec(REC.chart, fixed(0), fixed(0), fixed(width), fixed(height));

/** Series: sdtX (1 numbers, 3 texts), sdtY (1), cValx, cValy, sdtBSize (1), cValBSize. */
export const seriesRec = (categories: number, values: number, textCategories = true, sizes = 0): Uint8Array =>
  rec(REC.series, u16(textCategories ? 3 : 1), u16(1), u16(categories), u16(values), u16(1), u16(sizes));

/** AI (BRAI): id (0 name, 1 values, 2 categories, 3 bubble sizes), rt (0 default, 1 literal, 2 reference), flags, ifmt, cce, rgce. */
export const ai = (id: number, rt: number, tokens: number[] = [], ifmt = 0, flags = 0): Uint8Array => rec(REC.ai, [id, rt], u16(flags), u16(ifmt), u16(tokens.length), tokens);

/** SeriesText: reserved (2 bytes), cch (1 byte), then the characters behind a byte that says how wide they are. */
export const seriesText = (text: string): Uint8Array => rec(REC.seriesText, u16(0), str8(text));

/** DataFormat: xi (the point; 0xFFFF for the whole series), yi (the series), iss (its number among the colours), flags. */
export const dataFormat = (point = 0xffff, series = 0, number = series): Uint8Array => rec(REC.dataFormat, u16(point), u16(series), u16(number), u16(0));

export interface LineSpec {
  color?: string;
  /** lns: 0 solid, 1 dash, 2 dot, 3 dash-dot, 4 dash-dot-dot, 5 none. */
  style?: number;
  /** we: -1 hairline, 0 narrow, 1 medium, 2 wide. */
  weight?: number;
  auto?: boolean;
  /** fAxisOn: of the line of an axis (AxisLine 0), that the axis is shown. */
  axisOn?: boolean;
}
/** LineFormat: rgb, lns, we, flags (bit 0 fAuto, bit 2 fAxisOn), icv. */
export const lineFormat = (spec: LineSpec = {}): Uint8Array =>
  rec(REC.lineFormat, rgb(spec.color || '000000'), u16(spec.style || 0), u16((spec.weight === undefined ? 0 : spec.weight) & 0xffff), u16((spec.auto ? 1 : 0) | (spec.axisOn ? 4 : 0)), u16(0x4d));

export interface AreaSpec {
  color?: string;
  /** fls: 0 no fill, 1 solid. */
  pattern?: number;
  auto?: boolean;
  invert?: boolean;
}
/** AreaFormat: rgbFore, rgbBack, fls, flags (bit 0 fAuto, bit 1 fInvertNeg), icvFore, icvBack. */
export const areaFormat = (spec: AreaSpec = {}): Uint8Array =>
  rec(REC.areaFormat, rgb(spec.color || 'FFFFFF'), rgb('FFFFFF'), u16(spec.pattern === undefined ? 1 : spec.pattern), u16((spec.auto ? 1 : 0) | (spec.invert ? 2 : 0)), u16(0x4e), u16(0x4d));

export interface MarkerSpec {
  /** imk: 0 none, 1 square, 2 diamond, 3 triangle, 4 x, 5 star, 6 short bar, 7 long bar, 8 circle, 9 plus. */
  type?: number;
  line?: string;
  fill?: string;
  /** Twips: 20 to a point. */
  size?: number;
  auto?: boolean;
  noFill?: boolean;
  noLine?: boolean;
}
/** MarkerFormat: rgbFore, rgbBack, imk, flags (bit 0 fAuto, bit 4 fNotShowInt, bit 5 fNotShowBrd), icvFore, icvBack, miSize. */
export const markerFormat = (spec: MarkerSpec = {}): Uint8Array =>
  rec(
    REC.markerFormat,
    rgb(spec.line || '000000'),
    rgb(spec.fill || '000000'),
    u16(spec.type === undefined ? 1 : spec.type),
    u16((spec.auto ? 1 : 0) | (spec.noFill ? 0x10 : 0) | (spec.noLine ? 0x20 : 0)),
    u16(0x4d),
    u16(0x4d),
    u32(spec.size === undefined ? 100 : spec.size),
  );

/** PieFormat: pcExplode, how far a slice is pulled out, per cent of the radius. */
export const pieFormat = (explode: number): Uint8Array => rec(REC.pieFormat, u16(explode));
/** AttachedLabel: flags (bit 0 value, 1 per cent, 2 category and per cent, 4 category, 5 bubble size, 6 series name). */
export const attachedLabel = (flags: number): Uint8Array => rec(REC.attachedLabel, u16(flags));
/** SerFmt: flags (bit 0 fSmoothedLine). */
export const serFmt = (flags: number): Uint8Array => rec(REC.serFmt, u16(flags));
/** SerToCrt: the chart group of the series. */
export const serToCrt = (group: number): Uint8Array => rec(REC.serToCrt, u16(group));
/** SerParent: the series (counted from 1) this one is a trend line or an error bar of. */
export const serParent = (series: number): Uint8Array => rec(REC.serParent, u16(series));
/** SerAuxTrend: regt, ordUser, numIntercept, fEquation, fRSquared, numForecast, numBackcast. */
export const serAuxTrend = (): Uint8Array => rec(REC.serAuxTrend, [0, 1], f64(0), [0, 0], f64(0), f64(0));

/** ShtProps: flags (bit 1 fPlotVisOnly), mdBlank (0 gap, 1 zero, 2 span), reserved. */
export const shtProps = (visibleOnly = true, blanks = 0): Uint8Array => rec(REC.shtProps, u16(visibleOnly ? 2 : 0), [blanks, 0]);

/** AxisParent: iax (0 primary, 1 secondary), 16 bytes that are not used. */
export const axisParentRec = (index = 0): Uint8Array => rec(REC.axisParent, u16(index), new Array(16).fill(0));
/** Axis: wType (0 category or horizontal, 1 value, 2 series), 16 reserved bytes. */
export const axisRec = (type: number): Uint8Array => rec(REC.axis, u16(type), new Array(16).fill(0));

export interface ValueRangeSpec {
  min?: number;
  max?: number;
  major?: number;
  cross?: number;
  log?: boolean;
  reversed?: boolean;
  maxCross?: boolean;
}
/** ValueRange: numMin, numMax, numMajor, numMinor, numCross, flags (bits 0..4 the automatic ones, 5 fLog, 6 fReversed, 7 fMaxCross). */
export const valueRange = (spec: ValueRangeSpec = {}): Uint8Array =>
  rec(
    REC.valueRange,
    f64(spec.min || 0),
    f64(spec.max || 0),
    f64(spec.major || 0),
    f64(0),
    f64(spec.cross || 0),
    u16(
      (spec.min === undefined ? 1 : 0) |
        (spec.max === undefined ? 2 : 0) |
        (spec.major === undefined ? 4 : 0) |
        8 |
        (spec.cross === undefined ? 0x10 : 0) |
        (spec.log ? 0x20 : 0) |
        (spec.reversed ? 0x40 : 0) |
        (spec.maxCross ? 0x80 : 0) |
        0x100,
    ),
  );

/** CatSerRange: catCross, catLabel, catMark, flags (bit 0 fBetween, bit 1 fMaxCross, bit 2 fReverse). */
export const catSerRange = (flags = 1, cross = 1): Uint8Array => rec(REC.catSerRange, u16(cross), u16(1), u16(1), u16(flags));

/** AxcExt: catMin, catMax, catMajor, duMajor, catMinor, duMinor, duBase, catCrossDate, flags (bit 4 fDateAxis, bit 5 fAutoBase). */
export const axcExt = (flags: number, base = 0): Uint8Array => rec(REC.axcExt, u16(0), u16(0), u16(1), u16(0), u16(1), u16(0), u16(base), u16(0), u16(flags));

export interface TickSpec {
  /** tlt: 0 none, 1 low, 2 high, 3 next to the axis. */
  labels?: number;
  color?: string;
  /** rot: 0 none, 1 stacked, 2 a quarter turn counter-clockwise, 3 a quarter turn clockwise. */
  rotation?: number;
}
/** Tick: tktMajor, tktMinor, tlt, wBkgMode, rgb, 16 reserved bytes, flags (bit 0 fAutoCo, bits 2..4 rot, bit 5 fAutoRot), icv, trot. */
export const tick = (spec: TickSpec = {}): Uint8Array =>
  rec(
    REC.tick,
    [2, 0, spec.labels === undefined ? 3 : spec.labels, 1],
    rgb(spec.color || '000000'),
    new Array(16).fill(0),
    u16((spec.color ? 0 : 1) | 2 | ((spec.rotation || 0) << 2) | (spec.rotation === undefined ? 0x20 : 0)),
    u16(0x4d),
    u16(0),
  );

/** AxisLine: id (0 the axis itself, 1 major grid lines, 2 minor grid lines); the LineFormat after it is its look. */
export const axisLine = (id: number): Uint8Array => rec(REC.axisLine, u16(id));
/** IFmtRecord: the number format of the labels of an axis. */
export const ifmt = (format: number): Uint8Array => rec(REC.ifmt, u16(format));

/** ChartFormat: 16 reserved bytes, flags (bit 0 fVaried), icrt (the number of the chart group). */
export const chartFormatRec = (order = 0, varied = false): Uint8Array => rec(REC.chartFormat, new Array(16).fill(0), u16(varied ? 1 : 0), u16(order));

/** Bar: pcOverlap (signed), pcGap, flags (bit 0 fTranspose, bit 1 fStacked, bit 2 f100). */
export const bar = (flags = 0, gap = 150, overlap = 0): Uint8Array => rec(REC.bar, u16(overlap & 0xffff), u16(gap), u16(flags));
/** Line: flags (bit 0 fStacked, bit 1 f100). */
export const line = (flags = 0): Uint8Array => rec(REC.line, u16(flags));
/** Pie: anStart, pcDonut, flags. */
export const pie = (start = 0, donut = 0): Uint8Array => rec(REC.pie, u16(start), u16(donut), u16(0));
/** Area: flags (bit 0 fStacked, bit 1 f100). */
export const area = (flags = 0): Uint8Array => rec(REC.area, u16(flags));
/** Scatter: pcBubbleSizeRatio, wBubbleSize (1 area, 2 width), flags (bit 0 fBubbles, bit 1 fShowNegBubbles). */
export const scatter = (flags = 0, ratio = 100, size = 1): Uint8Array => rec(REC.scatter, u16(ratio), u16(size), u16(flags));
/** Radar and RadarArea: flags, two bytes that are not used. */
export const radar = (): Uint8Array => rec(REC.radar, u16(0), u16(0));
export const radarArea = (): Uint8Array => rec(REC.radarArea, u16(0), u16(0));
/** Surf: flags. */
export const surf = (): Uint8Array => rec(REC.surf, u16(0));
/** Chart3d: anRot, anElev, pcDist, pcHeight, pcDepth, pcGap, flags. */
export const chart3d = (): Uint8Array => rec(REC.chart3d, u16(20), u16(15), u16(30), u16(100), u16(100), u16(150), u16(0x14));
/** CrtLine: id (0 drop lines, 1 high-low lines); the LineFormat after it is its look. */
export const crtLine = (id: number): Uint8Array => rec(REC.crtLine, u16(id));
/** DropBar: pcGap. */
export const dropBar = (gap = 150): Uint8Array => rec(REC.dropBar, u16(gap));

/** Legend: x, y, dx, dy, wType (0 bottom, 1 corner, 2 top, 3 right, 4 left, 7 not docked), wSpace, flags. */
export const legendRec = (type = 3): Uint8Array => rec(REC.legend, u32(0), u32(0), u32(0), u32(0), [type, 1], u16(0x1f));

export interface TextSpec {
  color?: string;
  /** Bits of the flags beside fAutoColor: 0x4 value, 0x10 fAutoText, 0x40 fDeleted, 0x800 category and per cent, 0x1000 per cent, 0x2000 bubble size, 0x4000 category. */
  flags?: number;
  /** dlp: where a data label stands (1 outside end, 2 inside end, 3 centre, 4 inside base, 5 above, 6 below, 7 left, 8 right, 9 best fit). */
  position?: number;
}
/** Text: at, vat, wBkgMode, rgbText, x, y, dx, dy, flags, icvText, dlp and more, trot. */
export const textRec = (spec: TextSpec = {}): Uint8Array =>
  rec(REC.text, [2, 2], u16(1), rgb(spec.color || '000000'), i32(0), i32(0), i32(0), i32(0), u16((spec.color ? 0 : 1) | (spec.flags || 0)), u16(0x4d), u16(spec.position || 0), u16(0));

/** FontX: the font of the text, an index into the fonts of the workbook. */
export const fontX = (font: number): Uint8Array => rec(REC.fontX, u16(font));
/** ObjectLink: wLinkObj (1 chart title, 2 value axis, 3 category axis, 4 series or point), wLinkVar1 (series), wLinkVar2 (point). */
export const objectLink = (object: number, series = 0, point = 0): Uint8Array => rec(REC.objectLink, u16(object), u16(series), u16(point));
/** DefaultText: id; the Text after it is the default. */
export const defaultText = (id: number): Uint8Array => rec(REC.defaultText, u16(id));
/** Frame: frt, flags. */
export const frameRec = (): Uint8Array => rec(REC.frame, u16(0), u16(3));
/** A frame with its look. */
export const frame = (lineSpec?: LineSpec, areaSpec?: AreaSpec): Uint8Array => block(frameRec(), lineFormat(lineSpec || { auto: true }), areaFormat(areaSpec || { auto: true }));
export const plotArea = (): Uint8Array => rec(REC.plotArea);
/** SIIndex: numIndex (1 values, 2 categories, 3 bubble sizes); the cell records after it are the cache. */
export const siIndex = (index: number): Uint8Array => rec(REC.siIndex, u16(index));
/** A cached number: a Number record whose row is the point and whose column is the series. */
export const cachedNumber = (point: number, series: number, value: number): Uint8Array => rec(0x0203, u16(point), u16(series), u16(0), f64(value));
/** A cached text: a Label record (row, column, XF, then the text with a count of two bytes). */
export const cachedLabel = (point: number, series: number, text: string): Uint8Array => rec(0x0204, u16(point), u16(series), u16(0), u16(text.length), [0], Array.from(text, (c) => c.charCodeAt(0)));

/** A title or a label: Text with its font, its words and what it belongs to. */
export function textBlock(link: Bytes, words: string | number[] | undefined, spec: TextSpec & { font?: number } = {}): Uint8Array {
  const children: Bytes[] = [];
  if (spec.font !== undefined) children.push(fontX(spec.font));
  if (typeof words === 'string') children.push(ai(0, 1), seriesText(words));
  else if (words) children.push(ai(0, 2, words));
  else children.push(ai(0, 1));
  children.push(link);
  return block(textRec(spec), ...children);
}

export interface XlsSeriesSpec {
  /** The name: a text, or the tokens of a reference. */
  name?: string | number[];
  /** Tokens of the reference of the values; without them the values are literal (the cache alone). */
  values?: number[];
  categories?: number[];
  sizes?: number[];
  /** cValx and cValy of the Series record. Default 3 and 3. */
  counts?: [categories: number, values: number];
  /** The categories are numbers (sdtX = 1). */
  numericCategories?: boolean;
  /** The chart group the series belongs to. Default 0; null leaves the SerToCrt record out. */
  group?: number | null;
  /** iss of the DataFormat of the series. Default: its place. */
  number?: number;
  /** Records inside the DataFormat of the whole series. */
  format?: Bytes[];
  /** DataFormat blocks of single points: [point, records]. */
  points?: [point: number, records: Bytes[]][];
  /** Records after the SerToCrt record (SerParent, SerAuxTrend). */
  tail?: Bytes[];
}

export function seriesBlock(spec: XlsSeriesSpec, index: number): Uint8Array {
  const counts = spec.counts || [3, 3];
  const children: Bytes[] = [];
  if (typeof spec.name === 'string') children.push(ai(0, 1), seriesText(spec.name));
  else children.push(ai(0, spec.name ? 2 : 1, spec.name || []));
  children.push(ai(1, spec.values ? 2 : 1, spec.values || []));
  children.push(ai(2, spec.categories ? 2 : 0, spec.categories || []));
  children.push(ai(3, spec.sizes ? 2 : 1, spec.sizes || []));
  children.push(block(dataFormat(0xffff, index, spec.number === undefined ? index : spec.number), ...(spec.format || [])));
  for (const [point, records] of spec.points || []) children.push(block(dataFormat(point, index, spec.number === undefined ? index : spec.number), ...records));
  if (spec.group !== null) children.push(serToCrt(spec.group || 0));
  children.push(...(spec.tail || []));
  return block(seriesRec(counts[0], counts[1], !spec.numericCategories, spec.sizes ? counts[1] : 0), ...children);
}

export interface XlsGroupSpec {
  /** The record of the chart type: bar(), line(), pie(), ... */
  type: Bytes;
  varied?: boolean;
  /** icrt of the ChartFormat. Default: its place among all groups. */
  order?: number;
  /** Records after the type record (Chart3d, Legend, CrtLine, DropBar). */
  children?: Bytes[];
}

export interface XlsAxesSpec {
  /** Records inside the AxisParent before the chart groups: axes, their titles, the plot area. Default: a category and a value axis. */
  axes?: Bytes[];
  groups: XlsGroupSpec[];
}

export interface XlsChartSpec {
  series: XlsSeriesSpec[];
  /** The chart groups of the primary axes; several entries of XlsAxesSpec in `parents` instead give a secondary set. */
  groups?: XlsGroupSpec[];
  axes?: Bytes[];
  parents?: XlsAxesSpec[];
  /** Records inside the Chart block before the series (the frame of the chart). */
  head?: Bytes[];
  /** ShtProps. Default: only visible cells, gaps for blanks. */
  sheet?: Bytes;
  /** Records inside the Chart block after the axes: the title, data labels. */
  tail?: Bytes[];
  /** Records after the Chart block: the cache (SIIndex and cell records). */
  cache?: Bytes[];
  /** Records between BOF and the Chart record. */
  before?: Bytes[];
  /** The width and the height of the chart in points. */
  size?: [number, number];
}

/** A category axis and a value axis with what Excel writes for them. */
export const defaultAxes = (): Bytes[] => [
  block(axisRec(0), catSerRange(), axcExt(0xef), tick()),
  block(axisRec(1), valueRange(), tick(), axisLine(1), lineFormat({ color: '000000', weight: -1 })),
  plotArea(),
  frame({ color: '808080' }, { color: 'C0C0C0' }),
];

export function axisParentBlock(index: number, spec: XlsAxesSpec, firstGroup: number): Uint8Array {
  const groups = spec.groups.map((group, i) => block(chartFormatRec(group.order === undefined ? firstGroup + i : group.order, group.varied), group.type, ...(group.children || [])));
  return block(axisParentRec(index), rec(REC.pos, u16(2), u16(2), u32(0), u32(0), u32(4000), u32(4000)), ...(spec.axes || defaultAxes()), ...groups);
}

/** The records of a chart between BOF and EOF. */
export function chartRecords(spec: XlsChartSpec): Uint8Array {
  const parents = spec.parents || [{ axes: spec.axes, groups: spec.groups || [{ type: bar() }] }];
  let first = 0;
  const axes = parents.map((parent, i) => {
    const out = axisParentBlock(i, parent, first);
    first += parent.groups.length;
    return out;
  });
  const size = spec.size || [400, 300];
  return cat(
    ...(spec.before || []),
    rec(REC.units, u16(0)),
    block(
      chartRec(size[0], size[1]),
      ...(spec.head || [frame({ color: '000000' }, { color: 'FFFFFF' })]),
      ...spec.series.map((series, i) => seriesBlock(series, i)),
      spec.sheet || shtProps(),
      rec(REC.axesUsed, u16(parents.length)),
      ...axes,
      ...(spec.tail || []),
    ),
    ...(spec.cache || []),
  );
}

/** The substream of a chart: BOF (dt = 0x0020), the records, EOF. */
export function chartSubstream(spec: XlsChartSpec): Uint8Array {
  return cat(rec(0x0809, u16(0x0600), u16(0x0020), u16(0x0dbb), u16(0x07cc), u32(0), u32(6)), chartRecords(spec), rec(0x000a));
}

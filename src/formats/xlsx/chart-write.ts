// Writes charts: the part of a chart that is in no file (xl/charts/chartN.xml), the changes
// of a chart of the file, and the frame of a chart in a drawing part.
//
// A chart the engine makes is plain: one <c:chartSpace> with what the model holds and
// nothing else, so that reading the part gives the model back (the cached values aside:
// those are the cells as they are when the file is written). chartWrittenHere() tells such
// a part when it is read again; a chart whose part is one is written anew from its model
// whatever is changed in it.
//
// Any other chart of the file is changed in place (patchChartPart()): the elements the
// change is about are rewritten, the cached values of the series are brought up to date,
// and every other node of the part (extensions, mc:AlternateContent, text properties,
// formatting, elements nobody reads) stays as it is.

import { formulaNamesSheetsUnread } from '../../calc/refs';
import { cellDisplayText, isoToSerial } from '../../format/cell-format';
import {
  chartKindWritten,
  chartPartProblem,
  chartRefAreas,
  chartRefs,
  chartWritable,
  MAX_CHART_POINTS,
  MAX_CHART_SERIES,
  type Chart,
  type ChartAxis,
  type ChartCache,
  type ChartData,
  type ChartDataLabels,
  type ChartGroup,
  type ChartLegend,
  type ChartSeries,
  type ChartShape,
  type ChartTitle,
} from '../../model/sheet/chart';
import { resolveCellStyle } from '../../model/sheet/color';
import type { Cell, ChartOrigin, DrawingAnchor, DrawingMarker, Workbook } from '../../model/sheet/types';
import { localName, parseXml, serializeXml, type XmlDocument, type XmlElement, type XmlNode } from '../../xml';

export const CHART_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
export const DRAWING_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawing+xml';
const NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const EMU_PER_PIXEL = 9525;
const COLOR = /^#[0-9A-Fa-f]{6}$/;
/** Elements of one chart part that are looked at when it is compared or counted; a part with more is not known. */
const MAX_PART_ELEMENTS = 2000000;
/**
 * How deep the elements of a part may be nested for the part to be changed in place. The
 * functions that change a part walk it by calling themselves; a part nested deeper is read and
 * shown, and what it holds is left as it is. What applications write is nested some twenty deep.
 */
export const MAX_PATCH_DEPTH = 128;

/** Whether elements of a tree are nested deeper than `limit` (the root is at depth 0). Does not call itself. */
export function nestedBeyond(root: XmlElement, limit: number): boolean {
  let level: XmlElement[] = [root];
  for (let depth = 0; level.length > 0; depth++) {
    if (depth > limit) return true;
    const next: XmlElement[] = [];
    for (const el of level) for (const node of el.children) if (node.type === 'element') next.push(node);
    level = next;
  }
  return false;
}

const TEXT_ESCAPES: Record<string, string> = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '\r': '&#13;' };
const ATTR_ESCAPES: Record<string, string> = { '&': '&amp;', '<': '&lt;', '"': '&quot;', '\t': '&#9;', '\n': '&#10;', '\r': '&#13;' };
/** Characters an XML file cannot hold are left out. */
const clean = (text: string): string => text.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?:^|[^\uD800-\uDBFF])[\uDC00-\uDFFF]/g, (m) => (m.length === 2 ? m.charAt(0) : ''));
const text = (value: string): string => clean(value).replace(/[&<>\r]/g, (c) => TEXT_ESCAPES[c]);
const attr = (value: string): string => clean(value).replace(/[&<"\t\n\r]/g, (c) => ATTR_ESCAPES[c]);
const val = (name: string, value: string | number | boolean): string => `<c:${name} val="${typeof value === 'boolean' ? (value ? '1' : '0') : attr(String(value))}"/>`;

// ------------------------------------------------------------------ cached values

/** The values the cells of a reference hold now, in the form a chart part caches them; undefined for a reference that cannot be followed. */
export type ChartCells = (reference: string, kind: 'num' | 'str') => { cache: ChartCache; count: number; formatCode?: string } | undefined;

function numberOf(cell: Cell | undefined, date1904: boolean): number | null {
  if (!cell) return null;
  if (cell.kind === 'number' && typeof cell.value === 'number' && Number.isFinite(cell.value)) return cell.value;
  if (cell.kind === 'date' && typeof cell.value === 'string') {
    const serial = isoToSerial(cell.value, date1904);
    return serial === undefined ? null : serial;
  }
  return null;
}

/**
 * Reads the cells a chart of `workbook` refers to. `live`: the references name sheets as they are
 * called now (Chart.live).
 * UNSURE: the cells of hidden rows and columns are cached like the others (Excel may leave them
 * out of the cache of a chart that only plots what is visible; it works the cache out again
 * when it opens the file).
 */
export function chartCellsOf(workbook: Workbook, live: boolean): ChartCells {
  const date1904 = workbook.date1904 === true;
  return (reference, kind) => {
    const areas = chartRefAreas(workbook, reference, live);
    if (!areas) return undefined;
    const index: number[] = [];
    const values: (number | string | null)[] = [];
    let formatCode: string | undefined;
    let position = 0;
    for (const { sheet, range } of areas) {
      const width = range.right - range.left + 1;
      for (let row = range.top; row <= range.bottom && position < MAX_CHART_POINTS; row++) {
        const held = sheet.rows.get(row);
        if (!held) {
          position += width;
          continue;
        }
        for (let col = range.left; col <= range.right && position < MAX_CHART_POINTS; col++, position++) {
          const cell = held.cells[col];
          if (!cell) continue;
          if (kind === 'num') {
            const value = numberOf(cell, date1904);
            if (value === null) continue;
            if (formatCode === undefined) formatCode = resolveCellStyle(workbook.styles, cell.style).numFmt || 'General';
            index.push(position);
            values.push(value);
          } else {
            const shown = cellDisplayText(workbook, cell);
            if (shown === '') continue;
            index.push(position);
            values.push(shown);
          }
        }
      }
    }
    return { cache: { index, values }, count: Math.min(position, MAX_CHART_POINTS), formatCode };
  };
}

// ------------------------------------------------------------------ a chart part from the model

function pointsXml(cache: ChartCache, count: number): string {
  let out = val('ptCount', Math.max(0, Math.floor(count)));
  for (let i = 0; i < cache.index.length; i++) {
    const value = cache.values[i];
    if (value === null || value === undefined) continue;
    out += `<c:pt idx="${cache.index[i]}"><c:v>${text(String(value))}</c:v></c:pt>`;
  }
  return out;
}

function cacheXml(data: ChartData, cells: ChartCells | undefined): string {
  const now = cells && data.ref !== undefined ? cells(data.ref, data.kind) : undefined;
  const cache = now ? now.cache : data.cache;
  const count = now ? now.count : data.count;
  if (data.kind === 'num') {
    const code = (now ? now.formatCode : undefined) || data.formatCode || 'General';
    return `<c:formatCode>${text(code)}</c:formatCode>${pointsXml(cache, count)}`;
  }
  return pointsXml(cache, count);
}

/** A data source inside <c:cat>, <c:val>, <c:xVal> or <c:yVal>. */
function dataXml(name: string, data: ChartData, cells: ChartCells | undefined): string {
  const inner = cacheXml(data, cells);
  if (data.ref === undefined) return `<c:${name}><c:${data.kind}Lit>${inner}</c:${data.kind}Lit></c:${name}>`;
  return `<c:${name}><c:${data.kind}Ref><c:f>${text(data.ref)}</c:f><c:${data.kind}Cache>${inner}</c:${data.kind}Cache></c:${data.kind}Ref></c:${name}>`;
}

function nameXml(name: NonNullable<ChartSeries['name']>, cells: ChartCells | undefined): string {
  if (name.ref === undefined) return `<c:tx><c:v>${text(name.text || '')}</c:v></c:tx>`;
  const now = cells ? cells(name.ref, 'str') : undefined;
  const cached = now ? (now.cache.index[0] === 0 ? String(now.cache.values[0]) : undefined) : name.text;
  const point = cached === undefined ? '' : `<c:pt idx="0"><c:v>${text(cached)}</c:v></c:pt>`;
  return `<c:tx><c:strRef><c:f>${text(name.ref)}</c:f><c:strCache>${val('ptCount', 1)}${point}</c:strCache></c:strRef></c:tx>`;
}

const solid = (color: string | null): string => (color === null || !COLOR.test(color) ? '<a:noFill/>' : `<a:solidFill><a:srgbClr val="${color.slice(1).toUpperCase()}"/></a:solidFill>`);

function shapeXml(shape: ChartShape | undefined): string {
  if (!shape) return '';
  let inner = '';
  if (shape.fill !== undefined) inner += solid(shape.fill);
  const line = shape.line;
  if (line) {
    const width = line.width !== undefined ? ` w="${Math.round(line.width * EMU_PER_PIXEL)}"` : '';
    inner += `<a:ln${width}>${line.color !== undefined ? solid(line.color) : ''}${line.dash ? '<a:prstDash val="dash"/>' : ''}</a:ln>`;
  }
  return inner === '' ? '' : `<c:spPr>${inner}</c:spPr>`;
}

function markerXml(marker: ChartSeries['marker']): string {
  if (!marker) return '';
  const look: ChartShape = {};
  if (marker.fill !== undefined) look.fill = marker.fill;
  if (marker.line) look.line = marker.line;
  return `<c:marker>${val('symbol', marker.symbol)}${marker.size !== undefined ? val('size', marker.size) : ''}${shapeXml(look)}</c:marker>`;
}

function labelsXml(labels: ChartDataLabels | undefined): string {
  if (!labels) return '';
  if (labels.deleted) return `<c:dLbls>${val('delete', true)}</c:dLbls>`;
  return (
    '<c:dLbls>' +
    val('showLegendKey', labels.showLegendKey === true) +
    val('showVal', labels.showValue === true) +
    val('showCatName', labels.showCategory === true) +
    val('showSerName', labels.showSeries === true) +
    val('showPercent', labels.showPercent === true) +
    val('showBubbleSize', labels.showBubbleSize === true) +
    '</c:dLbls>'
  );
}

/** One series of a chart type of `kind`. */
function seriesXml(kind: ChartGroup['kind'], series: ChartSeries, cells: ChartCells | undefined): string {
  const scatter = kind === 'scatter' || kind === 'bubble';
  let out = `<c:ser>${val('idx', series.index)}${val('order', series.order)}`;
  if (series.name) out += nameXml(series.name, cells);
  out += shapeXml(series.shape);
  if (kind === 'bar') out += val('invertIfNegative', series.invertIfNegative === true);
  if (kind === 'line' || kind === 'scatter') out += markerXml(series.marker);
  out += labelsXml(series.dataLabels);
  if (series.categories) out += dataXml(scatter ? 'xVal' : 'cat', series.categories, cells);
  if (series.values) out += dataXml(scatter ? 'yVal' : 'val', series.values, cells);
  if (kind === 'line' || kind === 'scatter') out += val('smooth', series.smooth === true);
  return `${out}</c:ser>`;
}

const GROUP_ELEMENTS: { [kind: string]: string } = { bar: 'barChart', line: 'lineChart', area: 'areaChart', pie: 'pieChart', doughnut: 'doughnutChart', scatter: 'scatterChart' };

function groupXml(group: ChartGroup, cells: ChartCells | undefined): string {
  const name = GROUP_ELEMENTS[group.kind];
  let out = `<c:${name}>`;
  if (group.kind === 'bar') out += val('barDir', group.barDir === 'bar' ? 'bar' : 'col') + val('grouping', group.grouping === 'standard' ? 'clustered' : group.grouping);
  if (group.kind === 'line' || group.kind === 'area') out += val('grouping', group.grouping);
  if (group.kind === 'scatter') out += val('scatterStyle', group.scatterStyle || 'lineMarker');
  out += val('varyColors', group.varyColors === true);
  for (const series of group.series) out += seriesXml(group.kind, series, cells);
  out += labelsXml(group.dataLabels);
  if (group.kind === 'bar') {
    if (group.gapWidth !== undefined) out += val('gapWidth', group.gapWidth);
    if (group.overlap !== undefined) out += val('overlap', group.overlap);
  }
  if (group.kind === 'line') out += val('marker', group.markers !== false);
  if (group.kind === 'pie' || group.kind === 'doughnut') out += val('firstSliceAng', group.firstSliceAngle || 0);
  if (group.kind === 'doughnut') out += val('holeSize', group.holeSize || 50);
  for (const id of group.axisIds) out += val('axId', id);
  return `${out}</c:${name}>`;
}

/**
 * The rich text of a title: one paragraph for each line. No language is named (`lang` is optional): the workbook
 * says none, and a wrong one has the spelling checker of the application mark the whole title.
 */
function richXml(value: string): string {
  let out = '<c:rich><a:bodyPr/><a:lstStyle/>';
  for (const line of value.split('\n')) out += `<a:p><a:r><a:t>${text(line)}</a:t></a:r></a:p>`;
  return `${out}</c:rich>`;
}

function titleXml(title: ChartTitle | undefined): string {
  if (!title || title.text === undefined) return '';
  return `<c:title><c:tx>${richXml(title.text)}</c:tx>${val('overlay', false)}</c:title>`;
}

function axisXml(axis: ChartAxis): string {
  const name = axis.kind === 'cat' ? 'catAx' : 'valAx';
  let out = `<c:${name}>${val('axId', axis.id)}<c:scaling>${val('orientation', axis.reversed ? 'maxMin' : 'minMax')}`;
  if (axis.max !== undefined) out += val('max', axis.max);
  if (axis.min !== undefined) out += val('min', axis.min);
  out += `</c:scaling>${val('delete', axis.deleted === true)}${val('axPos', axis.position)}`;
  if (axis.majorGridlines) out += '<c:majorGridlines/>';
  out += titleXml(axis.title);
  out += `<c:numFmt formatCode="${attr(axis.formatCode || 'General')}" sourceLinked="${axis.sourceLinked === false ? '0' : '1'}"/>`;
  out += val('majorTickMark', 'out') + val('minorTickMark', 'none') + val('tickLblPos', axis.tickLabels);
  if (axis.crossAxis !== undefined) out += val('crossAx', axis.crossAxis);
  out += val('crosses', axis.crosses || 'autoZero');
  if (axis.kind === 'cat') out += val('auto', true) + val('lblAlgn', 'ctr') + val('lblOffset', 100) + val('noMultiLvlLbl', false);
  else out += val('crossBetween', axis.crossBetween || 'between');
  return `${out}</c:${name}>`;
}

function legendXml(legend: ChartLegend | undefined): string {
  return legend ? `<c:legend>${val('legendPos', legend.position)}${val('overlay', false)}</c:legend>` : '';
}

export { chartWritable };

/**
 * The part of a chart, from its model. With `cells` the values of the series are cached as the
 * cells hold them now; without, as the model holds them (what the file cached).
 * Throws for a chart of a kind that is not written (see chartWritable()).
 */
export function chartSpaceXml(chart: Chart, cells?: ChartCells): string {
  if (!chartWritable(chart)) throw new Error('Cannot write the chart: its type is not one that is written from the model');
  let out = `${DECLARATION}<c:chartSpace xmlns:c="${NS_C}" xmlns:a="${NS_A}" xmlns:r="${NS_R}">`;
  out += val('date1904', chart.date1904 === true) + val('roundedCorners', chart.roundedCorners === true);
  out += `<c:chart>${titleXml(chart.title)}${val('autoTitleDeleted', chart.autoTitleDeleted === true)}<c:plotArea><c:layout/>`;
  for (const group of chart.groups) out += groupXml(group, cells);
  for (const axis of chart.axes) out += axisXml(axis);
  out += `</c:plotArea>${legendXml(chart.legend)}${val('plotVisOnly', chart.plotVisOnly !== false)}${val('dispBlanksAs', chart.blanks)}</c:chart></c:chartSpace>`;
  return out;
}

// ------------------------------------------------------------------ what is known of a part

const isSpace = (node: XmlNode): boolean => node.type === 'text' && /^[ \t\r\n]*$/.test(node.value);

/**
 * Whether two elements say the same: their names as they are written (prefixes included: a part
 * that spells its namespaces another way is not one the engine wrote), their attributes
 * (declarations of namespaces aside) and their content, white space between elements aside.
 * `a` is what the engine writes now, `b` the part of a file.
 */
function sameTree(a: XmlElement, b: XmlElement, budget: { left: number }): boolean {
  if (--budget.left < 0 || a.name !== b.name) return false;
  const attrs = (el: XmlElement): Map<string, string> => {
    const out = new Map<string, string>();
    for (const [name, value] of el.attrs) if (name !== 'xmlns' && name.slice(0, 6) !== 'xmlns:') out.set(localName(name), value);
    return out;
  };
  const left = attrs(a);
  const right = attrs(b);
  if (left.size !== right.size) return false;
  for (const [name, value] of left) if (right.get(name) !== value) return false;
  const kids = (el: XmlElement): XmlNode[] => el.children.filter((node) => (node.type === 'element' || node.type === 'text') && !isSpace(node));
  const mine = kids(a);
  // Until the titles stopped naming a language every run of a title began with <a:rPr lang="en-US"/>: a part
  // that has that and nothing else there is one the engine wrote.
  const theirs = kids(b).filter((node) => !(node.type === 'element' && node.name === 'a:rPr' && a.name === 'a:r' && node.children.length === 0 && node.attrs.size === 1 && node.attrs.get('lang') === 'en-US'));
  if (mine.length !== theirs.length) return false;
  for (let i = 0; i < mine.length; i++) {
    const x = mine[i];
    const y = theirs[i];
    if (x.type !== y.type) return false;
    if (x.type === 'text' && y.type === 'text' && x.value !== y.value) return false;
    if (x.type === 'element' && y.type === 'element' && !sameTree(x, y, budget)) return false;
  }
  return true;
}

/** The children of the root of a chart part that lead out of it and that a copy of the chart can do without: see DrawingChart.loose. */
const LOOSE_NAMES: { [local: string]: true } = { userShapes: true, externalData: true };

/**
 * The elements that hold a formula (<c:f>, and <c15:f>, <cx:f> of extensions), whether the part
 * has what no model holds a reference for, and which children of its root lead out of it
 * (LOOSE_NAMES: those are not looked into).
 */
function formulas(root: XmlElement): { count: number; foreign: boolean; loose: string[] } {
  const out = { count: 0, foreign: false, loose: [] as string[] };
  const stack: XmlElement[] = [];
  for (const node of root.children) {
    if (node.type !== 'element') continue;
    const local = localName(node.name);
    // Shapes drawn over the chart (which may show cells) and data kept outside the workbook are not ranges the model follows.
    if (Object.prototype.hasOwnProperty.call(LOOSE_NAMES, local)) {
      if (out.loose.indexOf(local) < 0) out.loose.push(local);
    } else stack.push(node);
  }
  let seen = 0;
  while (stack.length > 0) {
    const el = stack.pop() as XmlElement;
    if (++seen > MAX_PART_ELEMENTS) {
      out.foreign = true;
      break;
    }
    const local = localName(el.name);
    if (local === 'f' || local === 'sqref' || local === 'formulaRef') out.count++;
    // The source of a pivot chart is not a range the model follows; neither is what leads out of the part from anywhere but its root.
    if (local === 'pivotSource' || local === 'externalData' || local === 'userShapes') out.foreign = true;
    for (const node of el.children) if (node.type === 'element') stack.push(node);
  }
  return out;
}

/** What the reader notes of a chart part beside its model: see DrawingChart.known, DrawingChart.seriesFixed and DrawingChart.loose. */
export interface ChartPartFacts {
  known?: 'all' | 'refs';
  seriesFixed?: boolean;
  loose?: string[];
}

/** Whether a part holds content bound to the number and order of its series that no model holds. */
function boundToSeries(root: XmlElement, chart: Chart): boolean {
  const chartEl = child(root, 'chart');
  if (child(chartEl, 'pivotFmts')) return true;
  let series = 0;
  let trends = false;
  for (const group of chart.groups) {
    series += group.series.length;
    if (group.series.some((item) => item.trendlines && item.trendlines.length > 0)) trends = true;
  }
  const entries = elements(child(chartEl, 'legend'), 'legendEntry');
  // The entries of trend lines are numbered behind those of the series.
  if (entries.length > 0 && trends) return true;
  for (const entry of entries) {
    const stated = (valOf(entry, 'idx') || '').trim();
    if (!/^\d{1,9}$/.test(stated) || Number(stated) >= series) return true;
  }
  return false;
}

/**
 * Whether a part holds a wrapper of alternatives (mc:AlternateContent) where the writer puts
 * and looks for what it changes: in the chart, its plot area, a chart type, a series, an axis, a
 * legend or a title. What stands inside such a wrapper is there twice, or not where it is looked
 * for: a title would be put beside the one in the wrapper.
 */
function wrapped(root: XmlElement): boolean {
  const holds = (el: XmlElement | undefined): boolean => !!el && el.children.some((node) => node.type === 'element' && localName(node.name) === 'AlternateContent');
  const chartEl = child(root, 'chart');
  const plot = child(chartEl, 'plotArea');
  if (holds(chartEl) || holds(plot) || holds(child(chartEl, 'title')) || holds(child(chartEl, 'legend'))) return true;
  for (const el of elements(plot)) {
    const local = localName(el.name);
    if (Object.prototype.hasOwnProperty.call(GROUP_NAMES, local)) {
      if (holds(el) || elements(el, 'ser').some(holds)) return true;
    } else if (Object.prototype.hasOwnProperty.call(AXIS_NAMES, local) && (holds(el) || holds(child(el, 'title')))) return true;
  }
  return false;
}

/**
 * What is known of a chart part beside its model: `known` "all" when the part is exactly what
 * chartSpaceXml() makes of the model (prefixes and namespaces included) and it has no
 * relationships; "refs" when every formula of the part is a reference of the model and the
 * series of the part are the series of the model. Undefined otherwise: also for a part nested
 * too deeply to be changed (MAX_PATCH_DEPTH) and one that wraps what is changed in alternatives.
 * Never throws.
 */
export function chartKnown(root: XmlElement, chart: Chart, related: boolean): ChartPartFacts | undefined {
  try {
    const found = formulas(root);
    if (found.loose.length > 0 || !refsKnown(root, chart, found)) return undefined;
    if (!related && chartWritable(chart) && root.attrs.get('xmlns:c') === NS_C && root.attrs.get('xmlns:a') === NS_A) {
      const made = parseXml(chartSpaceXml(chart)).root;
      if (sameTree(made, root, { left: MAX_PART_ELEMENTS })) return { known: 'all' };
    }
    return boundToSeries(root, chart) ? { known: 'refs', seriesFixed: true } : { known: 'refs' };
  } catch (error) {
    return undefined;
  }
}

/** Whether every formula of a part is a reference of the model, the series of the part are those of the model, and the part can be changed in place. */
function refsKnown(root: XmlElement, chart: Chart, found: { count: number; foreign: boolean }): boolean {
  if (nestedBeyond(root, MAX_PATCH_DEPTH)) return false;
  if (wrapped(root)) return false;
  const refs = chartRefs(chart);
  if (found.foreign || found.count !== refs.length) return false;
  // A reference that cannot be read to its end cannot follow a sheet that is renamed.
  if (refs.some((ref) => formulaNamesSheetsUnread(ref))) return false;
  // Amounts of error bars stated point by point are references the writer does not change in a part.
  if (chart.groups.some((group) => group.series.some((series) => (series.errorBars || []).some((bars) => (bars.plus && bars.plus.ref !== undefined) || (bars.minus && bars.minus.ref !== undefined))))) return false;
  const sets = seriesElements(root);
  return !!sets && sets.length === chart.groups.length && !sets.some((set, i) => set.length !== chart.groups[i].series.length);
}

/**
 * For a part of which chartKnown() knows nothing: whether it would be "refs" without the children
 * of its root that lead out of it (DrawingChart.loose), and which those are. Undefined for a part
 * that has none of them, and for one that is not known for another reason. Never throws.
 */
export function chartLoose(root: XmlElement, chart: Chart): ChartPartFacts | undefined {
  try {
    const found = formulas(root);
    if (found.loose.length === 0 || !refsKnown(root, chart, found)) return undefined;
    return boundToSeries(root, chart) ? { loose: found.loose, seriesFixed: true } : { loose: found.loose };
  } catch (error) {
    return undefined;
  }
}

// ------------------------------------------------------------------ a chart part changed in place

const GROUP_NAMES: { [local: string]: true } = {
  barChart: true,
  bar3DChart: true,
  lineChart: true,
  line3DChart: true,
  pieChart: true,
  pie3DChart: true,
  ofPieChart: true,
  doughnutChart: true,
  areaChart: true,
  area3DChart: true,
  scatterChart: true,
  radarChart: true,
  bubbleChart: true,
  stockChart: true,
};
const AXIS_NAMES: { [local: string]: true } = { catAx: true, dateAx: true, serAx: true, valAx: true };

const elements = (parent: XmlElement | undefined, local?: string): XmlElement[] => {
  const out: XmlElement[] = [];
  if (parent) for (const node of parent.children) if (node.type === 'element' && (local === undefined || localName(node.name) === local)) out.push(node);
  return out;
};
const child = (parent: XmlElement | undefined, local: string): XmlElement | undefined => elements(parent, local)[0];
const valOf = (parent: XmlElement | undefined, local: string): string | undefined => {
  const el = child(parent, local);
  return el ? el.attrs.get('val') : undefined;
};
const textOf = (el: XmlElement | undefined): string => {
  let out = '';
  if (el) for (const node of el.children) if (node.type === 'text') out += node.value;
  return out;
};

/** The chart types of a part in the order the reader takes them, as far as it reads them. */
function groupElements(root: XmlElement): XmlElement[] {
  const plot = child(child(root, 'chart'), 'plotArea');
  return elements(plot)
    .filter((el) => Object.prototype.hasOwnProperty.call(GROUP_NAMES, localName(el.name)))
    .slice(0, 16);
}

/**
 * The <c:ser> elements of each chart type of a part, in the order the reader puts the series of
 * the model in (by <c:order>, then as they stand); undefined for a part with more series than are read.
 */
function seriesElements(root: XmlElement): XmlElement[][] | undefined {
  const out: XmlElement[][] = [];
  let position = 0;
  for (const group of groupElements(root)) {
    const list: { el: XmlElement; order: number; at: number }[] = [];
    for (const el of elements(group, 'ser')) {
      if (position >= MAX_CHART_SERIES) return undefined;
      const stated = valOf(el, 'order');
      const order = stated !== undefined && /^\d{1,9}$/.test(stated.trim()) ? Number(stated.trim()) : position;
      list.push({ el, order, at: list.length });
      position++;
    }
    list.sort((a, b) => a.order - b.order || a.at - b.at);
    out.push(list.map((item) => item.el));
  }
  return out;
}

const STRICT_C = 'http://purl.oclc.org/ooxml/drawingml/chart';
const STRICT_A = 'http://purl.oclc.org/ooxml/drawingml/main';
const STRICT_R = 'http://purl.oclc.org/ooxml/officeDocument/relationships';
const STRICT_XDR = 'http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing';

/** Where an element of a part stands in the text of the part, and what it held when the part was read. */
interface Span {
  start: number;
  /** Just behind its start tag. */
  openEnd: number;
  end: number;
  selfClosing: boolean;
  name: string;
  attrs: string;
  children: XmlNode[];
}

/**
 * A part that is being changed: the text it was read from with the place of every element in
 * it (so that what is not changed is written with the bytes it has), which element each
 * element stands in (so that the namespaces in scope are known where something is put in),
 * and the namespaces the part writes its elements in.
 */
interface PartContext {
  source: string | undefined;
  spans: WeakMap<XmlElement, Span>;
  parents: WeakMap<XmlElement, XmlElement>;
  /** The namespaces generated text stands for by its prefixes ("c", "a", "xdr"): those of the part, strict where the part is. */
  spaces: { [prefix: string]: string };
}

/** The part that is being changed; the functions below are only called while one is. */
let part: PartContext | undefined;

const attrsKey = (el: XmlElement): string => {
  let out = '';
  for (const [name, value] of el.attrs) out += `${name}\u0000${value}\u0001`;
  return out;
};

/** The places of the elements of a text of XML, in document order; undefined for a text this simple scan cannot follow. */
function elementSpans(source: string): { name: string; start: number; openEnd: number; end: number; selfClosing: boolean }[] | undefined {
  const out: { name: string; start: number; openEnd: number; end: number; selfClosing: boolean }[] = [];
  const open: number[] = [];
  const n = source.length;
  let i = 0;
  while (i < n) {
    const at = source.indexOf('<', i);
    if (at < 0) break;
    if (source.startsWith('<!--', at)) {
      const end = source.indexOf('-->', at + 4);
      if (end < 0) return undefined;
      i = end + 3;
    } else if (source.startsWith('<![CDATA[', at)) {
      const end = source.indexOf(']]>', at + 9);
      if (end < 0) return undefined;
      i = end + 3;
    } else if (source.startsWith('<?', at)) {
      const end = source.indexOf('?>', at + 2);
      if (end < 0) return undefined;
      i = end + 2;
    } else if (source.charCodeAt(at + 1) === 33) {
      // A declaration of a document type: the reader refuses such a part before it comes here.
      return undefined;
    } else if (source.charCodeAt(at + 1) === 47) {
      const end = source.indexOf('>', at);
      const index = open.pop();
      if (end < 0 || index === undefined) return undefined;
      if (source.slice(at + 2, end).trim() !== out[index].name) return undefined;
      out[index].end = end + 1;
      i = end + 1;
    } else {
      // A start tag: its end is the first ">" outside a quoted value.
      let end = at + 1;
      let quote = 0;
      for (; end < n; end++) {
        const code = source.charCodeAt(end);
        if (quote !== 0) {
          if (code === quote) quote = 0;
        } else if (code === 34 || code === 39) quote = code;
        else if (code === 62) break;
      }
      if (end >= n) return undefined;
      const name = /^[^\s/>]+/.exec(source.slice(at + 1, end + 1));
      if (!name) return undefined;
      const selfClosing = source.charCodeAt(end - 1) === 47;
      out.push({ name: name[0], start: at, openEnd: end + 1, end: end + 1, selfClosing });
      if (!selfClosing) open.push(out.length - 1);
      i = end + 1;
    }
  }
  return open.length === 0 ? out : undefined;
}

/** Reads a part for changing: its tree, and the context that goes with it (see PartContext). */
function openPart(bytes: Uint8Array, spaces: (root: XmlElement, scope: Map<string, string>) => { [prefix: string]: string }): { doc: XmlDocument; context: PartContext } {
  const doc = parseXml(bytes);
  const context: PartContext = { source: undefined, spans: new WeakMap(), parents: new WeakMap(), spaces: {} };
  const elements: XmlElement[] = [];
  const walk = (el: XmlElement, depth: number): void => {
    elements.push(el);
    // Never met for a part the reader let be changed (see nestedBeyond()): what stops a stack that runs out.
    if (depth > 2000) throw new Error('Cannot write the part: its content is nested too deeply');
    for (const node of el.children) {
      if (node.type !== 'element') continue;
      context.parents.set(node, el);
      walk(node, depth + 1);
    }
  };
  walk(doc.root, 0);
  // The text of the part, when it is UTF-8 and says nothing else (what every application writes): only
  // then are its bytes kept. A part that declares another encoding is written out whole, as UTF-8
  // with a declaration that says so: new text beside the old declaration would be read as something else.
  const declared = doc.declaration ? /\sencoding\s*=\s*["']([^"']*)["']/.exec(doc.declaration) : null;
  let source: string | undefined;
  try {
    source = (declared && !/^utf-?8$/i.test(declared[1])) || (bytes.length >= 2 && ((bytes[0] === 0xff && bytes[1] === 0xfe) || (bytes[0] === 0xfe && bytes[1] === 0xff))) ? undefined : new TextDecoder('utf-8', { fatal: true, ignoreBOM: true }).decode(bytes);
  } catch (error) {
    source = undefined;
  }
  const found = source === undefined ? undefined : elementSpans(source);
  if (source !== undefined && found && found.length === elements.length && found.every((span, i) => span.name === elements[i].name)) {
    context.source = source;
    found.forEach((span, i) => {
      const el = elements[i];
      context.spans.set(el, { start: span.start, openEnd: span.openEnd, end: span.end, selfClosing: span.selfClosing, name: el.name, attrs: attrsKey(el), children: el.children.slice() });
    });
  }
  part = context;
  context.spaces = spaces(doc.root, scopeOf(doc.root));
  return { doc, context };
}

const escapeText = (value: string): string => value.replace(/[&<>\r]/g, (c) => TEXT_ESCAPES[c]);
const escapeAttr = (value: string): string => value.replace(/[&<"\t\n\r]/g, (c) => ATTR_ESCAPES[c]);

function startTag(el: XmlElement, close: boolean): string {
  let out = `<${el.name}`;
  for (const [name, value] of el.attrs) out += ` ${name}="${escapeAttr(value)}"`;
  return out + (close ? '/>' : '>');
}

/** A node as text: an element of the part that was not changed with the very text it has in the part, anything else written out. */
function emit(node: XmlNode, context: PartContext): string {
  if (node.type === 'text') return node.cdata ? `<![CDATA[${node.value}]]>` : escapeText(node.value);
  if (node.type === 'comment') return `<!--${node.value}-->`;
  if (node.type === 'pi') return `<?${node.target}${node.value ? ` ${node.value}` : ''}?>`;
  const source = context.source;
  const span = source === undefined ? undefined : context.spans.get(node);
  if (source !== undefined && span) {
    const sameTag = node.name === span.name && attrsKey(node) === span.attrs;
    const sameChildren = node.children.length === span.children.length && node.children.every((child, i) => child === span.children[i]);
    if (sameTag && sameChildren) {
      // Nothing of the element itself changed: its text, with every child that did change in its place.
      let out = '';
      let cursor = span.start;
      for (const child of node.children) {
        if (child.type !== 'element') continue;
        const inner = context.spans.get(child);
        if (!inner) return fresh(node, context, undefined);
        out += source.slice(cursor, inner.start) + emit(child, context);
        cursor = inner.end;
      }
      return out + source.slice(cursor, span.end);
    }
    return fresh(node, context, sameTag && !span.selfClosing ? source.slice(span.start, span.openEnd) : undefined);
  }
  return fresh(node, context, undefined);
}

/** An element written out: `open` is its start tag as the part has it, when that still holds. */
function fresh(el: XmlElement, context: PartContext, open: string | undefined): string {
  if (el.children.length === 0) return open !== undefined ? `${open}</${el.name}>` : startTag(el, true);
  let out = open !== undefined ? open : startTag(el, false);
  for (const child of el.children) out += emit(child, context);
  return `${out}</${el.name}>`;
}

/** The part as bytes: everything outside what was changed is what the part was read from. */
function closePart(doc: XmlDocument, context: PartContext): Uint8Array {
  const source = context.source;
  const span = source === undefined ? undefined : context.spans.get(doc.root);
  if (source === undefined || !span) return new TextEncoder().encode(serializeXml(doc));
  return new TextEncoder().encode(source.slice(0, span.start) + emit(doc.root, context) + source.slice(span.end));
}

/** The namespaces in scope at an element, by prefix ('' for the default one): the nearest declaration of each. */
function scopeOf(el: XmlElement | undefined): Map<string, string> {
  const scope = new Map<string, string>();
  const parents = part ? part.parents : undefined;
  for (let at = el, depth = 0; at && depth < 4000; at = parents ? parents.get(at) : undefined, depth++) {
    for (const [name, value] of at.attrs) {
      if (name !== 'xmlns' && name.slice(0, 6) !== 'xmlns:') continue;
      const prefix = name === 'xmlns' ? '' : name.slice(6);
      if (!scope.has(prefix)) scope.set(prefix, value);
    }
  }
  return scope;
}

const prefixOf = (name: string): string => (name.indexOf(':') < 0 ? '' : name.slice(0, name.indexOf(':')));

/**
 * An element for `parent`, made from text written with the prefixes of `part.spaces` ("c", "a",
 * "xdr"): every element of it is named with the prefix the part has in scope there for its
 * namespace (the default namespace included), and a namespace no prefix stands for there is
 * declared on the element itself. The element is not put anywhere: the caller does that.
 */
function made(parent: XmlElement, xml: string): XmlElement {
  const context = part as PartContext;
  const spaces = context.spaces;
  let declarations = '';
  for (const prefix of Object.keys(spaces)) declarations += ` xmlns:${prefix}="urn:x-generated:${prefix}"`;
  const el = parseXml(xml.replace(/^<([A-Za-z0-9:]+)/, `<$1${declarations}`)).root;
  for (const prefix of Object.keys(spaces)) el.attrs.delete(`xmlns:${prefix}`);
  const scope = scopeOf(parent);
  const own = prefixOf(parent.name);
  const chosen = new Map<string, string>();
  const declare: [string, string][] = [];
  const prefixFor = (written: string): string => {
    const known = chosen.get(written);
    if (known !== undefined) return known;
    const space = spaces[written];
    let prefix: string | undefined;
    if (scope.get(own) === space) prefix = own;
    else for (const [candidate, bound] of scope) if (bound === space && prefix === undefined) prefix = candidate;
    if (prefix === undefined) {
      // No prefix stands for the namespace here: one that is free is declared on the element.
      prefix = written;
      for (let n = 1; scope.has(prefix); n++) prefix = `${written}${n}`;
      declare.push([prefix, space]);
    }
    chosen.set(written, prefix);
    return prefix;
  };
  const rename = (node: XmlElement, depth: number): void => {
    const written = prefixOf(node.name);
    if (Object.prototype.hasOwnProperty.call(spaces, written)) {
      const prefix = prefixFor(written);
      node.name = (prefix === '' ? '' : `${prefix}:`) + localName(node.name);
    }
    if (depth < 200) {
      for (const child of node.children) {
        if (child.type !== 'element') continue;
        context.parents.set(child, node);
        rename(child, depth + 1);
      }
    }
  };
  rename(el, 0);
  for (const [prefix, space] of declare) el.attrs.set(prefix === '' ? 'xmlns' : `xmlns:${prefix}`, space);
  context.parents.set(el, parent);
  return el;
}

/** The namespaces a chart part writes its own elements and those of drawings in: the transitional ones, or the strict ones where its root is strict. */
function chartSpaces(root: XmlElement, scope: Map<string, string>): { [prefix: string]: string } {
  const own = scope.get(prefixOf(root.name));
  const isStrict = own === STRICT_C;
  let drawings = isStrict ? STRICT_A : NS_A;
  // A part that names the other flavour of the namespace of drawings is written in that one.
  for (const bound of scope.values()) if (bound === NS_A || bound === STRICT_A) drawings = bound;
  return { c: own === STRICT_C || own === NS_C ? own : NS_C, a: drawings };
}

const CHART_ORDER = ['title', 'autoTitleDeleted', 'pivotFmts', 'view3D', 'floor', 'sideWall', 'backWall', 'plotArea', 'legend', 'plotVisOnly', 'dispBlanksAs', 'showDLblsOverMax', 'extLst'];
const AXIS_ORDER = ['axId', 'scaling', 'delete', 'axPos', 'majorGridlines', 'minorGridlines', 'title', 'numFmt', 'majorTickMark', 'minorTickMark', 'tickLblPos', 'spPr', 'txPr', 'crossAx', 'crosses', 'crossesAt', 'extLst'];
const SCALING_ORDER = ['logBase', 'orientation', 'max', 'min', 'extLst'];
const GROUP_ORDER = [
  'barDir',
  'radarStyle',
  'scatterStyle',
  'wireframe',
  'ofPieType',
  'grouping',
  'varyColors',
  'ser',
  'dLbls',
  'gapWidth',
  'splitType',
  'splitPos',
  'custSplit',
  'secondPieSize',
  'overlap',
  'dropLines',
  'hiLowLines',
  'upDownBars',
  'marker',
  'smooth',
  'firstSliceAng',
  'holeSize',
  'bubble3D',
  'bubbleScale',
  'showNegBubbles',
  'sizeRepresents',
  'serLines',
  'axId',
  'extLst',
];
const SERIES_ORDER = ['idx', 'order', 'tx', 'spPr', 'invertIfNegative', 'pictureOptions', 'marker', 'explosion', 'dPt', 'dLbls', 'trendline', 'errBars', 'cat', 'xVal', 'val', 'yVal', 'smooth', 'shape', 'bubbleSize', 'bubble3D', 'extLst'];
const LABELS_ORDER = ['dLbl', 'delete', 'numFmt', 'spPr', 'txPr', 'dLblPos', 'showLegendKey', 'showVal', 'showCatName', 'showSerName', 'showPercent', 'showBubbleSize', 'separator', 'showLeaderLines', 'leaderLines', 'extLst'];
const SHAPE_ORDER = ['xfrm', 'custGeom', 'prstGeom', 'noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill', 'ln', 'effectLst', 'effectDag', 'scene3d', 'sp3d', 'extLst'];
const FILLS = ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'];

/** Puts an element among the children of `parent` where the schema has it: before the first child that comes later in `order`. */
function insert(parent: XmlElement, el: XmlElement, order: string[]): void {
  const rank = order.indexOf(localName(el.name));
  let at = parent.children.length;
  for (let i = 0; i < parent.children.length; i++) {
    const node = parent.children[i];
    if (node.type !== 'element') continue;
    const other = order.indexOf(localName(node.name));
    if (other > rank) {
      at = i;
      break;
    }
  }
  parent.children.splice(at, 0, el);
}

function remove(parent: XmlElement, local: string): boolean {
  const kept = parent.children.filter((node) => !(node.type === 'element' && localName(node.name) === local));
  if (kept.length === parent.children.length) return false;
  parent.children = kept;
  return true;
}

/** Sets the `val` of a child that holds one, making the child where the schema has it when there is none. */
function setVal(parent: XmlElement, local: string, value: string | number | boolean, order: string[]): void {
  const stated = typeof value === 'boolean' ? (value ? '1' : '0') : String(value);
  const el = child(parent, local);
  if (el) el.attrs.set('val', stated);
  else insert(parent, made(parent, val(local, value)), order);
}

function setText(el: XmlElement, value: string): void {
  el.children = [{ type: 'text', value: clean(value) }];
}

const sameTitle = (a: ChartTitle | undefined, b: ChartTitle | undefined): boolean => (a ? a.text : undefined) === (b ? b.text : undefined) && (a ? a.ref : undefined) === (b ? b.ref : undefined);

/**
 * Brings the title of a chart or an axis (a child of `holder`) in line with the model. The text
 * of a title of the file is put into its first run, whose look stays; a title that reads a cell
 * keeps doing so when only its reference changed.
 */
function patchTitle(holder: XmlElement, before: ChartTitle | undefined, now: ChartTitle | undefined, order: string[]): void {
  if (sameTitle(before, now)) return;
  const title = child(holder, 'title');
  if (!now || (now.text === undefined && now.ref === undefined)) {
    remove(holder, 'title');
    return;
  }
  if (now.ref !== undefined) {
    // Only its reference can have changed (rows, columns or sheets moved): the cell it reads is another.
    const formula = child(child(child(title, 'tx'), 'strRef'), 'f');
    if (formula) setText(formula, now.ref);
    return;
  }
  const wanted = now.text as string;
  if (!title) {
    insert(holder, made(holder, titleXml({ text: wanted })), order);
    return;
  }
  const tx = child(title, 'tx');
  const rich = child(tx, 'rich');
  const first = rich ? elements(rich, 'p')[0] : undefined;
  if (!tx || !rich || !first) {
    // A title without a text of its own (it read a cell, or showed what the chart made up) gets one.
    if (tx) title.children = title.children.filter((node) => node !== tx);
    title.children.unshift(made(title, `<c:tx>${richXml(wanted)}</c:tx>`));
    return;
  }
  // One paragraph for each line, each made like the first: its properties and the look of its first run.
  const properties = child(first, 'pPr');
  const run = elements(first).find((el) => localName(el.name) === 'r');
  const runProps = run ? child(run, 'rPr') : undefined;
  const prefix = first.name.slice(0, first.name.length - localName(first.name).length);
  const paragraphs: XmlNode[] = wanted.split('\n').map((line) => {
    const kids: XmlNode[] = [];
    if (properties) kids.push(properties);
    const inner: XmlNode[] = [];
    if (runProps) inner.push(runProps);
    inner.push({ type: 'element', name: `${prefix}t`, attrs: new Map(), children: [{ type: 'text', value: clean(line) }] });
    kids.push({ type: 'element', name: `${prefix}r`, attrs: new Map(), children: inner });
    return { type: 'element', name: first.name, attrs: new Map(first.attrs), children: kids };
  });
  let placed = false;
  const kept: XmlNode[] = [];
  for (const node of rich.children) {
    if (node.type === 'element' && localName(node.name) === 'p') {
      if (!placed) for (const paragraph of paragraphs) kept.push(paragraph);
      placed = true;
    } else kept.push(node);
  }
  rich.children = kept;
}

/** Sets the fill among the children of `holder` (<c:spPr>, or <a:ln> for the colour of a line): what it had as a fill gives way. */
function setFill(holder: XmlElement, color: string | null | undefined, order: string[]): void {
  holder.children = holder.children.filter((node) => !(node.type === 'element' && FILLS.indexOf(localName(node.name)) >= 0));
  if (color === undefined) return;
  insert(holder, made(holder, solid(color)), order);
}

function patchShape(series: XmlElement, before: ChartShape | undefined, now: ChartShape | undefined): void {
  const fillBefore = before ? before.fill : undefined;
  const fillNow = now ? now.fill : undefined;
  const lineBefore = before && before.line ? before.line.color : undefined;
  const lineNow = now && now.line ? now.line.color : undefined;
  if (fillBefore === fillNow && lineBefore === lineNow) return;
  let props = child(series, 'spPr');
  if (!props) {
    if (fillNow === undefined && lineNow === undefined) return;
    props = made(series, '<c:spPr/>');
    insert(series, props, SERIES_ORDER);
  }
  if (fillBefore !== fillNow) setFill(props, fillNow, SHAPE_ORDER);
  if (lineBefore !== lineNow) {
    let line = child(props, 'ln');
    if (!line) {
      line = made(props, '<a:ln/>');
      insert(props, line, SHAPE_ORDER);
    }
    setFill(line, lineNow, FILLS.concat(['prstDash', 'custDash', 'round', 'bevel', 'miter', 'headEnd', 'tailEnd', 'extLst']));
  }
}

function patchLabels(holder: XmlElement, before: ChartDataLabels | undefined, now: ChartDataLabels | undefined, order: string[]): void {
  if (JSON.stringify(before || null) === JSON.stringify(now || null)) return;
  const labels = child(holder, 'dLbls');
  if (!now) {
    remove(holder, 'dLbls');
    return;
  }
  if (!labels || (before && before.deleted)) {
    // Labels that were taken away in the file say nothing the model holds: they are written anew.
    remove(holder, 'dLbls');
    insert(holder, made(holder, labelsXml(now)), order);
    return;
  }
  const flags: [string, boolean][] = [
    ['showLegendKey', now.showLegendKey === true],
    ['showVal', now.showValue === true],
    ['showCatName', now.showCategory === true],
    ['showSerName', now.showSeries === true],
    ['showPercent', now.showPercent === true],
    ['showBubbleSize', now.showBubbleSize === true],
  ];
  for (const [name, on] of flags) setVal(labels, name, on, LABELS_ORDER);
}

/** The element of a data source that holds its reference and its cache (<c:numRef>, <c:strRef>, <c:multiLvlStrRef>). */
function refElement(holder: XmlElement | undefined): XmlElement | undefined {
  return elements(holder).find((el) => /^(numRef|strRef|multiLvlStrRef)$/.test(localName(el.name)));
}

/** Whether a cache of a part (<c:numCache>, <c:strCache>) says what the cells say: the same points with the same values, however they are spelled. */
function cacheSays(cache: XmlElement | undefined, held: { cache: ChartCache; count: number }, kind: 'num' | 'str'): boolean {
  if (!cache) return false;
  const stated = (valOf(cache, 'ptCount') || '').trim();
  if (!/^\d{1,9}$/.test(stated) || Number(stated) !== held.count) return false;
  const points = elements(cache, 'pt');
  if (points.length !== held.cache.index.length) return false;
  for (let i = 0; i < points.length; i++) {
    const at = (points[i].attrs.get('idx') || '').trim();
    if (!/^\d{1,9}$/.test(at) || Number(at) !== held.cache.index[i]) return false;
    const written = textOf(child(points[i], 'v'));
    const value = held.cache.values[i];
    if (kind === 'num' ? !(written.trim() !== '' && Number(written) === value) : written !== value) return false;
  }
  return true;
}

/** Brings a data source of a series (<c:cat>, <c:val>, ...) in line with the model, and its cache up to date with the cells. */
function patchData(series: XmlElement, name: string, before: ChartData | undefined, now: ChartData | undefined, cells: ChartCells | undefined): void {
  const holder = child(series, name);
  if (!now) {
    if (before) remove(series, name);
    return;
  }
  const ref = refElement(holder);
  const local = ref ? localName(ref.name) : '';
  const fits = !!ref && now.ref !== undefined && (local === 'multiLvlStrRef' ? now.multiLevel === true : local === `${now.kind}Ref`);
  if (!holder || !ref || !fits) {
    // The model has it another way than the part (no source before, or one of another form): written anew.
    if (before === now) return;
    if (holder) series.children = series.children.filter((node) => node !== holder);
    insert(series, made(series, dataXml(name, now, cells)), SERIES_ORDER);
    return;
  }
  const formula = child(ref, 'f');
  if (formula && (!before || before.ref !== now.ref)) setText(formula, now.ref as string);
  if (local === 'multiLvlStrRef' || !cells) return;
  // The cells as they are now, where the reference can be followed; what the file cached otherwise.
  const held = cells(now.ref as string, now.kind);
  if (!held) return;
  const cacheName = `${now.kind}Cache`;
  const old = child(ref, cacheName);
  // A cache that says what the cells say stays as it is, with the formats the file gave its points.
  if (cacheSays(old, held, now.kind)) return;
  const made_ = made(ref, `<c:${cacheName}>${cacheXml(now, cells)}</c:${cacheName}>`);
  if (old) {
    // What the cache holds beside its points (an extension list) stays.
    const extra = old.children.filter((node) => node.type === 'element' && localName(node.name) === 'extLst');
    made_.children = made_.children.concat(extra);
    ref.children = ref.children.map((node) => (node === old ? made_ : node));
  } else insert(ref, made_, ['f', cacheName, 'extLst']);
}

function patchName(series: XmlElement, before: ChartSeries['name'], now: ChartSeries['name'], cells: ChartCells | undefined): void {
  const same = (before ? before.ref : undefined) === (now ? now.ref : undefined) && (before ? before.text : undefined) === (now ? now.text : undefined);
  const tx = child(series, 'tx');
  if (!now) {
    if (before) remove(series, 'tx');
    return;
  }
  const ref = child(tx, 'strRef');
  if (same || (ref && now.ref !== undefined)) {
    // The name is read from a cell as before: its reference and what the cell holds now.
    if (!ref || now.ref === undefined) return;
    const formula = child(ref, 'f');
    if (formula && !same) setText(formula, now.ref);
    const held = cells ? cells(now.ref, 'str') : undefined;
    if (!held) return;
    // The first cell is the name; a cache that says so already stays as it is.
    const first = held.cache.index[0] === 0 ? { cache: { index: [0], values: [held.cache.values[0]] }, count: 1 } : { cache: { index: [], values: [] }, count: 1 };
    const old = child(ref, 'strCache');
    if (cacheSays(old, first, 'str')) return;
    const point = first.cache.index.length > 0 ? `<c:pt idx="0"><c:v>${text(String(first.cache.values[0]))}</c:v></c:pt>` : '';
    const cache = made(ref, `<c:strCache>${val('ptCount', 1)}${point}</c:strCache>`);
    if (old) {
      cache.children = cache.children.concat(old.children.filter((node) => node.type === 'element' && localName(node.name) === 'extLst'));
      ref.children = ref.children.map((node) => (node === old ? cache : node));
    } else insert(ref, cache, ['f', 'strCache', 'extLst']);
    return;
  }
  if (tx) series.children = series.children.filter((node) => node !== tx);
  insert(series, made(series, nameXml(now, cells)), SERIES_ORDER);
}

/** The ids of relationships an element and what stands in it name (r:id, r:embed, r:link, ...), added to `out`. */
function relationIdsOf(el: XmlElement, out: Set<string>, depth = 0): void {
  for (const [name, value] of el.attrs) if (name.indexOf(':') > 0 && name.slice(0, 6) !== 'xmlns:' && /^(id|embed|link|pict|dm|lo|qs|cs)$/.test(localName(name))) out.add(value);
  if (depth <= MAX_PATCH_DEPTH) for (const node of el.children) if (node.type === 'element') relationIdsOf(node, out, depth + 1);
}

/**
 * A chart part without the children of its root named in `drops` (DrawingChart.loose): for a
 * copy of the chart, which cannot take along what those lead to. Every other byte of the part
 * stays. Also says which relationships only what went named (`unrelated`). The same bytes when
 * the part has none of them.
 */
export function stripChartPart(bytes: Uint8Array, drops: readonly string[]): { bytes: Uint8Array; unrelated: string[] } {
  const before = part;
  try {
    const { doc, context } = openPart(bytes, chartSpaces);
    const root = doc.root;
    const gone = new Set<string>();
    const kept: XmlNode[] = [];
    for (const node of root.children) {
      if (node.type === 'element' && drops.indexOf(localName(node.name)) >= 0) relationIdsOf(node, gone);
      else kept.push(node);
    }
    if (kept.length === root.children.length) return { bytes, unrelated: [] };
    root.children = kept;
    const stay = new Set<string>();
    relationIdsOf(root, stay);
    return { bytes: closePart(doc, context), unrelated: [...gone].filter((id) => !stay.has(id)) };
  } finally {
    part = before;
  }
}

/**
 * The part of a chart of the file after the chart was changed: `origin` is the chart as the part
 * was read, `chart` the chart as it is now. What differs between the two is changed in the part,
 * in the prefixes and namespaces the part writes its elements with; the cached values of the
 * series are brought up to date where the cells say something else (`cells`); every other byte
 * of the part stays. `fixed`: DrawingChart.seriesFixed. Throws for a change the part cannot take:
 * chartPartProblem() says which, and the commands refuse those before they are operations.
 */
export function patchChartPart(bytes: Uint8Array, origin: ChartOrigin, chart: Chart, cells: ChartCells | undefined, fixed = false): Uint8Array {
  const problem = chartPartProblem(origin.chart, chart, origin.series, fixed);
  if (problem) throw new Error(`Cannot write the chart: ${problem}`);
  const before = part;
  try {
    const { doc, context } = openPart(bytes, chartSpaces);
    patchChart(doc.root, origin, chart, cells);
    return closePart(doc, context);
  } finally {
    part = before;
  }
}

function patchChart(root: XmlElement, origin: ChartOrigin, chart: Chart, cells: ChartCells | undefined): void {
  const chartEl = child(root, 'chart');
  const plot = child(chartEl, 'plotArea');
  const before = origin.chart;
  const sets = seriesElements(root);
  const groups = groupElements(root);
  if (!chartEl || !plot || !sets || groups.length !== before.groups.length || chart.groups.length !== before.groups.length || sets.some((set, i) => set.length !== before.groups[i].series.length)) {
    throw new Error('Cannot write the chart: its part is not what the chart was read from');
  }

  patchTitle(chartEl, before.title, chart.title, CHART_ORDER);
  // A chart without a title says whether the application makes one up: as the model says.
  if ((before.autoTitleDeleted === true) !== (chart.autoTitleDeleted === true)) setVal(chartEl, 'autoTitleDeleted', chart.autoTitleDeleted === true, CHART_ORDER);
  // The legend.
  const legendBefore = before.legend ? before.legend.position : undefined;
  const legendNow = chart.legend ? chart.legend.position : undefined;
  if (legendBefore !== legendNow) {
    const legend = child(chartEl, 'legend');
    if (legendNow === undefined) remove(chartEl, 'legend');
    else if (legend) {
      const position = child(legend, 'legendPos');
      if (position) position.attrs.set('val', legendNow);
      else legend.children.unshift(made(legend, val('legendPos', legendNow)));
    } else insert(chartEl, made(chartEl, legendXml(chart.legend)), CHART_ORDER);
  }
  // The entries of the legend are numbered by the series as they stand: each follows its series, that of a series that is gone goes.
  const legend = child(chartEl, 'legend');
  const entries = elements(legend, 'legendEntry');
  if (legend && entries.length > 0) {
    const moved = new Map<number, number>();
    let same = true;
    let was = 0;
    let is = 0;
    for (let g = 0; g < before.groups.length; g++) {
      const places = origin.series[g] || [];
      places.forEach((place, i) => {
        if (place >= 0) moved.set(was + place, is + i);
        if (place !== i) same = false;
      });
      if (places.length !== before.groups[g].series.length) same = false;
      was += before.groups[g].series.length;
      is += places.length;
    }
    if (!same) {
      const gone = new Set<XmlNode>();
      for (const entry of entries) {
        const stated = (valOf(entry, 'idx') || '').trim();
        const index = /^\d{1,9}$/.test(stated) ? Number(stated) : -1;
        if (index < 0 || index >= was) throw new Error('Cannot write the chart: an entry of its legend belongs to no series of the model');
        const now = moved.get(index);
        if (now === undefined) gone.add(entry);
        else if (now !== index) (child(entry, 'idx') as XmlElement).attrs.set('val', String(now));
      }
      if (gone.size > 0) legend.children = legend.children.filter((node) => !gone.has(node));
    }
  }
  // The axes, each known by its id.
  for (const axis of chart.axes) {
    const old = before.axes.find((other) => other.id === axis.id);
    if (!old || old === axis) continue;
    const el = elements(plot).find((item) => Object.prototype.hasOwnProperty.call(AXIS_NAMES, localName(item.name)) && (valOf(item, 'axId') || '').trim() === String(axis.id));
    if (!el) throw new Error('Cannot write the chart: an axis of the model is not in its part');
    patchTitle(el, old.title, axis.title, AXIS_ORDER);
    if (old.min !== axis.min || old.max !== axis.max) {
      let scaling = child(el, 'scaling');
      if (!scaling) {
        scaling = made(el, '<c:scaling/>');
        insert(el, scaling, AXIS_ORDER);
      }
      for (const key of ['max', 'min'] as const) {
        if (old[key] === axis[key]) continue;
        if (axis[key] === undefined) remove(scaling, key);
        else setVal(scaling, key, axis[key] as number, SCALING_ORDER);
      }
    }
  }
  // The chart types and their series.
  for (let g = 0; g < groups.length; g++) {
    const el = groups[g];
    const old = before.groups[g];
    const now = chart.groups[g];
    if (old.grouping !== now.grouping) setVal(el, 'grouping', now.kind === 'bar' && now.grouping === 'standard' ? 'clustered' : now.grouping, GROUP_ORDER);
    if (old.overlap !== now.overlap) {
      if (now.overlap === undefined) remove(el, 'overlap');
      else setVal(el, 'overlap', now.overlap, GROUP_ORDER);
    }
    patchLabels(el, old.dataLabels, now.dataLabels, GROUP_ORDER);
    const scatter = now.kind === 'scatter' || now.kind === 'bubble';
    const places = origin.series[g] || [];
    const kept = new Set<XmlElement>();
    now.series.forEach((series, i) => {
      const place = places[i];
      const source = place !== undefined && place >= 0 ? sets[g][place] : undefined;
      if (!source) {
        const added = made(el, seriesXml(now.kind, series, cells));
        insert(el, added, GROUP_ORDER);
        kept.add(added);
        return;
      }
      if (kept.has(source)) throw new Error('Cannot write the chart: one series of the file stands for two of the model');
      kept.add(source);
      const was = old.series[place];
      if (was.order !== series.order) setVal(source, 'order', series.order, SERIES_ORDER);
      patchName(source, was.name, series.name, cells);
      patchShape(source, was.shape, series.shape);
      if (scatter && (was.marker ? was.marker.fill : undefined) !== (series.marker ? series.marker.fill : undefined)) {
        // The colour of the points of a scatter series is that of its markers.
        let marker = child(source, 'marker');
        if (!marker) {
          marker = made(source, markerXml(series.marker) || '<c:marker/>');
          insert(source, marker, SERIES_ORDER);
        } else {
          let props = child(marker, 'spPr');
          if (!props) {
            props = made(marker, '<c:spPr/>');
            insert(marker, props, ['symbol', 'size', 'spPr', 'extLst']);
          }
          setFill(props, series.marker ? series.marker.fill : undefined, SHAPE_ORDER);
        }
      }
      patchLabels(source, was.dataLabels, series.dataLabels, SERIES_ORDER);
      if (was.pointLabels && !series.pointLabels && !series.dataLabels) remove(source, 'dLbls');
      patchData(source, scatter ? 'xVal' : 'cat', was.categories, series.categories, cells);
      patchData(source, scatter ? 'yVal' : 'val', was.values, series.values, cells);
      if (series.bubbleSizes) patchData(source, 'bubbleSize', was.bubbleSizes, series.bubbleSizes, cells);
    });
    // A series of the file that the model no longer has goes.
    if (el.children.some((node) => node.type === 'element' && localName(node.name) === 'ser' && !kept.has(node))) {
      el.children = el.children.filter((node) => !(node.type === 'element' && localName(node.name) === 'ser' && !kept.has(node)));
    }
  }
}

// ------------------------------------------------------------------ drawing parts

function markerXmlOf(name: string, marker: DrawingMarker): string {
  return `<xdr:${name}><xdr:col>${marker.col}</xdr:col><xdr:colOff>${Math.round(marker.colOff)}</xdr:colOff><xdr:row>${marker.row}</xdr:row><xdr:rowOff>${Math.round(marker.rowOff)}</xdr:rowOff></xdr:${name}>`;
}

/** The elements that say where an anchor is, in the order the schema has them, and the local name of the anchor they belong to. */
function placeXml(anchor: DrawingAnchor): { name: string; inner: string } {
  if (anchor.kind === 'twoCell' && anchor.from && anchor.to) return { name: 'twoCellAnchor', inner: markerXmlOf('from', anchor.from) + markerXmlOf('to', anchor.to) };
  const size = `<xdr:ext cx="${Math.round(anchor.width || 0)}" cy="${Math.round(anchor.height || 0)}"/>`;
  if (anchor.kind === 'oneCell' && anchor.from) return { name: 'oneCellAnchor', inner: markerXmlOf('from', anchor.from) + size };
  return { name: 'absoluteAnchor', inner: `<xdr:pos x="${Math.round(anchor.x || 0)}" y="${Math.round(anchor.y || 0)}"/>${size}` };
}

/**
 * The anchor of a chart in a drawing part: a frame that names the chart part by a relationship.
 * `strict`: for a part of the strict flavour of the format, whose namespaces are others.
 */
export function chartAnchorXml(anchor: DrawingAnchor, relId: string, shapeId: number, name: string, description?: string, strict = false): string {
  const place = placeXml(anchor);
  const editAs = place.name === 'twoCellAnchor' && anchor.editAs && anchor.editAs !== 'twoCell' ? ` editAs="${anchor.editAs}"` : '';
  const described = description !== undefined && description !== '' ? ` descr="${attr(description)}"` : '';
  const charts = strict ? STRICT_C : NS_C;
  return (
    `<xdr:${place.name}${editAs}>${place.inner}` +
    `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${shapeId}" name="${attr(name)}"${described}/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' +
    `<a:graphic><a:graphicData uri="${charts}"><c:chart xmlns:c="${charts}" xmlns:r="${strict ? STRICT_R : NS_R}" r:id="${attr(relId)}"/></a:graphicData></a:graphic>` +
    `</xdr:graphicFrame><xdr:clientData/></xdr:${place.name}>`
  );
}

/**
 * The anchor of a picture in a drawing part: the picture names its bytes by a relationship of the
 * part (`relId`). Its aspect is locked, as Excel locks it for a picture it adds. `box` is where it
 * stands on the sheet in EMU, for the transform of the picture (an application places it by the
 * anchor; for a picture turned near a quarter it is the box turned back, as the picture is drawn
 * in it). `look` is how it is turned (`a:xfrm rot`, in 60,000ths of a degree), flipped and
 * cropped (`a:srcRect`, in 1,000ths of a percent). `strict`: for a part of the strict flavour of
 * the format.
 */
export function pictureAnchorXml(
  anchor: DrawingAnchor,
  relId: string,
  shapeId: number,
  name: string,
  description?: string,
  strict = false,
  box?: { x: number; y: number; width: number; height: number },
  look: PictureLook = {},
): string {
  const place = placeXml(anchor);
  const editAs = place.name === 'twoCellAnchor' && anchor.editAs && anchor.editAs !== 'twoCell' ? ` editAs="${anchor.editAs}"` : '';
  const described = description !== undefined && description !== '' ? ` descr="${attr(description)}"` : '';
  const emu = (value: number | undefined): number => (typeof value === 'number' && value > 0 && value < 2e12 ? Math.round(value) : 0);
  const at = box || { x: 0, y: 0, width: 0, height: 0 };
  const rot = rotationValue(look.rotation);
  const transform = `${rot !== undefined ? ` rot="${rot}"` : ''}${look.flipH ? ' flipH="1"' : ''}${look.flipV ? ' flipV="1"' : ''}`;
  return (
    `<xdr:${place.name}${editAs}>${place.inner}` +
    `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${shapeId}" name="${attr(name)}"${described}/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>` +
    `<xdr:blipFill><a:blip xmlns:r="${strict ? STRICT_R : NS_R}" r:embed="${attr(relId)}"/>${cropXml(look.crop)}<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
    `<xdr:spPr><a:xfrm${transform}><a:off x="${emu(at.x)}" y="${emu(at.y)}"/><a:ext cx="${emu(at.width)}" cy="${emu(at.height)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>` +
    `</xdr:pic><xdr:clientData/></xdr:${place.name}>`
  );
}

/** How a picture is turned, flipped and cropped, as its anchor in a drawing part says. */
export interface PictureLook {
  /** Degrees, clockwise. */
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
  /** Shares (0..1) cut off each side. */
  crop?: { left: number; top: number; right: number; bottom: number };
}

/** `rot` of `a:xfrm` for a rotation in degrees: 60,000ths of a degree from 0 to 21,599,999; undefined for none. */
function rotationValue(degrees: number | undefined): number | undefined {
  if (typeof degrees !== 'number' || !isFinite(degrees)) return undefined;
  const steps = Math.round((((degrees % 360) + 360) % 360) * 60000);
  return steps <= 0 || steps >= 21600000 ? undefined : steps;
}

/** The sides of `a:srcRect` for a crop: each share in 1,000ths of a percent (100,000 is all of it); undefined for a side that cuts nothing. */
function cropSides(crop: PictureLook['crop']): [name: string, value: string][] {
  if (!crop) return [];
  const out: [string, string][] = [];
  const side = (name: string, share: number) => {
    const value = typeof share === 'number' && isFinite(share) ? Math.round(Math.min(Math.max(share, 0), 1) * 100000) : 0;
    if (value > 0) out.push([name, String(value)]);
  };
  side('l', crop.left);
  side('t', crop.top);
  side('r', crop.right);
  side('b', crop.bottom);
  return out;
}

const cropXml = (crop: PictureLook['crop']): string => {
  const sides = cropSides(crop);
  return sides.length === 0 ? '' : `<a:srcRect ${sides.map(([name, value]) => `${name}="${value}"`).join(' ')}/>`;
};

/**
 * What a save changes of a picture of the file besides its anchor: each field that is there is
 * set (null takes it away). Everything else of the picture stays as the file has it.
 */
export interface PicturePatch {
  rotation?: number | null;
  crop?: PictureLook['crop'] | null;
  description?: string | null;
}

/** Applies a patch to the picture of an anchor of a drawing part that is being changed (see editDrawingPart()). */
function patchPicture(anchor: XmlElement, patch: PicturePatch): void {
  const context = part as PartContext;
  const pic = child(anchor, 'pic');
  if (!pic) throw new Error('Cannot write the drawing: the picture of the model is not a picture in its part');
  const link = (node: XmlElement, parent: XmlElement): XmlElement => {
    context.parents.set(node, parent);
    for (const kid of elements(node)) link(kid, node);
    return node;
  };
  if (patch.description !== undefined) {
    const props = child(child(pic, 'nvPicPr'), 'cNvPr');
    if (!props) throw new Error('Cannot write the drawing: a picture of its part has no properties');
    if (patch.description === null || patch.description === '') props.attrs.delete('descr');
    else props.attrs.set('descr', clean(patch.description));
  }
  if (patch.crop !== undefined) {
    const fill = child(pic, 'blipFill');
    if (!fill) throw new Error('Cannot write the drawing: a picture of its part has no fill');
    const sides = cropSides(patch.crop === null ? undefined : patch.crop);
    const rect = child(fill, 'srcRect');
    if (sides.length === 0) {
      if (rect) fill.children = fill.children.filter((node) => node !== rect);
    } else if (rect) {
      // The sides that are given; whatever else the element says stays.
      for (const name of ['l', 't', 'r', 'b']) rect.attrs.delete(name);
      const attrs = new Map<string, string>(sides);
      for (const [name, value] of rect.attrs) attrs.set(name, value);
      rect.attrs = attrs;
    } else {
      // After the blip, where the schema has it.
      const made_ = link(made(fill, cropXml(patch.crop as PictureLook['crop'])), fill);
      const blip = child(fill, 'blip');
      const at = blip ? fill.children.indexOf(blip) + 1 : 0;
      fill.children = fill.children.slice(0, at).concat([made_ as XmlNode], fill.children.slice(at));
    }
  }
  if (patch.rotation !== undefined) {
    const rot = rotationValue(patch.rotation === null ? undefined : patch.rotation);
    let props = child(pic, 'spPr');
    if (!props) {
      if (rot === undefined) return;
      props = link(made(pic, '<xdr:spPr/>'), pic);
      const fill = child(pic, 'blipFill');
      const at = fill ? pic.children.indexOf(fill) + 1 : pic.children.length;
      pic.children = pic.children.slice(0, at).concat([props as XmlNode], pic.children.slice(at));
    }
    const transform = child(props, 'xfrm');
    if (transform) {
      if (rot === undefined) transform.attrs.delete('rot');
      else {
        // `rot` first, where applications write it; whatever else the element says stays.
        const attrs = new Map<string, string>([['rot', String(rot)]]);
        for (const [name, value] of transform.attrs) if (name !== 'rot') attrs.set(name, value);
        transform.attrs = attrs;
      }
    } else if (rot !== undefined) {
      // A transform of the turn alone (off and ext may be left out): first of the properties, where the schema has it.
      const made_ = link(made(props, `<a:xfrm rot="${rot}"/>`), props);
      props.children = [made_ as XmlNode].concat(props.children);
    }
  }
}

/** A drawing part that holds the given anchors and nothing else. */
export function drawingPartXml(anchors: string[], strict = false): string {
  return `${DECLARATION}<xdr:wsDr xmlns:xdr="${strict ? STRICT_XDR : NS_XDR}" xmlns:a="${strict ? STRICT_A : NS_A}">${anchors.join('')}</xdr:wsDr>`;
}

/** The namespaces a drawing part writes its own elements and those of drawings in. */
function drawingSpaces(root: XmlElement, scope: Map<string, string>): { [prefix: string]: string } {
  const own = scope.get(prefixOf(root.name));
  const isStrict = own === STRICT_XDR;
  let drawings = isStrict ? STRICT_A : NS_A;
  for (const bound of scope.values()) if (bound === NS_A || bound === STRICT_A) drawings = bound;
  return { xdr: own === STRICT_XDR || own === NS_XDR ? own : NS_XDR, a: drawings };
}

const ANCHOR_NAMES: { [local: string]: true } = { twoCellAnchor: true, oneCellAnchor: true, absoluteAnchor: true };

interface FoundAnchor {
  el: XmlElement;
  /** The child of the root the anchor stands in: itself, or the wrapper (mc:AlternateContent) that holds it. */
  top: XmlElement;
}

/** The anchors of a drawing part in the order the reader takes them (see anchorElements() of the reader). */
function anchorsOf(parent: XmlElement, depth: number, top: XmlElement | undefined, out: FoundAnchor[]): void {
  for (const node of parent.children) {
    if (node.type !== 'element') continue;
    const local = localName(node.name);
    if (Object.prototype.hasOwnProperty.call(ANCHOR_NAMES, local)) out.push({ el: node, top: top || node });
    else if (local === 'AlternateContent' && depth < 8) {
      const before = out.length;
      for (const branch of node.children) {
        if (branch.type !== 'element' || localName(branch.name) !== 'Choice') continue;
        anchorsOf(branch, depth + 1, top || node, out);
        if (out.length > before) break;
      }
      if (out.length === before) {
        const fallback = child(node, 'Fallback');
        if (fallback) anchorsOf(fallback, depth + 1, top || node, out);
      }
    }
  }
}

/** What a save does to one drawing part of the file. */
export interface DrawingPartEdit {
  /** Anchors of the part, by their place among the anchors the reader takes, that get another place on the sheet. */
  moved: Map<number, DrawingAnchor>;
  /** Anchors that go. */
  removed: Set<number>;
  /**
   * Charts that are added: where each stands, the id of its relationship, its name. With `from`
   * (the anchor of the chart it is a copy of: see framedAnchorXml()) the anchor is made from
   * that one, and keeps everything of it but its place, its id, its name, its relationship and
   * the macro it is assigned.
   */
  added: { anchor: DrawingAnchor; relId: string; name: string; description?: string; from?: string; picture?: { box?: { x: number; y: number; width: number; height: number }; look?: PictureLook } }[];
  /** Pictures of the part, by their place among the anchors the reader takes, that are turned, cropped or described anew. */
  patched?: Map<number, PicturePatch>;
}

/** The <c:chart r:id> of a frame: the element that names the chart part, and the name of the attribute that does. */
function chartReference(anchor: XmlElement): { el: XmlElement; attr: string } | undefined {
  const stack: { el: XmlElement; depth: number; inData: boolean }[] = [{ el: anchor, depth: 0, inData: false }];
  while (stack.length > 0) {
    const { el, depth, inData } = stack.pop() as { el: XmlElement; depth: number; inData: boolean };
    const local = localName(el.name);
    if (inData && local === 'chart') {
      for (const name of el.attrs.keys()) if (name.indexOf(':') > 0 && name.slice(0, 6) !== 'xmlns:' && localName(name) === 'id') return { el, attr: name };
    }
    if (depth < 16) for (const node of el.children) if (node.type === 'element') stack.push({ el: node, depth: depth + 1, inData: local === 'graphicData' });
  }
  return undefined;
}

/**
 * The anchor of a drawing part whose frame shows the chart of the relationship `relId`, as XML
 * that stands on its own (every namespace its names use is declared on it): what the frame of a
 * copy of that chart is made from. An anchor inside a wrapper of alternatives is given without
 * the wrapper. Undefined when the part has no such anchor, or cannot be read.
 */
export function framedAnchorXml(bytes: Uint8Array, relId: string): string | undefined {
  let root: XmlElement;
  try {
    root = parseXml(bytes).root;
  } catch (error) {
    return undefined;
  }
  if (nestedBeyond(root, MAX_PATCH_DEPTH)) return undefined;
  const declared = (el: XmlElement, scope: Map<string, string>): Map<string, string> => {
    let own = scope;
    for (const [name, value] of el.attrs) {
      if (name !== 'xmlns' && name.slice(0, 6) !== 'xmlns:') continue;
      if (own === scope) own = new Map(scope);
      own.set(name === 'xmlns' ? '' : name.slice(6), value);
    }
    return own;
  };
  let found: { el: XmlElement; scope: Map<string, string> } | undefined;
  const walk = (parent: XmlElement, scope: Map<string, string>, depth: number): void => {
    for (const node of parent.children) {
      if (found || node.type !== 'element') continue;
      const local = localName(node.name);
      if (Object.prototype.hasOwnProperty.call(ANCHOR_NAMES, local)) {
        const reference = chartReference(node);
        if (reference && reference.el.attrs.get(reference.attr) === relId) found = { el: node, scope };
      } else if ((local === 'AlternateContent' || local === 'Choice' || local === 'Fallback') && depth < 16) walk(node, declared(node, scope), depth + 1);
    }
  };
  walk(root, declared(root, new Map()), 0);
  if (!found) return undefined;
  // The prefixes its names use, and the default namespace when a name has no prefix.
  const used = new Set<string>();
  const names = (el: XmlElement, depth: number): void => {
    used.add(prefixOf(el.name));
    for (const name of el.attrs.keys()) if (name.indexOf(':') > 0 && name.slice(0, 6) !== 'xmlns:') used.add(prefixOf(name));
    if (depth <= MAX_PATCH_DEPTH) for (const node of el.children) if (node.type === 'element') names(node, depth + 1);
  };
  names(found.el, 0);
  const attrs = new Map<string, string>();
  for (const prefix of used) {
    const key = prefix === '' ? 'xmlns' : `xmlns:${prefix}`;
    const space = found.scope.get(prefix);
    if (space !== undefined && prefix !== 'xml' && !found.el.attrs.has(key)) attrs.set(key, space);
  }
  for (const [name, value] of found.el.attrs) attrs.set(name, value);
  return serializeXml({ declaration: null, leading: [], root: { type: 'element', name: found.el.name, attrs, children: found.el.children }, trailing: [] });
}

/**
 * The anchor of a chart that is a copy, made from the anchor of the chart it was copied from
 * (`item.from`): the place, the id of the shape, its name and the relationship of the chart are
 * the copy's; the macro the frame was assigned is not carried (a copy is not bound to a macro
 * behind the back of whoever made it); what else of the frame names a relationship of the
 * drawing it came from (a link) is left out, since the copy may stand in another drawing.
 * Everything else stays. Undefined when the anchor is not one of a chart.
 */
function copiedAnchor(root: XmlElement, item: DrawingPartEdit['added'][number], shapeId: number): XmlElement | undefined {
  const context = part as PartContext;
  let el: XmlElement;
  try {
    el = parseXml(item.from as string).root;
  } catch (error) {
    return undefined;
  }
  if (!Object.prototype.hasOwnProperty.call(ANCHOR_NAMES, localName(el.name)) || nestedBeyond(el, MAX_PATCH_DEPTH)) return undefined;
  const reference = chartReference(el);
  if (!reference) return undefined;
  const link = (node: XmlElement, parent: XmlElement, depth: number): void => {
    context.parents.set(node, parent);
    if (depth <= MAX_PATCH_DEPTH) for (const child_ of node.children) if (child_.type === 'element') link(child_, node, depth + 1);
  };
  link(el, root, 0);
  reference.el.attrs.set(reference.attr, item.relId);
  // Whatever else names a relationship named one of the drawing the frame came from.
  const prune = (node: XmlElement, depth: number): void => {
    node.children = node.children.filter((kid) => {
      if (kid.type !== 'element' || kid === reference.el) return true;
      for (const name of kid.attrs.keys()) if (name.indexOf(':') > 0 && name.slice(0, 6) !== 'xmlns:' && /^(id|embed|link|pict|dm|lo|qs|cs)$/.test(localName(name))) return false;
      return true;
    });
    if (depth <= MAX_PATCH_DEPTH) for (const kid of node.children) if (kid.type === 'element') prune(kid, depth + 1);
  };
  prune(el, 0);
  const frame = child(el, 'graphicFrame');
  if (frame && frame.attrs.has('macro')) frame.attrs.set('macro', '');
  const props = child(child(frame, 'nvGraphicFramePr'), 'cNvPr');
  if (props) {
    props.attrs.set('id', String(shapeId));
    props.attrs.set('name', item.name);
  }
  // Where it stands, in the prefix the anchor has; whether it moves and sizes with the cells is the copy's as well.
  const place = placeXml(item.anchor);
  const holder = made(el, `<xdr:w>${place.inner}</xdr:w>`);
  const kept = el.children.filter((node) => !(node.type === 'element' && /^(from|to|pos|ext)$/.test(localName(node.name))));
  const placed = elements(holder);
  for (const node of placed) context.parents.set(node, el);
  el.children = (placed as XmlNode[]).concat(kept);
  el.name = el.name.slice(0, el.name.length - localName(el.name).length) + place.name;
  el.attrs.delete('editAs');
  if (place.name === 'twoCellAnchor' && item.anchor.editAs && item.anchor.editAs !== 'twoCell') {
    // Before what the anchor declares, where applications write it.
    const attrs = new Map<string, string>([['editAs', item.anchor.editAs]]);
    for (const [name, value] of el.attrs) attrs.set(name, value);
    el.attrs = attrs;
  }
  return el;
}

/**
 * A drawing part with anchors moved, taken out and added; every other node of it stays. Also
 * says which relationships the anchors that went named and nothing that stays names (`unrelated`),
 * and whether anything is left to draw (`empty`: no anchor is). Throws when the part does not
 * hold an anchor that is to be changed.
 */
export function editDrawingPart(bytes: Uint8Array, edit: DrawingPartEdit): { bytes: Uint8Array; unrelated: string[]; empty: boolean } {
  const before = part;
  try {
    const { doc, context } = openPart(bytes, drawingSpaces);
    const root = doc.root;
    const strict = context.spaces.xdr === STRICT_XDR;
    const found: FoundAnchor[] = [];
    anchorsOf(root, 0, undefined, found);
    const relIds = relationIdsOf;
    const gone = new Set<string>();
    for (const [index, anchor] of edit.moved) {
      const item = found[index];
      if (!item) throw new Error('Cannot write the drawing: an anchor of the model is not in its part');
      const el = item.el;
      const place = placeXml(anchor);
      // The elements that say where it is, in the prefix and namespace the anchor has; everything else of it stays.
      const holder = made(el, `<xdr:w>${place.inner}</xdr:w>`);
      const kept = el.children.filter((node) => !(node.type === 'element' && /^(from|to|pos|ext)$/.test(localName(node.name))));
      const placed = elements(holder);
      for (const node of placed) context.parents.set(node, el);
      el.children = (placed as XmlNode[]).concat(kept);
      el.name = el.name.slice(0, el.name.length - localName(el.name).length) + place.name;
      if (place.name !== 'twoCellAnchor') el.attrs.delete('editAs');
    }
    if (edit.patched) {
      for (const [index, patch] of edit.patched) {
        const item = found[index];
        if (!item) throw new Error('Cannot write the drawing: an anchor of the model is not in its part');
        patchPicture(item.el, patch);
      }
    }
    const removed = new Set<XmlNode>();
    for (const index of edit.removed) {
      const item = found[index];
      if (!item) throw new Error('Cannot write the drawing: an anchor of the model is not in its part');
      // The wrapper goes with its anchor, and with what stands in for the anchor in its other branches;
      // an anchor that shares its wrapper with another one is not taken out (the reader says so: SheetDrawing.wrapped).
      if (found.some((other, at) => other.top === item.top && other !== item && !edit.removed.has(at))) throw new Error('Cannot write the drawing: an anchor that goes shares its wrapper with one that stays');
      relIds(item.top, gone);
      removed.add(item.top);
    }
    if (removed.size > 0) root.children = root.children.filter((node) => !removed.has(node));
    if (edit.added.length > 0) {
      let top = 1;
      const ids = (el: XmlElement, depth: number): void => {
        if (localName(el.name) === 'cNvPr') {
          const id = Number(el.attrs.get('id'));
          if (id > top && id < 4294967295) top = id;
        }
        if (depth <= MAX_PATCH_DEPTH) for (const node of el.children) if (node.type === 'element') ids(node, depth + 1);
      };
      ids(root, 0);
      const added: XmlNode[] = edit.added.map((item) => {
        const id = ++top;
        if (item.picture) return made(root, pictureAnchorXml(item.anchor, item.relId, id, item.name, item.description, strict, item.picture.box, item.picture.look));
        return (item.from !== undefined ? copiedAnchor(root, item, id) : undefined) || made(root, chartAnchorXml(item.anchor, item.relId, id, item.name, item.description, strict));
      });
      root.children = root.children.concat(added);
    }
    const stay = new Set<string>();
    relIds(root, stay);
    const left: FoundAnchor[] = [];
    anchorsOf(root, 0, undefined, left);
    return { bytes: closePart(doc, context), unrelated: [...gone].filter((id) => !stay.has(id)), empty: left.length === 0 };
  } finally {
    part = before;
  }
}

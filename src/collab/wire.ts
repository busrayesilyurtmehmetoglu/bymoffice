// Wire operations: made from the operations of the model, checked when they arrive, and
// applied. Applying one is a function of the workbook and the operation alone, so that
// participants who apply the same operations in the same order have the same workbook.
//
// Nothing that arrives is trusted: every field is read by its type and its range into a new
// object, and what is applied goes through applyOp() and the commands, which keep the model
// whole (the edges of the sheet, array formulas, merged cells, the names of sheets).

import { adjustFormula, adjustRange, renameSheetsInFormula, type StructuralEdit } from '../calc/refs';
import {
  MAX_COLS,
  MAX_ROWS,
  cellName,
  foldSheetName,
  parseCellRef,
  parseRange,
  type PaneState,
  type SheetNameMap,
  type Alignment,
  type AutoFilter,
  type Border,
  type BorderEdge,
  type Cell,
  type CellFormat,
  type CellRange,
  type Color,
  type ColumnRange,
  type Fill,
  type FilterColumn,
  type Font,
  type Protection,
  type RichRun,
  type Sheet,
  type SheetComment,
  type SortState,
  type StyleSheet,
  type Workbook,
} from '../model/sheet/types';
import { addChart, moveChart, pasteChartOf, removeChart, updateChart } from '../ops/charts';
import { DEFAULT_NOTE_AUTHOR, MAX_NOTE_TEXT, MAX_SHEET_NOTES, cellNow, notesBlocked } from '../ops/notes';
import { pivotSourceLostBy } from '../ops/pivot-places';
import { MAX_CELL_FORMATS, applyOp, ensureNumberFormat, explicitFormula, sameFormat, sharedMasters, shiftAutoFilter, type CellChange, type Op, type RowProps, type SharedMasters } from '../ops/ops';
import { addConditionalRule, clearConditionalRules, moveConditionalRule, removeConditionalRule, setValidation, updateConditionalRule } from '../ops/rules';
import { CollabUnsupportedError, WireError, type Rec, type Shift, type WireCellPatch, type WireColumn, type WireContent, type WireNote, type WireOp, type WireRange, type WireRowProps } from './types';

/** Most operations one step may hold (the relay refuses more). */
export const MAX_STEP_OPS = 10000;
/** Most cells the undoing of a deletion of rows or columns brings back; above it the deletion cannot be undone. */
export const MAX_UNDO_CELLS = 200000;
const MAX_CELL_TEXT = 32767;
const MAX_FORMULA_TEXT = 8192;
const MAX_SHEETS = 4096;
/** Most merged ranges, notes, runs of columns and formats one operation or step may bring: applying them is not free. */
export const MAX_STEP_MERGES = 2000;
export const MAX_STEP_NOTES = 5000;
export const MAX_STEP_RUNS = 1024;
export const MAX_STEP_FORMATS = 256;
/** What a value that is passed on as it is (the arguments of a command) may be made of. */
const MAX_PLAIN_DEPTH = 12;
const MAX_PLAIN_NODES = 100000;

// ---------------------------------------------------------------------------
// Reading what arrived

const fail = (what: string): never => {
  throw new WireError(`Invalid ${what}`);
};

function int(value: unknown, low: number, high: number, what: string): number {
  if (typeof value !== 'number' || Math.floor(value) !== value || value < low || value > high) return fail(what);
  return value;
}

function num(value: unknown, low: number, high: number, what: string): number {
  if (typeof value !== 'number' || !(value >= low && value <= high)) return fail(what);
  return value;
}

/**
 * Whether a text can stand in an XML file as it is: no control character but tab and line
 * breaks, no half of a surrogate pair, none of the two characters that are no characters.
 * Every text that arrives ends up in a file; what a file cannot hold is no text of a step.
 */
export function xmlText(value: string): boolean {
  for (let i = 0; i < value.length; i++) {
    const c = value.charCodeAt(i);
    if (c < 0x20) {
      if (c !== 9 && c !== 10 && c !== 13) return false;
    } else if (c >= 0xd800 && c <= 0xdbff) {
      const next = value.charCodeAt(i + 1);
      if (!(next >= 0xdc00 && next <= 0xdfff)) return false;
      i++;
    } else if ((c >= 0xdc00 && c <= 0xdfff) || c === 0xfffe || c === 0xffff) return false;
  }
  return true;
}

function str(value: unknown, max: number, what: string): string {
  if (typeof value !== 'string' || value.length > max || !xmlText(value)) return fail(what);
  return value;
}

/** Whether a cell is one that scrolls with these rows and columns frozen (none frozen: no such cell is stated). */
function scrollsAt(rows: number, cols: number, ref: string): boolean {
  const at = parseCellRef(ref);
  return !!at && (rows > 0 || cols > 0) && at.row >= rows && at.col >= cols;
}

/**
 * The panes a "pane" step sets: the first cell that scrolls is the one it states, or the one right below and beside
 * the frozen ones (as freezePanes() makes it).
 */
export function paneOfWire(op: { r: number; c: number; tl?: string; fv?: string }): PaneState {
  const pane: PaneState = { frozenRows: op.r, frozenCols: op.c };
  if (op.r > 0 || op.c > 0) pane.topLeftCell = op.tl !== undefined ? op.tl : cellName(op.r, op.c);
  if (op.fv !== undefined) pane.firstVisibleCell = op.fv;
  return pane;
}

/** The name of one cell ("B2"). */
function cellRef(value: unknown, what: string): string {
  if (typeof value !== 'string' || value.length > 10 || !parseCellRef(value)) return fail(what);
  return value;
}

function bool(value: unknown, what: string): boolean {
  if (typeof value !== 'boolean') return fail(what);
  return value;
}

function arr(value: unknown, max: number, what: string): unknown[] {
  if (!Array.isArray(value) || value.length > max) return fail(what);
  return value;
}

function rec(value: unknown, what: string): Record<string, unknown> {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) return fail(what);
  return value as Record<string, unknown>;
}

const has = (object: Record<string, unknown>, key: string): boolean => Object.prototype.hasOwnProperty.call(object, key) && object[key] !== undefined;

function oneOf<T extends string>(value: unknown, allowed: readonly T[], what: string): T {
  if (typeof value !== 'string' || allowed.indexOf(value as T) < 0) return fail(what);
  return value as T;
}

/** A copy of a value made of nothing but what JSON has, within limits; names that reach into prototypes are refused. */
export function plainCopy(value: unknown, what: string): unknown {
  let nodes = 0;
  const copy = (item: unknown, depth: number): unknown => {
    if (++nodes > MAX_PLAIN_NODES || depth > MAX_PLAIN_DEPTH) return fail(what);
    if (item === null || typeof item === 'boolean') return item;
    if (typeof item === 'number') return isFinite(item) ? item : fail(what);
    if (typeof item === 'string') return item.length <= MAX_CELL_TEXT && xmlText(item) ? item : fail(what);
    if (Array.isArray(item)) return item.map((child) => (child === undefined ? null : copy(child, depth + 1)));
    if (typeof item !== 'object') return fail(what);
    const proto = Object.getPrototypeOf(item);
    if (proto !== Object.prototype && proto !== null) return fail(what);
    const out: Record<string, unknown> = {};
    for (const key of Object.keys(item as object)) {
      if (key === '__proto__' || key === 'constructor' || key === 'prototype' || key.length > 64 || !xmlText(key)) return fail(what);
      const child = (item as Record<string, unknown>)[key];
      if (child !== undefined) out[key] = copy(child, depth + 1);
    }
    return out;
  };
  return copy(value, 0);
}

function readColor(value: unknown): Color {
  const from = rec(value, 'colour');
  const color: Color = {};
  if (has(from, 'rgb')) {
    color.rgb = str(from.rgb, 8, 'colour');
    if (!/^[0-9A-Fa-f]{6}([0-9A-Fa-f]{2})?$/.test(color.rgb)) fail('colour');
  }
  if (has(from, 'theme')) color.theme = int(from.theme, 0, 255, 'colour');
  if (has(from, 'tint')) color.tint = num(from.tint, -1, 1, 'colour');
  if (has(from, 'indexed')) color.indexed = int(from.indexed, 0, 255, 'colour');
  if (has(from, 'auto')) color.auto = bool(from.auto, 'colour');
  return color;
}

const UNDERLINES = ['none', 'single', 'double', 'singleAccounting', 'doubleAccounting'] as const;

function readFont(value: unknown): Font {
  const from = rec(value, 'font');
  const font: Font = { bold: bool(from.bold, 'font'), italic: bool(from.italic, 'font'), strike: bool(from.strike, 'font'), underline: oneOf(from.underline, UNDERLINES, 'font') };
  if (has(from, 'name')) font.name = str(from.name, 255, 'font');
  if (has(from, 'size')) font.size = num(from.size, 1, 409.55, 'font');
  if (has(from, 'color')) font.color = readColor(from.color);
  if (has(from, 'vertAlign')) font.vertAlign = oneOf(from.vertAlign, ['superscript', 'subscript'] as const, 'font');
  if (has(from, 'family')) font.family = int(from.family, 0, 255, 'font');
  if (has(from, 'charset')) font.charset = int(from.charset, 0, 255, 'font');
  if (has(from, 'scheme')) font.scheme = oneOf(from.scheme, ['major', 'minor', 'none'] as const, 'font');
  return font;
}

function readRuns(value: unknown): RichRun[] {
  return arr(value, 4096, 'runs').map((item) => {
    const from = rec(item, 'run');
    const run: RichRun = { text: str(from.text, MAX_CELL_TEXT, 'run') };
    if (has(from, 'font')) run.font = readFont(from.font);
    return run;
  });
}

/** The patterns of a fill and the styles of a border edge, as the schema names them: nothing else is written into a file. */
const PATTERNS = ['none', 'solid', 'mediumGray', 'darkGray', 'lightGray', 'darkHorizontal', 'darkVertical', 'darkDown', 'darkUp', 'darkGrid', 'darkTrellis', 'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp', 'lightGrid', 'lightTrellis', 'gray125', 'gray0625'] as const;
const EDGES = ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot'] as const;

function readFill(value: unknown): Fill {
  const from = rec(value, 'fill');
  // A gradient is kept as the element the file has, which nothing that arrives may bring: it does not travel.
  if (from.type !== 'pattern') return fail('fill');
  const fill: Fill = { type: 'pattern', pattern: oneOf(from.pattern, PATTERNS, 'fill') };
  if (has(from, 'fg')) fill.fg = readColor(from.fg);
  if (has(from, 'bg')) fill.bg = readColor(from.bg);
  return fill;
}

function readEdge(value: unknown): BorderEdge {
  const from = rec(value, 'border');
  const edge: BorderEdge = { style: oneOf(from.style, EDGES, 'border') };
  if (has(from, 'color')) edge.color = readColor(from.color);
  return edge;
}

function readBorder(value: unknown): Border {
  const from = rec(value, 'border');
  const border: Border = { diagonalUp: bool(from.diagonalUp, 'border'), diagonalDown: bool(from.diagonalDown, 'border') };
  for (const side of ['left', 'right', 'top', 'bottom', 'diagonal'] as const) if (has(from, side)) border[side] = readEdge(from[side]);
  return border;
}

const HORIZONTAL = ['general', 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'] as const;
const VERTICAL = ['top', 'center', 'bottom', 'justify', 'distributed'] as const;

function readAlignment(value: unknown): Alignment {
  const from = rec(value, 'alignment');
  const alignment: Alignment = {};
  if (has(from, 'horizontal')) alignment.horizontal = oneOf(from.horizontal, HORIZONTAL, 'alignment');
  if (has(from, 'vertical')) alignment.vertical = oneOf(from.vertical, VERTICAL, 'alignment');
  if (has(from, 'wrapText')) alignment.wrapText = bool(from.wrapText, 'alignment');
  if (has(from, 'shrinkToFit')) alignment.shrinkToFit = bool(from.shrinkToFit, 'alignment');
  if (has(from, 'textRotation')) alignment.textRotation = int(from.textRotation, 0, 255, 'alignment');
  if (has(from, 'indent')) alignment.indent = int(from.indent, 0, 255, 'alignment');
  if (has(from, 'readingOrder')) alignment.readingOrder = int(from.readingOrder, 0, 2, 'alignment');
  return alignment;
}

function readProtection(value: unknown): Protection {
  const from = rec(value, 'protection');
  const protection: Protection = {};
  if (has(from, 'locked')) protection.locked = bool(from.locked, 'protection');
  if (has(from, 'hidden')) protection.hidden = bool(from.hidden, 'protection');
  return protection;
}

// ---------------------------------------------------------------------------
// Formats

/** A cell format written out: what a message brings along for the cells it formats. */
export interface WireFormat {
  /** The number format: one of the built-in ones by its number, any other by its code. */
  n: number | string;
  f: Font;
  l: Fill;
  b: Border;
  /** The named style the format belongs to, by its place in the file. */
  x?: number;
  a?: Alignment;
  p?: Protection;
  q?: true;
}

const FIRST_CUSTOM_FORMAT = 164;
const DEFAULT_FONT: Font = { bold: false, italic: false, strike: false, underline: 'none' };
const DEFAULT_FILL: Fill = { type: 'pattern', pattern: 'none' };
const DEFAULT_BORDER: Border = { diagonalUp: false, diagonalDown: false };

/** A value as text, with the members of every object in the order of their names. */
export function stable(value: unknown): string {
  if (Array.isArray(value)) return `[${value.map(stable).join(',')}]`;
  if (value instanceof Map) return stable(Array.from(value.entries()));
  if (!value || typeof value !== 'object') return JSON.stringify(value === undefined ? null : value);
  const object = value as Record<string, unknown>;
  return `{${Object.keys(object)
    .filter((name) => object[name] !== undefined)
    .sort()
    .map((name) => `${JSON.stringify(name)}:${stable(object[name])}`)
    .join(',')}}`;
}

/** The format at an index of the cell formats, written out. */
export function wireFormatOf(styles: StyleSheet, index: number): WireFormat {
  const format = styles.cellXfs[index] || styles.cellXfs[0];
  if (!format) return { n: 0, f: DEFAULT_FONT, l: DEFAULT_FILL, b: DEFAULT_BORDER };
  const code = styles.numFmts.get(format.numFmtId);
  const fill = styles.fills[format.fillId] || DEFAULT_FILL;
  // Kept as the element of the file, which is not sent: a cell with such a fill keeps its format while the workbook is edited together.
  if (fill.type === 'gradient') throw new CollabUnsupportedError('other', 'The format of a cell with a gradient fill cannot be changed while the workbook is edited together');
  const out: WireFormat = {
    n: code !== undefined ? code : format.numFmtId < FIRST_CUSTOM_FORMAT ? format.numFmtId : 0,
    f: styles.fonts[format.fontId] || DEFAULT_FONT,
    l: fill,
    b: styles.borders[format.borderId] || DEFAULT_BORDER,
  };
  if (format.xfId !== undefined) out.x = format.xfId;
  if (format.alignment) out.a = format.alignment;
  if (format.protection) out.p = format.protection;
  if (format.quotePrefix) out.q = true;
  return out;
}

function readFormat(value: unknown): WireFormat {
  const from = rec(value, 'format');
  const out: WireFormat = {
    n: typeof from.n === 'number' ? int(from.n, 0, FIRST_CUSTOM_FORMAT - 1, 'format') : str(from.n, 255, 'format'),
    f: readFont(from.f),
    l: readFill(from.l),
    b: readBorder(from.b),
  };
  if (has(from, 'x')) out.x = int(from.x, 0, 65535, 'format');
  if (has(from, 'a')) out.a = readAlignment(from.a);
  if (has(from, 'p')) out.p = readProtection(from.p);
  if (from.q === true) out.q = true;
  return out;
}

/** The table of cell formats of this workbook has no room for the formats of a step. */
export class StyleTableFull extends Error {
  constructor() {
    super('The workbook has too many cell formats');
    this.name = 'StyleTableFull';
  }
}

/** The texts of the records of a style table, for finding one that is there already; kept beside the table. */
const KEYS = new WeakMap<object, string[]>();

function internRecord<T>(list: T[], value: T, styles: StyleSheet): number {
  let keys = KEYS.get(list);
  // A table that got shorter (formats taken back, see resetStyles()) is read anew.
  if (!keys || keys.length > list.length) {
    keys = [];
    KEYS.set(list, keys);
  }
  for (let i = keys.length; i < list.length; i++) keys.push(stable(list[i]));
  const key = stable(value);
  const found = keys.indexOf(key);
  if (found >= 0) return found;
  list.push(value);
  keys.push(key);
  styles.modified = true;
  return list.length - 1;
}

const XF_KEYS = new WeakMap<object, { count: number; at: Map<string, number> }>();
const flat = (record: object | undefined): string => stable(record || {});
const formatKey = (format: CellFormat): string => `${format.numFmtId}|${format.fontId}|${format.fillId}|${format.borderId}|${format.xfId}|${format.quotePrefix}|${flat(format.alignment)}|${flat(format.protection)}`;

/** The index of a format that arrived among the cell formats of the workbook; added when it is not there. */
export function internFormat(styles: StyleSheet, format: WireFormat): number {
  const numFmtId = typeof format.n === 'number' ? format.n : ensureNumberFormat(styles, format.n);
  const wanted: CellFormat = {
    numFmtId,
    fontId: internRecord(styles.fonts, format.f, styles),
    fillId: internRecord(styles.fills, format.l, styles),
    borderId: internRecord(styles.borders, format.b, styles),
    quotePrefix: format.q === true,
  };
  if (format.x !== undefined && format.x < styles.cellStyleXfs.length) wanted.xfId = format.x;
  if (format.a) wanted.alignment = format.a;
  if (format.p) wanted.protection = format.p;
  // Found by a text of the record, kept beside the table (as for fonts): a step may bring hundreds of formats.
  const list = styles.cellXfs;
  let known = XF_KEYS.get(list);
  if (!known || known.count > list.length) {
    known = { count: 0, at: new Map() };
    XF_KEYS.set(list, known);
  }
  for (; known.count < list.length; known.count++) {
    const key = formatKey(list[known.count]);
    if (!known.at.has(key)) known.at.set(key, known.count);
  }
  const found = known.at.get(formatKey(wanted));
  if (found !== undefined && sameFormat(list[found], wanted)) return found;
  if (list.length >= MAX_CELL_FORMATS) throw new StyleTableFull();
  styles.modified = true;
  return list.push(wanted) - 1;
}

/** Every place of an operation that names a format, read and written through `map`. */
function mapStyles(op: WireOp, map: (index: number) => number): WireOp {
  if (op.t === 'cells') {
    let changed = false;
    const cells = op.c.map((entry) => {
      const patch = entry[2];
      if (patch === 0 || patch.s === undefined) return entry;
      changed = true;
      return [entry[0], entry[1], { ...patch, s: map(patch.s) }] as [number, number, WireCellPatch];
    });
    return changed ? { ...op, c: cells } : op;
  }
  if (op.t === 'rows') {
    return { ...op, r: op.r.map((entry) => (typeof entry[1].s === 'number' ? ([entry[0], { ...entry[1], s: map(entry[1].s) }] as [number, WireRowProps]) : entry)) };
  }
  if (op.t === 'cols') {
    return { ...op, p: op.p.map((entry) => (entry[2] !== 0 && entry[2].s !== undefined ? ([entry[0], entry[1], { ...entry[2], s: map(entry[2].s) }] as [number, number, WireColumn | 0]) : entry)) };
  }
  return op;
}

/**
 * The operations of a step as they go on the wire: an array of JSON values. Formats are
 * written out once, in a first element `{ t: "st", f: [...] }`, and named by their place in it.
 */
export function encodeStep(ops: readonly WireOp[], styles: StyleSheet): unknown[] {
  const table: WireFormat[] = [];
  const places = new Map<number, number>();
  const out: unknown[] = ops.map((op) =>
    mapStyles(op, (index) => {
      let place = places.get(index);
      if (place === undefined) {
        place = table.length;
        table.push(wireFormatOf(styles, index));
        places.set(index, place);
      }
      return place;
    }),
  );
  if (table.length > 0) out.unshift({ t: 'st', f: table });
  return out;
}

// ---------------------------------------------------------------------------
// Reading operations

const KINDS = ['blank', 'number', 'string', 'boolean', 'error', 'date'] as const;

/** The values an error cell holds. */
const ERRORS = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA', '#SPILL!', '#CALC!', '#FIELD!', '#BLOCKED!', '#CONNECT!', '#UNKNOWN!'] as const;
/** The attributes of a formula that travel: flags, which say nothing about other cells. */
const FORMULA_FLAGS = ['aca', 'ca', 'bx'] as const;

function readContent(value: unknown): WireContent {
  const from = rec(value, 'cell');
  const kind = oneOf(from.k, KINDS, 'cell');
  let held: WireContent['v'] = null;
  if (kind === 'number') held = num(from.v, -Number.MAX_VALUE, Number.MAX_VALUE, 'cell');
  else if (kind === 'boolean') held = bool(from.v, 'cell');
  else if (kind === 'string') held = str(from.v, MAX_CELL_TEXT, 'cell');
  else if (kind === 'error') held = oneOf(from.v, ERRORS, 'cell');
  else if (kind === 'date') {
    held = str(from.v, 40, 'cell');
    if (!/^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}(:\d{2}(\.\d{1,9})?)?)?Z?$/.test(held)) fail('cell');
  }
  else if (from.v !== null && from.v !== undefined) fail('cell');
  const content: WireContent = { k: kind, v: held };
  if (has(from, 'f')) {
    const f = rec(from.f, 'formula');
    const formula: WireContent['f'] = { x: str(f.x, MAX_FORMULA_TEXT, 'formula') };
    // A formula travels as the formula of its cell alone: what makes it one of a group (an array, a data
    // table, a shared formula) is built by the model and is taken from nobody.
    if (has(f, 't') || has(f, 'r')) fail('formula');
    if (has(f, 'a')) {
      formula.a = arr(f.a, FORMULA_FLAGS.length, 'formula').map((pair) => {
        const item = arr(pair, 2, 'formula');
        return [oneOf(item[0], FORMULA_FLAGS, 'formula'), oneOf(item[1], ['1', '0', 'true', 'false'] as const, 'formula')] as [string, string];
      });
    }
    if (has(f, 'd')) {
      if (f.d !== 1) fail('formula');
      formula.d = 1;
    }
    content.f = formula;
  }
  if (has(from, 'r')) content.r = readRuns(from.r);
  if (from.ph === true) content.ph = true;
  return content;
}

function readRange(value: unknown): WireRange {
  const item = arr(value, 4, 'range');
  if (item.length !== 4) fail('range');
  const top = int(item[0], 0, MAX_ROWS - 1, 'range');
  const left = int(item[1], 0, MAX_COLS - 1, 'range');
  const bottom = int(item[2], top, MAX_ROWS - 1, 'range');
  const right = int(item[3], left, MAX_COLS - 1, 'range');
  return [top, left, bottom, right];
}

const toRange = (range: WireRange): CellRange => ({ top: range[0], left: range[1], bottom: range[2], right: range[3] });
const fromRange = (range: CellRange): WireRange => [range.top, range.left, range.bottom, range.right];

function readModelRange(value: unknown): CellRange {
  const from = rec(value, 'range');
  return toRange(readRange([from.top, from.left, from.bottom, from.right]));
}

const FILTER_KINDS = ['values', 'custom', 'top10', 'dynamic', 'color', 'icon', 'none', 'unknown'] as const;

/** A filter as it travels: the fields of the model, without what a file holds beside them. */
function readFilter(value: unknown): AutoFilter {
  const from = rec(value, 'filter');
  const range = readModelRange(from.range);
  const width = range.right - range.left + 1;
  const seen = new Set<number>();
  const columns: FilterColumn[] = arr(from.columns, MAX_COLS, 'filter').map((item) => {
    const c = rec(item, 'filter');
    const column: FilterColumn = { colId: int(c.colId, 0, width - 1, 'filter'), kind: oneOf(c.kind, FILTER_KINDS, 'filter') };
    if (seen.has(column.colId)) fail('filter');
    seen.add(column.colId);
    if (has(c, 'values')) column.values = arr(c.values, 100000, 'filter').map((text) => str(text, MAX_CELL_TEXT, 'filter'));
    if (has(c, 'blank')) column.blank = bool(c.blank, 'filter');
    if (has(c, 'dateGroups')) {
      column.dateGroups = arr(c.dateGroups, 100000, 'filter').map((group) => {
        const g = rec(group, 'filter');
        const out: NonNullable<FilterColumn['dateGroups']>[number] = { grouping: oneOf(g.grouping, ['year', 'month', 'day', 'hour', 'minute', 'second'] as const, 'filter'), year: int(g.year, 0, 9999, 'filter') };
        for (const part of ['month', 'day', 'hour', 'minute', 'second'] as const) if (has(g, part)) out[part] = int(g[part], 0, 99, 'filter');
        return out;
      });
    }
    if (has(c, 'and')) column.and = bool(c.and, 'filter');
    if (has(c, 'conditions')) {
      column.conditions = arr(c.conditions, 2, 'filter').map((condition) => {
        const k = rec(condition, 'filter');
        return { operator: oneOf(k.operator, ['equal', 'notEqual', 'greaterThan', 'greaterThanOrEqual', 'lessThan', 'lessThanOrEqual'] as const, 'filter'), value: str(k.value, MAX_CELL_TEXT, 'filter') };
      });
    }
    if (has(c, 'top')) column.top = bool(c.top, 'filter');
    if (has(c, 'percent')) column.percent = bool(c.percent, 'filter');
    if (has(c, 'value')) column.value = num(c.value, -Number.MAX_VALUE, Number.MAX_VALUE, 'filter');
    if (has(c, 'dynamicType')) {
      column.dynamicType = str(c.dynamicType, 64, 'filter');
      if (!/^[A-Za-z0-9]+$/.test(column.dynamicType)) fail('filter');
    }
    if (has(c, 'hiddenButton')) column.hiddenButton = bool(c.hiddenButton, 'filter');
    return column;
  });
  const filter: AutoFilter = { range, columns };
  if (has(from, 'sortState')) {
    const s = rec(from.sortState, 'filter');
    const state: SortState = {
      range: readModelRange(s.range),
      conditions: arr(s.conditions, 64, 'filter').map((condition) => {
        const k = rec(condition, 'filter');
        return { range: readModelRange(k.range), descending: bool(k.descending, 'filter') };
      }),
    };
    if (has(s, 'caseSensitive')) state.caseSensitive = bool(s.caseSensitive, 'filter');
    if (has(s, 'columnSort')) state.columnSort = bool(s.columnSort, 'filter');
    filter.sortState = state;
  }
  if (has(from, 'held')) filter.held = arr(from.held, MAX_ROWS, 'filter').map((row) => int(row, 0, MAX_ROWS - 1, 'filter'));
  return filter;
}

/** The fields of a filter that travel (see readFilter()). */
function filterFields(filter: AutoFilter): AutoFilter {
  const columns = filter.columns.map((column) => {
    const out = { ...column };
    delete out.raw;
    return out;
  });
  const out: AutoFilter = { range: { ...filter.range }, columns };
  if (filter.sortState) {
    const state = filter.sortState;
    out.sortState = { range: { ...state.range }, conditions: state.conditions.map((condition) => ({ range: { ...condition.range }, descending: condition.descending })) };
    if (state.caseSensitive !== undefined) out.sortState.caseSensitive = state.caseSensitive;
    if (state.columnSort !== undefined) out.sortState.columnSort = state.columnSort;
  }
  if (filter.held && filter.held.length > 0) out.held = filter.held.slice();
  return out;
}

function readRows(value: unknown, what: string): number[] {
  return arr(value, MAX_ROWS, what).map((row) => int(row, 0, MAX_ROWS - 1, what));
}

/** The commands of pictures (see ops/pictures.ts): refused in a room. */
const PICTURE_COMMANDS = ['addPicture', 'movePicture', 'removePicture'];

/** The names of the commands that travel, and how many arguments each takes at most. */
const COMMANDS: { [name: string]: { family: 'validation' | 'conditional' | 'chart'; run: (workbook: Workbook, sheet: number, args: unknown[]) => Op[] } } = {
  setValidation: { family: 'validation', run: (wb, s, a) => setValidation(wb, s, a[0] as never, (a[1] === null ? undefined : a[1]) as never) },
  addConditionalRule: { family: 'conditional', run: (wb, s, a) => addConditionalRule(wb, s, a[0] as never, a[1] as never) },
  updateConditionalRule: { family: 'conditional', run: (wb, s, a) => updateConditionalRule(wb, s, a[0] as never, a[1] as never) },
  removeConditionalRule: { family: 'conditional', run: (wb, s, a) => removeConditionalRule(wb, s, a[0] as never) },
  moveConditionalRule: { family: 'conditional', run: (wb, s, a) => moveConditionalRule(wb, s, a[0] as never, a[1] as never) },
  clearConditionalRules: { family: 'conditional', run: (wb, s, a) => clearConditionalRules(wb, s, (a[0] === null ? undefined : a[0]) as never) },
  addChart: { family: 'chart', run: (wb, s, a) => addChart(wb, s, a[0] as never) },
  updateChart: { family: 'chart', run: (wb, s, a) => updateChart(wb, s, a[0] as never, a[1] as never) },
  moveChart: { family: 'chart', run: (wb, s, a) => moveChart(wb, s, a[0] as never, a[1] as never) },
  removeChart: { family: 'chart', run: (wb, s, a) => removeChart(wb, s, a[0] as never) },
  // A copy of a chart (a paste, a duplicate): the chart is named by its sheet and its place, and each participant makes the copy from the file they read.
  pasteChart: { family: 'chart', run: (wb, s, a) => pasteChartOf(wb, s, a[0] as never, a[1] as never, a[2] as never) },
};

/**
 * Every kind of operation of a workbook that travels: with the names of the commands, what the
 * fingerprint of the wire is taken from (see PROTOCOL_VERSION). A kind of WireOp that is left
 * out here is an error of the type checker.
 */
export const WIRE_KINDS: { readonly [K in WireOp['t']]: true } = { cells: true, rows: true, cols: true, merges: true, filter: true, sort: true, ins: true, del: true, notes: true, 'sheet+': true, 'sheet-': true, name: true, move: true, state: true, tab: true, pane: true, cmd: true };

/** The names of the commands that travel. */
export function commandNames(): string[] {
  return Object.keys(COMMANDS);
}

/** What a command is about: its rules or its charts are one list, which two commands of a kind change one after the other. */
export function commandFamily(name: string): 'validation' | 'conditional' | 'chart' | undefined {
  return Object.prototype.hasOwnProperty.call(COMMANDS, name) ? COMMANDS[name].family : undefined;
}

function readOp(value: unknown, style: (place: unknown) => number): WireOp {
  const from = rec(value, 'operation');
  const s = int(from.s, 0, MAX_SHEETS, 'sheet');
  switch (from.t) {
    case 'cells':
      return {
        t: 'cells',
        s,
        c: arr(from.c, MAX_ROWS, 'cells').map((entry) => {
          const item = arr(entry, 3, 'cells');
          const row = int(item[0], 0, MAX_ROWS - 1, 'row');
          const col = int(item[1], 0, MAX_COLS - 1, 'column');
          if (item[2] === 0) return [row, col, 0] as [number, number, WireCellPatch];
          const p = rec(item[2], 'cell');
          const patch: { v?: WireContent; s?: number } = {};
          if (has(p, 'v')) patch.v = readContent(p.v);
          if (has(p, 's')) patch.s = style(p.s);
          return [row, col, patch] as [number, number, WireCellPatch];
        }),
      };
    case 'rows':
      return {
        t: 'rows',
        s,
        r: arr(from.r, MAX_ROWS, 'rows').map((entry) => {
          const item = arr(entry, 2, 'rows');
          const p = rec(item[1], 'row');
          const props: WireRowProps = {};
          if (Object.prototype.hasOwnProperty.call(p, 'h')) props.h = p.h === null ? null : num(p.h, 0, 409.55, 'row');
          if (has(p, 'ch')) props.ch = bool(p.ch, 'row');
          if (has(p, 'hid')) props.hid = bool(p.hid, 'row');
          if (Object.prototype.hasOwnProperty.call(p, 's')) props.s = p.s === null ? null : style(p.s);
          return [int(item[0], 0, MAX_ROWS - 1, 'row'), props] as [number, WireRowProps];
        }),
      };
    case 'cols':
      return {
        t: 'cols',
        s,
        p: arr(from.p, MAX_STEP_RUNS, 'columns').map((entry) => {
          const item = arr(entry, 3, 'columns');
          const min = int(item[0], 0, MAX_COLS - 1, 'column');
          const max = int(item[1], min, MAX_COLS - 1, 'column');
          if (item[2] === 0) return [min, max, 0] as [number, number, WireColumn | 0];
          const c = rec(item[2], 'column');
          const column: WireColumn = { cw: bool(c.cw, 'column'), hid: bool(c.hid, 'column'), ol: int(c.ol, 0, 7, 'column'), col: bool(c.col, 'column') };
          if (has(c, 'w')) column.w = num(c.w, 0, 255, 'column');
          if (has(c, 's')) column.s = style(c.s);
          return [min, max, column] as [number, number, WireColumn | 0];
        }),
      };
    case 'merges':
      return { t: 'merges', s, add: arr(from.add, MAX_STEP_MERGES, 'merges').map(readRange), del: arr(from.del, MAX_STEP_MERGES, 'merges').map(readRange) };
    case 'filter': {
      const op: WireOp = { t: 'filter', s, f: from.f === 0 ? 0 : readFilter(from.f) };
      if (has(from, 'show')) op.show = readRows(from.show, 'filter');
      if (has(from, 'hide')) op.hide = readRows(from.hide, 'filter');
      return op;
    }
    case 'sort':
      return { t: 'sort', s, g: readRange(from.g), o: arr(from.o, MAX_ROWS, 'sort').map((row) => int(row, 0, MAX_ROWS - 1, 'sort')) };
    case 'ins':
    case 'del': {
      const x = oneOf(from.x, ['row', 'col'] as const, 'axis');
      const limit = x === 'row' ? MAX_ROWS : MAX_COLS;
      return { t: from.t, s, x, at: int(from.at, 0, limit - 1, 'position'), n: int(from.n, 1, limit, 'count') };
    }
    case 'notes':
      return {
        t: 'notes',
        s,
        n: arr(from.n, MAX_STEP_NOTES, 'notes').map((entry) => {
          const item = arr(entry, 3, 'notes');
          const row = int(item[0], 0, MAX_ROWS - 1, 'row');
          const col = int(item[1], 0, MAX_COLS - 1, 'column');
          if (item[2] === 0) return [row, col, 0] as [number, number, WireNote | 0];
          const n = rec(item[2], 'note');
          const note: WireNote = { x: str(n.x, MAX_NOTE_TEXT, 'note') };
          if (has(n, 'a')) note.a = str(n.a, 255, 'note');
          if (has(n, 'r')) note.r = readRuns(n.r);
          if (n.vis === true) note.vis = true;
          return [row, col, note] as [number, number, WireNote | 0];
        }),
      };
    case 'sheet+':
      return { t: 'sheet+', s, name: str(from.name, 31, 'name') };
    case 'sheet-':
      return { t: 'sheet-', s };
    case 'name':
      return { t: 'name', s, name: str(from.name, 31, 'name') };
    case 'move':
      return { t: 'move', s, to: int(from.to, 0, MAX_SHEETS, 'sheet') };
    case 'state':
      return { t: 'state', s, state: oneOf(from.state, ['visible', 'hidden', 'veryHidden'] as const, 'state') };
    case 'tab':
      return { t: 'tab', s, color: from.color === 0 ? 0 : readColor(from.color) };
    case 'pane': {
      const op: WireOp = { t: 'pane', s, r: int(from.r, 0, MAX_ROWS - 1, 'pane'), c: int(from.c, 0, MAX_COLS - 1, 'pane') };
      // The first cell that scrolls is one of those that scroll: below and right of the frozen rows and columns.
      if (has(from, 'tl')) {
        const tl = cellRef(from.tl, 'pane');
        if (!scrollsAt(op.r, op.c, tl)) fail('pane');
        op.tl = tl;
      }
      if (has(from, 'fv')) op.fv = cellRef(from.fv, 'pane');
      return op;
    }
    case 'cmd': {
      const name = str(from.name, 64, 'command');
      if (!commandFamily(name)) fail('command');
      return { t: 'cmd', s, name, a: plainCopy(arr(from.a, 4, 'command'), 'command') as unknown[] };
    }
    default:
      return fail('operation');
  }
}

/**
 * Reads the operations of a step that arrived (see encodeStep()) and gives its formats places
 * among those of the workbook. Throws WireError for anything that is not as the contract says.
 */
export function decodeStep(raw: unknown, styles: StyleSheet): WireOp[] {
  const items = arr(raw, MAX_STEP_OPS + 1, 'step');
  // Room for every format the step brings, before any is added: a table that is full here is this workbook's
  // matter (another participant's may have room), and the step is neither applied in part nor called no step.
  if (items.length > 0 && items[0] !== null && typeof items[0] === 'object' && Array.isArray((items[0] as { f?: unknown }).f)) {
    const count = (items[0] as { f: unknown[] }).f.length;
    if (count <= MAX_STEP_FORMATS && styles.cellXfs.length + count > MAX_CELL_FORMATS) throw new StyleTableFull();
  }
  let formats: WireFormat[] = [];
  let first = 0;
  if (items.length > 0 && items[0] !== null && typeof items[0] === 'object' && (items[0] as { t?: unknown }).t === 'st') {
    formats = arr((items[0] as { f?: unknown }).f, MAX_STEP_FORMATS, 'formats').map(readFormat);
    first = 1;
  }
  // Everything is read before anything is kept: a step that is refused leaves no format behind.
  const named = (place: unknown): number => int(place, 0, formats.length - 1, 'format');
  const read: WireOp[] = [];
  for (let i = first; i < items.length; i++) read.push(readOp(items[i], named));
  if (read.length === 0) fail('step');
  // A format is given its place when an operation names it: one that nothing names adds nothing.
  const places: (number | undefined)[] = [];
  return read.map((op) =>
    mapStyles(op, (at) => {
      let index = places[at];
      if (index === undefined) {
        index = internFormat(styles, formats[at]);
        places[at] = index;
      }
      return index;
    }),
  );
}

// ---------------------------------------------------------------------------
// From the operations of the model

function contentOf(cell: Cell | undefined, row: number, col: number, masters: () => SharedMasters): WireContent {
  if (!cell) return { k: 'blank', v: null };
  // What a dynamic array formula spilled is no content: the others work it out themselves (Cell.spilled).
  if (cell.spilled === true && !cell.formula) return { k: 'blank', v: null };
  const explicit = cell.formula && cell.formula.type === 'shared' ? explicitFormula(cell, row, col, masters()) : cell;
  const content: WireContent = { k: explicit.kind, v: explicit.value };
  const formula = explicit.formula;
  if (formula) {
    // What a formula is beside the formula of its own cell does not travel (see readContent()).
    if (formula.type === 'array' || formula.type === 'dataTable') throw new CollabUnsupportedError('other', 'Array formulas and data tables cannot be changed while the workbook is edited together');
    content.f = { x: formula.text };
    if (formula.dynamic === true) content.f.d = 1;
    if (formula.attrs && formula.attrs.size > 0) {
      const flags: [string, string][] = [];
      formula.attrs.forEach((value, name) => {
        if ((FORMULA_FLAGS as readonly string[]).indexOf(name) >= 0 && /^(1|0|true|false)$/.test(value)) flags.push([name, value]);
      });
      if (flags.length > 0) content.f.a = flags;
    }
  }
  if (explicit.runs) content.r = explicit.runs;
  if (explicit.ph) content.ph = true;
  return content;
}

/**
 * What turns the cell that is there (`now`) into `next`: undefined when they hold the same and
 * look the same.
 */
function cellPatch(now: Cell | undefined, next: Cell | undefined, row: number, col: number, masters: () => SharedMasters): WireCellPatch | undefined {
  if (!next) return now ? 0 : undefined;
  if (!now) return { v: contentOf(next, row, col, masters), s: next.style };
  const patch: { v?: WireContent; s?: number } = {};
  // The very same content (only the format is another one) is not looked into: it may be something that does not travel.
  const kept = now.formula === next.formula && now.kind === next.kind && now.value === next.value && now.runs === next.runs && now.ph === next.ph;
  if (!kept) {
    const content = contentOf(next, row, col, masters);
    // What is there may be something that does not travel (an array formula): then what comes is another content, whatever it is.
    let was: string | undefined;
    try {
      was = stable(contentOf(now, row, col, masters));
    } catch (error) {
      if (!(error instanceof CollabUnsupportedError)) throw error;
    }
    if (was !== stable(content)) patch.v = content;
  }
  if (now.style !== next.style) patch.s = next.style;
  return patch.v === undefined && patch.s === undefined ? undefined : patch;
}

function lazyMasters(sheet: Sheet, more?: readonly CellChange[]): () => SharedMasters {
  let masters: SharedMasters | undefined;
  return () => {
    if (!masters) {
      masters = sharedMasters(sheet);
      // A master that the change itself took away is among the cells it put back.
      if (more) {
        for (const change of more) {
          const formula = change.cell ? change.cell.formula : undefined;
          if (formula && formula.type === 'shared' && formula.text !== '' && formula.si !== undefined && !masters.has(formula.si)) masters.set(formula.si, { text: formula.text, row: change.row, col: change.col });
        }
      }
    }
    return masters;
  };
}

function cellsOp(sheet: Sheet, index: number, changes: readonly CellChange[]): WireOp | undefined {
  const masters = lazyMasters(sheet, changes);
  const cells: [number, number, WireCellPatch][] = [];
  for (const change of changes) {
    const row = sheet.rows.get(change.row);
    const patch = cellPatch(row ? row.cells[change.col] : undefined, change.cell, change.row, change.col, masters);
    if (patch !== undefined) cells.push([change.row, change.col, patch]);
  }
  return cells.length > 0 ? { t: 'cells', s: index, c: cells } : undefined;
}

const columnKey = (column: ColumnRange | undefined): string => (column ? `${column.width}|${column.customWidth}|${column.hidden}|${column.style}|${column.outlineLevel}|${column.collapsed}` : '');

function wireColumn(column: ColumnRange): WireColumn {
  const out: WireColumn = { cw: column.customWidth === true, hid: column.hidden === true, ol: column.outlineLevel || 0, col: column.collapsed === true };
  if (column.width !== undefined) out.w = column.width;
  if (column.style !== undefined) out.s = column.style;
  return out;
}

/** The runs of columns whose settings differ between two lists, with the settings of the second. */
function columnsDiff(from: readonly ColumnRange[], to: readonly ColumnRange[]): [number, number, WireColumn | 0][] {
  const points = new Set<number>();
  for (const list of [from, to]) {
    for (const range of list) {
      points.add(range.min);
      points.add(range.max + 1);
    }
  }
  const sorted = Array.from(points).sort((a, b) => a - b);
  const at = (list: readonly ColumnRange[], col: number): ColumnRange | undefined => {
    for (const range of list) if (range.min <= col && range.max >= col) return range;
    return undefined;
  };
  const out: [number, number, WireColumn | 0][] = [];
  let lastKey = '';
  for (let i = 0; i + 1 < sorted.length; i++) {
    const a = sorted[i];
    const b = sorted[i + 1] - 1;
    const next = at(to, a);
    const key = columnKey(next);
    if (columnKey(at(from, a)) === key) {
      lastKey = '\u0000';
      continue;
    }
    const last = out[out.length - 1];
    if (last && last[1] === a - 1 && lastKey === key) last[1] = b;
    else out.push([a, b, next ? wireColumn(next) : 0]);
    lastKey = key;
  }
  return out;
}

function columnsPatched(list: readonly ColumnRange[], patch: readonly [number, number, WireColumn | 0][]): ColumnRange[] {
  let out: ColumnRange[] = list.slice();
  for (const [min, max, set] of patch) {
    const next: ColumnRange[] = [];
    for (const range of out) {
      if (range.max < min || range.min > max) next.push(range);
      else {
        if (range.min < min) next.push({ ...range, max: min - 1 });
        if (range.max > max) next.push({ ...range, min: max + 1 });
      }
    }
    if (set !== 0) {
      const column: ColumnRange = { min, max, customWidth: set.cw, hidden: set.hid, outlineLevel: set.ol, collapsed: set.col };
      if (set.w !== undefined) column.width = set.w;
      if (set.s !== undefined) column.style = set.s;
      next.push(column);
    }
    out = next.sort((a, b) => a.min - b.min);
  }
  return out;
}

const sameRange = (a: CellRange, b: CellRange): boolean => a.top === b.top && a.left === b.left && a.bottom === b.bottom && a.right === b.right;
const intersects = (a: CellRange, b: CellRange): boolean => a.top <= b.bottom && a.bottom >= b.top && a.left <= b.right && a.right >= b.left;

function rowPropsOf(props: RowProps): WireRowProps {
  const out: WireRowProps = {};
  for (const key of Object.keys(props) as (keyof RowProps)[]) {
    const value = props[key];
    if (key === 'height') out.h = typeof value === 'number' ? value : null;
    else if (key === 'style') out.s = typeof value === 'number' ? value : null;
    else if (key === 'customHeight') out.ch = value === true;
    else if (key === 'hidden') out.hid = value === true;
  }
  return out;
}

function noteOf(comment: SheetComment): WireNote {
  const note: WireNote = { x: String(comment.text) };
  if (comment.author !== undefined) note.a = comment.author;
  if (comment.runs) note.r = comment.runs;
  if (comment.visible) note.vis = true;
  return note;
}

/** What encodeOp() made of an operation of the model. */
export interface Encoded {
  /** Missing when the operation changes nothing, and when it was left out (`skipped`). */
  op?: WireOp;
  /** The operation does not travel and was left out: rules or drawings that no command stands for (a paste brings them). */
  skipped?: 'rules' | 'drawings';
}

/**
 * The wire operation that says what an operation of the model would do to the workbook as it
 * is now. Throws CollabUnsupportedError for an operation that has none.
 */
export function encodeOp(workbook: Workbook, op: Op): Encoded {
  const sheet = workbook.sheets[op.sheet];
  switch (op.type) {
    case 'setCells':
      if (!sheet) throw new RangeError(`No sheet at index ${op.sheet}`);
      return { op: cellsOp(sheet, op.sheet, op.cells) };
    case 'setColumns': {
      if (!sheet) throw new RangeError(`No sheet at index ${op.sheet}`);
      const patch = columnsDiff(sheet.columns, columnsPatched([], op.columns.map((column) => [column.min, column.max, wireColumn(column)] as [number, number, WireColumn])));
      return { op: patch.length > 0 ? { t: 'cols', s: op.sheet, p: patch } : undefined };
    }
    case 'setRows':
      return { op: op.rows.length > 0 ? { t: 'rows', s: op.sheet, r: op.rows.map((change) => [change.row, rowPropsOf(change.props)] as [number, WireRowProps]) } : undefined };
    case 'setMerges': {
      if (!sheet) throw new RangeError(`No sheet at index ${op.sheet}`);
      const add = op.merges.filter((range) => !sheet.merges.some((other) => sameRange(other, range))).map(fromRange);
      const del = sheet.merges.filter((range) => !op.merges.some((other) => sameRange(other, range))).map(fromRange);
      return { op: add.length > 0 || del.length > 0 ? { t: 'merges', s: op.sheet, add, del } : undefined };
    }
    case 'setAutoFilter': {
      const out: WireOp = { t: 'filter', s: op.sheet, f: op.autoFilter ? filterFields(op.autoFilter) : 0 };
      if (op.show && op.show.length > 0) out.show = op.show.slice();
      if (op.hide && op.hide.length > 0) out.hide = op.hide.slice();
      return { op: out };
    }
    case 'sortRange':
      return { op: { t: 'sort', s: op.sheet, g: fromRange(op.range), o: op.order.slice() } };
    case 'setComments': {
      if (!sheet) throw new RangeError(`No sheet at index ${op.sheet}`);
      const before = sheet.comments || [];
      const after = op.comments || [];
      const kept = new Set(after);
      const had = new Set(before);
      const byCell = new Map<number, [number, number, WireNote | 0]>();
      for (const comment of before) {
        if (kept.has(comment)) continue;
        const cell = cellNow(sheet, comment);
        if (cell) byCell.set(cell.row * MAX_COLS + cell.col, [cell.row, cell.col, 0]);
      }
      for (const comment of after) {
        if (had.has(comment)) continue;
        if (comment.threaded) throw new CollabUnsupportedError('other', 'Threaded comments cannot be changed while the workbook is edited together');
        const cell = cellNow(sheet, comment);
        if (cell) byCell.set(cell.row * MAX_COLS + cell.col, [cell.row, cell.col, noteOf(comment)]);
      }
      return { op: byCell.size > 0 ? { t: 'notes', s: op.sheet, n: Array.from(byCell.values()) } : undefined };
    }
    case 'setDataValidations':
    case 'setConditionalFormats':
    case 'setDrawings': {
      // A picture would have to travel with its bytes: pictures are not added, moved or deleted in a room. Refused, before anything is changed.
      if (op.type === 'setDrawings' && op.by && PICTURE_COMMANDS.indexOf(op.by.name) >= 0) throw new CollabUnsupportedError('picture', 'Pictures cannot be added, moved or deleted while the workbook is edited together');
      if (!op.by || !commandFamily(op.by.name)) return { skipped: op.type === 'setDrawings' ? 'drawings' : 'rules' };
      // As JSON has it: what is left out and what is undefined are the same on the wire.
      const args = JSON.parse(JSON.stringify(op.by.args, (_key, value) => (value === undefined ? null : value))) as unknown[];
      return { op: { t: 'cmd', s: op.sheet, name: op.by.name, a: args } };
    }
    case 'insert':
    case 'delete':
      return { op: { t: op.type === 'insert' ? 'ins' : 'del', s: op.sheet, x: op.axis, at: op.at, n: op.count } };
    case 'insertSheet':
      return { op: { t: 'sheet+', s: op.sheet, name: op.name } };
    case 'deleteSheet':
      return { op: { t: 'sheet-', s: op.sheet } };
    case 'renameSheet':
      return { op: { t: 'name', s: op.sheet, name: op.name } };
    case 'moveSheet':
      return { op: { t: 'move', s: op.sheet, to: op.to } };
    case 'setSheetState':
      return { op: { t: 'state', s: op.sheet, state: op.state } };
    case 'setTabColor':
      return { op: { t: 'tab', s: op.sheet, color: op.color ? op.color : 0 } };
    case 'setPane': {
      const out: WireOp = { t: 'pane', s: op.sheet, r: op.pane.frozenRows, c: op.pane.frozenCols };
      if (op.pane.topLeftCell !== undefined && scrollsAt(out.r, out.c, op.pane.topLeftCell)) out.tl = op.pane.topLeftCell;
      if (op.pane.firstVisibleCell !== undefined) out.fv = op.pane.firstVisibleCell;
      return { op: out };
    }
    case 'copySheet':
      throw new CollabUnsupportedError('copySheet', 'A sheet cannot be copied while the workbook is edited together');
    case 'setPageSetup':
    case 'setDefinedNames':
      // How a sheet is printed does not travel: refused, before anything is changed.
      throw new CollabUnsupportedError('other', 'The page setup cannot be changed while the workbook is edited together');
    case 'setPivot':
      // A refresh reads the source as this participant has it, and its cache is not sent: refused, before anything is changed.
      throw new CollabUnsupportedError('other', 'A pivot table cannot be refreshed while the workbook is edited together');
    default:
      throw new CollabUnsupportedError('other', `The operation "${op.type}" cannot be shared`);
  }
}

// ---------------------------------------------------------------------------
// Applying

/** What applying one wire operation did. */
export interface Applied {
  /** The operations of the model that take it back exactly, in the order they have to be applied in. */
  inverse: Op[];
  /**
   * The wire operations that undo it as a change of its own, in order; null when it cannot be
   * undone that way (a deleted sheet, a rule, a chart).
   */
  undo: Rec[] | null;
  /** Undoing brings back the cells but not everything that went with them (notes, links, rules). */
  lossy?: boolean;
  /** The name the operation is read by (see Rec.nm). */
  nm?: string;
  /** A sort took every used column of the sheet along (see Rec.wide). */
  wide?: boolean;
  /** The cells whose content changed; missing when cells moved, which leaves every formula to be calculated again. */
  cells?: { sheet: number; row: number; col: number }[];
  /** Sizes, hidden rows or merged cells of the sheet changed. */
  layout: boolean;
  /** The list of sheets, or the name of one, changed. */
  sheets: boolean;
  /** The rules or the charts of a sheet changed: what names one of them by its place is out of date. */
  lists?: boolean;
}

function worksheet(workbook: Workbook, index: number): Sheet {
  const sheet = workbook.sheets[index];
  if (!sheet) throw new RangeError(`No sheet at index ${index}`);
  return sheet;
}

function cellFrom(content: WireContent, style: number): Cell {
  const cell: Cell = { kind: content.k, value: content.v, style };
  if (content.f) {
    cell.formula = { text: content.f.x, type: content.f.t || 'normal' };
    if (content.f.r !== undefined && content.f.t) cell.formula.ref = content.f.r;
    if (content.f.a) cell.formula.attrs = new Map(content.f.a);
    if (content.f.d === 1 && !content.f.t) cell.formula.dynamic = true;
  }
  if (content.r) cell.runs = content.r;
  if (content.ph) cell.ph = true;
  return cell;
}

/** The format a position has without a cell of its own: its row's, else its column's. */
function inheritedStyle(sheet: Sheet, row: number, col: number): number {
  const rowData = sheet.rows.get(row);
  if (rowData && rowData.style !== undefined) return rowData.style;
  for (const range of sheet.columns) {
    if (range.min > col) break;
    if (range.max >= col) return range.style !== undefined ? range.style : 0;
  }
  return 0;
}

function patched(workbook: Workbook, sheet: Sheet, row: number, col: number, patch: WireCellPatch): Cell | undefined {
  if (patch === 0) return undefined;
  const rowData = sheet.rows.get(row);
  const now = rowData ? rowData.cells[col] : undefined;
  let style = patch.s !== undefined ? patch.s : now ? now.style : inheritedStyle(sheet, row, col);
  if (!(style >= 0 && style < workbook.styles.cellXfs.length)) style = 0;
  if (patch.v) return cellFrom(patch.v, style);
  // Only the format: what the cell holds, and what the file has beside it, stays.
  return now ? { ...now, style } : { kind: 'blank', value: null, style };
}

function rowProps(props: WireRowProps, workbook: Workbook): RowProps {
  const out: RowProps = {};
  if (Object.prototype.hasOwnProperty.call(props, 'h')) out.height = props.h === null ? undefined : props.h;
  if (props.ch !== undefined) out.customHeight = props.ch;
  if (props.hid !== undefined) out.hidden = props.hid;
  if (Object.prototype.hasOwnProperty.call(props, 's')) out.style = typeof props.s === 'number' && props.s < workbook.styles.cellXfs.length ? props.s : undefined;
  return out;
}

/** A filter that arrived, with what the filter that is there holds beside the fields of the model where they say the same. */
function filterWith(filter: AutoFilter, local: AutoFilter | undefined): AutoFilter {
  if (!local) return filter;
  const out: AutoFilter = { ...filter };
  const plain = (column: FilterColumn): string => stable({ ...column, raw: undefined });
  out.columns = filter.columns.map((column) => {
    const mine = local.columns.find((other) => other.colId === column.colId);
    return mine && mine.raw !== undefined && plain(mine) === plain(column) ? mine : column;
  });
  if (local.attrs) out.attrs = local.attrs;
  if (local.extra !== undefined) out.extra = local.extra;
  if (filter.sortState && local.sortState && stable(filterFields(local).sortState) === stable(filter.sortState)) out.sortState = local.sortState;
  return out;
}

/** Whether a formula of a cell or a defined name of the workbook names a sheet of this name (there need be no such sheet). */
function namesSheet(workbook: Workbook, name: string): boolean {
  const wanted = foldSheetName(name);
  const map: SheetNameMap = (written) => (foldSheetName(written) === wanted ? null : undefined);
  const names = (text: string): boolean => text.indexOf('!') >= 0 && renameSheetsInFormula(text, map) !== text;
  for (const sheet of workbook.sheets) {
    for (const row of sheet.rows.values()) {
      let found = false;
      row.cells.forEach((cell) => {
        if (!found && cell.formula && cell.formula.text !== '' && names(cell.formula.text)) found = true;
      });
      if (found) return true;
    }
  }
  return workbook.definedNames.some((defined) => names(defined.value));
}

/** What undoing a deletion of rows or columns has to put back beside the rows or columns themselves; worked out before they go. */
function deletionUndo(workbook: Workbook, index: number, axis: 'row' | 'col', at: number, n: number): { recs: Rec[] | null; lossy: boolean } {
  const sheet = worksheet(workbook, index);
  const end = at + n - 1;
  const gone: StructuralEdit = { kind: 'delete', axis, at, count: n, sheetName: sheet.name };
  const back: StructuralEdit = { kind: 'insert', axis, at, count: n, sheetName: sheet.name };
  const inBand = (row: number, col: number): boolean => (axis === 'row' ? row >= at && row <= end : col >= at && col <= end);
  let budget = MAX_UNDO_CELLS;
  const recs: Rec[] = [{ w: { t: 'ins', s: index, x: axis, at, n }, nm: sheet.name }];

  // The settings of the rows or the columns that go.
  if (axis === 'row') {
    const rows: [number, WireRowProps][] = [];
    for (const [r, row] of sheet.rows) {
      if (r < at || r > end) continue;
      const props: WireRowProps = {};
      if (row.height !== undefined) props.h = row.height;
      if (row.customHeight) props.ch = true;
      if (row.hidden) props.hid = true;
      if (row.style !== undefined) props.s = row.style;
      if (Object.keys(props).length > 0) rows.push([r, props]);
    }
    if (rows.length > 0) recs.push({ w: { t: 'rows', s: index, r: rows } });
  } else {
    const columns: [number, number, WireColumn | 0][] = [];
    for (const range of sheet.columns) if (range.max >= at && range.min <= end) columns.push([Math.max(range.min, at), Math.min(range.max, end), wireColumn(range)]);
    if (columns.length > 0) recs.push({ w: { t: 'cols', s: index, p: columns } });
  }

  // Merged ranges that neither the deletion nor putting the rows back leaves as they were.
  const add: WireRange[] = [];
  const del: WireRange[] = [];
  for (const merge of sheet.merges) {
    let left = adjustRange(merge, gone);
    if (left && left.top === left.bottom && left.left === left.right) left = undefined;
    const again = left ? adjustRange(left, back) : undefined;
    if (again && sameRange(again, merge)) continue;
    if (again) del.push(fromRange(again));
    add.push(fromRange(merge));
  }
  if (add.length > 0) recs.push({ w: { t: 'merges', s: index, add, del } });

  // The cells that go, and the formulas anywhere that the two edits together do not leave as they were.
  let over = false;
  workbook.sheets.forEach((other, otherIndex) => {
    if (over) return;
    const masters = lazyMasters(other);
    const cells: [number, number, WireCellPatch][] = [];
    for (const [r, row] of other.rows) {
      if (over) break;
      row.cells.forEach((cell, c) => {
        if (over) return;
        if (other === sheet && inBand(r, c)) {
          cells.push([r, c, { v: contentOf(cell, r, c, masters), s: cell.style }]);
        } else if (cell.formula) {
          const content = contentOf(cell, r, c, masters);
          const text = content.f ? content.f.x : '';
          if (text === '') return;
          const after = adjustFormula(text, other.name, gone);
          if (after === text || adjustFormula(after, other.name, back) === text) return;
          cells.push([r, c, { v: content }]);
        } else return;
        if (--budget < 0) over = true;
      });
    }
    if (cells.length > 0) recs.push({ w: { t: 'cells', s: otherIndex, c: cells } });
  });
  if (over) return { recs: null, lossy: false };

  let lossy = false;
  for (const comment of sheet.comments || []) {
    const cell = cellNow(sheet, comment);
    if (cell && inBand(cell.row, cell.col)) lossy = true;
  }
  for (const link of sheet.hyperlinks) if (!adjustRange(link.range, gone)) lossy = true;
  if (sheet.autoFilter && !shiftAutoFilter(sheet.autoFilter, gone)) lossy = true;
  // Rules and charts of any sheet may read the rows that go, and are not put back as they were.
  for (const other of workbook.sheets) {
    if ((other.dataValidations && other.dataValidations.length > 0) || (other.conditionalFormats && other.conditionalFormats.length > 0) || (other.drawings && other.drawings.length > 0)) lossy = true;
  }
  // The source of a pivot table that loses its last rows does not grow back by them.
  if (pivotSourceLostBy(workbook, sheet, gone)) lossy = true;
  for (const name of workbook.definedNames) {
    const home = name.localSheet !== undefined && workbook.sheets[name.localSheet] ? workbook.sheets[name.localSheet].name : '';
    const after = adjustFormula(name.value, home, gone);
    if (after !== name.value && adjustFormula(after, home, back) !== name.value) lossy = true;
  }
  return { recs, lossy };
}

/**
 * Applies a wire operation to the workbook. Throws RangeError when the workbook does not take
 * it (a position outside the sheet, a merged range across another, a name that is taken,
 * whatever the commands refuse); nothing was changed then.
 */
export function applyWire(workbook: Workbook, op: WireOp): Applied {
  switch (op.t) {
    case 'cells': {
      const sheet = worksheet(workbook, op.s);
      const changes: CellChange[] = op.c.map(([row, col, patch]) => ({ row, col, cell: patched(workbook, sheet, row, col, patch) }));
      const inverse = applyOp(workbook, { type: 'setCells', sheet: op.s, cells: changes });
      const before = inverse.type === 'setCells' ? inverse.cells : [];
      // What was there may be something that does not travel (an array formula): the step is then not one to undo.
      let undo: Rec[] | null = [];
      try {
        const back = cellsOp(sheet, op.s, before);
        if (back) undo = [{ w: back }];
      } catch (error) {
        if (!(error instanceof CollabUnsupportedError)) throw error;
        undo = null;
      }
      return { inverse: [inverse], undo, cells: before.map((change) => ({ sheet: op.s, row: change.row, col: change.col })), layout: false, sheets: false };
    }
    case 'rows': {
      worksheet(workbook, op.s);
      const inverse = applyOp(workbook, { type: 'setRows', sheet: op.s, rows: op.r.map(([row, props]) => ({ row, props: rowProps(props, workbook) })) });
      const undo: Rec[] = inverse.type === 'setRows' && inverse.rows.length > 0 ? [{ w: { t: 'rows', s: op.s, r: inverse.rows.map((change) => [change.row, rowPropsOf(change.props)] as [number, WireRowProps]) } }] : [];
      return { inverse: [inverse], undo, cells: [], layout: true, sheets: false };
    }
    case 'cols': {
      const sheet = worksheet(workbook, op.s);
      const before = sheet.columns;
      const patch = op.p.map((entry) => (entry[2] !== 0 && entry[2].s !== undefined && !(entry[2].s < workbook.styles.cellXfs.length) ? ([entry[0], entry[1], { ...entry[2], s: undefined }] as [number, number, WireColumn | 0]) : entry));
      const inverse = applyOp(workbook, { type: 'setColumns', sheet: op.s, columns: columnsPatched(before, patch) });
      const back = columnsDiff(sheet.columns, before);
      return { inverse: [inverse], undo: back.length > 0 ? [{ w: { t: 'cols', s: op.s, p: back } }] : [], cells: [], layout: true, sheets: false };
    }
    case 'merges': {
      const sheet = worksheet(workbook, op.s);
      const gone = op.del.map(toRange);
      const removed: WireRange[] = [];
      const next = sheet.merges.filter((range) => {
        if (!gone.some((other) => sameRange(other, range))) return true;
        removed.push(fromRange(range));
        return false;
      });
      for (const item of op.add) {
        const range = toRange(item);
        if (range.top === range.bottom && range.left === range.right) throw new RangeError('A merged range has more than one cell');
        if (next.some((other) => intersects(other, range))) throw new RangeError('Cannot merge cells across a merged range');
        next.push(range);
      }
      const inverse = applyOp(workbook, { type: 'setMerges', sheet: op.s, merges: next });
      const undo: Rec[] = removed.length > 0 || op.add.length > 0 ? [{ w: { t: 'merges', s: op.s, add: removed, del: op.add.slice() } }] : [];
      return { inverse: [inverse], undo, cells: [], layout: true, sheets: false };
    }
    case 'filter': {
      const sheet = worksheet(workbook, op.s);
      if (sheet.kind !== 'worksheet') throw new RangeError(`Cannot filter a ${sheet.kind}`);
      const filter = op.f === 0 ? undefined : filterWith(op.f, sheet.autoFilter);
      const inverse = applyOp(workbook, { type: 'setAutoFilter', sheet: op.s, autoFilter: filter, show: op.show, hide: op.hide });
      const undo: WireOp = { t: 'filter', s: op.s, f: inverse.type === 'setAutoFilter' && inverse.autoFilter ? filterFields(inverse.autoFilter) : 0 };
      if (inverse.type === 'setAutoFilter') {
        if (inverse.show && inverse.show.length > 0) undo.show = inverse.show;
        if (inverse.hide && inverse.hide.length > 0) undo.hide = inverse.hide;
      }
      return { inverse: [inverse], undo: [{ w: undo }], cells: [], layout: true, sheets: false };
    }
    case 'sort': {
      const sorted = worksheet(workbook, op.s);
      // Whether whole rows move: nothing of the rows stands outside the columns that are sorted.
      let last = sorted.colCount - 1;
      for (const row of sorted.rows.values()) if (row.cells.length - 1 > last) last = row.cells.length - 1;
      const wide = op.g[1] === 0 && op.g[3] >= last;
      // The formulas of the rows that move are written anew for where they go, and what a formula reads may fall off
      // the sheet on the way: putting the rows back does not bring such a formula back. They are kept as they are.
      let formulas: [number, number, WireCellPatch][] | null = [];
      try {
        const masters = lazyMasters(sorted);
        for (let i = 0; i < op.o.length && formulas; i++) {
          if (op.o[i] === i) continue;
          const row = sorted.rows.get(op.g[0] + i);
          if (!row) continue;
          for (let col = op.g[1]; col <= op.g[3] && col < row.cells.length; col++) {
            const cell = row.cells[col];
            if (!cell || !cell.formula) continue;
            formulas.push([op.g[0] + i, col, { v: contentOf(cell, op.g[0] + i, col, masters) }]);
            if (formulas.length > MAX_UNDO_CELLS) {
              formulas = null;
              break;
            }
          }
        }
      } catch (error) {
        if (!(error instanceof CollabUnsupportedError)) throw error;
        formulas = null;
      }
      const inverse = applyOp(workbook, { type: 'sortRange', sheet: op.s, range: toRange(op.g), order: op.o });
      const back: number[] = new Array(op.o.length);
      for (let i = 0; i < op.o.length; i++) back[op.o[i]] = i;
      let undo: Rec[] | null = null;
      if (formulas) {
        undo = [{ w: { t: 'sort', s: op.s, g: op.g, o: back }, wide }];
        if (formulas.length > 0) undo.push({ w: { t: 'cells', s: op.s, c: formulas } });
      }
      return { inverse: [inverse], undo, wide, layout: true, sheets: false };
    }
    case 'ins':
    case 'del': {
      const sheet = worksheet(workbook, op.s);
      const name = sheet.name;
      let material: { recs: Rec[] | null; lossy: boolean } | undefined;
      if (op.t === 'del') {
        try {
          material = deletionUndo(workbook, op.s, op.x, op.at, op.n);
        } catch (error) {
          if (!(error instanceof CollabUnsupportedError)) throw error;
          material = { recs: null, lossy: false };
        }
      }
      const inverse = applyOp(workbook, { type: op.t === 'ins' ? 'insert' : 'delete', sheet: op.s, axis: op.x, at: op.at, count: op.n });
      const undo: Rec[] | null = material ? material.recs : [{ w: { t: 'del', s: op.s, x: op.x, at: op.at, n: op.n }, nm: name }];
      return { inverse: [inverse], undo, lossy: material ? material.lossy : false, nm: name, layout: true, sheets: false };
    }
    case 'notes': {
      const sheet = worksheet(workbook, op.s);
      if (sheet.kind !== 'worksheet') throw new RangeError(`Cannot change comments in a ${sheet.kind}`);
      if (notesBlocked(sheet)) throw new RangeError('Cannot change the comments of this sheet');
      const next = (sheet.comments || []).slice();
      const edits = sheet.structuralEdits ? sheet.structuralEdits.length : 0;
      const back: [number, number, WireNote | 0][] = [];
      // The first comment of every cell, found once: a step may bring thousands of notes.
      const byCell = new Map<number, SheetComment>();
      for (const comment of next) {
        const cell = cellNow(sheet, comment);
        if (cell && !byCell.has(cell.row * MAX_COLS + cell.col)) byCell.set(cell.row * MAX_COLS + cell.col, comment);
      }
      const gone = new Set<SheetComment>();
      const edited = new Map<SheetComment, SheetComment>();
      const added: SheetComment[] = [];
      for (const [row, col, note] of op.n) {
        const at = row * MAX_COLS + col;
        const found = byCell.get(at);
        if (found && found.threaded) throw new RangeError('Cannot change a threaded comment');
        if (note === 0) {
          if (!found) continue;
          gone.add(found);
          byCell.delete(at);
        } else {
          const made: SheetComment = found ? { ...found, text: note.x } : { row, col, text: note.x, author: note.a !== undefined && note.a !== '' ? note.a : DEFAULT_NOTE_AUTHOR, since: edits };
          if (found && note.a !== undefined) made.author = note.a;
          if (note.r) made.runs = note.r;
          else delete made.runs;
          if (note.vis) made.visible = true;
          else delete made.visible;
          if (found) {
            gone.add(found);
            edited.set(found, made);
          } else added.push(made);
          byCell.set(at, made);
        }
        back.unshift([row, col, found ? noteOf(found) : 0]);
      }
      // A note that was written twice in one operation: the last one stands, in the place of the first.
      const live = new Set(byCell.values());
      const list: SheetComment[] = [];
      for (const comment of next) {
        if (!gone.has(comment)) list.push(comment);
        else {
          let now = edited.get(comment);
          while (now && edited.has(now)) now = edited.get(now);
          if (now && live.has(now)) list.push(now);
        }
      }
      for (const comment of added) {
        let now: SheetComment | undefined = comment;
        while (now && edited.has(now)) now = edited.get(now);
        if (now && live.has(now)) list.push(now);
      }
      next.length = 0;
      for (const comment of list) next.push(comment);
      if (next.length > MAX_SHEET_NOTES) throw new RangeError('The sheet has too many comments');
      const inverse = applyOp(workbook, { type: 'setComments', sheet: op.s, comments: next.length > 0 ? next : undefined });
      return { inverse: [inverse], undo: back.length > 0 ? [{ w: { t: 'notes', s: op.s, n: back } }] : [], cells: [], layout: false, sheets: false };
    }
    case 'sheet+': {
      // A formula that names a sheet of this name already reads the new sheet from now on, and would be an error
      // once the sheet is taken away again, which it was not before: such a step is not one to undo.
      const named = namesSheet(workbook, op.name);
      const inverse = applyOp(workbook, { type: 'insertSheet', sheet: op.s, name: op.name });
      return { inverse: [inverse], undo: named ? null : [{ w: { t: 'sheet-', s: op.s }, nm: op.name }], nm: op.name, cells: [], layout: false, sheets: true };
    }
    case 'sheet-': {
      const name = worksheet(workbook, op.s).name;
      const inverse = applyOp(workbook, { type: 'deleteSheet', sheet: op.s });
      return { inverse: [inverse], undo: null, nm: name, layout: false, sheets: true };
    }
    case 'name': {
      const old = worksheet(workbook, op.s).name;
      // As for a new sheet: a formula that named the new name before would be written with the old one when the name is taken back.
      const named = foldSheetName(old) !== foldSheetName(op.name) && namesSheet(workbook, op.name);
      const inverse = applyOp(workbook, { type: 'renameSheet', sheet: op.s, name: op.name });
      let undo: Rec[] | null = named ? null : [{ w: { t: 'name', s: op.s, name: old }, nm: op.name }];
      // The formulas that named the sheet, as they were written: taking the name back writes the name the usual way,
      // which need not be the way they had it ("'Data'!A1").
      if (undo && inverse.type === 'restoreSheets') {
        let budget = MAX_UNDO_CELLS;
        try {
          for (const entry of inverse.state.cells) {
            const at = workbook.sheets.indexOf(entry.sheet);
            if (at < 0 || !undo) continue;
            const masters = lazyMasters(entry.sheet, entry.cells);
            const cells: [number, number, WireCellPatch][] = entry.cells.map((change) => [change.row, change.col, change.cell ? { v: contentOf(change.cell, change.row, change.col, masters) } : 0]);
            budget -= cells.length;
            if (budget < 0) undo = null;
            else if (cells.length > 0) undo.push({ w: { t: 'cells', s: at, c: cells } });
          }
        } catch (error) {
          if (!(error instanceof CollabUnsupportedError)) throw error;
          undo = null;
        }
      }
      return { inverse: [inverse], undo, nm: old, layout: false, sheets: true };
    }
    case 'move': {
      const inverse = applyOp(workbook, { type: 'moveSheet', sheet: op.s, to: op.to });
      return { inverse: [inverse], undo: [{ w: { t: 'move', s: op.to, to: op.s } }], cells: [], layout: false, sheets: true };
    }
    case 'state': {
      const old = worksheet(workbook, op.s).state;
      const inverse = applyOp(workbook, { type: 'setSheetState', sheet: op.s, state: op.state });
      return { inverse: [inverse], undo: [{ w: { t: 'state', s: op.s, state: old } }], cells: [], layout: false, sheets: true };
    }
    case 'pane': {
      const pane = paneOfWire(op);
      const inverse = applyOp(workbook, { type: 'setPane', sheet: op.s, pane });
      if (inverse.type !== 'setPane') throw new WireError('Invalid panes');
      // The panes as they were, the file's first cell that scrolls with them (Excel writes where the pane was scrolled
      // to): undone, the sheet is written as the file has it. One no participant would take is made of the rows and columns.
      const back: WireOp = { t: 'pane', s: op.s, r: inverse.pane.frozenRows, c: inverse.pane.frozenCols };
      if (inverse.pane.topLeftCell !== undefined && scrollsAt(back.r, back.c, inverse.pane.topLeftCell)) back.tl = inverse.pane.topLeftCell;
      if (inverse.pane.firstVisibleCell !== undefined) back.fv = inverse.pane.firstVisibleCell;
      return { inverse: [inverse], undo: [{ w: back }], cells: [], layout: true, sheets: false };
    }
    case 'tab': {
      const old = worksheet(workbook, op.s).tabColor;
      const inverse = applyOp(workbook, { type: 'setTabColor', sheet: op.s, color: op.color === 0 ? undefined : op.color });
      return { inverse: [inverse], undo: [{ w: { t: 'tab', s: op.s, color: old ? old : 0 } }], cells: [], layout: false, sheets: true };
    }
    case 'cmd': {
      worksheet(workbook, op.s);
      if (!commandFamily(op.name)) throw new WireError('Invalid command');
      let ops: Op[];
      try {
        // The command gets a copy: it may keep what it is given in the model.
        ops = COMMANDS[op.name].run(workbook, op.s, plainCopy(op.a, 'command') as unknown[]);
      } catch (error) {
        if (error instanceof RangeError || error instanceof WireError) throw error;
        // Arguments of a kind the command never expected: refused, like any it refuses itself.
        throw new RangeError(`The command "${op.name}" cannot be applied`);
      }
      const inverse: Op[] = [];
      try {
        for (const made of ops) inverse.unshift(applyOp(workbook, made));
      } catch (error) {
        for (const done of inverse) applyOp(workbook, done);
        throw error;
      }
      return { inverse, undo: null, cells: [], layout: false, sheets: false, lists: true };
    }
  }
}

/** How an operation that was applied moves what stands around it (see Shift); `nm` as applyWire() gave it. */
export function shiftOf(op: WireOp, nm: string | undefined, wide = false): Shift | undefined {
  switch (op.t) {
    case 'ins':
    case 'del':
      return { k: 'rc', s: op.s, axis: op.x, ins: op.t === 'ins', at: op.at, n: op.n, name: nm || '' };
    case 'sort': {
      const target: number[] = new Array(op.o.length);
      for (let i = 0; i < op.o.length; i++) target[op.o[i]] = i;
      return { k: 'perm', s: op.s, top: op.g[0], left: op.g[1], bottom: op.g[2], right: op.g[3], target, wide };
    }
    case 'sheet+':
      return { k: 'sIns', at: op.s, name: op.name };
    case 'sheet-':
      return { k: 'sDel', at: op.s, name: nm || '' };
    case 'move':
      return op.s === op.to ? undefined : { k: 'sMove', from: op.s, to: op.to };
    case 'name':
      return { k: 'sName', s: op.s, old: nm || '', name: op.name };
    default:
      return undefined;
  }
}

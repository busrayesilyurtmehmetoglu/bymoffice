import assert from 'node:assert/strict';
import { test } from 'node:test';
import { createLayout } from '../src/layout/grid-geometry';
import { createSheetMetrics } from '../src/layout/sheet-metrics';
import { createSheet, emptyStyleSheet } from '../src/model/sheet/types';
import type { Alignment, Border, Cell, CellRange, Fill, Font, Row, Sheet, StyleSheet, Workbook } from '../src/model/sheet/types';
import { DEFAULT_GRID_THEME, renderGrid } from '../src/render/grid';
import type { GridSelection } from '../src/render/grid';
import { clearTextMeasureCache, FONT_FALLBACK } from '../src/render/text';
import { LOCALE_TR, type FormatLocale } from '../src/format/locale';
import { ConditionalFormats, type CellConditional, type ConditionalSource } from '../src/format/conditional';

// Default geometry: columns are 64 px wide, rows 20 px high, the row header is
// 72 px wide and the column header 22 px high. So A1 is at (72, 22).
const HX = 72;
const HY = 22;
const CW = 64;
const RH = 20;
/** Calibri 11 pt in pixels; the fake context makes every character half of that wide. */
const PX = 14.67;
const CHAR = PX / 2;
const DEFAULT_FONT = `${PX}px "Calibri", ${FONT_FALLBACK}`;
const HEADER_FONT = `12px ${FONT_FALLBACK}`;
/** Baseline of bottom-aligned default text in row 1. */
const BASELINE_ROW1 = 37;

interface Box {
  x: number;
  y: number;
  width: number;
  height: number;
}

interface TextCall {
  text: string;
  x: number;
  y: number;
  font: string;
  fill: string;
  align: string;
  /** Clip rectangles in force, outermost first. */
  clips: Box[];
}

interface RectCall extends Box {
  fill: string;
  clips: Box[];
}

interface StrokeCall {
  style: string;
  width: number;
  dash: number[];
  alpha: number;
  /** [x1, y1, x2, y2] of each line of the path. */
  lines: number[][];
  clips: Box[];
}

interface StrokeRectCall extends Box {
  style: string;
  lineWidth: number;
}

/** A path that was filled: its corners (lines) and circles (arcs: centre, radius, start and end angle). */
interface PathCall {
  fill: string;
  points: number[][];
  arcs: number[][];
}

/** What createLinearGradient() hands out; it stands in for a fill style. */
class FakeGradient {
  stops: [number, string][] = [];
  constructor(readonly line: number[]) {}
  addColorStop(offset: number, color: string): void {
    this.stops.push([offset, color]);
  }
}

/** A stand-in for CanvasRenderingContext2D that records what is drawn. Each character is half the font size wide. */
class FakeContext {
  fillStyle: string | FakeGradient = '#000000';
  strokeStyle = '#000000';
  lineWidth = 1;
  globalAlpha = 1;
  textAlign = 'left';
  textBaseline = 'alphabetic';
  transform: number[] = [];
  texts: TextCall[] = [];
  rects: RectCall[] = [];
  strokes: StrokeCall[] = [];
  strokeRects: StrokeRectCall[] = [];
  /** Paths filled with fill(), and rectangles filled with a gradient. */
  paths: PathCall[] = [];
  gradientRects: (Box & { gradient: FakeGradient })[] = [];
  /** Every value assigned to `font`. */
  fonts: string[] = [];
  measured: string[] = [];
  saves = 0;
  restores = 0;

  private currentFont = '10px sans-serif';
  private dash: number[] = [];
  private clips: Box[] = [];
  private stack: { clips: Box[]; font: string; fill: string | FakeGradient }[] = [];
  private pathRect: Box | undefined;
  private pathLines: number[][] = [];
  private pathPoints: number[][] = [];
  private pathArcs: number[][] = [];
  private pen: number[] = [0, 0];

  get font(): string {
    return this.currentFont;
  }
  set font(value: string) {
    this.currentFont = value;
    this.fonts.push(value);
  }

  setTransform(...values: number[]): void {
    this.transform = values;
  }
  save(): void {
    this.saves++;
    this.stack.push({ clips: this.clips.slice(), font: this.currentFont, fill: this.fillStyle });
  }
  restore(): void {
    this.restores++;
    const state = this.stack.pop();
    assert.ok(state, 'restore() without save()');
    this.clips = state.clips;
    this.currentFont = state.font;
    this.fillStyle = state.fill;
  }
  beginPath(): void {
    this.pathRect = undefined;
    this.pathLines = [];
    this.pathPoints = [];
    this.pathArcs = [];
  }
  closePath(): void {}
  arc(x: number, y: number, radius: number, start: number, end: number): void {
    this.pathArcs.push([x, y, radius, start, end]);
  }
  fill(): void {
    assert.equal(typeof this.fillStyle, 'string', 'fill() with a gradient');
    this.paths.push({ fill: this.fillStyle as string, points: this.pathPoints.slice(), arcs: this.pathArcs.slice() });
  }
  createLinearGradient(x0: number, y0: number, x1: number, y1: number): FakeGradient {
    return new FakeGradient([x0, y0, x1, y1]);
  }
  rect(x: number, y: number, width: number, height: number): void {
    this.pathRect = { x, y, width, height };
  }
  clip(): void {
    assert.ok(this.pathRect, 'clip() without a rect');
    this.clips = this.clips.concat([this.pathRect]);
  }
  moveTo(x: number, y: number): void {
    this.pen = [x, y];
    this.pathPoints.push([x, y]);
  }
  lineTo(x: number, y: number): void {
    this.pathLines.push([this.pen[0], this.pen[1], x, y]);
    this.pen = [x, y];
    this.pathPoints.push([x, y]);
  }
  stroke(): void {
    this.strokes.push({
      style: this.strokeStyle,
      width: this.lineWidth,
      dash: this.dash.slice(),
      alpha: this.globalAlpha,
      lines: this.pathLines.slice(),
      clips: this.clips,
    });
  }
  setLineDash(dash: number[]): void {
    this.dash = dash.slice();
  }
  fillRect(x: number, y: number, width: number, height: number): void {
    if (typeof this.fillStyle !== 'string') this.gradientRects.push({ x, y, width, height, gradient: this.fillStyle });
    else this.rects.push({ x, y, width, height, fill: this.fillStyle, clips: this.clips });
  }
  strokeRect(x: number, y: number, width: number, height: number): void {
    this.strokeRects.push({ x, y, width, height, style: this.strokeStyle, lineWidth: this.lineWidth });
  }
  fillText(text: string, x: number, y: number): void {
    assert.equal(typeof this.fillStyle, 'string', 'text drawn with a gradient');
    this.texts.push({ text, x, y, font: this.currentFont, fill: this.fillStyle as string, align: this.textAlign, clips: this.clips });
  }
  measureText(text: string): { width: number } {
    this.measured.push(text);
    const match = /(\d+(?:\.\d+)?)px/.exec(this.currentFont);
    return { width: text.length * (match ? Number(match[1]) : 10) * 0.5 };
  }

  /** Text drawn for cells (the headers have their own font). */
  get cellTexts(): TextCall[] {
    return this.texts.filter((call) => call.font !== HEADER_FONT);
  }
  get headerTexts(): TextCall[] {
    return this.texts.filter((call) => call.font === HEADER_FONT);
  }
  textCall(text: string): TextCall {
    const found = this.cellTexts.filter((call) => call.text === text);
    assert.equal(found.length, 1, `"${text}" drawn ${found.length} times`);
    return found[0];
  }
  fills(colour: string): RectCall[] {
    return this.rects.filter((call) => call.fill === colour);
  }
  strokesOf(colour: string): StrokeCall[] {
    return this.strokes.filter((call) => call.style === colour);
  }
}

const newSheet = (): Sheet => createSheet({ name: 'S', sheetId: 1, state: 'visible', kind: 'worksheet', part: '/x' });
const newRow = (): Row => ({ customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells: [] });
const baseFont = (init: Partial<Font> = {}): Font => ({ name: 'Calibri', size: 11, bold: false, italic: false, strike: false, underline: 'none', ...init });

function rowOf(sheet: Sheet, index: number): Row {
  let row = sheet.rows.get(index);
  if (!row) {
    row = newRow();
    sheet.rows.set(index, row);
  }
  return row;
}

function put(sheet: Sheet, row: number, col: number, value: Cell['value'], style = 0, kind?: Cell['kind']): Cell {
  const cell: Cell = {
    kind: kind || (typeof value === 'number' ? 'number' : typeof value === 'boolean' ? 'boolean' : value === null ? 'blank' : 'string'),
    value,
    style,
  };
  rowOf(sheet, row).cells[col] = cell;
  sheet.rowCount = Math.max(sheet.rowCount, row + 1);
  sheet.colCount = Math.max(sheet.colCount, col + 1);
  return cell;
}

/** Adds a cell format and returns its index. */
function xf(styles: StyleSheet, init: { font?: Font; fill?: Fill; border?: Partial<Border>; alignment?: Alignment }): number {
  let fontId = 0;
  let fillId = 0;
  let borderId = 0;
  if (init.font) fontId = styles.fonts.push(init.font) - 1;
  if (init.fill) fillId = styles.fills.push(init.fill) - 1;
  if (init.border) borderId = styles.borders.push({ diagonalUp: false, diagonalDown: false, ...init.border }) - 1;
  return styles.cellXfs.push({ numFmtId: 0, fontId, fillId, borderId, alignment: init.alignment, quotePrefix: false }) - 1;
}

const solid = (rgb: string): Fill => ({ type: 'pattern', pattern: 'solid', fg: { rgb } });

interface DrawOptions {
  styles?: StyleSheet;
  width?: number;
  height?: number;
  scrollX?: number;
  scrollY?: number;
  selection?: GridSelection;
  dpr?: number;
  locale?: FormatLocale;
  date1904?: boolean;
  /** Conditional formats; `true` for the ones the sheet itself holds. */
  conditional?: ConditionalSource | true;
}

function draw(sheet: Sheet, options: DrawOptions = {}): FakeContext {
  const workbook: Workbook = {
    sheets: [sheet],
    styles: options.styles || emptyStyleSheet(),
    sharedStrings: [],
    definedNames: [],
    date1904: options.date1904 === true,
    activeSheet: 0,
  };
  const layout = createLayout(createSheetMetrics(sheet), sheet.view, {
    width: options.width === undefined ? 800 : options.width,
    height: options.height === undefined ? 600 : options.height,
    scrollX: options.scrollX || 0,
    scrollY: options.scrollY || 0,
  });
  const ctx = new FakeContext();
  renderGrid(ctx as unknown as CanvasRenderingContext2D, {
    workbook,
    sheet,
    layout,
    selection: options.selection,
    locale: options.locale,
    devicePixelRatio: options.dpr === undefined ? 1 : options.dpr,
    conditional: options.conditional === true ? new ConditionalFormats(workbook, sheet) : options.conditional,
  });
  assert.equal(ctx.saves, ctx.restores, 'save() and restore() are balanced');
  return ctx;
}

const box = (x: number, y: number, width: number, height: number): Box => ({ x, y, width, height });
const range = (top: number, left: number, bottom: number, right: number): CellRange => ({ top, left, bottom, right });
const hasLine = (calls: StrokeCall[], expected: number[]) => calls.some((call) => call.lines.some((l) => l.join() === expected.join()));

test('grid: transform, background and an empty sheet', () => {
  const ctx = draw(newSheet(), { dpr: 2 });
  assert.deepEqual(ctx.transform, [2, 0, 0, 2, 0, 0]);
  assert.deepEqual({ ...ctx.rects[0], clips: undefined }, { ...box(0, 0, 800, 600), fill: DEFAULT_GRID_THEME.background, clips: undefined });
  assert.equal(ctx.cellTexts.length, 0);
  // A damaged ratio falls back to 1.
  for (const dpr of [0, -1, Number.NaN, Number.POSITIVE_INFINITY]) assert.deepEqual(draw(newSheet(), { dpr }).transform, [1, 0, 0, 1, 0, 0]);
});

test('grid: only visible cells are drawn in a sheet with a million rows', () => {
  const sheet = newSheet();
  for (let r = 0; r < 1000000; r++) put(sheet, r, 0, r);
  for (let c = 1; c < 200; c++) put(sheet, 500000, c, 'far');
  const started = Date.now();
  const ctx = draw(sheet, { scrollY: 500000 * RH });
  assert.ok(Date.now() - started < 1000);
  // 578 px of cells: 29 rows of 20 px, the last one cut. 728 px across: 12 columns.
  assert.equal(ctx.cellTexts.filter((call) => call.text !== 'far').length, 29);
  assert.equal(ctx.cellTexts.filter((call) => call.text === 'far').length, 11);
  assert.equal(ctx.cellTexts[0].text, '500000');
  assert.equal(ctx.headerTexts.length, 29 + 12);
  assert.ok(ctx.measured.length <= 29 + 12);
});

test('grid: horizontal alignment', () => {
  const styles = emptyStyleSheet();
  const left = xf(styles, { alignment: { horizontal: 'left' } });
  const centre = xf(styles, { alignment: { horizontal: 'center' } });
  const right = xf(styles, { alignment: { horizontal: 'right' } });
  const indented = xf(styles, { alignment: { horizontal: 'left', indent: 2 } });
  const indentedRight = xf(styles, { alignment: { horizontal: 'right', indent: 1 } });
  const sheet = newSheet();
  // "General": text to the left, numbers to the right, booleans and errors in the middle.
  put(sheet, 0, 0, 'abc');
  put(sheet, 0, 1, 12);
  put(sheet, 0, 2, true);
  put(sheet, 0, 3, '#N/A', 0, 'error');
  put(sheet, 1, 0, 42, left);
  put(sheet, 1, 1, 'mid', centre);
  put(sheet, 1, 2, 'end', right);
  put(sheet, 1, 3, 'in', indented);
  put(sheet, 1, 4, 'in2', indentedRight);
  const ctx = draw(sheet, { styles });

  const at = (text: string) => ctx.textCall(text).x;
  assert.equal(at('abc'), HX + 3);
  assert.equal(at('12'), Math.round(HX + 2 * CW - 3 - 2 * CHAR));
  assert.equal(at('TRUE'), Math.round(HX + 2 * CW + (CW - 4 * CHAR) / 2));
  assert.equal(at('#N/A'), Math.round(HX + 3 * CW + (CW - 4 * CHAR) / 2));
  assert.equal(at('42'), HX + 3);
  assert.equal(at('mid'), Math.round(HX + CW + (CW - 3 * CHAR) / 2));
  assert.equal(at('end'), Math.round(HX + 3 * CW - 3 - 3 * CHAR));
  assert.equal(at('in'), HX + 3 * CW + 3 + 18);
  assert.equal(at('in2'), Math.round(HX + 5 * CW - 3 - 9 - 3 * CHAR));

  const abc = ctx.textCall('abc');
  assert.equal(abc.font, DEFAULT_FONT);
  assert.equal(abc.fill, DEFAULT_GRID_THEME.defaultText);
  assert.equal(abc.align, 'left');
  assert.equal(abc.y, BASELINE_ROW1);
  // Short text needs no clip of its own: only the pane clips it.
  assert.deepEqual(abc.clips, [box(HX, HY, 800 - HX, 600 - HY)]);
});

test('grid: vertical alignment', () => {
  const styles = emptyStyleSheet();
  const top = xf(styles, { alignment: { vertical: 'top' } });
  const centre = xf(styles, { alignment: { vertical: 'center' } });
  const sheet = newSheet();
  rowOf(sheet, 0).height = 45; // 60 px
  put(sheet, 0, 0, 'top', top);
  put(sheet, 0, 1, 'mid', centre);
  put(sheet, 0, 2, 'low');
  const ctx = draw(sheet, { styles });
  const line = PX * 1.2;
  const ascent = line - PX * 0.25;
  assert.equal(ctx.textCall('top').y, Math.round(HY + 1 + ascent));
  assert.equal(ctx.textCall('mid').y, Math.round(HY + (60 - line) / 2 + ascent));
  assert.equal(ctx.textCall('low').y, Math.round(HY + 60 - 1 - line + ascent));
});

test('grid: font, colour, underline and strike-through', () => {
  const styles = emptyStyleSheet();
  const fancy = xf(styles, {
    font: baseFont({ name: 'Arial', size: 12, bold: true, italic: true, underline: 'double', strike: true, color: { rgb: 'FFFF0000' } }),
  });
  const themed = xf(styles, { font: baseFont({ color: { theme: 1 } }) });
  styles.themeColors = ['FFFFFF', '112233'];
  const sheet = newSheet();
  put(sheet, 0, 0, 'ab', fancy);
  put(sheet, 1, 0, 'cd', themed);
  const ctx = draw(sheet, { styles });

  const call = ctx.textCall('ab');
  assert.equal(call.font, `italic bold 16px "Arial", ${FONT_FALLBACK}`);
  assert.equal(call.fill, '#FF0000');
  assert.equal(ctx.textCall('cd').fill, '#112233');
  // Two underline strokes and one strike-through, as wide as the text (2 characters of 8 px).
  const marks = ctx.fills('#FF0000').filter((rect) => rect.width === 16 && rect.height === 1);
  assert.equal(marks.length, 3);
  assert.ok(marks.every((rect) => rect.x === call.x));
  assert.equal(marks.filter((rect) => rect.y > call.y).length, 2);
  assert.equal(marks.filter((rect) => rect.y < call.y).length, 1);
});

test('grid: a font name from a file cannot break out of the font string', () => {
  const styles = emptyStyleSheet();
  const evil = xf(styles, { font: baseFont({ name: 'x"; } body { background: url(javascript:alert(1)) } /*', size: 11 }) });
  const sheet = newSheet();
  put(sheet, 0, 0, 'text', evil);
  const cell = put(sheet, 1, 0, 'rich');
  cell.runs = [
    { text: 'ri', font: baseFont({ name: "'; font: 99px evil" }) },
    { text: 'ch', font: baseFont({ name: '\\"}</style><script>' }) },
  ];
  const ctx = draw(sheet, { styles });
  assert.equal(ctx.textCall('text').font, `${PX}px "x body background urljavascriptalert1", ${FONT_FALLBACK}`);
  const safe = /^(italic )?(bold )?\d+(\.\d+)?px ("[\p{L}\p{M}\p{N} -]+", )?"Carlito", "Calibri", Arial, sans-serif$/u;
  assert.ok(ctx.fonts.length > 0);
  for (const font of ctx.fonts) assert.match(font, safe);
  // Colours are "#RRGGBB" or come from the theme of the grid.
  const theme = Object.values(DEFAULT_GRID_THEME);
  for (const call of ctx.rects) assert.ok(/^#[0-9A-F]{6}$/.test(call.fill) || theme.includes(call.fill), call.fill);
});

test('grid: text overflows into empty neighbours and stops at the first cell in use', () => {
  const long = 'abcdefghijklmnopqrst'; // 20 characters, about 147 px
  const styles = emptyStyleSheet();
  const right = xf(styles, { alignment: { horizontal: 'right' } });
  const centre = xf(styles, { alignment: { horizontal: 'center' } });
  const filled = xf(styles, { fill: solid('FFFFFF00') });
  const sheet = newSheet();
  put(sheet, 0, 0, long); // free to run over B1 and C1
  put(sheet, 1, 0, long + '1');
  put(sheet, 1, 1, 'stop'); // cut at the edge of A2
  put(sheet, 2, 0, long + '2');
  put(sheet, 2, 2, 7); // cut at the end of B3
  put(sheet, 3, 4, long + '3', right); // runs to the left over D4 and C4
  put(sheet, 4, 4, long + '4', right);
  put(sheet, 4, 2, 'x'); // cut at the start of D5
  put(sheet, 5, 3, long + '5', centre); // runs both ways
  put(sheet, 6, 3, long + '6', centre);
  put(sheet, 6, 2, 'l');
  put(sheet, 6, 4, 'r'); // cut on both sides
  put(sheet, 7, 0, long + '7');
  put(sheet, 7, 1, null, filled); // a blank cell with a format is still empty
  const ctx = draw(sheet, { styles });
  const pane = box(HX, HY, 800 - HX, 600 - HY);
  const rowBox = (row: number, x: number, width: number) => box(x, HY + row * RH, width, RH);

  assert.deepEqual(ctx.textCall(long).clips, [pane]);
  assert.equal(ctx.textCall(long).x, HX + 3);
  assert.deepEqual(ctx.textCall(long + '1').clips, [pane, rowBox(1, HX, CW)]);
  assert.deepEqual(ctx.textCall(long + '2').clips, [pane, rowBox(2, HX, 2 * CW)]);

  const width = 21 * CHAR;
  assert.deepEqual(ctx.textCall(long + '3').clips, [pane]);
  assert.equal(ctx.textCall(long + '3').x, Math.round(HX + 5 * CW - 3 - width));
  assert.deepEqual(ctx.textCall(long + '4').clips, [pane, rowBox(4, HX + 3 * CW, 2 * CW)]);

  assert.deepEqual(ctx.textCall(long + '5').clips, [pane]);
  assert.equal(ctx.textCall(long + '5').x, Math.round(HX + 3 * CW + (CW - width) / 2));
  assert.deepEqual(ctx.textCall(long + '6').clips, [pane, rowBox(6, HX + 3 * CW, CW)]);
  assert.deepEqual(ctx.textCall(long + '7').clips, [pane]);
});

test('grid: text of a cell scrolled out of view still overflows into view', () => {
  const sheet = newSheet();
  put(sheet, 0, 0, 'abcdefghijklmnopqrst');
  put(sheet, 1, 0, 'short');
  put(sheet, 2, 0, 123456);
  const ctx = draw(sheet, { scrollX: CW });
  // Column A is one column to the left of the cells; the pane clips what is under the header.
  const call = ctx.textCall('abcdefghijklmnopqrst');
  assert.equal(call.x, HX - CW + 3);
  assert.deepEqual(call.clips, [box(HX, HY, 800 - HX, 600 - HY)]);
  assert.equal(ctx.cellTexts.filter((entry) => entry.text === '123456').length, 0);
});

test('grid: numbers that do not fit show as #, other values are clipped', () => {
  const sheet = newSheet();
  put(sheet, 0, 0, 123456789012); // "1.23457E+11", 11 characters
  put(sheet, 1, 0, 1234567); // fits: 7 characters in 58 px
  put(sheet, 2, 0, '2024-01-01T00:00:00', 0, 'date');
  put(sheet, 3, 0, '#GETTING_DATA', 0, 'error');
  sheet.columns = [{ min: 2, max: 2, width: 0.5, customWidth: true, hidden: false, outlineLevel: 0, collapsed: false }];
  put(sheet, 0, 2, 5);
  const ctx = draw(sheet);
  const hashes = ctx.cellTexts.filter((call) => /^#+$/.test(call.text));
  assert.deepEqual(hashes.map((call) => call.text), ['#######', '#######']);
  // Right aligned like the number it stands for, and not clipped.
  assert.equal(hashes[0].x, Math.round(HX + CW - 3 - 7 * CHAR));
  assert.equal(hashes[0].clips.length, 1);
  assert.equal(ctx.textCall('1234567').clips.length, 1);
  assert.deepEqual(ctx.textCall('#GETTING_DATA').clips[1], box(HX, HY + 3 * RH, CW, RH));
  // A column too narrow for a single "#" shows nothing.
  assert.equal(ctx.cellTexts.filter((call) => call.text === '5').length, 0);
});

test('grid: wrapped text', () => {
  const styles = emptyStyleSheet();
  const wrap = xf(styles, { alignment: { wrapText: true, vertical: 'top' } });
  const sheet = newSheet();
  rowOf(sheet, 0).height = 60; // 80 px
  put(sheet, 0, 0, 'aaa bbb ccc ddd', wrap); // 7 characters fit into 58 px
  put(sheet, 0, 1, 'abcdefghijklmnop', wrap);
  put(sheet, 0, 2, 'one\ntwo', wrap);
  put(sheet, 0, 3, 123456789012, wrap); // numbers do not wrap
  put(sheet, 1, 0, 'aaa bbb ccc ddd', wrap); // too high for a 20 px row: clipped
  put(sheet, 2, 0, 'one\ntwo'); // no wrap: the break is not shown
  const ctx = draw(sheet, { styles });

  const inCell = (col: number) => ctx.cellTexts.filter((call) => call.y < HY + 80 && call.x >= HX + col * CW && call.x < HX + (col + 1) * CW);
  assert.deepEqual(inCell(0).map((call) => call.text), ['aaa bbb', 'ccc ddd']);
  assert.deepEqual(inCell(1).map((call) => call.text), ['abcdefg', 'hijklmn', 'op']);
  assert.deepEqual(inCell(2).map((call) => call.text), ['one', 'two']);
  assert.deepEqual(inCell(3).map((call) => call.text), ['#######']);

  const line = PX * 1.2;
  const ascent = line - PX * 0.25;
  assert.deepEqual(inCell(0).map((call) => call.y), [Math.round(HY + 1 + ascent), Math.round(HY + 1 + line + ascent)]);
  assert.ok(inCell(0).every((call) => call.x === HX + 3 && call.clips.length === 1));

  const clipped = ctx.cellTexts.filter((call) => call.y > HY + 80 && call.y < HY + 80 + 2 * RH && call.clips.length === 2);
  assert.ok(clipped.length >= 1);
  assert.deepEqual(clipped[0].clips[1], box(HX, HY + 80, CW, RH));
  assert.equal(ctx.textCall('onetwo').x, HX + 3);
});

test('grid: rich text draws each run in its own font and colour', () => {
  const sheet = newSheet();
  const cell = put(sheet, 0, 0, 'plain bold red');
  cell.runs = [
    { text: 'plain ' },
    { text: 'bold ', font: baseFont({ bold: true, size: 20 }) },
    { text: 'red', font: baseFont({ color: { rgb: 'FF0000' } }) },
  ];
  const ctx = draw(sheet);
  const calls = ctx.cellTexts;
  assert.deepEqual(calls.map((call) => call.text), ['plain ', 'bold ', 'red']);
  assert.deepEqual(calls.map((call) => call.fill), ['#000000', '#000000', '#FF0000']);
  assert.equal(calls[1].font, `bold 26.67px "Calibri", ${FONT_FALLBACK}`);
  // One after the other on a shared baseline.
  assert.equal(calls[0].x, HX + 3);
  assert.equal(calls[1].x, HX + 3 + 6 * CHAR);
  assert.equal(calls[2].x, HX + 3 + 6 * CHAR + 5 * (26.67 / 2));
  assert.ok(calls.every((call) => call.y === calls[0].y));
});

test('grid: fills, and the precedence of cell, row and column formats', () => {
  const styles = emptyStyleSheet();
  const yellow = xf(styles, { fill: solid('FFFFFF00') });
  const red = xf(styles, { fill: solid('FF0000') });
  const blue = xf(styles, { fill: solid('0000FF') });
  const pattern = xf(styles, { fill: { type: 'pattern', pattern: 'darkGray', fg: { rgb: '00FF00' }, bg: { rgb: '123456' } } });
  const none = xf(styles, { fill: { type: 'pattern', pattern: 'none', fg: { rgb: 'ABCDEF' } } });
  const gradient = xf(styles, { fill: { type: 'gradient', raw: { name: 'gradientFill' } as never } });
  const sheet = newSheet();
  put(sheet, 0, 0, 'a', yellow);
  put(sheet, 0, 1, null, pattern);
  put(sheet, 0, 2, null, none);
  put(sheet, 0, 3, null, gradient);
  sheet.columns = [{ min: 5, max: 6, customWidth: false, hidden: false, style: blue, outlineLevel: 0, collapsed: false }];
  rowOf(sheet, 3).style = red;
  put(sheet, 3, 5, null, yellow);
  const ctx = draw(sheet, { styles });

  const cellBox = (row: number, col: number) => ({ x: HX + col * CW - 1, y: HY + row * RH - 1, width: CW + 1, height: RH + 1 });
  const has = (colour: string, row: number, col: number) => {
    const b = cellBox(row, col);
    return ctx.fills(colour).some((rect) => rect.x === b.x && rect.y === b.y && rect.width === b.width && rect.height === b.height);
  };
  // The fill reaches one pixel up and left, over the grid lines of the neighbours.
  assert.ok(has('#FFFF00', 0, 0));
  assert.ok(has('#00FF00', 0, 1));
  assert.equal(ctx.fills('#ABCDEF').length, 0);
  assert.equal(ctx.fills('#123456').length, 0);

  // Column F and G are blue where nothing else says otherwise.
  assert.ok(has('#0000FF', 0, 5) && has('#0000FF', 0, 6) && has('#0000FF', 10, 5));
  assert.ok(!has('#0000FF', 0, 4) && !has('#0000FF', 0, 7));
  // Row 4 is red, also where it crosses the blue column; the cell F4 has its own format.
  assert.ok(has('#FF0000', 3, 0) && has('#FF0000', 3, 6));
  assert.ok(!has('#0000FF', 3, 6) && !has('#FF0000', 3, 5) && !has('#0000FF', 3, 5));
  assert.ok(has('#FFFF00', 3, 5));
});

test('grid: grid lines are drawn before the fills and can be switched off', () => {
  const styles = emptyStyleSheet();
  const yellow = xf(styles, { fill: solid('FFFF00') });
  const sheet = newSheet();
  put(sheet, 1, 1, null, yellow);
  const ctx = draw(sheet, { styles });
  const lines = ctx.strokesOf(DEFAULT_GRID_THEME.gridLine);
  assert.equal(lines.length, 1);
  // On the last pixel of each column and row, across the pane.
  assert.ok(hasLine(lines, [HX + CW - 0.5, HY, HX + CW - 0.5, 600]));
  assert.ok(hasLine(lines, [HX, HY + RH - 0.5, 800, HY + RH - 0.5]));
  assert.equal(lines[0].width, 1);
  assert.deepEqual(lines[0].clips, [box(HX, HY, 800 - HX, 600 - HY)]);

  sheet.view.showGridLines = false;
  const plain = draw(sheet, { styles });
  assert.equal(plain.strokesOf(DEFAULT_GRID_THEME.gridLine).length, 0);
  assert.equal(plain.fills('#FFFF00').length, 1);
});

test('grid: border styles, colours and the heavier of two shared edges', () => {
  const styles = emptyStyleSheet();
  const b2 = xf(styles, {
    border: {
      left: { style: 'thin' },
      top: { style: 'medium', color: { rgb: 'FF0000' } },
      right: { style: 'thick' },
      bottom: { style: 'dashed' },
    },
  });
  const c2 = xf(styles, { border: { left: { style: 'thin', color: { rgb: '0000FF' } }, right: { style: 'double' }, bottom: { style: 'hair' } } });
  const odd = xf(styles, { border: { top: { style: 'constructor' }, bottom: { style: 'none' }, left: { style: 'dotted' } } });
  const sheet = newSheet();
  put(sheet, 1, 1, null, b2);
  put(sheet, 1, 2, null, c2);
  put(sheet, 4, 1, null, odd);
  const ctx = draw(sheet, { styles });
  const borders = ctx.strokes.filter((call) => call.style !== DEFAULT_GRID_THEME.gridLine && call.clips.length === 1 && call.clips[0].x === HX && call.clips[0].y === HY);
  const x = HX + CW; // left edge of B
  const y = HY + RH; // top edge of row 2
  const find = (width: number, colour: string, dash: number[] = []) =>
    borders.filter((call) => call.width === width && call.style === colour && call.dash.join() === dash.join());

  // Thin: one pixel, on the grid line (the last pixel of the neighbour).
  assert.ok(hasLine(find(1, '#000000'), [x - 0.5, y - 1, x - 0.5, y + RH]));
  // Medium: two pixels, centred on the pixel boundary.
  assert.ok(hasLine(find(2, '#FF0000'), [x - 1, y, x + CW, y]));
  // Thick wins over the thin left edge of C2: three pixels, and no blue line at all.
  assert.ok(hasLine(find(3, '#000000'), [x + CW - 0.5, y - 1, x + CW - 0.5, y + RH]));
  assert.equal(borders.filter((call) => call.style === '#0000FF').length, 0);
  assert.ok(hasLine(find(1, '#000000', [3, 2]), [x - 1, y + RH - 0.5, x + CW, y + RH - 0.5]));
  // Double: two thin lines with a gap.
  const right = x + 2 * CW;
  const double = find(1, '#000000').filter((call) => call.lines.some((l) => l[0] === right - 1.5));
  assert.equal(double.length, 1);
  assert.ok(hasLine(double, [right - 1.5, y - 1, right - 1.5, y + RH]) && hasLine(double, [right + 0.5, y - 1, right + 0.5, y + RH]));
  // Hair: one pixel, lighter.
  const hair = borders.filter((call) => call.alpha === 0.5);
  assert.equal(hair.length, 1);
  assert.ok(hasLine(hair, [x + CW - 1, y + RH - 0.5, x + 2 * CW, y + RH - 0.5]));
  // An unknown style name draws as thin, "none" draws nothing, dotted has its own dash.
  const y5 = HY + 4 * RH;
  assert.ok(hasLine(find(1, '#000000'), [x - 1, y5 - 0.5, x + CW, y5 - 0.5]));
  assert.ok(!borders.some((call) => call.lines.some((l) => l[1] === l[3] && l[1] >= y5 + RH - 2 && l[1] <= y5 + RH + 1 && l[0] === x - 1)));
  assert.ok(hasLine(find(1, '#000000', [1, 1]), [x - 0.5, y5 - 1, x - 0.5, y5 + RH]));
  // Dash and alpha do not leak into what is drawn next.
  assert.ok(ctx.strokes[ctx.strokes.length - 1].dash.length === 0 && ctx.strokes[ctx.strokes.length - 1].alpha === 1);
});

test('grid: borders on the first row and column are inside the pane', () => {
  const styles = emptyStyleSheet();
  const boxed = xf(styles, { border: { left: { style: 'thin' }, top: { style: 'thin' } } });
  const sheet = newSheet();
  put(sheet, 0, 0, null, boxed);
  const ctx = draw(sheet, { styles });
  const borders = ctx.strokesOf('#000000').filter((call) => call.clips.length === 1 && call.clips[0].y === HY);
  // The grid line pixel before A1 belongs to the headers, so the edges move one pixel in.
  assert.ok(hasLine(borders, [HX + 0.5, HY - 1, HX + 0.5, HY + RH]));
  assert.ok(hasLine(borders, [HX - 1, HY + 0.5, HX + CW, HY + 0.5]));
});

test('grid: merged cells', () => {
  const styles = emptyStyleSheet();
  const centre = xf(styles, { fill: solid('FFFF00'), alignment: { horizontal: 'center', vertical: 'center' } });
  const edges = xf(styles, { border: { left: { style: 'thin' }, right: { style: 'thin' }, top: { style: 'thin' }, bottom: { style: 'thin' } }, fill: solid('00FF00') });
  const sheet = newSheet();
  put(sheet, 1, 1, 'merged', centre);
  put(sheet, 1, 2, 'hidden', edges);
  put(sheet, 2, 1, 'hidden', edges);
  put(sheet, 2, 2, 'hidden', edges);
  put(sheet, 5, 0, 'abcdefghijklmnopqrstuvwxyz');
  put(sheet, 6, 0, 'abcdefghijklmnopqrstuvwxyz0123456789'); // 264 px, the merge is 192 px wide
  sheet.merges = [range(1, 1, 2, 2), range(5, 1, 5, 2), range(6, 0, 6, 2), range(8, 8, 8, 8)];
  for (let i = 0; i < 20000; i++) sheet.merges.push(range(1000 + i * 2, 0, 1000 + i * 2, 3));
  const started = Date.now();
  const ctx = draw(sheet, { styles });
  assert.ok(Date.now() - started < 500);

  const x = HX + CW;
  const y = HY + RH;
  // Only the top-left cell is shown, laid out in the whole area.
  assert.equal(ctx.cellTexts.filter((call) => call.text === 'hidden').length, 0);
  const call = ctx.textCall('merged');
  assert.equal(call.x, Math.round(x + (2 * CW - 6 * CHAR) / 2));
  const line = PX * 1.2;
  assert.equal(call.y, Math.round(y + (2 * RH - line) / 2 + line - PX * 0.25));
  // Its fill covers the area; the fills of the other cells are not drawn.
  assert.deepEqual(ctx.fills('#FFFF00').map((rect) => [rect.x, rect.y, rect.width, rect.height]), [[x - 1, y - 1, 2 * CW + 1, 2 * RH + 1]]);
  assert.equal(ctx.fills('#00FF00').length, 0);
  // Borders: only on the outline, each from the cell that lies there. B2 has none; C2, B3 and C3 have.
  const borders = ctx.strokesOf('#000000').filter((entry) => entry.clips.length === 1 && entry.clips[0].y === HY);
  assert.ok(hasLine(borders, [x + 2 * CW - 0.5, y - 1, x + 2 * CW - 0.5, y + RH])); // right of C2
  assert.ok(hasLine(borders, [x - 1, y + 2 * RH - 0.5, x + CW, y + 2 * RH - 0.5])); // bottom of B3
  assert.ok(!hasLine(borders, [x + CW - 0.5, y - 1, x + CW - 0.5, y + RH])); // between B2 and C2
  assert.ok(!hasLine(borders, [x + CW - 1, y + RH - 0.5, x + 2 * CW, y + RH - 0.5])); // between C2 and C3
  // An unfilled merge is painted in the background colour to take its inner grid lines out.
  const inner = ctx.fills(DEFAULT_GRID_THEME.background).filter((rect) => rect.clips.length === 1);
  assert.deepEqual(inner.map((rect) => [rect.x, rect.y, rect.width, rect.height]), [
    [x, HY + 5 * RH, 2 * CW - 1, RH - 1],
    [HX, HY + 6 * RH, 3 * CW - 1, RH - 1],
  ]);
  // Text does not overflow into a merge, and the text of a merge is clipped to it.
  assert.deepEqual(ctx.textCall('abcdefghijklmnopqrstuvwxyz').clips[1], box(HX, HY + 5 * RH, CW, RH));
  assert.deepEqual(ctx.textCall('abcdefghijklmnopqrstuvwxyz0123456789').clips[1], box(HX, HY + 6 * RH, 3 * CW, RH));
});

test('grid: a merge whose top-left cell is scrolled out of view', () => {
  const sheet = newSheet();
  put(sheet, 0, 0, 'title');
  sheet.merges = [range(0, 0, 9, 5)];
  const ctx = draw(sheet, { scrollX: 2 * CW, scrollY: 3 * RH });
  // Drawn once, at the position the scrolled area gives it.
  assert.equal(ctx.textCall('title').x, HX - 2 * CW + 3);
  assert.equal(ctx.textCall('title').y, Math.round(HY - 3 * RH + 10 * RH - 1 - PX * 0.25));
});

test('grid: frozen panes', () => {
  const sheet = newSheet();
  sheet.view.frozenRows = 1;
  sheet.view.frozenCols = 1;
  put(sheet, 0, 0, 'corner');
  put(sheet, 0, 3, 'top');
  put(sheet, 3, 0, 'side');
  put(sheet, 3, 3, 'body');
  put(sheet, 1, 1, 'gone'); // scrolled under the frozen parts
  const ctx = draw(sheet, { scrollX: 2 * CW, scrollY: 2 * RH });

  const corner = box(HX, HY, CW, RH);
  const top = box(HX + CW, HY, 800 - HX - CW, RH);
  const side = box(HX, HY + RH, CW, 600 - HY - RH);
  const body = box(HX + CW, HY + RH, 800 - HX - CW, 600 - HY - RH);
  const xScrolled = HX + 3 * CW - 2 * CW + 3;
  const yScrolled = BASELINE_ROW1 + 3 * RH - 2 * RH;
  assert.deepEqual(ctx.textCall('corner'), { ...ctx.textCall('corner'), x: HX + 3, y: BASELINE_ROW1, clips: [corner] });
  assert.deepEqual(ctx.textCall('top'), { ...ctx.textCall('top'), x: xScrolled, y: BASELINE_ROW1, clips: [top] });
  assert.deepEqual(ctx.textCall('side'), { ...ctx.textCall('side'), x: HX + 3, y: yScrolled, clips: [side] });
  assert.deepEqual(ctx.textCall('body'), { ...ctx.textCall('body'), x: xScrolled, y: yScrolled, clips: [body] });
  assert.equal(ctx.cellTexts.filter((call) => call.text === 'gone').length, 0);

  // Headers: the frozen ones stay, the others follow the scrolling.
  const header = (text: string) => {
    const found = ctx.headerTexts.filter((call) => call.text === text);
    assert.equal(found.length, 1, text);
    return found[0];
  };
  assert.equal(header('A').x, HX + CW / 2);
  assert.deepEqual(header('A').clips, [box(HX, 0, CW, HY)]);
  assert.equal(header('D').x, HX + CW + CW / 2);
  assert.deepEqual(header('D').clips, [box(HX + CW, 0, 800 - HX - CW, HY)]);
  assert.equal(header('1').y, HY + RH / 2);
  assert.deepEqual(header('1').clips, [box(0, HY, HX, RH)]);
  assert.equal(header('4').y, HY + RH + RH / 2);
  assert.deepEqual(header('4').clips, [box(0, HY + RH, HX, 600 - HY - RH)]);
  assert.equal(ctx.headerTexts.filter((call) => call.text === 'B' || call.text === '2').length, 0);

  // A line marks the frozen boundary, over the headers too.
  const frozen = ctx.strokesOf(DEFAULT_GRID_THEME.frozenLine);
  assert.equal(frozen.length, 1);
  assert.deepEqual(frozen[0].lines, [
    [HX + CW - 0.5, 0, HX + CW - 0.5, 600],
    [0, HY + RH - 0.5, 800, HY + RH - 0.5],
  ]);
  assert.equal(frozen[0].clips.length, 0);
  assert.equal(draw(newSheet()).strokesOf(DEFAULT_GRID_THEME.frozenLine).length, 0);
});

test('grid: a merge across the frozen boundary is drawn by each pane under its own clip', () => {
  const styles = emptyStyleSheet();
  const yellow = xf(styles, { fill: solid('FFFF00') });
  const sheet = newSheet();
  sheet.view.frozenCols = 1;
  put(sheet, 0, 0, 'wide', yellow);
  sheet.merges = [range(0, 0, 0, 3)];
  const ctx = draw(sheet, { styles, scrollX: CW });
  const calls = ctx.cellTexts.filter((call) => call.text === 'wide');
  assert.equal(calls.length, 2);
  // The frozen pane lays the merge out unscrolled, the other one scrolled.
  assert.deepEqual(calls.map((call) => call.x), [HX + 3, HX - CW + 3]);
  assert.deepEqual(calls.map((call) => call.clips[0]), [box(HX, HY, CW, 600 - HY), box(HX + CW, HY, 800 - HX - CW, 600 - HY)]);
  assert.deepEqual(ctx.fills('#FFFF00').map((rect) => [rect.x, rect.width, rect.clips[0].x]), [
    [HX - 1, 4 * CW + 1, HX],
    [HX - CW - 1, 4 * CW + 1, HX + CW],
  ]);
});

test('grid: selection', () => {
  const sheet = newSheet();
  const ctx = draw(sheet, { selection: { active: { row: 1, col: 1 }, range: range(1, 1, 2, 2) } });
  const x = HX + CW;
  const y = HY + RH;
  // B2:C3 with B2 active: the fill leaves the active cell out.
  const fill = ctx.fills(DEFAULT_GRID_THEME.selectionFill);
  assert.deepEqual(fill.map((rect) => [rect.x, rect.y, rect.width, rect.height]), [
    [x, y + RH, 2 * CW, RH],
    [x + CW, y, CW, RH],
  ]);
  assert.deepEqual(fill[0].clips, [box(HX, HY, 800 - HX, 600 - HY)]);
  const frames = ctx.strokeRects.filter((rect) => rect.style === DEFAULT_GRID_THEME.selectionBorder);
  assert.deepEqual(frames, [
    { x: x - 0.5, y: y - 0.5, width: CW, height: RH, style: DEFAULT_GRID_THEME.selectionBorder, lineWidth: 1 },
    { x, y, width: 2 * CW - 1, height: 2 * RH - 1, style: DEFAULT_GRID_THEME.selectionBorder, lineWidth: 2 },
  ]);
  // The headers of the selected rows and columns are highlighted.
  const marked = ctx.fills(DEFAULT_GRID_THEME.headerSelectedBackground);
  assert.deepEqual(marked.map((rect) => [rect.x, rect.y, rect.width, rect.height]), [
    [x, 0, CW, HY],
    [x + CW, 0, CW, HY],
    [0, y, HX, RH],
    [0, y + RH, HX, RH],
  ]);

  // A single cell: a frame and no fill.
  const one = draw(sheet, { selection: { active: { row: 0, col: 0 }, range: range(0, 0, 0, 0) } });
  assert.equal(one.fills(DEFAULT_GRID_THEME.selectionFill).length, 0);
  assert.deepEqual(one.strokeRects.map((rect) => [rect.x, rect.y, rect.width, rect.height, rect.lineWidth]), [[HX, HY, CW - 1, RH - 1, 2]]);

  // An active cell inside a merge leaves the whole merge unfilled.
  sheet.merges = [range(1, 1, 2, 2)];
  const merged = draw(sheet, { selection: { active: { row: 2, col: 2 }, range: range(1, 1, 2, 3) } });
  assert.deepEqual(merged.fills(DEFAULT_GRID_THEME.selectionFill).map((rect) => [rect.x, rect.y, rect.width, rect.height]), [[x + 2 * CW, y, CW, 2 * RH]]);
  sheet.merges = [];

  // A selection out of view draws nothing in the cells; a whole column or a damaged one does no harm.
  const far = draw(sheet, { selection: { active: { row: 5000, col: 0 }, range: range(5000, 0, 5001, 0) } });
  assert.equal(far.strokeRects.length, 0);
  const column = draw(sheet, { selection: { active: { row: 0, col: 1 }, range: range(0, 1, 1048575, 1) } });
  assert.equal(column.strokeRects.length, 2);
  for (const bad of [range(Number.NaN, 0, 1, 1), range(5, 5, 1, 1), range(-10, -10, -5, -5)]) {
    assert.equal(draw(sheet, { selection: { active: { row: 0, col: 0 }, range: bad } }).strokeRects.length, 0);
  }
  assert.equal(draw(sheet, { selection: { active: { row: Number.NaN, col: 0 }, range: range(0, 0, 0, 0) } }).strokeRects.length, 0);
});

test('grid: the selection is clipped by the frozen panes', () => {
  const sheet = newSheet();
  sheet.view.frozenCols = 1;
  const ctx = draw(sheet, { scrollX: CW, selection: { active: { row: 0, col: 0 }, range: range(0, 0, 0, 2) } });
  const fill = ctx.fills(DEFAULT_GRID_THEME.selectionFill);
  // Each pane draws the range in its own coordinates and clips it.
  assert.deepEqual(fill.map((rect) => [rect.x, rect.width, rect.clips[0].x, rect.clips[0].width]), [
    [HX + CW, 2 * CW, HX, CW],
    [HX, 2 * CW, HX + CW, 800 - HX - CW],
  ]);
});

test('grid: header labels, and hidden rows and columns', () => {
  const sheet = newSheet();
  const plain = draw(sheet, { width: HX + 3 * CW, height: HY + 3 * RH });
  assert.deepEqual(plain.headerTexts.map((call) => call.text), ['A', 'B', 'C', '1', '2', '3']);
  assert.deepEqual(plain.headerTexts.map((call) => call.x), [HX + 32, HX + CW + 32, HX + 2 * CW + 32, 36, 36, 36]);
  assert.deepEqual(plain.headerTexts.map((call) => call.y), [11, 11, 11, HY + 10, HY + RH + 10, HY + 2 * RH + 10]);
  assert.ok(plain.headerTexts.every((call) => call.align === 'center' && call.fill === DEFAULT_GRID_THEME.headerText));
  // The header strips and their closing lines, which also frame the corner box.
  const back = plain.fills(DEFAULT_GRID_THEME.headerBackground);
  assert.deepEqual(back.map((rect) => [rect.x, rect.y, rect.width, rect.height]), [
    [0, 0, HX + 3 * CW, HY],
    [0, 0, HX, HY + 3 * RH],
  ]);
  const lines = plain.strokesOf(DEFAULT_GRID_THEME.headerBorder);
  assert.ok(hasLine(lines, [0, HY - 0.5, HX + 3 * CW, HY - 0.5]) && hasLine(lines, [HX - 0.5, 0, HX - 0.5, HY + 3 * RH]));
  assert.ok(hasLine(lines, [HX + CW - 0.5, 0, HX + CW - 0.5, HY]) && hasLine(lines, [0, HY + RH - 0.5, HX, HY + RH - 0.5]));
  // Headers are drawn after the cells.
  put(sheet, 0, 0, 'cell');
  const order = draw(sheet).texts.map((call) => call.text);
  assert.ok(order.indexOf('cell') < order.indexOf('A'));

  sheet.columns = [{ min: 1, max: 2, customWidth: false, hidden: true, outlineLevel: 0, collapsed: false }];
  rowOf(sheet, 1).hidden = true;
  put(sheet, 0, 1, 'hidden col');
  put(sheet, 1, 0, 'hidden row');
  put(sheet, 2, 3, 'seen');
  const ctx = draw(sheet, { width: HX + 3 * CW, height: HY + 3 * RH });
  assert.deepEqual(ctx.headerTexts.map((call) => call.text), ['A', 'D', 'E', '1', '3', '4']);
  assert.deepEqual(ctx.cellTexts.map((call) => call.text), ['cell', 'seen']);
  // D follows A directly, row 3 follows row 1.
  assert.equal(ctx.textCall('seen').x, HX + CW + 3);
  assert.equal(ctx.textCall('seen').y, BASELINE_ROW1 + RH);
});

test('grid: zeros can be hidden', () => {
  const sheet = newSheet();
  put(sheet, 0, 0, 0);
  put(sheet, 0, 1, 10);
  assert.deepEqual(draw(sheet).cellTexts.map((call) => call.text), ['0', '10']);
  sheet.view.showZeros = false;
  assert.deepEqual(draw(sheet).cellTexts.map((call) => call.text), ['10']);
});

test('grid: a huge cell text does not stall a frame', () => {
  const styles = emptyStyleSheet();
  const wrap = xf(styles, { alignment: { wrapText: true, vertical: 'top' } });
  const sheet = newSheet();
  const huge = 'lorem ipsum '.repeat(90000); // about 1 MB
  put(sheet, 0, 0, huge);
  put(sheet, 1, 0, huge, wrap);
  put(sheet, 2, 0, 'z'.repeat(1000000), wrap);
  const cell = put(sheet, 3, 0, 'runs');
  cell.runs = [];
  for (let i = 0; i < 200000; i++) cell.runs.push({ text: i % 2 ? 'ab' : '' });
  clearTextMeasureCache();
  const started = Date.now();
  const ctx = draw(sheet, { styles });
  assert.ok(Date.now() - started < 1500, `took ${Date.now() - started} ms`);
  // Nothing longer than the cut reaches the canvas, and only lines that can be seen are drawn.
  assert.ok(ctx.measured.every((text) => text.length <= 2000), `measured ${Math.max(...ctx.measured.map((t) => t.length))}`);
  assert.ok(ctx.cellTexts.every((call) => call.text.length <= 2000), 'drawn text too long');
  assert.equal(ctx.cellTexts[0].text.length, 2000);
  assert.ok(ctx.cellTexts.length < 20, `${ctx.cellTexts.length} text calls`);
  clearTextMeasureCache();
});

test('grid: a typical frame is fast and measures each text once', () => {
  const styles = emptyStyleSheet();
  const boxed = xf(styles, {
    fill: solid('EEEEEE'),
    border: { left: { style: 'thin' }, right: { style: 'thin' }, top: { style: 'thin' }, bottom: { style: 'thin' } },
  });
  const sheet = newSheet();
  for (let r = 0; r < 50; r++) for (let c = 0; c < 30; c++) put(sheet, r, c, c % 2 ? r * c : `r${r}c${c}`, r % 3 ? 0 : boxed);
  const options = { styles, width: HX + 30 * CW, height: HY + 50 * RH };
  clearTextMeasureCache();
  const first = draw(sheet, options);
  assert.equal(first.cellTexts.length, 1500);
  // Thin black borders all go into one path.
  assert.equal(first.strokesOf('#000000').length, 1);
  // The font is assigned when it changes, not per cell.
  assert.ok(first.fonts.length < 20, `${first.fonts.length} font assignments`);

  const started = Date.now();
  let again = first;
  for (let i = 0; i < 20; i++) again = draw(sheet, options);
  const perFrame = (Date.now() - started) / 20;
  assert.ok(perFrame < 50, `${perFrame} ms per frame`);
  assert.equal(again.measured.length, 0);
  clearTextMeasureCache();
});

test('grid: empty and degenerate canvases, sheets without cells', () => {
  const sheet = newSheet();
  put(sheet, 0, 0, 'a');
  for (const [width, height] of [[0, 0], [0, 600], [800, 0], [-50, -50], [Number.NaN, 100]]) {
    const ctx = draw(sheet, { width, height });
    assert.equal(ctx.texts.length, 0);
    assert.equal(ctx.rects.length, 0);
    assert.equal(ctx.transform.length, 6);
  }
  // Smaller than the headers: no pane, no cell text.
  assert.equal(draw(sheet, { width: 30, height: 10 }).cellTexts.length, 0);
  assert.equal(draw(sheet, { width: HX + 1, height: HY + 1 }).cellTexts.length, 1);

  const chart = createSheet({ name: 'Chart', sheetId: 2, state: 'visible', kind: 'chartsheet', part: '/c' });
  assert.equal(draw(chart).cellTexts.length, 0);
});

test('grid: a damaged model does not throw', () => {
  const styles = emptyStyleSheet();
  const sheet = newSheet();
  put(sheet, 0, 0, 'style out of range', 999);
  put(sheet, 1, 0, 'negative style', -3);
  put(sheet, 2, 0, 'nan style', Number.NaN);
  const odd = xf(styles, {
    font: baseFont({ size: Number.NaN, color: { rgb: 'not a colour"; x' } }),
    fill: { type: 'pattern', pattern: 'solid', fg: { rgb: 'red; }' } },
    border: { left: { style: '__proto__', color: { rgb: 'javascript:' } } },
    alignment: { indent: Number.POSITIVE_INFINITY, horizontal: 'weird' as never, vertical: 'odd' as never },
  });
  put(sheet, 3, 0, 'odd', odd);
  sheet.merges = [range(Number.NaN, 0, 2, 2), range(9, 9, 3, 3), range(-5, -5, 2000000, 20000), range(4, 0, 4, 0)];
  sheet.merges.push(range(0, 0, 1, 1), range(1, 1, 2, 2)); // overlapping
  const ctx = draw(sheet, { styles });
  assert.ok(ctx.cellTexts.length >= 1);
  for (const call of ctx.texts) assert.match(call.fill, /^#[0-9A-Fa-f]{6}$/);
  for (const call of ctx.strokes) assert.match(call.style, /^#[0-9A-Fa-f]{6}$/);
});

test('grid: number formats decide the text, its colour and the date system', () => {
  const styles = emptyStyleSheet();
  const format = (code: string) => {
    const id = 200 + styles.numFmts.size;
    styles.numFmts.set(id, code);
    return styles.cellXfs.push({ numFmtId: id, fontId: 0, fillId: 0, borderId: 0, quotePrefix: false }) - 1;
  };
  const money = format('#,##0.00;[Red]-#,##0.00');
  const date = format('yyyy-mm-dd');
  const percent = format('0%');
  const sheet = newSheet();
  // 140 px: wide enough for every text below.
  sheet.columns = [{ min: 0, max: 0, width: 20, customWidth: true, hidden: false, outlineLevel: 0, collapsed: false }];
  put(sheet, 0, 0, 1234.5, money);
  put(sheet, 1, 0, -1234.5, money);
  put(sheet, 2, 0, 45000, date);
  put(sheet, 3, 0, -1, date);
  put(sheet, 4, 0, 0.256, percent);
  put(sheet, 5, 0, true);

  const en = draw(sheet, { styles });
  assert.equal(en.textCall('1,234.50').fill, DEFAULT_GRID_THEME.defaultText);
  assert.equal(en.textCall('-1,234.50').fill, '#FF0000');
  assert.ok(en.textCall('2023-03-15'));
  assert.ok(en.textCall('26%'));
  assert.ok(en.textCall('TRUE'));
  // A negative date has no text: it fills the cell with "#", right aligned like a number.
  const hashes = en.cellTexts.filter((call) => /^#+$/.test(call.text));
  assert.equal(hashes.length, 1);
  assert.equal(hashes[0].text.length, Math.floor((140 - 6) / CHAR));
  assert.equal(hashes[0].y, en.textCall('2023-03-15').y + RH);

  // Two format colours in the same font are two text styles, not one.
  const blue = format('[Blue]0.0;[Red]-0.0');
  put(sheet, 6, 0, 7, blue);
  put(sheet, 7, 0, -7, blue);
  const colours = draw(sheet, { styles });
  assert.deepEqual([colours.textCall('7.0').fill, colours.textCall('-7.0').fill, colours.textCall('-1,234.50').fill], ['#0000FF', '#FF0000', '#FF0000']);

  const tr = draw(sheet, { styles, locale: LOCALE_TR });
  assert.ok(tr.textCall('1.234,50'));
  assert.equal(tr.textCall('-1.234,50').fill, '#FF0000');
  assert.ok(tr.textCall('DOĞRU'));

  // The same serial in the 1904 date system is four years and a day later.
  assert.ok(draw(sheet, { styles, date1904: true }).textCall('2027-03-16'));
});

// --- conditional formats -----------------------------------------------------

/** A source that answers from a table of cells: "row,col" -> what the cell gets. */
function fixedConditional(cells: Record<string, CellConditional>): ConditionalSource & { asked: string[]; frames: number } {
  const source = {
    asked: [] as string[],
    frames: 0,
    begin() {
      source.frames++;
    },
    resolve(row: number, col: number) {
      source.asked.push(`${row},${col}`);
      return cells[`${row},${col}`];
    },
  };
  return source;
}

test('grid: a conditional format replaces the fill and goes on top of font, border and number format', () => {
  const styles = emptyStyleSheet();
  const yellow = xf(styles, { fill: solid('FFFF00'), font: baseFont({ italic: true, color: { rgb: '00FF00' } }), border: { bottom: { style: 'thin', color: { rgb: '0000FF' } } } });
  const sheet = newSheet();
  put(sheet, 0, 0, 'plain', yellow);
  put(sheet, 1, 0, 'styled', yellow);
  put(sheet, 2, 0, 12.5, yellow);
  put(sheet, 3, 0, 'nofill', yellow);
  const format = { fill: '#FF0000', bold: true, color: '#123456', strike: true, underline: 'double' as const, top: { style: 'medium', color: { rgb: 'AA00AA' } } };
  const source = fixedConditional({
    '1,0': { style: format },
    '2,0': { style: { numFmt: { id: 164, code: '0.00"!"' }, italic: false } },
    '3,0': { style: { fill: null } },
    // An empty cell inside a range is filled as well.
    '1,2': { style: format },
    // Colours that are not "#RRGGBB" never reach the canvas.
    '5,0': { style: { fill: 'red; drop', color: 'url(x)' }, fill: 'javascript:1' },
  });
  put(sheet, 5, 0, 'unsafe');
  const ctx = draw(sheet, { styles, conditional: source });
  assert.equal(source.frames, 1);

  // Row 1 keeps its own format.
  assert.equal(ctx.fills('#FFFF00').filter((call) => call.y === HY - 1).length, 1);
  assert.equal(ctx.textCall('plain').fill, '#00FF00');
  assert.match(ctx.textCall('plain').font, /^italic /);

  // Row 2: red instead of yellow, bold added to the italic of the cell, the colour replaced.
  assert.deepEqual(ctx.fills('#FF0000').map((call) => [call.x, call.y, call.width, call.height]), [
    [HX - 1, HY + RH - 1, CW + 1, RH + 1],
    [HX + 2 * CW - 1, HY + RH - 1, CW + 1, RH + 1],
  ]);
  assert.equal(ctx.fills('#FFFF00').filter((call) => call.y === HY + RH - 1).length, 0);
  const styled = ctx.textCall('styled');
  assert.equal(styled.fill, '#123456');
  assert.match(styled.font, /italic/);
  assert.match(styled.font, /bold/);
  // Strike-through and a double underline: three thin rectangles in the text colour under and through the text.
  assert.equal(ctx.fills('#123456').filter((call) => call.x === styled.x && call.height <= 2).length, 3);
  // The border of the format on top of the cell's own: medium purple above, the thin blue one below stays.
  assert.ok(ctx.strokesOf('#AA00AA').some((call) => call.width === 2 && call.lines.some((l) => l[1] === HY + RH && l[3] === HY + RH)));
  assert.ok(ctx.strokesOf('#0000FF').length > 0);

  // Row 3: the number format of the rule, the italic of the cell switched off.
  const number = ctx.textCall('12.50!');
  assert.doesNotMatch(number.font, /italic/);
  // Row 4: "no fill" takes the cell's fill away.
  assert.equal(ctx.fills('#FFFF00').filter((call) => call.y === HY + 3 * RH - 1).length, 0);
  assert.ok(ctx.textCall('nofill'));

  assert.equal(ctx.textCall('unsafe').fill, DEFAULT_GRID_THEME.defaultText);
  for (const call of ctx.rects) assert.match(call.fill, /^(#[0-9A-Fa-f]{6}|rgba\(.*\))$/);

  // Without a source nothing is asked and nothing changes.
  const without = draw(sheet, { styles });
  assert.equal(without.fills('#FF0000').length, 0);
  assert.equal(without.textCall('styled').fill, '#00FF00');
});

test('grid: colour scale fills, data bars and icons', () => {
  const sheet = newSheet();
  put(sheet, 0, 0, 5);
  put(sheet, 1, 0, 7);
  put(sheet, 2, 0, -3);
  put(sheet, 3, 0, 9);
  put(sheet, 4, 0, 11);
  put(sheet, 5, 0, 13);
  const styles = emptyStyleSheet();
  const source = fixedConditional({
    '0,0': { fill: '#ABCDEF', style: { fill: '#FF0000' } },
    '1,0': { bar: { from: 0, to: 0.5, color: '#00B050', gradient: false, negative: false, rightToLeft: false, border: '#005500' } },
    '2,0': { bar: { from: 0.25, to: 0.5, color: '#FF0000', gradient: true, negative: true, rightToLeft: false, axis: 0.5, axisColor: '#333333' } },
    '3,0': { bar: { from: 0, to: 0.5, color: '#00B050', gradient: false, negative: false, rightToLeft: true }, hideValue: true },
    '4,0': { icon: { set: '3TrafficLights1', index: 2 } },
    '5,0': { icon: { set: '3Arrows', index: 0 }, hideValue: true },
  });
  const ctx = draw(sheet, { styles, conditional: source });

  // The colour of the scale wins over the fill of a format.
  assert.equal(ctx.fills('#ABCDEF').length, 1);
  assert.equal(ctx.fills('#FF0000').length, 0);

  // A bar leaves two pixels to every side of the cell (the last pixel of a cell is its grid line): 59 x 15 to fill.
  const room = CW - 5;
  assert.deepEqual(
    ctx.fills('#00B050').map((call) => [call.x, call.y, call.width, call.height]),
    [
      [HX + 2, HY + RH + 2, Math.round(room / 2), RH - 5],
      // Right to left: the same bar from the other side.
      [HX + 2 + Math.round(room / 2), HY + 3 * RH + 2, room - Math.round(room / 2), RH - 5],
    ],
  );
  assert.deepEqual(ctx.strokeRects.filter((call) => call.style === '#005500').map((call) => [call.x, call.y, call.width, call.height]), [[HX + 2.5, HY + RH + 2.5, Math.round(room / 2) - 1, RH - 6]]);
  // The bar is drawn before the text of its cell.
  assert.ok(ctx.textCall('7'));

  // A negative gradient bar: solid at the axis, faded at its far end; the axis is a dashed line over the whole cell.
  assert.equal(ctx.gradientRects.length, 1);
  const negative = ctx.gradientRects[0];
  assert.deepEqual([negative.x, negative.y, negative.width, negative.height], [HX + 2 + Math.round(room * 0.25), HY + 2 * RH + 2, Math.round(room * 0.5) - Math.round(room * 0.25), RH - 5]);
  assert.deepEqual(negative.gradient.stops, [
    [1, '#FF0000'],
    [0, '#FFD9D9'],
  ]);
  const axis = ctx.strokesOf('#333333');
  assert.equal(axis.length, 1);
  assert.deepEqual(axis[0].dash, [1, 1]);
  assert.deepEqual(axis[0].lines, [[HX + 2 + Math.round(room / 2) + 0.5, HY + 2 * RH, HX + 2 + Math.round(room / 2) + 0.5, HY + 3 * RH - 1]]);
  assert.ok(ctx.textCall('-3'));

  // No value shown: the bar is there, the text is not.
  assert.equal(ctx.cellTexts.filter((call) => call.text === '9').length, 0);

  // An icon is a circle at the left of the cell; the number stays at the right.
  const circles = ctx.paths.filter((path) => path.arcs.length === 1 && path.fill === '#68A490');
  assert.equal(circles.length, 1);
  const size = Math.floor(PX * 1.05);
  assert.deepEqual(circles[0].arcs[0].slice(0, 3), [HX + 3 + size / 2, HY + 5 * RH - 2 - size + size / 2, size * 0.44]);
  assert.equal(ctx.textCall('11').x, Math.round(HX + CW - 3 - 2 * CHAR));
  // Icon only: an arrow (seven corners) in red, no text.
  const arrowsDrawn = ctx.paths.filter((path) => path.points.length === 7 && path.fill === '#D65532');
  assert.equal(arrowsDrawn.length, 1);
  assert.ok(arrowsDrawn[0].points.every(([x, y]) => x >= HX + 3 && x <= HX + 3 + size && y >= HY + 5 * RH && y <= HY + 6 * RH));
  assert.equal(ctx.cellTexts.filter((call) => call.text === '13').length, 0);
});

test('grid: an icon pushes left-aligned text in and narrows the room of a number', () => {
  const styles = emptyStyleSheet();
  const left = xf(styles, { alignment: { horizontal: 'left' } });
  const centred = xf(styles, { alignment: { horizontal: 'center', vertical: 'top' } });
  const sheet = newSheet();
  put(sheet, 0, 0, 42, left);
  put(sheet, 1, 0, 12345678);
  put(sheet, 2, 0, 5, centred);
  const size = Math.floor(PX * 1.05);
  const ctx = draw(sheet, {
    styles,
    conditional: fixedConditional({
      '0,0': { icon: { set: '5Rating', index: 2 } },
      '1,0': { icon: { set: '3Flags', index: 0 } },
      '2,0': { icon: { set: '3Signs', index: 0 }, hideValue: true },
      // A set or a position that does not exist draws nothing and moves nothing.
      '3,0': { icon: { set: 'constructor', index: 0 } },
    }),
  });
  assert.equal(ctx.textCall('42').x, HX + size + 3 + 3);
  // Eight digits are 58.7 px; they fit a 64 px cell but not what the icon leaves of it.
  assert.equal(ctx.cellTexts.filter((call) => call.text === '12345678').length, 0);
  assert.equal(ctx.cellTexts.filter((call) => /^#+$/.test(call.text)).length, 1);
  // Rating: four bars, two of them in the colour of the set.
  assert.equal(ctx.fills('#4472C4').length, 2);
  assert.equal(ctx.fills('#C8C8C8').length, 2);
  // The icon without a value follows the alignment of its cell: centred, at the top.
  // UNSURE: with "general" alignment such an icon is put at the left; Excel is believed to do the same.
  const diamond = ctx.paths.filter((path) => path.points.length === 4 && path.fill === '#D65532');
  assert.equal(diamond.length, 1);
  assert.deepEqual(diamond[0].points[0], [Math.round(HX + (CW - size) / 2) + size * 0.5, HY + 2 * RH + 1 + size * 0.04]);
});

test('grid: every icon of every set can be drawn', () => {
  const sets = ['3Arrows', '3ArrowsGray', '3Flags', '3TrafficLights1', '3TrafficLights2', '3Signs', '3Symbols', '3Symbols2', '3Stars', '3Triangles', '4Arrows', '4ArrowsGray', '4RedToBlack', '4Rating', '4TrafficLights', '5Arrows', '5ArrowsGray', '5Rating', '5Quarters', '5Boxes'];
  const sheet = newSheet();
  const cells: Record<string, CellConditional> = {};
  sets.forEach((set, r) => {
    for (let i = 0; i < Number(set[0]); i++) {
      put(sheet, r, i, i);
      cells[`${r},${i}`] = { icon: { set, index: i }, hideValue: true };
    }
  });
  const ctx = draw(sheet, { conditional: fixedConditional(cells) });
  assert.equal(ctx.cellTexts.length, 0);
  // Each of the 74 icons leaves something on the canvas, inside its own cell.
  for (let r = 0; r < sets.length; r++) {
    for (let i = 0; i < Number(sets[r][0]); i++) {
      const inCell = (x: number, y: number) => x >= HX + i * CW && x <= HX + (i + 1) * CW && y >= HY + r * RH && y <= HY + (r + 1) * RH;
      const drawn =
        ctx.paths.some((path) => path.points.some(([x, y]) => inCell(x, y)) || path.arcs.some(([x, y]) => inCell(x, y))) ||
        ctx.rects.some((rect) => rect.width < 20 && rect.height < 20 && inCell(rect.x, rect.y)) ||
        ctx.strokes.some((stroke) => stroke.lines.length > 0 && stroke.lines.every((l) => inCell(l[0], l[1]) && inCell(l[2], l[3])));
      assert.ok(drawn, `${sets[r]} icon ${i}`);
    }
  }
  // Nothing but "#RRGGBB" and the theme's own colours was used.
  for (const path of ctx.paths) assert.match(path.fill, /^#[0-9A-F]{6}$/);
});

test('grid: a merged area is formatted as its top-left cell', () => {
  const sheet = newSheet();
  put(sheet, 1, 1, 8);
  put(sheet, 2, 2, 'hidden by the merge');
  sheet.merges = [range(1, 1, 2, 3)];
  const source = fixedConditional({
    '1,1': { style: { fill: '#FF0000', bold: true }, bar: { from: 0, to: 1, color: '#00B050', gradient: false, negative: false, rightToLeft: false } },
    // What the other cells of the area would get does not count.
    '2,2': { style: { fill: '#0000FF' } },
  });
  const ctx = draw(sheet, { conditional: source });
  assert.deepEqual(ctx.fills('#FF0000').map((call) => [call.x, call.y, call.width, call.height]), [[HX + CW - 1, HY + RH - 1, 3 * CW + 1, 2 * RH + 1]]);
  assert.equal(ctx.fills('#0000FF').length, 0);
  assert.deepEqual(ctx.fills('#00B050').map((call) => [call.x, call.y, call.width, call.height]), [[HX + CW + 2, HY + RH + 2, 3 * CW - 5, 2 * RH - 5]]);
  assert.match(ctx.textCall('8').font, /bold/);
  assert.ok(!source.asked.includes('2,2'));
});

test('grid: the rules of a sheet reach the canvas through ConditionalFormats', () => {
  const styles = emptyStyleSheet();
  styles.dxfFormats = [{ fill: { type: 'pattern', pattern: 'solid', bg: { rgb: 'FFFFC7CE' } }, font: { color: { rgb: 'FF9C0006' } } }];
  const sheet = newSheet();
  for (let r = 0; r < 5; r++) put(sheet, r, 0, r * 10);
  const base = { stopIfTrue: false, percent: false, bottom: false, aboveAverage: true, equalAverage: false };
  sheet.conditionalFormats = [
    { ranges: [range(0, 0, 4, 0)], rules: [{ ...base, type: 'cellIs', priority: 1, operator: 'greaterThan', formulas: ['25'], dxfId: 0 }] },
    { ranges: [range(0, 0, 4, 0)], rules: [{ ...base, type: 'colorScale', priority: 2, formulas: [], colorScale: { values: [{ type: 'min', gte: true }, { type: 'max', gte: true }], colors: [{ rgb: 'FF000000' }, { rgb: 'FFFFFFFF' }] } }] },
  ];
  const ctx = draw(sheet, { styles, conditional: true });
  // 30 and 40 take the pink of the format; the cells before them the colour of the scale.
  assert.deepEqual(ctx.fills('#FFC7CE').map((call) => call.y), [HY + 3 * RH - 1, HY + 4 * RH - 1]);
  assert.equal(ctx.textCall('40').fill, '#9C0006');
  assert.equal(ctx.textCall('10').fill, DEFAULT_GRID_THEME.defaultText);
  assert.equal(ctx.fills('#000000').filter((call) => call.y === HY - 1).length, 1);
  assert.equal(ctx.fills('#404040').length, 1);
  assert.equal(ctx.fills('#808080').length, 1);
});

test('grid: a sheet without conditional formats is drawn without asking for any', () => {
  const sheet = newSheet();
  for (let r = 0; r < 30; r++) for (let c = 0; c < 12; c++) put(sheet, r, c, r * c);
  const before = draw(sheet);
  const source = fixedConditional({});
  const after = draw(sheet, { conditional: source });
  // A source that has nothing to say changes nothing on the canvas.
  assert.deepEqual(after.texts, before.texts);
  assert.deepEqual(after.rects, before.rects);
  assert.equal(after.paths.length, 0);
  // A source that throws at construction time or is not a source at all is ignored.
  assert.doesNotThrow(() => draw(sheet, { conditional: {} as ConditionalSource }));
});

// --- conditional formats: what the first tests let through ----------------------

test('grid: a right-to-left data bar mirrors its gradient and its axis', () => {
  const sheet = newSheet();
  put(sheet, 0, 0, 5);
  put(sheet, 1, 0, -5);
  put(sheet, 2, 0, 5);
  const ctx = draw(sheet, {
    conditional: fixedConditional({
      '0,0': { bar: { from: 0.25, to: 0.75, color: '#00B050', gradient: true, negative: false, rightToLeft: true, axis: 0.25, axisColor: '#111111' } },
      '1,0': { bar: { from: 0, to: 0.25, color: '#FF0000', gradient: true, negative: true, rightToLeft: true, axis: 0.25, axisColor: '#222222' } },
      '2,0': { bar: { from: 0.25, to: 0.75, color: '#0000FF', gradient: true, negative: false, rightToLeft: false, axis: 0.25, axisColor: '#333333' } },
    }),
  });
  const room = CW - 5;
  const x = (share: number) => HX + 2 + Math.round(room * share);
  assert.equal(ctx.gradientRects.length, 3);
  const [positive, negative, plain] = ctx.gradientRects;
  // Mirrored: the bar that ran from 25% to 75% is still there, but it starts at its right end.
  assert.deepEqual([positive.x, positive.width], [x(0.25), x(0.75) - x(0.25)]);
  assert.deepEqual(positive.gradient.stops, [[1, '#00B050'], [0, fadeOf('#00B050')]]);
  // A negative bar right-to-left lies right of the axis and is solid at its left end, the axis.
  assert.deepEqual([negative.x, negative.width], [x(0.75), x(1) - x(0.75)]);
  assert.deepEqual(negative.gradient.stops, [[0, '#FF0000'], [1, fadeOf('#FF0000')]]);
  // Left to right, for comparison: solid at the left.
  assert.deepEqual(plain.gradient.stops, [[0, '#0000FF'], [1, fadeOf('#0000FF')]]);
  // The axis at 25% from the start is at 75% from the left when mirrored.
  assert.deepEqual(ctx.strokesOf('#111111')[0].lines, [[x(0.75) + 0.5, HY, x(0.75) + 0.5, HY + RH - 1]]);
  assert.deepEqual(ctx.strokesOf('#222222')[0].lines, [[x(0.75) + 0.5, HY + RH, x(0.75) + 0.5, HY + 2 * RH - 1]]);
  assert.deepEqual(ctx.strokesOf('#333333')[0].lines, [[x(0.25) + 0.5, HY + 2 * RH, x(0.25) + 0.5, HY + 3 * RH - 1]]);
});

/** A colour faded towards white the way the end of a gradient bar is (three quarters of the way). */
function fadeOf(color: string): string {
  let out = '#';
  for (let i = 1; i < 7; i += 2) {
    const channel = parseInt(color.slice(i, i + 2), 16);
    const value = Math.round(channel + (255 - channel) * 0.85);
    out += (value < 16 ? '0' : '') + value.toString(16).toUpperCase();
  }
  return out;
}

test('grid: the colour of a colour scale fills a merged area', () => {
  const styles = emptyStyleSheet();
  const yellow = xf(styles, { fill: solid('FFFF00') });
  const sheet = newSheet();
  put(sheet, 1, 1, 8, yellow);
  sheet.merges = [range(1, 1, 2, 3)];
  const source = fixedConditional({
    // The scale wins over the fill of a format and over the cell's own fill.
    '1,1': { fill: '#ABCDEF', style: { fill: '#FF0000' } },
    '2,3': { fill: '#123456' },
  });
  const ctx = draw(sheet, { styles, conditional: source });
  assert.deepEqual(ctx.fills('#ABCDEF').map((call) => [call.x, call.y, call.width, call.height]), [[HX + CW - 1, HY + RH - 1, 3 * CW + 1, 2 * RH + 1]]);
  assert.equal(ctx.fills('#FF0000').length, 0);
  assert.equal(ctx.fills('#FFFF00').length, 0);
  assert.equal(ctx.fills('#123456').length, 0);
  // A colour that is not a colour leaves the fill of the format, and without one the cell's own.
  const unsafe = draw(sheet, { styles, conditional: fixedConditional({ '1,1': { fill: 'url(x)', style: { fill: '#FF0000' } } }) });
  assert.equal(unsafe.fills('#FF0000').length, 1);
  const own = draw(sheet, { styles, conditional: fixedConditional({ '1,1': { fill: 'url(x)' } }) });
  assert.equal(own.fills('#FFFF00').length, 1);
});

test('grid: the number format of a conditional format follows the date system of the workbook', () => {
  const sheet = newSheet();
  put(sheet, 0, 0, 2);
  put(sheet, 1, 0, 2);
  const source = () => fixedConditional({ '0,0': { style: { numFmt: { id: 164, code: 'yy-m-d' } } } });
  // Serial 2 is 2 January 1900, and 3 January 1904 in the 1904 system.
  const ctx1900 = draw(sheet, { conditional: source() });
  assert.ok(ctx1900.textCall('00-1-2'));
  assert.ok(ctx1900.textCall('2'));
  const ctx1904 = draw(sheet, { conditional: source(), date1904: true });
  assert.ok(ctx1904.textCall('04-1-3'));
  assert.ok(ctx1904.textCall('2'));
});

test('grid: when an icon leaves no room, the text is not drawn', () => {
  const sheet = newSheet();
  // Column A is 19 px wide: an icon of 13 px and its gap leave 3 px, less than the padding on both sides.
  sheet.columns = [{ min: 0, max: 0, width: 2, customWidth: true, hidden: false, outlineLevel: 0, collapsed: false }];
  // A text, which would otherwise be drawn clipped rather than replaced by "#".
  put(sheet, 0, 0, 'ab');
  put(sheet, 0, 1, 7);
  put(sheet, 1, 0, 7);
  const ctx = draw(sheet, {
    conditional: fixedConditional({
      '0,0': { icon: { set: '3TrafficLights1', index: 2 } },
      '0,1': { icon: { set: '3TrafficLights1', index: 2 } },
    }),
  });
  const circles = ctx.paths.filter((path) => path.arcs.length === 1 && path.fill === '#68A490');
  assert.equal(circles.length, 2);
  // Neither the number nor a row of "#" in the narrow cell; the wider one and the cell without an icon keep theirs.
  const inA1 = ctx.cellTexts.filter((call) => call.y < HY + RH && call.x < HX + 19);
  assert.deepEqual(inA1.map((call) => call.text), []);
  assert.equal(ctx.cellTexts.filter((call) => call.text === '7').length, 2);
});

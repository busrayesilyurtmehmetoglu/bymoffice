import assert from 'node:assert/strict';
import { test } from 'node:test';
import { LOCALE_TR } from '../src/format/locale';
import { createSheetMetrics } from '../src/layout/sheet-metrics';
import { MAX_ROWS, createSheet, emptyStyleSheet } from '../src/model/sheet/types';
import type { Cell, Row, Sheet, SheetPageSetup, Workbook } from '../src/model/sheet/types';
import { renderGridPrint } from '../src/render/grid';
import { DocumentPrintError } from '../src/ui/doc-print';
import { defaultPaperFor, headerFooterFor, headerFooterValues, paintSheetPage, planSheetPages, type SheetJobPlan } from '../src/ui/sheet-print';

interface TextCall {
  text: string;
  /** Position on the canvas, the transform applied. */
  x: number;
  y: number;
  font: string;
  fill: string;
  scale: number;
}

interface RectCall {
  x: number;
  y: number;
  width: number;
  height: number;
  fill: string;
}

interface StrokeCall {
  style: string;
  /** Width on the canvas. */
  width: number;
  /** [x1, y1, x2, y2] on the canvas. */
  lines: number[][];
}

/** A stand-in for CanvasRenderingContext2D that records what is drawn, in canvas pixels. Each character is half the font size wide. */
class FakeContext {
  fillStyle = '#000000';
  strokeStyle = '#000000';
  lineWidth = 1;
  globalAlpha = 1;
  textAlign = 'left';
  textBaseline = 'alphabetic';
  font = '10px sans-serif';
  texts: TextCall[] = [];
  rects: RectCall[] = [];
  strokes: StrokeCall[] = [];
  clips: RectCall[] = [];
  saves = 0;
  private matrix = [1, 0, 0, 1, 0, 0];
  private stack: { matrix: number[]; fill: string; stroke: string; font: string; lineWidth: number }[] = [];
  private path: number[][] = [];
  private pathRects: number[][] = [];
  private pen: number[] = [0, 0];

  private point(x: number, y: number): number[] {
    const m = this.matrix;
    return [m[0] * x + m[4], m[3] * y + m[5]];
  }
  setTransform(a: number, _b: number, _c: number, d: number, e: number, f: number): void {
    this.matrix = [a, 0, 0, d, e, f];
  }
  save(): void {
    this.saves++;
    this.stack.push({ matrix: this.matrix.slice(), fill: this.fillStyle, stroke: this.strokeStyle, font: this.font, lineWidth: this.lineWidth });
  }
  restore(): void {
    this.saves--;
    const state = this.stack.pop();
    if (!state) return;
    this.matrix = state.matrix;
    this.fillStyle = state.fill;
    this.strokeStyle = state.stroke;
    this.font = state.font;
    this.lineWidth = state.lineWidth;
  }
  setLineDash(): void {}
  beginPath(): void {
    this.path = [];
    this.pathRects = [];
  }
  closePath(): void {}
  moveTo(x: number, y: number): void {
    this.pen = this.point(x, y);
  }
  lineTo(x: number, y: number): void {
    const to = this.point(x, y);
    this.path.push([this.pen[0], this.pen[1], to[0], to[1]]);
    this.pen = to;
  }
  rect(x: number, y: number, width: number, height: number): void {
    const from = this.point(x, y);
    this.pathRects.push([from[0], from[1], width * this.matrix[0], height * this.matrix[3]]);
  }
  clip(): void {
    for (const r of this.pathRects) this.clips.push({ x: r[0], y: r[1], width: r[2], height: r[3], fill: '' });
  }
  stroke(): void {
    this.strokes.push({ style: this.strokeStyle, width: this.lineWidth * this.matrix[0], lines: this.path.slice() });
  }
  fill(): void {}
  arc(): void {}
  fillRect(x: number, y: number, width: number, height: number): void {
    const from = this.point(x, y);
    this.rects.push({ x: from[0], y: from[1], width: width * this.matrix[0], height: height * this.matrix[3], fill: this.fillStyle });
  }
  strokeRect(): void {}
  fillText(text: string, x: number, y: number): void {
    const at = this.point(x, y);
    this.texts.push({ text, x: at[0], y: at[1], font: this.font, fill: this.fillStyle, scale: this.matrix[0] });
  }
  measureText(text: string): { width: number } {
    const match = /([\d.]+)px/.exec(this.font);
    return { width: (text.length * (match ? Number(match[1]) : 10)) / 2 };
  }
}

const asContext = (fake: FakeContext) => fake as unknown as CanvasRenderingContext2D;

function row(cells: Cell[] = [], extra: Partial<Row> = {}): Row {
  return { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells, ...extra };
}

function makeSheet(name: string, rows: number, cols: number, setup?: SheetPageSetup): Sheet {
  const sheet = createSheet({ name, sheetId: 1, state: 'visible', kind: 'worksheet', part: `/xl/worksheets/${name}.xml` });
  for (let r = 0; r < rows; r++) {
    const cells: Cell[] = [];
    for (let c = 0; c < cols; c++) cells[c] = { kind: 'string', value: `r${r}c${c}`, style: 0 };
    sheet.rows.set(r, row(cells));
  }
  if (setup) sheet.pageSetup = setup;
  return sheet;
}

function book(sheets: Sheet[], names: Workbook['definedNames'] = []): Workbook {
  return { sheets, styles: emptyStyleSheet(), sharedStrings: [], definedNames: names, date1904: false, activeSheet: 0 };
}

const codeOf = (run: () => unknown): string => {
  try {
    run();
  } catch (error) {
    if (error instanceof DocumentPrintError) return error.code;
    throw error;
  }
  return 'none';
};

const numbers = (plan: SheetJobPlan) => plan.pages.map((page) => `${page.sheet}:${page.index}:${page.number}`);

test('sheet print: the default paper follows the language of the view', () => {
  for (const tag of ['en', 'en-US', 'en_US', 'fr-CA', 'es-US', 'EN']) assert.equal(defaultPaperFor(tag), 1, tag);
  for (const tag of ['tr', 'tr-TR', 'en-GB', 'de', 'de-DE', '', undefined]) assert.equal(defaultPaperFor(tag), 9, String(tag));
});

test('sheet print: the pages of a job run through its sheets and are numbered through', () => {
  const workbook = book([makeSheet('A', 60, 2), makeSheet('B', 3, 2), makeSheet('C', 100, 2)]);
  const one = planSheetPages(workbook, [0]);
  assert.deepEqual(numbers(one), ['0:0:1', '0:1:2']);
  assert.deepEqual(one.sheets, [0]);
  const all = planSheetPages(workbook, [0, 1, 2]);
  assert.deepEqual(numbers(all), ['0:0:1', '0:1:2', '1:0:3', '2:0:4', '2:1:5', '2:2:6']);
  assert.deepEqual(all.sheets, [0, 1, 2]);
  // A sheet that states the number of its first page starts there, and the next sheet goes on from it.
  workbook.sheets[1].pageSetup = { page: { useFirstPageNumber: true, firstPageNumber: 10 } };
  assert.deepEqual(numbers(planSheetPages(workbook, [0, 1, 2])), ['0:0:1', '0:1:2', '1:0:10', '2:0:11', '2:1:12', '2:2:13']);
  // The paper of a sheet without one is the one handed in.
  assert.deepEqual(planSheetPages(workbook, [1], { defaultPaper: 1 }).pages[0].layout.paper, { width: 816, height: 1056 });
  assert.deepEqual(planSheetPages(workbook, [1]).pages[0].layout.paper, { width: 793.7, height: 1122.52 });
});

test('sheet print: sheets with nothing to print, chart sheets and sheets that do not exist are left out; a job without pages is empty', () => {
  const empty = createSheet({ name: 'Empty', sheetId: 2, state: 'visible', kind: 'worksheet', part: '/e' });
  const chart = makeSheet('Chart', 2, 2);
  chart.kind = 'chartsheet';
  const workbook = book([empty, makeSheet('Data', 2, 2), chart]);
  const plan = planSheetPages(workbook, [0, 1, 2, 7, -1]);
  assert.deepEqual(numbers(plan), ['1:0:1']);
  assert.deepEqual(plan.sheets, [1]);
  assert.equal(codeOf(() => planSheetPages(workbook, [0])), 'empty');
  assert.equal(codeOf(() => planSheetPages(workbook, [2])), 'empty');
  assert.equal(codeOf(() => planSheetPages(workbook, [])), 'empty');
});

test('sheet print: the print area and the print titles of the defined names are used', () => {
  const workbook = book(
    [makeSheet('Data', 100, 5)],
    [
      { name: '_xlnm.Print_Area', value: 'Data!$A$1:$C$60', localSheet: 0, hidden: false },
      { name: '_xlnm.Print_Titles', value: 'Data!$1:$1', localSheet: 0, hidden: false },
    ],
  );
  const plan = planSheetPages(workbook, [0]);
  assert.deepEqual(plan.pages.map((entry) => [entry.page.rows, entry.page.cols, entry.page.titleRows]), [
    [{ first: 0, last: 47 }, { first: 0, last: 2 }, undefined],
    [{ first: 48, last: 59 }, { first: 0, last: 2 }, { first: 0, last: 0 }],
  ]);
});

test('sheet print: a job beyond the limit on pages is too long, whichever sheet goes over it', () => {
  const far = createSheet({ name: 'Far', sheetId: 1, state: 'visible', kind: 'worksheet', part: '/f' });
  far.rows.set(MAX_ROWS - 1, row([{ kind: 'number', value: 1, style: 0 }]));
  const start = Date.now();
  assert.equal(codeOf(() => planSheetPages(book([far]), [0])), 'too-long');
  assert.ok(Date.now() - start < 20000);
  const workbook = book([makeSheet('A', 100, 2), makeSheet('B', 100, 2)]);
  assert.equal(planSheetPages(workbook, [0, 1], { maxPages: 6 }).pages.length, 6);
  assert.equal(codeOf(() => planSheetPages(workbook, [0, 1], { maxPages: 5 })), 'too-long');
  assert.equal(codeOf(() => planSheetPages(workbook, [0, 1], { maxPages: 3 })), 'too-long');
  assert.equal(codeOf(() => planSheetPages(workbook, [0, 1], { maxPages: 2 })), 'too-long');
  // An empty sheet after the limit was reached does not make the job too long.
  const empty = createSheet({ name: 'E', sheetId: 3, state: 'visible', kind: 'worksheet', part: '/e' });
  assert.equal(planSheetPages(book([makeSheet('A', 100, 2), empty]), [0, 1], { maxPages: 3 }).pages.length, 3);
});

test('sheet print: which header and footer a page gets', () => {
  const setup = { oddHeader: 'odd', oddFooter: 'oddF', evenHeader: 'even', evenFooter: 'evenF', firstHeader: 'first', firstFooter: 'firstF' };
  assert.deepEqual(headerFooterFor(undefined, true, 1), { header: undefined, footer: undefined });
  // Without the switches the odd ones are those of every page.
  assert.deepEqual(headerFooterFor(setup, true, 1), { header: 'odd', footer: 'oddF' });
  assert.deepEqual(headerFooterFor(setup, false, 2), { header: 'odd', footer: 'oddF' });
  assert.deepEqual(headerFooterFor({ ...setup, differentOddEven: true }, false, 2), { header: 'even', footer: 'evenF' });
  assert.deepEqual(headerFooterFor({ ...setup, differentOddEven: true }, false, 3), { header: 'odd', footer: 'oddF' });
  // The first page of the sheet, whatever its number is.
  assert.deepEqual(headerFooterFor({ ...setup, differentFirst: true, differentOddEven: true }, true, 4), { header: 'first', footer: 'firstF' });
  assert.deepEqual(headerFooterFor({ ...setup, differentFirst: true, differentOddEven: true }, false, 4), { header: 'even', footer: 'evenF' });
  // A first page that is to differ and has no header of its own has none.
  assert.deepEqual(headerFooterFor({ oddHeader: 'odd', differentFirst: true }, true, 1), { header: undefined, footer: undefined });
});

test('sheet print: the values of the fields: page numbers, the sheet, the file, date and time in the locale of the view', () => {
  const workbook = book([makeSheet('Özet', 60, 2)]);
  const plan = planSheetPages(workbook, [0]);
  const now = new Date(2026, 9, 4, 14, 5, 9);
  assert.deepEqual(headerFooterValues({ workbook, job: plan.pages[1], total: 2, now, fileName: 'rapor.xlsx', locale: LOCALE_TR }), {
    page: 2,
    pages: 2,
    date: '04.10.2026',
    time: '14:05',
    file: 'rapor.xlsx',
    path: '',
    sheet: 'Özet',
  });
  const english = headerFooterValues({ workbook, job: plan.pages[0], total: 2, now });
  assert.equal(english.date, '10/4/2026');
  assert.equal(english.time, '14:05');
  assert.equal(english.file, '');
  // The 1904 calendar of a workbook shows the same day.
  workbook.date1904 = true;
  assert.equal(headerFooterValues({ workbook, job: plan.pages[0], total: 2, now }).date, '10/4/2026');
  // A moment that is none shows nothing.
  assert.equal(headerFooterValues({ workbook, job: plan.pages[0], total: 2, now: new Date(NaN) }).date, '');
});

test('sheet print: a page is white paper with its cells at the margins, scaled to the canvas', () => {
  const workbook = book([makeSheet('Data', 3, 2)]);
  const plan = planSheetPages(workbook, [0]);
  const ctx = new FakeContext();
  // Twice the px of the paper: 793.7 px -> 1587.4 pixels.
  paintSheetPage(asContext(ctx), { workbook, job: plan.pages[0], total: 1, metrics: createSheetMetrics(workbook.sheets[0]), pixelWidth: 1587.4, now: new Date(2026, 0, 1) });
  // The first thing painted is the paper, all of it.
  assert.equal(ctx.rects[0].fill, '#ffffff');
  assert.ok(ctx.rects[0].x === 0 && ctx.rects[0].y === 0 && ctx.rects[0].width >= 1587.4 && ctx.rects[0].height >= 2245.04);
  // The cells are clipped to their block: 2 columns by 3 rows at the left and top margin (67.2 px and 72 px, doubled).
  const clip = ctx.clips[0];
  assert.ok(Math.abs(clip.x - 134.4) < 1e-6 && Math.abs(clip.y - 144) < 1e-6 && Math.abs(clip.width - 256) < 1e-6 && Math.abs(clip.height - 120) < 1e-6);
  assert.deepEqual(ctx.texts.map((call) => call.text), ['r0c0', 'r0c1', 'r1c0', 'r1c1', 'r2c0', 'r2c1']);
  // Text of A1 starts 3 px into the cell; that of B1 one column (64 px) further, twice that on the canvas.
  assert.ok(Math.abs(ctx.texts[0].x - (134.4 + 6)) < 1e-6, String(ctx.texts[0].x));
  assert.ok(Math.abs(ctx.texts[1].x - ctx.texts[0].x - 128) < 1e-6);
  assert.ok(ctx.texts[0].y > 144 && ctx.texts[0].y < 144 + 40);
  assert.equal(ctx.texts[0].scale, 2);
  // No grid lines unless the sheet asks for them, no headings, and every save has its restore.
  assert.deepEqual(ctx.strokes, []);
  assert.equal(ctx.saves, 0);
});

test('sheet print: the scale of the sheet, grid lines and headings', () => {
  const workbook = book([makeSheet('Data', 2, 2, { page: { scale: 50 }, print: { gridLines: true, headings: true } })]);
  const plan = planSheetPages(workbook, [0]);
  const ctx = new FakeContext();
  paintSheetPage(asContext(ctx), { workbook, job: plan.pages[0], total: 1, metrics: createSheetMetrics(workbook.sheets[0]), pixelWidth: 793.7, now: new Date(2026, 0, 1) });
  // Half size: a column is 32 px on the paper, the headings 20 px wide and 10 px high.
  assert.equal(ctx.texts.filter((call) => call.text === 'r0c0')[0].scale, 0.5);
  const labels = ctx.texts.filter((call) => /^(A|B|1|2)$/.test(call.text)).map((call) => call.text);
  assert.deepEqual(labels, ['A', 'B', '1', '2']);
  const a = ctx.texts.filter((call) => call.text === 'A')[0];
  assert.ok(Math.abs(a.x - (67.2 + 20 + 16)) < 1e-6 && Math.abs(a.y - (72 + 5)) < 1e-6);
  // Grid lines: black hairlines, a quarter of a px wide at this scale.
  const grid = ctx.strokes.filter((stroke) => Math.abs(stroke.width - 0.25) < 1e-9);
  assert.ok(grid.length >= 2);
  assert.equal(grid.every((stroke) => stroke.style === '#000000'), true);
  // The line that closes column A: at 20 + 64 px of the sheet, less half a px, halved.
  const lines = grid[0].lines;
  assert.ok(lines.some((line) => Math.abs(line[0] - (67.2 + (40 + 64 - 0.5) / 2)) < 1e-6 && line[0] === line[2]));
});

test('sheet print: header and footer are drawn at their margins with the values of the page', () => {
  const setup: SheetPageSetup = {
    headerFooter: { oddHeader: '&L&A&C&"Arial,Bold"&14Title\nSecond&RPage &P of &N', oddFooter: '&C&F &KFF0000red' },
    margins: { left: 1, right: 1, top: 1, bottom: 1, header: 0.5, footer: 0.5 },
  };
  const workbook = book([makeSheet('Data', 60, 2, setup)]);
  const plan = planSheetPages(workbook, [0]);
  const ctx = new FakeContext();
  paintSheetPage(asContext(ctx), { workbook, job: plan.pages[1], total: 2, metrics: createSheetMetrics(workbook.sheets[0]), pixelWidth: 793.7, now: new Date(2026, 0, 1), fileName: 'book.xlsx' });
  const find = (text: string) => ctx.texts.filter((call) => call.text === text)[0];
  // Left section at the left margin, in the default font of the workbook.
  const left = find('Data');
  assert.ok(Math.abs(left.x - 96) < 1e-6);
  assert.ok(left.y > 48 && left.y < 48 + 20);
  assert.ok(/14\.67px "Calibri"/.test(left.font), left.font);
  // The centre in its own font and size, around the middle of the room between the margins; the second line below the first.
  const title = find('Title');
  assert.ok(/^bold 18\.67px "Arial"/.test(title.font), title.font);
  const width = (5 * 18.67) / 2;
  assert.ok(Math.abs(title.x - (793.7 / 2 - width / 2)) < 1e-6);
  assert.ok(find('Second').y > title.y + 18);
  // The right section ends at the right margin.
  const right = find('Page 2 of 2');
  assert.ok(Math.abs(right.x + (11 * 14.67) / 2 - (793.7 - 96)) < 1e-6);
  // The footer ends at its margin above the bottom edge; a colour of the code is used for what follows it.
  const file = find('book.xlsx ');
  assert.ok(file.y < 1122.52 - 48 && file.y > 1122.52 - 48 - 20);
  assert.equal(file.fill, '#000000');
  assert.equal(find('red').fill, '#FF0000');
  // No code of the header reached the canvas as text.
  assert.equal(ctx.texts.some((call) => call.text.indexOf('&') >= 0), false);
  assert.equal(ctx.saves, 0);
});

test('sheet print: the header is scaled with the sheet unless the file says otherwise, and can leave the margins', () => {
  const paint = (headerFooter: SheetPageSetup['headerFooter']) => {
    const workbook = book([makeSheet('Data', 2, 2, { page: { scale: 50 }, headerFooter, margins: { left: 1.5, right: 1.5, top: 1, bottom: 1, header: 0.5, footer: 0.5 } })]);
    const ctx = new FakeContext();
    paintSheetPage(asContext(ctx), { workbook, job: planSheetPages(workbook, [0]).pages[0], total: 1, metrics: createSheetMetrics(workbook.sheets[0]), pixelWidth: 793.7, now: new Date(2026, 0, 1) });
    return ctx.texts.filter((call) => call.text === 'Head')[0];
  };
  const scaled = paint({ oddHeader: '&LHead' });
  assert.ok(/7\.33px/.test(scaled.font), scaled.font);
  assert.ok(Math.abs(scaled.x - 144) < 1e-6);
  const fixed = paint({ oddHeader: '&LHead', scaleWithDoc: false, alignWithMargins: false });
  assert.ok(/14\.67px/.test(fixed.font), fixed.font);
  assert.ok(Math.abs(fixed.x - 67.2) < 1e-6);
});

test('sheet print: print titles are drawn before the cells of the page, each band under its own clip', () => {
  const workbook = book([makeSheet('Data', 100, 2)], [{ name: '_xlnm.Print_Titles', value: 'Data!$1:$2', localSheet: 0, hidden: false }]);
  const plan = planSheetPages(workbook, [0]);
  const ctx = new FakeContext();
  paintSheetPage(asContext(ctx), { workbook, job: plan.pages[1], total: plan.pages.length, metrics: createSheetMetrics(workbook.sheets[0]), pixelWidth: 793.7, now: new Date(2026, 0, 1) });
  const texts = ctx.texts.map((call) => call.text);
  assert.deepEqual(texts.slice(0, 6), ['r0c0', 'r0c1', 'r1c0', 'r1c1', 'r48c0', 'r48c1']);
  // Row 49 of the sheet follows the two title rows: 40 px below the top margin.
  const first = ctx.texts.filter((call) => call.text === 'r48c0')[0];
  const title = ctx.texts.filter((call) => call.text === 'r0c0')[0];
  assert.ok(Math.abs(first.y - title.y - 40) < 1e-6);
  assert.equal(texts.indexOf('r47c0'), -1);
});

test('grid print: only what belongs on paper is drawn, and a scene without cells draws nothing', () => {
  const sheet = makeSheet('Data', 2, 2);
  sheet.view.showGridLines = true;
  const workbook = book([sheet]);
  const metrics = createSheetMetrics(sheet);
  const ctx = new FakeContext();
  renderGridPrint(asContext(ctx), { workbook, sheet, metrics, rows: [{ first: 0, last: 1 }], cols: [{ first: 0, last: 1 }], headings: false, gridLines: false, scale: 1 });
  // The grid lines of the screen are not those of the paper; nothing is filled (no background, no selection).
  assert.deepEqual(ctx.strokes, []);
  assert.deepEqual(ctx.rects, []);
  assert.equal(ctx.texts.length, 4);
  // The transform of the caller is kept: the block starts at the origin.
  assert.ok(Math.abs(ctx.texts[0].x - 3) < 1e-6);
  const none = new FakeContext();
  renderGridPrint(asContext(none), { workbook, sheet, metrics, rows: [], cols: [{ first: 0, last: 1 }], headings: true, gridLines: true, scale: 1 });
  renderGridPrint(asContext(none), { workbook, sheet, metrics, rows: [{ first: 5, last: 2 }], cols: [{ first: 0, last: 1 }], headings: true, gridLines: true, scale: NaN });
  assert.deepEqual([none.texts, none.rects, none.strokes], [[], [], []]);
});

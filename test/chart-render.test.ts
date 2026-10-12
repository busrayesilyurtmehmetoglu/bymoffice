// Painting a chart: what reaches the canvas for each chart type, that nothing
// of a file reaches a style of the context unchecked, what is kept from one
// frame to the next, and the chart as an object of the layer above the cells.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { open, type Chart, type DrawingChart, type Workbook } from '../src/index';
import { chartValues, layoutChart, type ChartLayout } from '../src/layout/chart';
import { createSheetMetrics } from '../src/layout/sheet-metrics';
import { setCellInput } from '../src/ops/commands';
import { History } from '../src/ops/ops';
import { CHART_KEEP_BUDGET, ChartRenderer, chartCssFont, paintChartLayout } from '../src/render/chart';
import { notifyFontsChanged } from '../src/render/fonts';
import { SheetDrawings } from '../src/render/sheet-drawings';
import { catAx, chartFrameOf, chartWorkbook, chartXml, fakeMeasure, group, legend, numLit, numRef, readChart, ser, sheetData, solid, spPr, strLit, strRef, title, valAx, xyAxes, xySer } from './chart-fixtures';
import { absolute, sp, solid as shapeFill } from './sheet-drawing-fixtures';

interface Box {
  x: number;
  y: number;
  width: number;
  height: number;
}

/** A stand-in for CanvasRenderingContext2D that records what is drawn, in the coordinates of the canvas. Each character is half the font size wide. */
class FakeContext {
  fillStyle = '#000000';
  strokeStyle = '#000000';
  lineWidth = 1;
  lineJoin = 'miter';
  lineCap = 'butt';
  globalAlpha = 1;
  textAlign = 'left';
  textBaseline = 'alphabetic';
  font = '10px sans-serif';
  rects: (Box & { fill: string })[] = [];
  fills: { fill: string; path: string[]; arcs: { x: number; y: number; radius: number; start: number; end: number }[] }[] = [];
  strokes: { style: string; width: number; dash: number[]; path: string[] }[] = [];
  texts: { text: string; x: number; y: number; font: string; fill: string; align: string; rotated: boolean }[] = [];
  clips: string[][] = [];
  fonts: string[] = [];
  styles: string[] = [];
  saves = 0;
  restores = 0;
  /** Makes fillRect() fail from the n-th call on, as a canvas might. */
  failAfter = Infinity;

  private dx = 0;
  private dy = 0;
  private rotated = false;
  private dash: number[] = [];
  private path: string[] = [];
  private arcs: { x: number; y: number; radius: number; start: number; end: number }[] = [];
  private stack: { dx: number; dy: number; rotated: boolean; fill: string; stroke: string; font: string }[] = [];

  save(): void {
    this.saves++;
    this.stack.push({ dx: this.dx, dy: this.dy, rotated: this.rotated, fill: this.fillStyle, stroke: this.strokeStyle, font: this.font });
  }
  restore(): void {
    this.restores++;
    const state = this.stack.pop();
    assert.ok(state, 'restore() without save()');
    this.dx = state.dx;
    this.dy = state.dy;
    this.rotated = state.rotated;
    this.fillStyle = state.fill;
    this.strokeStyle = state.stroke;
    this.font = state.font;
  }
  translate(x: number, y: number): void {
    this.dx += x;
    this.dy += y;
  }
  rotate(): void {
    this.rotated = true;
  }
  scale(): void {}
  setTransform(): void {}
  beginPath(): void {
    this.path = [];
    this.arcs = [];
  }
  closePath(): void {
    this.path.push('Z');
  }
  rect(x: number, y: number, width: number, height: number): void {
    this.path.push(`R${x + this.dx},${y + this.dy},${width},${height}`);
  }
  moveTo(x: number, y: number): void {
    this.path.push(`M${x + this.dx},${y + this.dy}`);
  }
  lineTo(x: number, y: number): void {
    this.path.push(`L${x + this.dx},${y + this.dy}`);
  }
  quadraticCurveTo(): void {
    this.path.push('Q');
  }
  bezierCurveTo(): void {
    this.path.push('C');
  }
  arc(x: number, y: number, radius: number, start: number, end: number): void {
    this.path.push('A');
    this.arcs.push({ x: x + this.dx, y: y + this.dy, radius, start, end });
  }
  clip(): void {
    this.clips.push(this.path.slice());
  }
  fill(): void {
    this.note();
    this.fills.push({ fill: this.fillStyle, path: this.path.slice(), arcs: this.arcs.slice() });
  }
  stroke(): void {
    this.note();
    this.strokes.push({ style: this.strokeStyle, width: this.lineWidth, dash: this.dash.slice(), path: this.path.slice() });
  }
  fillRect(x: number, y: number, width: number, height: number): void {
    if (this.rects.length >= this.failAfter) throw new Error('the canvas refuses');
    this.note();
    this.rects.push({ x: x + this.dx, y: y + this.dy, width, height, fill: this.fillStyle });
  }
  strokeRect(): void {
    this.note();
  }
  fillText(text: string, x: number, y: number): void {
    this.note();
    this.texts.push({ text, x: x + this.dx, y: y + this.dy, font: this.font, fill: this.fillStyle, align: this.textAlign, rotated: this.rotated });
  }
  strokeText(): void {
    assert.fail('text is only ever filled');
  }
  measureText(text: string): { width: number } {
    this.fonts.push(this.font);
    const size = Number((/(\d+(?:\.\d+)?)px/.exec(this.font) || [])[1]) || 10;
    return { width: text.length * size * 0.5 };
  }
  setLineDash(dash: number[]): void {
    this.dash = dash;
  }
  drawImage(): void {}

  /** Every style and font that was in force when something was drawn. */
  private note(): void {
    this.styles.push(String(this.fillStyle), String(this.strokeStyle));
    this.fonts.push(this.font);
  }
}

const context = () => new FakeContext();
const asCanvas = (ctx: FakeContext) => ctx as unknown as CanvasRenderingContext2D;

function model(xml: string): Chart {
  const chart = readChart(xml);
  assert.ok(chart, 'the chart is read');
  return chart as Chart;
}

function layoutOf(xml: string, width = 400, height = 300): ChartLayout {
  const chart = model(xml);
  return layoutChart(chart, chartValues(chart), width, height, fakeMeasure);
}

const CATS = strLit(['a', 'b', 'c']);
const columns = (values: number[], props = '', rest: { head?: string; tail?: string } = {}) =>
  chartXml({ plot: group('barChart', '<c:barDir val="col"/><c:grouping val="clustered"/>', ser(0, { tx: '<c:v>Bir</c:v>', cat: CATS, val: numLit(values), props })) + catAx() + valAx(), ...rest });

test('columns: one filled rectangle per value in the colour of the series, as high as the value, at the place of the chart', () => {
  const layout = layoutOf(columns([10, 20, 40], spPr(solid('00B050'))));
  const ctx = context();
  paintChartLayout(asCanvas(ctx), layout, 100, 50);
  assert.deepEqual(ctx.rects[0], { x: 100, y: 50, width: 400, height: 300, fill: '#FFFFFF' }, 'the frame is filled first, where the chart is');
  const bars = ctx.rects.filter((rect) => rect.fill === '#00B050');
  assert.equal(bars.length, 3);
  assert.ok(Math.abs(bars[1].height / bars[0].height - 2) < 1e-9 && Math.abs(bars[2].height / bars[0].height - 4) < 1e-9, 'heights in the proportion of the values');
  const floor = bars[0].y + bars[0].height;
  for (const bar of bars) {
    assert.ok(Math.abs(bar.y + bar.height - floor) < 1e-9, 'all stand on the same line');
    assert.ok(bar.x >= 100 && bar.x + bar.width <= 500 && bar.y >= 50 && floor <= 350, 'inside the frame of the chart');
  }
  assert.ok(bars[0].x < bars[1].x && bars[1].x < bars[2].x, 'from left to right');
  assert.equal(ctx.saves, ctx.restores, 'every save() has its restore()');
  assert.equal(ctx.clips.length, 1, 'clipped to the frame');
  assert.deepEqual(ctx.clips[0], ['R100,50,400,300']);
});

test('words reach the canvas through fillText only: the title, the labels of both axes, the legend', () => {
  const layout = layoutOf(columns([10, 20, 40], '', { head: title('Satış <b>'), tail: legend('r') }));
  const ctx = context();
  paintChartLayout(asCanvas(ctx), layout, 0, 0);
  const words = ctx.texts.map((item) => item.text);
  for (const wanted of ['Satış <b>', 'a', 'b', 'c', '0', 'Bir']) assert.ok(words.indexOf(wanted) >= 0, `"${wanted}" is drawn`);
  const heading = ctx.texts.filter((item) => item.text === 'Satış <b>')[0];
  assert.equal(heading.align, 'center');
  assert.match(heading.font, /^18\.67px /, '14 points');
  assert.equal(heading.fill, '#595959');
});

test('a pie: one slice per value, in the colours of the theme, the angles together a full turn', () => {
  const layout = layoutOf(chartXml({ plot: group('pieChart', '', ser(0, { cat: CATS, val: numLit([1, 2, 5]) }), '', false) }));
  const ctx = context();
  paintChartLayout(asCanvas(ctx), layout, 0, 0);
  const slices = ctx.fills.filter((fill) => fill.arcs.length === 1);
  assert.deepEqual(slices.map((slice) => slice.fill), ['#4472C4', '#ED7D31', '#A5A5A5']);
  let turn = 0;
  for (const slice of slices) turn += slice.arcs[0].end - slice.arcs[0].start;
  assert.ok(Math.abs(turn - Math.PI * 2) < 1e-9, `a full turn (${turn})`);
  assert.ok(Math.abs(slices[2].arcs[0].end - slices[2].arcs[0].start - (Math.PI * 2 * 5) / 8) < 1e-9, 'five eighths for the largest');
  assert.ok(slices.every((slice) => slice.path[0].charAt(0) === 'M' && slice.path[slice.path.length - 1] === 'Z'), 'each slice starts at the centre and is closed');
  // A doughnut: two arcs per slice, the inner one back.
  const ring = context();
  paintChartLayout(asCanvas(ring), layoutOf(chartXml({ plot: group('doughnutChart', '', ser(0, { val: numLit([1, 1]) }), '', false) })), 0, 0);
  assert.equal(ring.fills.filter((fill) => fill.arcs.length === 2).length, 2);
});

test('a line: one stroked path through the points, markers filled in the colour of the line', () => {
  const layout = layoutOf(chartXml({ plot: group('lineChart', '<c:grouping val="standard"/>', ser(0, { val: numLit([1, 3, 2]), props: spPr(`<a:ln w="28575">${solid('C00000')}<a:prstDash val="dash"/></a:ln>`) }), '<c:marker val="1"/>') + catAx() + valAx('') }));
  const ctx = context();
  paintChartLayout(asCanvas(ctx), layout, 0, 0);
  const lines = ctx.strokes.filter((stroke) => stroke.style === '#C00000' && stroke.path.length === 3);
  assert.equal(lines.length, 1);
  assert.deepEqual(lines[0].path.map((step) => step.charAt(0)), ['M', 'L', 'L']);
  assert.equal(lines[0].width, 3);
  assert.deepEqual(lines[0].dash, [12, 9], 'dashed as the file says');
  const markers = ctx.fills.filter((fill) => fill.fill === '#C00000' && fill.arcs.length === 3);
  assert.equal(markers.length, 1, 'three round markers in one path');
});

test('an area is one closed, filled path; a scatter chart draws its markers at their places', () => {
  const area = context();
  paintChartLayout(asCanvas(area), layoutOf(chartXml({ plot: group('areaChart', '', ser(0, { val: numLit([1, 3, 2]) })) + catAx() + valAx('') })), 0, 0);
  const shape = area.fills.filter((fill) => fill.fill === '#4472C4');
  assert.equal(shape.length, 1);
  assert.deepEqual(shape[0].path.map((step) => step.charAt(0)).join(''), 'MLLLLLZ');

  const cloud = context();
  paintChartLayout(asCanvas(cloud), layoutOf(chartXml({ plot: group('scatterChart', '<c:scatterStyle val="marker"/>', xySer(0, { cat: numLit([1, 2, 3]), val: numLit([3, 1, 2]), props: '<c:marker><c:symbol val="square"/></c:marker>' })) + xyAxes() })), 0, 0);
  const squares = cloud.fills.filter((fill) => fill.fill === '#4472C4' && fill.path.length === 3 && fill.path.every((step) => step.charAt(0) === 'R'));
  assert.equal(squares.length, 1, 'three squares');
});

test('every kind of marker can be drawn', () => {
  for (const symbol of ['circle', 'square', 'diamond', 'triangle', 'x', 'plus', 'star', 'dash', 'dot', 'auto', 'picture']) {
    const ctx = context();
    paintChartLayout(asCanvas(ctx), layoutOf(chartXml({ plot: group('lineChart', '', ser(0, { val: numLit([1, 2]), props: `<c:marker><c:symbol val="${symbol}"/></c:marker>` })) + catAx() + valAx('') })), 0, 0);
    const drawn = ctx.fills.filter((fill) => fill.fill === '#4472C4').length + ctx.strokes.filter((stroke) => stroke.style === '#4472C4' && stroke.path.length !== 2).length;
    assert.ok(drawn >= 1, `${symbol} leaves a mark`);
  }
});

test('nothing but colours of six hex digits reaches a style, and no font name reaches the font unquoted', () => {
  const chart = model(columns([1, 2, 3], '', { head: title('T'), tail: legend('r') }));
  const evil = 'red; background: url(x)';
  chart.groups[0].series[0].shape = { fill: evil, line: { color: 'rgba(0,0,0,1)', width: 2 } };
  chart.area = { fill: 'transparent', line: { color: '#12345' } };
  chart.plotArea = { fill: '#GGGGGG', line: { color: evil } };
  chart.palette = [evil, evil, evil, evil, evil, evil];
  chart.title = { text: 'T', style: { color: evil, font: 'Arial", serif; x: "', size: 20, bold: true } };
  chart.textStyle = { color: 'blue', font: "x'); }", size: 1e9 };
  chart.axes[1].majorGridlines = { color: evil, width: 1e9 };
  chart.axes[0].line = { color: evil };
  const ctx = context();
  paintChartLayout(asCanvas(ctx), layoutChart(chart, chartValues(chart), 400, 300, (text, font) => ctx.measureText.call({ font: chartCssFont(font), fonts: [] }, text).width), 0, 0);
  assert.ok(ctx.styles.length > 10, 'something was drawn');
  for (const style of ctx.styles) assert.match(style, /^#[0-9A-Fa-f]{6}$/, `a style of the canvas: ${style}`);
  for (const font of ctx.fonts) {
    assert.ok(font.indexOf(';') < 0 && font.indexOf('}') < 0 && font.indexOf(')') < 0, `a font of the canvas: ${font}`);
    assert.match(font, /^(italic )?(bold )?\d+(\.\d+)?px /, 'a size and families');
  }
  assert.ok(ctx.texts.some((item) => item.text === 'T'), 'the title is still drawn');
  assert.ok(ctx.texts.every((item) => Number(/(\d+(\.\d+)?)px/.exec(item.font)![1]) <= 128), 'a font size from the file is kept within bounds');
});

test('the title of the value axis is turned, around its own place, and the state is given back', () => {
  const layout = layoutOf(chartXml({ plot: group('barChart', '<c:barDir val="col"/>', ser(0, { val: numLit([1, 2, 3]) })) + catAx() + valAx(title('Tutar')) }));
  const ctx = context();
  paintChartLayout(asCanvas(ctx), layout, 10, 20);
  const turned = ctx.texts.filter((item) => item.rotated);
  assert.deepEqual(turned.map((item) => item.text), ['Tutar']);
  assert.ok(turned[0].x > 10 && turned[0].x < 60, 'at the left of the chart');
  assert.ok(ctx.texts.filter((item) => !item.rotated).length > 3, 'what is drawn after it is upright again');
  assert.equal(ctx.saves, ctx.restores);
});

test('a canvas that fails: the error is passed on and the state of the context is as it was', () => {
  const layout = layoutOf(columns([1, 2, 3]));
  const ctx = context();
  ctx.failAfter = 2;
  ctx.fillStyle = '#123456';
  assert.throws(() => paintChartLayout(asCanvas(ctx), layout, 0, 0), /refuses/);
  assert.equal(ctx.saves, ctx.restores, 'restored');
  assert.equal(ctx.fillStyle, '#123456');
});

test('a layout without a size paints nothing', () => {
  const ctx = context();
  paintChartLayout(asCanvas(ctx), layoutOf(columns([1, 2, 3]), 0, 0), 0, 0);
  assert.equal(ctx.saves, 0);
  assert.equal(ctx.rects.length, 0);
});

// ---------------------------------------------------------------- the renderer

test('the renderer keeps what it worked out while nothing changed, and works it out anew when something did', () => {
  // A series without a name asks for one every time the values are worked out.
  const chart = model(chartXml({ plot: group('barChart', '<c:barDir val="col"/>', ser(0, { val: numLit([1, 2, 3]) })) + catAx() + valAx() }));
  let version = 1;
  let asked = 0;
  const renderer = new ChartRenderer({ version: () => version, seriesName: (n) => (asked++, `S${n}`) });
  const ctx = context();
  const first = renderer.layout(asCanvas(ctx), chart, 400, 300);
  assert.equal(asked, 1);
  assert.equal(renderer.layout(asCanvas(ctx), chart, 400, 300), first, 'the same layout for the same frame');
  assert.equal(asked, 1);
  const resized = renderer.layout(asCanvas(ctx), chart, 500, 300);
  assert.notEqual(resized, first, 'another size: another layout');
  assert.equal(asked, 1, 'but the same values');
  version++;
  assert.notEqual(renderer.layout(asCanvas(ctx), chart, 500, 300), resized, 'a change: everything anew');
  assert.equal(asked, 2);

  const forgetful = new ChartRenderer({ seriesName: (n) => (asked++, `S${n}`) });
  forgetful.layout(asCanvas(ctx), chart, 400, 300);
  forgetful.layout(asCanvas(ctx), chart, 400, 300);
  assert.equal(asked, 4, 'without a version nothing is kept');
});

test('the renderer works a chart out anew when the fonts changed, and lays it out anew for another height', () => {
  const chart = model(chartXml({ plot: group('barChart', '<c:barDir val="col"/>', ser(0, { val: numLit([1, 2, 3]) })) + catAx() + valAx() }));
  const renderer = new ChartRenderer({ version: () => 1 });
  const ctx = context();
  const first = renderer.layout(asCanvas(ctx), chart, 400, 300);
  assert.deepEqual([renderer.state().values, renderer.state().layouts], [1, 1]);
  const higher = renderer.layout(asCanvas(ctx), chart, 400, 301);
  assert.notEqual(higher, first, 'a frame one pixel higher is another layout');
  assert.equal(higher.height, 301);
  assert.deepEqual([renderer.state().values, renderer.state().layouts], [1, 2], 'from the same values');
  assert.equal(renderer.layout(asCanvas(ctx), chart, 400, 300), first, 'the first size is still kept');
  // A font arrived: what was measured with the old ones is measured anew.
  const measured = ctx.fonts.length;
  notifyFontsChanged();
  const after = renderer.layout(asCanvas(ctx), chart, 400, 300);
  assert.notEqual(after, first, 'new fonts: a new layout');
  assert.deepEqual([renderer.state().values, renderer.state().layouts], [2, 3]);
  assert.ok(ctx.fonts.length > measured, 'the texts were measured again');
  assert.equal(renderer.layout(asCanvas(ctx), chart, 400, 300), after, 'and kept again');
});

test('the same chart in frames of two sizes: neither makes the other be laid out again', () => {
  const chart = model(chartXml({ plot: group('barChart', '<c:barDir val="col"/>', ser(0, { val: numLit([1, 2, 3]) })) + catAx() + valAx() }));
  const renderer = new ChartRenderer({ version: () => 1 });
  const ctx = context();
  for (let frame = 0; frame < 20; frame++) {
    renderer.paint(asCanvas(ctx), chart, 0, 0, 400, 300);
    renderer.paint(asCanvas(ctx), chart, 500, 0, 200, 150);
  }
  assert.deepEqual([renderer.state().values, renderer.state().layouts], [1, 2], 'twenty frames: one set of values, one layout per size');
  // Only a few sizes of one chart are kept: the one used longest ago goes.
  for (let width = 300; width < 310; width++) renderer.layout(asCanvas(ctx), chart, width, 200);
  assert.equal(renderer.layouts().length, 4, 'four sizes at most');
  assert.equal(renderer.layouts()[3].width, 309, 'the one used last is kept');
});

test('what a renderer keeps is limited for all charts together: the chart drawn longest ago is forgotten first', () => {
  const make = () => model(chartXml({ plot: group('lineChart', '<c:grouping val="standard"/>', ser(0, { val: numLit(Array.from({ length: 500 }, (_unused, i) => i % 7)) })) + catAx() + valAx() }));
  const charts = [make(), make(), make(), make()];
  const ctx = context();
  const roomy = new ChartRenderer({ version: () => 1 });
  for (const chart of charts) roomy.layout(asCanvas(ctx), chart, 400, 300);
  const each = roomy.state().weight / 4;
  assert.equal(roomy.state().charts, 4, 'the usual budget keeps them all');
  assert.ok(each > 500 && each < 5000, `a chart of 500 points weighs its numbers and what is drawn of it (${each})`);
  assert.ok(CHART_KEEP_BUDGET >= 1000000, 'the usual budget is that of a few very large charts');

  // Room for two and a half of them.
  const tight = new ChartRenderer({ version: () => 1 }, each * 2.5);
  for (const chart of charts) tight.layout(asCanvas(ctx), chart, 400, 300);
  assert.equal(tight.state().charts, 2, 'two are kept');
  assert.ok(tight.state().weight <= each * 2.5, 'within the budget');
  assert.equal(tight.state().values, 4);
  tight.layout(asCanvas(ctx), charts[3], 400, 300);
  tight.layout(asCanvas(ctx), charts[2], 400, 300);
  assert.equal(tight.state().values, 4, 'the two drawn last are still there');
  tight.layout(asCanvas(ctx), charts[0], 400, 300);
  assert.equal(tight.state().values, 5, 'the first was forgotten and is worked out again');
  tight.layout(asCanvas(ctx), charts[2], 400, 300);
  assert.equal(tight.state().values, 5, 'the one used just before is kept over the one used before that');
  tight.layout(asCanvas(ctx), charts[3], 400, 300);
  assert.equal(tight.state().values, 6);
  // A budget of nothing still draws: the chart in use is never the one that goes.
  const none = new ChartRenderer({ version: () => 1 }, 0);
  const layout = none.layout(asCanvas(ctx), charts[0], 400, 300);
  assert.equal(none.layout(asCanvas(ctx), charts[0], 400, 300), layout);
  none.layout(asCanvas(ctx), charts[1], 400, 300);
  assert.equal(none.state().charts, 1);
  // clear() forgets everything.
  roomy.clear();
  assert.deepEqual([roomy.state().charts, roomy.state().weight], [0, 0]);
  roomy.layout(asCanvas(ctx), charts[0], 400, 300);
  assert.equal(roomy.state().values, 5, 'and what is drawn next is worked out again');
});

test('dates count from 1904 for a renderer when its workbook says so, or else when the chart does', () => {
  const xml = (before: string) => chartXml({ before, plot: group('barChart', '<c:barDir val="col"/>', ser(0, { val: numLit([1, 2]) })) + catAx() + valAx('<c:numFmt formatCode="yyyy-mm-dd" sourceLinked="0"/>', '<c:orientation val="minMax"/><c:max val="2"/><c:min val="1"/>') });
  const ticks = (renderer: ChartRenderer, chart: Chart) => renderer.layout(asCanvas(context()), chart, 480, 320).items.filter((item) => item.kind === 'text' && item.role === 'tick').map((item) => (item as { text: string }).text);
  assert.equal(ticks(new ChartRenderer(), model(xml('')))[0], '1900-01-01');
  assert.equal(ticks(new ChartRenderer(), model(xml('<c:date1904 val="1"/>')))[0], '1904-01-02', 'no workbook: what the chart part says');
  const workbook = { date1904: true, sheets: [], styles: {} } as unknown as Workbook;
  assert.equal(ticks(new ChartRenderer({ workbook }), model(xml('')))[0], '1904-01-02', 'the workbook counts from 1904');
  const usual = { date1904: false, sheets: [], styles: {} } as unknown as Workbook;
  assert.equal(ticks(new ChartRenderer({ workbook: usual }), model(xml('<c:date1904 val="1"/>')))[0], '1900-01-01', 'the workbook decides over the chart part');
});

// ------------------------------------------------------- in the layer of a sheet

const CELLS = sheetData([['Ürün', 'Adet'], ['Çay', 10], ['Kahve', 20], ['Süt', 40]]);
const LINKED = chartXml({
  plot: group('barChart', '<c:barDir val="col"/><c:grouping val="clustered"/>', ser(0, { tx: strRef('Sayfa1!$B$1', ['x']), cat: strRef('Sayfa1!$A$2:$A$4', ['x', 'y', 'z']), val: numRef('Sayfa1!$B$2:$B$4', [1, 1, 1]), props: spPr(solid('00B050')) })) + catAx() + valAx(),
});

async function sheetWith(drawing: string, charts: string[]): Promise<Workbook> {
  return (await open(await chartWorkbook({ sheet: CELLS, drawing, charts }))).workbook();
}

function paintLayer(workbook: Workbook, layer: SheetDrawings): FakeContext {
  const ctx = context();
  const sheet = workbook.sheets[0];
  layer.paint(asCanvas(ctx), { metrics: createSheetMetrics(sheet), clip: { x: 0, y: 0, width: 2000, height: 2000 }, originX: 30, originY: 20, devicePixelRatio: 2 });
  assert.equal(ctx.saves, ctx.restores, 'every save() has its restore()');
  return ctx;
}

test('a chart of a sheet is drawn in its anchor, from the cells as they are now; an edit shows once the version moved on', async () => {
  const workbook = await sheetWith(absolute(100, 60, 400, 300, chartFrameOf('rIdC1')), [LINKED]);
  let version = 0;
  const layer = new SheetDrawings(workbook.sheets[0], { charts: { workbook, version: () => version } });
  const before = paintLayer(workbook, layer);
  assert.deepEqual(before.rects[0], { x: 130, y: 80, width: 400, height: 300, fill: '#FFFFFF' }, 'the frame of the chart at its anchor');
  const bars = before.rects.filter((rect) => rect.fill === '#00B050');
  assert.equal(bars.length, 3);
  assert.ok(Math.abs(bars[2].height / bars[0].height - 4) < 1e-9, 'the cells (10, 20, 40), not the cache (1, 1, 1)');
  assert.ok(before.texts.some((item) => item.text === 'Kahve'), 'the categories of the cells');
  assert.equal(before.texts.filter((item) => item.text === 'Chart').length, 0, 'no placeholder');

  const history = new History(workbook);
  history.apply(setCellInput(workbook, 0, 1, 1, '80'), undefined, undefined);
  version++;
  const after = paintLayer(workbook, layer).rects.filter((rect) => rect.fill === '#00B050');
  assert.ok(Math.abs(after[0].height / after[2].height - 2) < 1e-9, 'the first column is now twice the third');
  history.undo();
  version++;
  const undone = paintLayer(workbook, layer).rects.filter((rect) => rect.fill === '#00B050');
  assert.deepEqual(undone, bars, 'undone: as before');
});

test('scrolling by parts of a pixel does not lay a chart out again, whatever the size of its frame', async () => {
  // A frame 400.4 by 300.6 pixels (in EMU: whole numbers), and one of whole pixels.
  const workbook = await sheetWith(absolute(100, 60, 400.4, 300.6, chartFrameOf('rIdC1')) + absolute(600, 60, 320, 200, chartFrameOf('rIdC1')), [LINKED]);
  const sheet = workbook.sheets[0];
  const renderer = new ChartRenderer({ workbook, version: () => 1 });
  const layer = new SheetDrawings(sheet, { charts: renderer });
  const frames: number[] = [];
  for (let step = 0; step <= 40; step++) {
    const ctx = context();
    // As on a screen of 2.625 device pixels per pixel that is scrolled a device pixel at a time.
    layer.paint(asCanvas(ctx), { metrics: createSheetMetrics(sheet), clip: { x: 0, y: 0, width: 2000, height: 2000 }, originX: 30 - step / 2.625, originY: 20 - step * 0.3, devicePixelRatio: 2.625 });
    const frame = ctx.rects[0];
    assert.deepEqual([frame.width, frame.height], [400, 301], `step ${step}: the frame is as large as ever`);
    assert.ok(Number.isInteger(frame.x) && Number.isInteger(frame.y), `step ${step}: on whole pixels`);
    frames.push(frame.x);
  }
  assert.ok(new Set(frames).size > 5, 'the chart did move');
  assert.deepEqual([renderer.state().values, renderer.state().layouts], [1, 2], 'one set of values and one layout per size for forty-one frames');
});

test('without a workbook the layer draws a chart from the values of the file', async () => {
  const workbook = await sheetWith(absolute(0, 0, 400, 300, chartFrameOf('rIdC1')), [LINKED]);
  const bars = paintLayer(workbook, new SheetDrawings(workbook.sheets[0])).rects.filter((rect) => rect.fill === '#00B050');
  assert.equal(bars.length, 3);
  assert.ok(Math.abs(bars[0].height - bars[2].height) < 1e-9, 'the cache: 1, 1, 1');
});

test('a chart that cannot be drawn is shown as the placeholder of a chart, and what lies above it is still drawn', async () => {
  const workbook = await sheetWith(absolute(0, 0, 300, 200, chartFrameOf('rIdC1')) + absolute(400, 0, 100, 50, sp({ props: shapeFill('FF0000') })), [LINKED]);
  const object = workbook.sheets[0].drawings![0].object as DrawingChart;
  // A model that no reader makes: the layout runs into it.
  (object.chart as unknown as { groups: unknown }).groups = null;
  const ctx = paintLayer(workbook, new SheetDrawings(workbook.sheets[0], { charts: { workbook }, labels: { chart: 'Grafik' } }));
  assert.ok(ctx.texts.some((item) => item.text === 'Grafik'), 'the word of the placeholder');
  assert.ok(ctx.rects.some((rect) => rect.fill === '#F1F3F4' && rect.width === 300), 'its grey box');
  assert.ok(ctx.fills.some((fill) => fill.fill === '#FF0000'), 'the shape after it');
});

test('a chart whose anchor is squeezed to nothing is not drawn; a tiny one shows its series alone', async () => {
  const workbook = await sheetWith(absolute(0, 0, 0, 100, chartFrameOf('rIdC1')) + absolute(100, 100, 40, 30, chartFrameOf('rIdC1')), [LINKED]);
  const ctx = paintLayer(workbook, new SheetDrawings(workbook.sheets[0], { charts: { workbook } }));
  assert.equal(ctx.rects.filter((rect) => rect.fill === '#FFFFFF').length, 1, 'one frame');
  assert.equal(ctx.texts.length, 0, 'no words in 40 by 30 pixels');
  const bars = ctx.rects.filter((rect) => rect.fill === '#00B050');
  assert.equal(bars.length, 3);
  assert.ok(bars.every((bar) => bar.x >= 130 && bar.x + bar.width <= 170 && bar.y >= 120 && bar.y + bar.height <= 150), 'inside the 40 by 30 pixels');
});

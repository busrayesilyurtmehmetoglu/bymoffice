// The dialog of a chart without its form: the fields a chart is shown with, and the patch
// the fields ask for (nothing for what was only looked at). And the selection of a drawing:
// which drawing is under a point, where its handles are and what dragging one does to its box.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { getLocale } from '../src/format/locale';
import { createSheetMetrics } from '../src/layout/sheet-metrics';
import { addChart, sheetCharts } from '../src/ops/charts';
import { applyOp } from '../src/ops/ops';
import { SheetDrawings, drawingHandleAt, resizedRect } from '../src/render/sheet-drawings';
import { chartFields, chartPatch, type ChartFields } from '../src/ui/chart-dialog';
import { FieldError } from '../src/ui/rules-dialog';
import { STRINGS_EN, STRINGS_TR } from '../src/ui/strings';
import { bookOf } from './chart-edit-fixtures';

const EN = getLocale('en');
const TR = getLocale('tr');
const WORDS = STRINGS_EN.charts;
const SALES = [
  ['Month', 'North', 'South'],
  ['Jan', 10, 30],
  ['Feb', 20, 20],
];

function fieldsOf(spec: Parameters<typeof addChart>[2]): ChartFields {
  const wb = bookOf(SALES);
  applyOp(wb, addChart(wb, 0, spec)[0]);
  return chartFields(sheetCharts(wb, 0)[0], EN, 'Data');
}

test('chart dialog: the fields show the chart, with the references of its own sheet written short', () => {
  const fields = fieldsOf({ type: 'column', range: { top: 0, left: 0, bottom: 2, right: 2 }, title: 'Sales', legend: 'b' });
  assert.deepEqual(fields, {
    type: 'column',
    grouping: 'standard',
    title: 'Sales',
    range: 'A1:C3',
    seriesIn: 'columns',
    legend: 'b',
    dataLabels: false,
    categoryTitle: '',
    valueTitle: '',
    minimum: '',
    maximum: '',
    categories: '$A$2:$A$3',
    series: [
      { from: 0, name: '=$B$1', values: '$B$2:$B$3', color: '' },
      { from: 1, name: '=$C$1', values: '$C$2:$C$3', color: '' },
    ],
  });
});

test('chart dialog: fields that were only looked at ask for nothing; each field that changed asks for its own change', () => {
  const initial = fieldsOf({ type: 'column', range: { top: 0, left: 0, bottom: 2, right: 2 }, title: 'Sales' });
  const copy = (): ChartFields => JSON.parse(JSON.stringify(initial));
  assert.deepEqual(chartPatch(initial, copy(), EN, WORDS), {});
  let now = copy();
  now.type = 'pie';
  now.title = '';
  now.legend = 'none';
  now.dataLabels = true;
  assert.deepEqual(chartPatch(initial, now, EN, WORDS), { type: 'pie', title: null, legend: 'none', dataLabels: true });
  now = copy();
  now.categoryTitle = 'Month';
  now.minimum = '0';
  now.maximum = ' 1.5 ';
  assert.deepEqual(chartPatch(initial, now, EN, WORDS), { axisTitles: { category: 'Month' }, axisBounds: { min: 0, max: 1.5 } });
  // In Turkish the comma is the decimal sign.
  now = copy();
  now.maximum = '1.234,5';
  assert.deepEqual(chartPatch(initial, now, TR, STRINGS_TR.charts), { axisBounds: { max: 1234.5 } });
  now.maximum = 'çok';
  assert.throws(
    () => chartPatch(initial, now, TR, STRINGS_TR.charts),
    (error: unknown) => error instanceof FieldError && error.field === 'maximum' && error.message === STRINGS_TR.charts.badNumber,
  );
});

test('chart dialog: another range makes the series anew; a changed list of series names each by where it came from', () => {
  const initial = fieldsOf({ type: 'column', range: { top: 0, left: 0, bottom: 2, right: 2 } });
  const copy = (): ChartFields => JSON.parse(JSON.stringify(initial));
  let now = copy();
  now.range = 'A1:B3';
  now.seriesIn = 'rows';
  assert.deepEqual(chartPatch(initial, now, EN, WORDS), { range: { top: 0, left: 0, bottom: 2, right: 1 }, seriesIn: 'rows' });
  now.range = 'A1:B3, D1:D3';
  assert.throws(() => chartPatch(initial, now, EN, WORDS), (error: unknown) => error instanceof FieldError && error.field === 'range');
  // The second series first, recoloured; the first renamed; a third one added; shared categories from another column.
  now = copy();
  now.series = [
    { from: 1, name: '=$C$1', values: '$C$2:$C$3', color: '#aa0000' },
    { from: 0, name: 'Kuzey', values: '$B$2:$B$3', color: '' },
    { name: '', values: '=Other!$A$1:$A$2', color: '' },
  ];
  now.categories = '$C$2:$C$3';
  assert.deepEqual(chartPatch(initial, now, EN, WORDS), {
    series: [
      { from: 1, color: '#aa0000', categories: '$C$2:$C$3' },
      { from: 0, name: 'Kuzey', categories: '$C$2:$C$3' },
      { name: null, nameRef: null, values: 'Other!$A$1:$A$2', color: null, categories: '$C$2:$C$3' },
    ],
  });
  now.series[2].values = '  ';
  assert.throws(() => chartPatch(initial, now, EN, WORDS), (error: unknown) => error instanceof FieldError && error.field === 'values2');
  now.series = [];
  assert.throws(() => chartPatch(initial, now, EN, WORDS), (error: unknown) => error instanceof FieldError && error.field === 'series');
});

test('chart selection: the drawing under a point is the one on top; its handles are its corners and the middles of its sides', () => {
  const wb = bookOf(SALES);
  applyOp(wb, addChart(wb, 0, { type: 'line', range: { top: 0, left: 0, bottom: 2, right: 2 }, anchor: { top: 4, left: 1, bottom: 9, right: 4 } })[0]);
  applyOp(wb, addChart(wb, 0, { type: 'pie', range: { top: 0, left: 0, bottom: 2, right: 1 }, anchor: { top: 6, left: 2, bottom: 12, right: 6 } })[0]);
  const sheet = wb.sheets[0];
  const metrics = createSheetMetrics(sheet);
  const layer = new SheetDrawings(sheet);
  const first = layer.boxOf(metrics, 0);
  assert.deepEqual(first, { x: 64, y: 80, width: 4 * 64, height: 6 * 20 });
  assert.equal(layer.hit(metrics, 70, 90), 0, 'only the first chart is there');
  assert.equal(layer.hit(metrics, 200, 150), 1, 'where both are, the one on top');
  assert.equal(layer.hit(metrics, 10, 10), -1);
  assert.equal(layer.hit(metrics, 200, 150, () => false), -1, 'a drawing that is not wanted is not found');
  assert.equal(layer.boxOf(metrics, 5), undefined);
  const box = { x: 100, y: 50, width: 200, height: 100 };
  assert.equal(drawingHandleAt(box, 101, 49), 'nw');
  assert.equal(drawingHandleAt(box, 200, 150), 's');
  assert.equal(drawingHandleAt(box, 300, 100), 'e');
  assert.equal(drawingHandleAt(box, 200, 100), undefined, 'the middle of the box is no handle');
  assert.deepEqual(resizedRect(box, 'move', 30, -20, 24), { x: 130, y: 30, width: 200, height: 100 });
  assert.deepEqual(resizedRect(box, 'move', -500, -500, 24), { x: 0, y: 0, width: 200, height: 100 }, 'a drawing does not leave the sheet');
  assert.deepEqual(resizedRect(box, 'se', 50, 10, 24), { x: 100, y: 50, width: 250, height: 110 });
  assert.deepEqual(resizedRect(box, 'nw', 50, 10, 24), { x: 150, y: 60, width: 150, height: 90 });
  assert.deepEqual(resizedRect(box, 'w', 500, 0, 24), { x: 276, y: 50, width: 24, height: 100 }, 'at least the smallest size stays');
  assert.deepEqual(resizedRect(box, 'n', 30, 500, 24), { x: 100, y: 126, width: 200, height: 24 });
});

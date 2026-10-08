import assert from 'node:assert/strict';
import { test } from 'node:test';
import { cellRect, colX, createLayout, gridPanes, hitTest, maxScroll, rangeRect, rowY, scrollToReveal } from '../src/layout/grid-geometry';
import { AxisMetrics, columnWidthToPixels, createSheetMetrics, rowHeightToPixels } from '../src/layout/sheet-metrics';
import { createSheet, type Row } from '../src/model/sheet/types';

const row = (init: Partial<Row>): Row => ({ customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells: [], ...init });
const sheetWith = () => createSheet({ name: 'S', sheetId: 1, state: 'visible', kind: 'worksheet', part: '/x' });

test('metrics: Excel unit conversions', () => {
  // Stored widths for Calibri 11: 9.140625 is the default column (64 px), 10.7109375 is "10" in the Excel dialog (75 px).
  assert.deepEqual([9.140625, 10.7109375, 10, 1, 20, 0, -1].map((w) => columnWidthToPixels(w)), [64, 75, 70, 7, 140, 0, 0]);
  assert.equal(columnWidthToPixels(1000), columnWidthToPixels(255));
  assert.deepEqual([15, 12.75, 30, 0, 409.5, 5000].map(rowHeightToPixels), [20, 17, 40, 0, 546, 546]);
});

test('metrics: axis offsets, sizes and lookups with custom and hidden items', () => {
  const axis = new AxisMetrics(1000, 20, [
    [5, 50],
    [2, 0],
    [3, 0],
    [999, 10],
    [5, 60],
    [7, 20],
    [5000, 99],
  ]);
  assert.equal(axis.total, 1000 * 20 - 40 + 40 - 10);
  assert.deepEqual([0, 1, 2, 3, 4, 5, 6, 999, 1000, -1].map((i) => axis.size(i)), [20, 20, 0, 0, 20, 60, 20, 10, 0, 0]);
  assert.deepEqual([0, 1, 2, 3, 4, 5, 6, 7].map((i) => axis.offset(i)), [0, 20, 40, 40, 40, 60, 120, 140]);
  assert.equal(axis.offset(1000), axis.total);
  assert.equal(axis.offset(5000), axis.total);
  // Hidden items never win a lookup: position 40 belongs to item 4.
  assert.deepEqual([0, 19, 20, 39, 40, 59, 60, 119, 120].map((px) => axis.indexAt(px)), [0, 0, 1, 1, 4, 4, 5, 5, 6]);
  assert.equal(axis.indexAt(-5), 0);
  assert.equal(axis.indexAt(axis.total), 999);
  assert.equal(axis.indexAt(axis.total + 500), 999);
  assert.deepEqual([axis.next(2, 1), axis.next(3, -1), axis.next(4, 1)], [4, 1, 4]);

  // Brute-force agreement on a random axis.
  let seed = 7;
  const random = () => (seed = (seed * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff;
  const custom: [number, number][] = Array.from({ length: 60 }, () => [Math.floor(random() * 300), Math.floor(random() * 4) * 15]);
  const fuzz = new AxisMetrics(300, 20, custom);
  const sizes = Array.from({ length: 300 }, () => 20);
  for (const [index, size] of custom) sizes[index] = size;
  let offset = 0;
  for (let i = 0; i < 300; i++) {
    assert.equal(fuzz.size(i), sizes[i]);
    assert.equal(fuzz.offset(i), offset);
    for (let px = offset; px < offset + sizes[i]; px += 7) assert.equal(fuzz.indexAt(px), i, `px ${px}`);
    offset += sizes[i];
  }
  assert.equal(fuzz.total, offset);
});

test('metrics: an axis whose first or last items are hidden still answers lookups', () => {
  const head = new AxisMetrics(10, 20, [[0, 0], [1, 0]]);
  assert.equal(head.indexAt(0), 2);
  assert.equal(head.indexAt(-1), 2);
  const tail = new AxisMetrics(10, 20, [[8, 0], [9, 0]]);
  assert.equal(tail.indexAt(tail.total), 7);
  const all = new AxisMetrics(3, 20, [[0, 0], [1, 0], [2, 0]]);
  assert.equal(all.total, 0);
  assert.ok(all.indexAt(0) >= 0 && all.indexAt(0) < 3);
});

test('metrics: long runs of hidden items are skipped in one step', () => {
  // A filtered sheet: a million rows of which all but a few are hidden.
  const hidden: [number, number][] = [];
  for (let i = 0; i < 1000000; i++) if (i % 250000 !== 5) hidden.push([i, 0]);
  const axis = new AxisMetrics(1048576, 20, hidden);
  const started = Date.now();
  for (let n = 0; n < 2000; n++) {
    assert.equal(axis.indexAt(0), 5);
    assert.equal(axis.next(6, 1), 250005);
    assert.equal(axis.next(250004, -1), 5);
    assert.equal(axis.skipHidden(0, -1), -1);
    assert.equal(axis.skipHidden(750006, 1), 1000000);
    assert.equal(axis.indexAt(45), 500005);
  }
  assert.ok(Date.now() - started < 10000, `took ${Date.now() - started} ms`);
  assert.equal(axis.skipHidden(-1, 1), -1);
  assert.equal(axis.skipHidden(1048576, -1), 1048576);
  assert.equal(axis.skipHidden(5, 1), 5);

  const tail = new AxisMetrics(10, 20, [[7, 0], [8, 0], [9, 0]]);
  assert.deepEqual([tail.skipHidden(7, 1), tail.next(7, 1), tail.next(9, -1)], [10, 6, 6]);
});

test('metrics: sheet rows and columns become pixel sizes', () => {
  const sheet = sheetWith();
  sheet.rows.set(0, row({ height: 30 }));
  sheet.rows.set(2, row({ hidden: true, height: 30 }));
  sheet.columns.push(
    { min: 0, max: 0, width: 20, customWidth: true, hidden: false, outlineLevel: 0, collapsed: false },
    { min: 2, max: 3, width: 9, customWidth: false, hidden: true, outlineLevel: 0, collapsed: false },
    { min: 5, max: 16383, customWidth: false, hidden: false, outlineLevel: 0, collapsed: false, style: 3 },
  );
  const metrics = createSheetMetrics(sheet);
  assert.deepEqual([0, 1, 2, 3].map((i) => metrics.rows.size(i)), [40, 20, 0, 20]);
  assert.deepEqual([0, 1, 2, 3, 4, 5].map((i) => metrics.cols.size(i)), [140, 64, 0, 0, 64, 64]);
  assert.deepEqual([metrics.rows.count, metrics.cols.count], [1048576, 16384]);
  assert.equal(metrics.rows.offset(1048576), metrics.rows.total);

  sheet.defaultRowHeight = 30;
  sheet.defaultColWidth = 20;
  const other = createSheetMetrics(sheet);
  assert.deepEqual([other.rows.size(1), other.cols.size(1), other.rows.size(0)], [40, 140, 40]);
});

test('geometry: cells, panes and hit testing without frozen panes', () => {
  const sheet = sheetWith();
  const metrics = createSheetMetrics(sheet);
  const { headerWidth: hw, headerHeight: hh } = metrics;
  const layout = createLayout(metrics, sheet.view, { width: hw + 64 * 5, height: hh + 20 * 10, scrollX: 64 * 2 + 10, scrollY: 25 });

  assert.deepEqual(cellRect(layout, 0, 0), { x: hw - 138, y: hh - 25, width: 64, height: 20 });
  assert.deepEqual(cellRect(layout, 2, 3), { x: hw + 64 - 10, y: hh + 15, width: 64, height: 20 });
  assert.deepEqual(rangeRect(layout, { top: 2, left: 3, bottom: 3, right: 4 }), { x: hw + 54, y: hh + 15, width: 128, height: 40 });

  const panes = gridPanes(layout);
  assert.equal(panes.length, 1);
  assert.deepEqual(
    [panes[0].firstRow, panes[0].lastRow, panes[0].firstCol, panes[0].lastCol],
    [1, 11, 2, 7],
  );
  assert.deepEqual(panes[0].rect, { x: hw, y: hh, width: 320, height: 200 });

  assert.deepEqual(hitTest(layout, hw + 1, hh + 1), { area: 'cell', row: 1, col: 2 });
  assert.deepEqual(hitTest(layout, hw + 54, hh + 15), { area: 'cell', row: 2, col: 3 });
  assert.deepEqual(hitTest(layout, 5, hh + 15), { area: 'rowHeader', row: 2 });
  assert.deepEqual(hitTest(layout, hw + 54, 3), { area: 'colHeader', col: 3 });
  assert.deepEqual(hitTest(layout, 3, 3), { area: 'corner' });
});

test('geometry: frozen rows and columns stay put and split the grid into four panes', () => {
  const sheet = sheetWith();
  sheet.view.frozenRows = 2;
  sheet.view.frozenCols = 1;
  const metrics = createSheetMetrics(sheet);
  const { headerWidth: hw, headerHeight: hh } = metrics;
  const layout = createLayout(metrics, sheet.view, { width: hw + 64 * 5, height: hh + 20 * 10, scrollX: 64 * 3, scrollY: 20 * 100 });

  assert.deepEqual([layout.frozenWidth, layout.frozenHeight], [64, 40]);
  assert.deepEqual([colX(layout, 0), rowY(layout, 0), rowY(layout, 1)], [hw, hh, hh + 20]);
  // Column 4 is the first scrolled column in view: it sits right after the frozen one.
  assert.equal(colX(layout, 4), hw + 64);
  assert.equal(rowY(layout, 102), hh + 40);

  const panes = gridPanes(layout);
  assert.deepEqual(
    panes.map((p) => [p.frozenRows, p.frozenCols, p.firstRow, p.lastRow, p.firstCol, p.lastCol]),
    [
      [true, true, 0, 1, 0, 0],
      [true, false, 0, 1, 4, 7],
      [false, true, 102, 109, 0, 0],
      [false, false, 102, 109, 4, 7],
    ],
  );
  assert.deepEqual(panes[3].rect, { x: hw + 64, y: hh + 40, width: 256, height: 160 });

  assert.deepEqual(hitTest(layout, hw + 10, hh + 10), { area: 'cell', row: 0, col: 0 });
  assert.deepEqual(hitTest(layout, hw + 70, hh + 30), { area: 'cell', row: 1, col: 4 });
  assert.deepEqual(hitTest(layout, hw + 10, hh + 45), { area: 'cell', row: 102, col: 0 });
  assert.deepEqual(hitTest(layout, hw + 70, hh + 45), { area: 'cell', row: 102, col: 4 });
  assert.deepEqual(hitTest(layout, 2, hh + 45), { area: 'rowHeader', row: 102 });
});

test('geometry: scroll limits and bringing a cell into view', () => {
  const sheet = sheetWith();
  sheet.view.frozenRows = 1;
  const metrics = createSheetMetrics(sheet);
  const { headerWidth: hw, headerHeight: hh } = metrics;
  const viewport = { width: hw + 320, height: hh + 200, scrollX: 0, scrollY: 0 };
  const layout = createLayout(metrics, sheet.view, viewport);

  assert.deepEqual(maxScroll(layout), { x: 16384 * 64 - 320, y: 1048576 * 20 - 200 });
  const clamped = createLayout(metrics, sheet.view, { ...viewport, scrollX: 1e12, scrollY: -50 });
  assert.deepEqual([clamped.scrollX, clamped.scrollY], [16384 * 64 - 320, 0]);
  // At the far end the last column ends exactly at the canvas edge.
  assert.equal(colX(clamped, 16383) + 64, hw + 320);

  assert.deepEqual(scrollToReveal(layout, 3, 2), { scrollX: 0, scrollY: 0 });
  // Row 20 is below the view: it becomes the last fully visible row.
  assert.deepEqual(scrollToReveal(layout, 20, 9), { scrollX: 64 * 10 - 320, scrollY: 20 * 21 - 20 - 180 });
  const scrolled = createLayout(metrics, sheet.view, { ...viewport, scrollX: 500, scrollY: 500 });
  assert.deepEqual(scrollToReveal(scrolled, 5, 1), { scrollX: 64, scrollY: 20 * 5 - 20 });
  // A frozen row never needs vertical scrolling.
  assert.equal(scrollToReveal(scrolled, 0, 20).scrollY, 500);
  assert.deepEqual(scrollToReveal(layout, 1048575, 16383), { scrollX: maxScroll(layout).x, scrollY: maxScroll(layout).y });
});

test('geometry: a canvas smaller than the headers or frozen area has no negative sizes', () => {
  const sheet = sheetWith();
  sheet.view.frozenRows = 50;
  sheet.view.frozenCols = 20;
  const metrics = createSheetMetrics(sheet);
  for (const [width, height] of [[0, 0], [10, 10], [metrics.headerWidth + 100, metrics.headerHeight + 100]]) {
    const layout = createLayout(metrics, sheet.view, { width, height, scrollX: 0, scrollY: 0 });
    for (const pane of gridPanes(layout)) {
      assert.ok(pane.rect.width > 0 && pane.rect.height > 0);
      assert.ok(pane.lastRow >= pane.firstRow && pane.lastCol >= pane.firstCol);
      assert.ok(pane.rect.x + pane.rect.width <= Math.max(width, 0) && pane.rect.y + pane.rect.height <= Math.max(height, 0));
    }
  }
});

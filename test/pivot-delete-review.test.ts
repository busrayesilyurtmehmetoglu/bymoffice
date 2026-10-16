// Deleting pivot tables after its review: what names the relationships that go (the selection of a
// sheet view, the caches of the workbook's extension list), the conditional formats of a table,
// merged cells across its edge, and sorting once no table is left.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { ZipArchive } from '../src/index';
import { deletePivot, deletePivotProblem } from '../src/ops/pivot-delete';
import { sortRange } from '../src/ops/sort';
import { hasPython, independentPivots, load, pivotBook, same, save, text, type Datum } from './pivot-fixtures';
import type { Op } from '../src/ops/ops';
import type { Loaded } from './pivot-fixtures';

const encoder = new TextEncoder();
const decoder = new TextDecoder();
const DATA: Datum[][] = [
  ['Region', 'Product', 'Sales'],
  ['East', 'A', 10],
  ['West', 'B', 20],
  ['East', 'B', 5],
  ['West', 'A', 7],
];

const apply = (doc: Loaded, ops: Op[]) => doc.history.apply(ops, undefined, undefined);

async function edit(bytes: Uint8Array, part: string, change: (xml: string) => string): Promise<Uint8Array> {
  const zip = ZipArchive.read(bytes);
  zip.set(part, encoder.encode(change(decoder.decode(await zip.read(part)))));
  return zip.write();
}

const base = (spec: Record<string, unknown> = {}) => pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }], ...spec });

const SELECTION =
  '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pivotSelection pane="topLeft" showHeader="1" activeRow="3" previousRow="3" click="1" r:id="rId1"><pivotArea type="data" outline="0" fieldPosition="0"><references count="1"><reference field="0" count="1"><x v="0"/></reference></references></pivotArea></pivotSelection><selection activeCell="A4" sqref="A4"/></sheetView></sheetViews>';

test('the selection of a sheet view that names the table goes with it; undone, the bytes of the file', async () => {
  const bytes = await edit(await base(), 'xl/worksheets/sheet2.xml', (xml) => xml.replace('<sheetData', `${SELECTION}<sheetData`));
  const doc = await load(bytes);
  assert.equal(deletePivotProblem(doc.wb, 0), undefined);
  apply(doc, deletePivot(doc.wb, 0));
  const deleted = await save(doc);
  const sheet = await text(deleted, 'xl/worksheets/sheet2.xml');
  assert.doesNotMatch(sheet, /pivotSelection/);
  assert.match(sheet, /<sheetView tabSelected="1" workbookViewId="0"><selection activeCell="A4" sqref="A4"\/><\/sheetView>/);
  if (hasPython()) assert.deepEqual(independentPivots({ deleted }), []);
  doc.history.undo();
  assert.ok(same(await save(doc), bytes));
});

test('another element of the sheet that names the relationship of the table refuses the deletion', async () => {
  const bytes = await edit(await base(), 'xl/worksheets/sheet2.xml', (xml) =>
    xml.replace('</worksheet>', '<extLst><ext uri="{00000000-0000-0000-0000-000000000000}"><thing xmlns="urn:x" r:id="rId1"/></ext></extLst></worksheet>'),
  );
  const doc = await load(bytes);
  assert.match(deletePivotProblem(doc.wb, 0) || '', /Cannot delete the pivot table "PivotTable1": an element of its sheet names it \(thing\)/);
});

const X14 =
  '<extLst><ext uri="{876F7934-8845-4945-9796-88D515C7AA90}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:pivotCaches><pivotCache cacheId="7" r:id="rId3"/></x14:pivotCaches></ext></extLst>';

test('a cache the extension list of the workbook names as well: the table that would take it along is not deleted; one that leaves it is', async () => {
  const one = await load(await edit(await base(), 'xl/workbook.xml', (xml) => xml.replace('</workbook>', `${X14}</workbook>`)));
  assert.match(deletePivotProblem(one.wb, 0) || '', /an element of the workbook names its cache \(x14:pivotCaches\/pivotCache\)/);
  const two = await load(await edit(await base({ twin: { row: 2, col: 5 } }), 'xl/workbook.xml', (xml) => xml.replace('</workbook>', `${X14}</workbook>`)));
  assert.equal(deletePivotProblem(two.wb, 0), undefined, 'the cache stays');
  apply(two, deletePivot(two.wb, 0));
  assert.match(deletePivotProblem(two.wb, 0) || '', /x14:pivotCaches/);
  const saved = await save(two);
  assert.match(await text(saved, 'xl/workbook.xml'), /<x14:pivotCaches><pivotCache cacheId="7" r:id="rId3"\/><\/x14:pivotCaches>/);
});

test('the conditional formats of the table go with it, in the same step; others stay', async () => {
  const tableFormats = '<conditionalFormats count="1"><conditionalFormat priority="2"><pivotAreas count="1"><pivotArea type="data" outline="0" fieldPosition="0"/></pivotAreas></conditionalFormat></conditionalFormats>';
  let bytes = await base({ tableTail: tableFormats });
  // The table's own rule (priority 2, pivot="1"), a rule of the sheet (priority 1) and another pivot rule no table names (priority 3).
  bytes = await edit(bytes, 'xl/worksheets/sheet2.xml', (xml) =>
    xml.replace(
      '</sheetData>',
      '</sheetData><conditionalFormatting sqref="H1"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>1</formula></cfRule></conditionalFormatting>' +
        '<conditionalFormatting sqref="B4" pivot="1"><cfRule type="cellIs" dxfId="0" priority="2" operator="greaterThan"><formula>1</formula></cfRule></conditionalFormatting>' +
        '<conditionalFormatting sqref="J1" pivot="1"><cfRule type="cellIs" dxfId="0" priority="3" operator="greaterThan"><formula>1</formula></cfRule></conditionalFormatting>',
    ),
  );
  bytes = await edit(bytes, 'xl/styles.xml', (xml) => xml.replace('</styleSheet>', '<dxfs count="1"><dxf><font><b/></font></dxf></dxfs></styleSheet>'));
  const doc = await load(bytes);
  assert.equal(doc.wb.sheets[1].conditionalFormats!.length, 3);
  const ops = deletePivot(doc.wb, 0);
  assert.ok(ops.some((op) => op.type === 'setConditionalFormats'));
  apply(doc, ops);
  assert.deepEqual(
    doc.wb.sheets[1].conditionalFormats!.map((format) => format.rules.map((rule) => rule.priority)),
    [[1], [3]],
  );
  const deleted = await save(doc);
  const sheet = await text(deleted, 'xl/worksheets/sheet2.xml');
  assert.doesNotMatch(sheet, /sqref="B4"/);
  assert.match(sheet, /sqref="H1"/);
  assert.match(sheet, /<conditionalFormatting sqref="J1" pivot="1">/);
  if (hasPython()) assert.deepEqual(independentPivots({ deleted }), []);
  doc.history.undo();
  assert.ok(same(await save(doc), bytes));
});

test('merged cells that the table takes are no longer merged; those beside it stay', async () => {
  const bytes = await edit(await base(), 'xl/worksheets/sheet2.xml', (xml) => xml.replace('</sheetData>', '</sheetData><mergeCells count="2"><mergeCell ref="B4:C5"/><mergeCell ref="H1:I2"/></mergeCells>'));
  const doc = await load(bytes);
  apply(doc, deletePivot(doc.wb, 0));
  assert.deepEqual(doc.wb.sheets[1].merges, [{ top: 0, left: 7, bottom: 1, right: 8 }]);
  const deleted = await save(doc);
  assert.match(await text(deleted, 'xl/worksheets/sheet2.xml'), /<mergeCells count="1"><mergeCell ref="H1:I2"\/><\/mergeCells>/);
  doc.history.undo();
  assert.deepEqual(doc.wb.sheets[1].merges.length, 2);
  assert.ok(same(await save(doc), bytes));
});

test('once no pivot table is left, the sheets are sorted; undone, they are not', async () => {
  const doc = await load(await base());
  const sort = () => sortRange(doc.wb, 0, { top: 0, left: 0, bottom: 4, right: 2 }, [{ col: 2 }], { header: true });
  const refused = () => {
    try {
      apply(doc, sort());
      return false;
    } catch (error) {
      assert.ok(error instanceof RangeError);
      return /pivotTable/.test((error as Error).message);
    }
  };
  assert.ok(refused(), 'a table is left');
  apply(doc, deletePivot(doc.wb, 0));
  apply(doc, sort());
  assert.deepEqual(
    [1, 2, 3, 4].map((r) => doc.wb.sheets[0].rows.get(r)!.cells[2].value),
    [5, 7, 10, 20],
  );
  doc.history.undo();
  doc.history.undo();
  assert.ok(refused(), 'the table is back');
});

test('what names a relationship is read in linear time: many ids in text, in one long tag, in comments', async () => {
  const { elementsNaming, withoutPivotSelections } = await import('../src/formats/xlsx/rel-names');
  const text = `<a>${'"rId1" '.repeat(200000)}</a><b ${Array.from({ length: 20000 }, (_, i) => `x${i}="rId1"`).join(' ')} r:id="rId1"/><!-- r:id="rId1" -->`;
  const started = Date.now();
  const found = elementsNaming(text, 'rId1');
  assert.ok(Date.now() - started < 2000, `${Date.now() - started} ms`);
  assert.deepEqual(found.map((el) => el.name), ['b']);
  const sheet = `<v><pivotSelection r:id="rId1"><x/></pivotSelection><pivotSelection r:id="rId2"/><x:pivotSelection r:id='rId1'/></v>`;
  assert.equal(withoutPivotSelections(sheet, ['rId1']), '<v><pivotSelection r:id="rId2"/></v>');
  assert.equal(withoutPivotSelections(sheet, ['rId9']), undefined);
});

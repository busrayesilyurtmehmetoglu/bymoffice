// The filter of a sheet in the file: what the reader makes of <autoFilter>,
// that an unchanged filter is written back as it was, what a changed one is
// written as, and a cross-check of both directions against openpyxl (skipped
// where python3 or openpyxl is not installed).

import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { writeXlsx } from '../src/formats/xlsx/write';
import { ZipArchive, open, rangeName } from '../src/index';
import type { OfficeFile, Sheet, Workbook } from '../src/index';
import { setCellInput } from '../src/ops/commands';
import { clearFilter, setAutoFilter, setFilterColumn } from '../src/ops/filter';
import { History, type Op } from '../src/ops/ops';
import { sortRange } from '../src/ops/sort';
import { buildXlsx, worksheetXml } from './fixtures';

const decoder = new TextDecoder();

interface Loaded {
  file: OfficeFile;
  wb: Workbook;
  history: History;
}

async function load(bytes: Uint8Array): Promise<Loaded> {
  const file = await open(bytes);
  const wb = await file.workbook();
  return { file, wb, history: new History(wb) };
}

async function save(doc: Loaded): Promise<Uint8Array> {
  await writeXlsx(doc.file.package!, doc.wb);
  return doc.file.package!.save();
}

const partText = async (bytes: Uint8Array, name: string) => decoder.decode(await ZipArchive.read(bytes).read(name));
const run = (doc: Loaded, ops: Op[]) => doc.history.apply(ops, undefined, undefined);

const num = (ref: string, value: number) => `<c r="${ref}"><v>${value}</v></c>`;
const str = (ref: string, text: string) => `<c r="${ref}" t="inlineStr"><is><t>${text}</t></is></c>`;

const FILTER =
  '<autoFilter ref="A1:D7" xr:uid="{00000000-0009-0000-0000-000000000000}" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">' +
  '<filterColumn colId="0"><filters blank="1"><filter val="Apple"/><filter val="Cherry"/></filters></filterColumn>' +
  '<filterColumn colId="1" hiddenButton="1"><colorFilter dxfId="0" cellColor="0"/></filterColumn>' +
  '<filterColumn colId="2"><customFilters and="1"><customFilter operator="greaterThanOrEqual" val="5"/><customFilter operator="lessThan" val="100"/></customFilters></filterColumn>' +
  '<filterColumn colId="3"><filters calendarType="gregorian"><dateGroupItem year="2024" month="2" dateTimeGrouping="month"/></filters></filterColumn>' +
  '<sortState ref="A2:D7"><sortCondition descending="1" ref="C2:C7"/><sortCondition ref="A2:A7" sortBy="cellColor" dxfId="1"/></sortState>' +
  '<extLst><ext uri="{TEST-EXT}"><x:thing xmlns:x="urn:test" keep="1"/></ext></extLst>' +
  '</autoFilter>';

const ROWS =
  `<row r="1">${str('A1', 'Name')}${str('B1', 'Colour')}${str('C1', 'Amount')}${str('D1', 'When')}</row>` +
  `<row r="2">${str('A2', 'Apple')}${str('B2', 'red')}${num('C2', 10)}${num('D2', 45323)}</row>` +
  `<row r="3" hidden="1">${str('A3', 'Banana')}${str('B3', 'yellow')}${num('C3', 25)}${num('D3', 45324)}</row>` +
  `<row r="4">${str('A4', 'Cherry')}${str('B4', 'red')}${num('C4', 5)}${num('D4', 45330)}</row>` +
  `<row r="5" hidden="1">${str('A5', 'Lemon')}${str('B5', 'yellow')}${num('C5', 400)}${num('D5', 45331)}</row>` +
  `<row r="6" hidden="1">${str('A6', 'Apple')}${str('B6', 'green')}${num('C6', 7)}${num('D6', 45300)}</row>` +
  `<row r="7" hidden="1">${str('A7', 'Cherry')}${str('B7', 'blue')}${num('C7', 8)}${num('D7', 45335)}</row>`;

const NAMES = "<definedNames><definedName name=\"Total\">Fruit!$C$2</definedName><definedName name=\"_xlnm._FilterDatabase\" localSheetId=\"0\" hidden=\"1\">Fruit!$A$1:$D$7</definedName></definedNames>";

const filtered = () =>
  buildXlsx({
    workbookTail: NAMES,
    sheets: [
      { name: 'Fruit', xml: worksheetXml(`<sheetPr filterMode="1"><tabColor rgb="FF00FF00"/></sheetPr><sheetData>${ROWS}</sheetData>${FILTER}<mergeCells count="1"><mergeCell ref="F1:G1"/></mergeCells>`) },
      {
        name: 'Other kinds',
        xml: worksheetXml(
          `<sheetData><row r="1">${str('A1', 'n')}${str('B1', 'm')}${str('C1', 'k')}</row><row r="2">${num('A2', 1)}${num('B2', 2)}${num('C2', 3)}</row></sheetData>` +
            '<autoFilter ref="A1:C2"><filterColumn colId="0"><top10 top="0" percent="1" val="20" filterVal="1"/></filterColumn>' +
            '<filterColumn colId="1"><dynamicFilter type="aboveAverage" val="2"/></filterColumn>' +
            '<filterColumn colId="2" showButton="0"/><filterColumn colId="99999"><filters/></filterColumn></autoFilter>',
        ),
      },
      { name: 'Plain', xml: worksheetXml(`<sheetData><row r="1">${str('A1', 'h')}${str('B1', 'k')}</row><row r="2">${num('A2', 2)}${num('B2', 1)}</row><row r="3">${num('A3', 1)}${num('B3', 2)}</row></sheetData><mergeCells count="1"><mergeCell ref="D1:E1"/></mergeCells>`) },
    ],
  });

const hidden = (sheet: Sheet) =>
  Array.from(sheet.rows.entries())
    .filter(([, row]) => row.hidden)
    .map(([index]) => index + 1)
    .sort((a, b) => a - b);

test('filter xlsx: the reader takes the range, every kind of criterion, the sort state and what it does not know', async () => {
  const { wb } = await load(await filtered());
  const filter = wb.sheets[0].autoFilter!;
  assert.equal(rangeName(filter.range), 'A1:D7');
  assert.deepEqual(filter.attrs, [
    ['xr:uid', '{00000000-0009-0000-0000-000000000000}'],
    ['xmlns:xr', 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision'],
  ]);
  assert.equal(filter.extra, '<extLst><ext uri="{TEST-EXT}"><x:thing xmlns:x="urn:test" keep="1"/></ext></extLst>');
  const columns = filter.columns.map((column) => ({ ...column, raw: undefined }));
  assert.deepEqual(columns, [
    { colId: 0, kind: 'values', values: ['Apple', 'Cherry'], blank: true, raw: undefined },
    { colId: 1, kind: 'color', hiddenButton: true, raw: undefined },
    { colId: 2, kind: 'custom', and: true, conditions: [{ operator: 'greaterThanOrEqual', value: '5' }, { operator: 'lessThan', value: '100' }], raw: undefined },
    { colId: 3, kind: 'values', values: [], dateGroups: [{ grouping: 'month', year: 2024, month: 2 }], raw: undefined },
  ]);
  assert.equal(filter.columns[1].raw, '<filterColumn colId="1" hiddenButton="1"><colorFilter dxfId="0" cellColor="0"/></filterColumn>');
  assert.deepEqual(filter.sortState, {
    range: { top: 1, left: 0, bottom: 6, right: 3 },
    conditions: [
      { range: { top: 1, left: 2, bottom: 6, right: 2 }, descending: true },
      { range: { top: 1, left: 0, bottom: 6, right: 0 }, descending: false, raw: '<sortCondition ref="A2:A7" sortBy="cellColor" dxfId="1"/>' },
    ],
  });
  assert.deepEqual(hidden(wb.sheets[0]), [3, 5, 6, 7]);
  assert.equal(wb.sheets[0].autoFilterModified, undefined);

  const other = wb.sheets[1].autoFilter!;
  assert.deepEqual(
    other.columns.map((column) => ({ ...column, raw: undefined })),
    [
      { colId: 0, kind: 'top10', top: false, percent: true, value: 20, raw: undefined },
      { colId: 1, kind: 'dynamic', dynamicType: 'aboveAverage', raw: undefined },
      { colId: 2, kind: 'none', hiddenButton: true, raw: undefined },
    ],
  );
  // A column that is none of the sheet's is kept as text.
  assert.equal(other.extra, '<filterColumn colId="99999"><filters/></filterColumn>');
  assert.equal(wb.sheets[2].autoFilter, undefined);
});

test('filter xlsx: a filter without a usable range is left to the source, and a second filter is ignored', async () => {
  const bytes = await buildXlsx({
    sheets: [
      { name: 'Bad', xml: worksheetXml('<sheetData/><autoFilter ref="nonsense"><filterColumn colId="0"/></autoFilter>') },
      { name: 'Twice', xml: worksheetXml('<sheetData/><autoFilter ref="A1:B2"/><autoFilter ref="C1:D2"/>') },
    ],
  });
  const { wb } = await load(bytes);
  assert.equal(wb.sheets[0].autoFilter, undefined);
  assert.equal(rangeName(wb.sheets[1].autoFilter!.range), 'A1:B2');
});

test('filter xlsx: nothing changed - every part is byte for byte what it was; a cell edit keeps the filter as written', async () => {
  const bytes = await filtered();
  const untouched = await load(bytes);
  const again = await save(untouched);
  const before = ZipArchive.read(bytes);
  const after = ZipArchive.read(again);
  for (const name of ['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml', 'xl/workbook.xml']) {
    assert.deepEqual(await after.read(name), await before.read(name), name);
  }

  const doc = await load(bytes);
  run(doc, setCellInput(doc.wb, 0, 8, 0, 'new'));
  const saved = await save(doc);
  const sheet = await partText(saved, 'xl/worksheets/sheet1.xml');
  assert.ok(sheet.includes(FILTER), 'the filter element is the source text');
  assert.ok(sheet.includes('<sheetPr filterMode="1"><tabColor rgb="FF00FF00"/></sheetPr>'));
  assert.ok(sheet.includes('<row r="3" hidden="1">'));
  assert.ok((await partText(saved, 'xl/workbook.xml')).includes(NAMES));
  assert.deepEqual(await ZipArchive.read(saved).read('xl/worksheets/sheet2.xml'), await before.read('xl/worksheets/sheet2.xml'));
});

test('filter xlsx: a changed criterion is written from the model; the other columns, attributes and extensions stay as they were', async () => {
  const doc = await load(await filtered());
  // The colour filter cannot be applied: rows it hides (6 and 7 pass everything else) stay hidden.
  run(doc, setFilterColumn(doc.wb, 0, 2, { kind: 'values', values: ['10', '5', '7', '<&"'] }));
  assert.deepEqual(hidden(doc.wb.sheets[0]), [3, 5, 6, 7]);
  const saved = await save(doc);
  const sheet = await partText(saved, 'xl/worksheets/sheet1.xml');
  const expected =
    '<autoFilter ref="A1:D7" xr:uid="{00000000-0009-0000-0000-000000000000}" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">' +
    '<filterColumn colId="0"><filters blank="1"><filter val="Apple"/><filter val="Cherry"/></filters></filterColumn>' +
    '<filterColumn colId="1" hiddenButton="1"><colorFilter dxfId="0" cellColor="0"/></filterColumn>' +
    '<filterColumn colId="2"><filters><filter val="10"/><filter val="5"/><filter val="7"/><filter val="&lt;&amp;&quot;"/></filters></filterColumn>' +
    '<filterColumn colId="3"><filters calendarType="gregorian"><dateGroupItem year="2024" month="2" dateTimeGrouping="month"/></filters></filterColumn>' +
    '<sortState ref="A2:D7"><sortCondition descending="1" ref="C2:C7"/><sortCondition ref="A2:A7" sortBy="cellColor" dxfId="1"/></sortState>' +
    '<extLst><ext uri="{TEST-EXT}"><x:thing xmlns:x="urn:test" keep="1"/></ext></extLst>' +
    '</autoFilter>';
  assert.ok(sheet.includes(`</sheetData>${expected}<mergeCells`), sheet.slice(sheet.indexOf('</sheetData>'), sheet.indexOf('</sheetData>') + 400));
  assert.ok((await partText(saved, 'xl/workbook.xml')).includes(NAMES));

  const reopened = await load(saved);
  const view = (model: Sheet) => model.autoFilter!.columns.map((column) => ({ ...column, raw: undefined }));
  assert.deepEqual(view(reopened.wb.sheets[0]), view(doc.wb.sheets[0]));
  assert.deepEqual(reopened.wb.sheets[0].autoFilter!.sortState, doc.wb.sheets[0].autoFilter!.sortState);
  assert.deepEqual(reopened.wb.sheets[0].autoFilter!.columns[2].values, ['10', '5', '7', '<&"']);
  assert.deepEqual(hidden(reopened.wb.sheets[0]), [3, 5, 6, 7]);
});

test('filter xlsx: clearing the criteria takes the filter mark off the sheet; switching the filter off removes element and name', async () => {
  const doc = await load(await filtered());
  // Every criterion goes, the colour filter included: all rows are shown.
  run(doc, clearFilter(doc.wb, 0));
  assert.deepEqual(hidden(doc.wb.sheets[0]), []);
  let saved = await save(doc);
  let sheet = await partText(saved, 'xl/worksheets/sheet1.xml');
  assert.ok(sheet.includes('<sheetPr><tabColor rgb="FF00FF00"/></sheetPr>'), sheet.slice(0, 400));
  assert.ok(sheet.includes('<filterColumn colId="1" hiddenButton="1"/>'));
  assert.ok(!sheet.includes('colorFilter') && !sheet.includes('hidden="1">'));
  assert.ok(sheet.includes('<sortState ref="A2:D7">'));

  run(doc, setAutoFilter(doc.wb, 0, undefined));
  saved = await save(doc);
  sheet = await partText(saved, 'xl/worksheets/sheet1.xml');
  assert.ok(!sheet.includes('autoFilter') && !sheet.includes('sortState'));
  const book = await partText(saved, 'xl/workbook.xml');
  assert.ok(book.includes('<definedNames><definedName name="Total">Fruit!$C$2</definedName></definedNames>'), book);
  assert.equal((await load(saved)).wb.sheets[0].autoFilter, undefined);

  // Undo all of it: the file is as it was read, apart from being regenerated.
  doc.history.undo();
  doc.history.undo();
  saved = await save(doc);
  sheet = await partText(saved, 'xl/worksheets/sheet1.xml');
  assert.ok(sheet.includes(FILTER));
  assert.ok(sheet.includes('<sheetPr filterMode="1"><tabColor rgb="FF00FF00"/></sheetPr>'));
  assert.ok((await partText(saved, 'xl/workbook.xml')).includes('<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Fruit!$A$1:$D$7</definedName>'));
  assert.deepEqual(hidden((await load(saved)).wb.sheets[0]), [3, 5, 6, 7]);
});

test('filter xlsx: a filter switched on in a sheet without one is put where the schema has it, with its name and mark', async () => {
  const doc = await load(await filtered());
  run(doc, setAutoFilter(doc.wb, 2, { top: 0, left: 0, bottom: 2, right: 1 }));
  let saved = await save(doc);
  let sheet = await partText(saved, 'xl/worksheets/sheet3.xml');
  assert.ok(sheet.includes('</sheetData><autoFilter ref="A1:B3"/><mergeCells'), sheet);
  assert.ok(!sheet.includes('sheetPr'));
  let book = await partText(saved, 'xl/workbook.xml');
  assert.ok(book.includes('<definedName name="_xlnm._FilterDatabase" localSheetId="2" hidden="1">Plain!$A$1:$B$3</definedName></definedNames>'), book);

  run(doc, setFilterColumn(doc.wb, 2, 1, { kind: 'values', values: ['2'], blank: true }));
  run(doc, sortRange(doc.wb, 2, { top: 0, left: 0, bottom: 2, right: 1 }, [{ col: 0 }], { header: true }));
  saved = await save(doc);
  sheet = await partText(saved, 'xl/worksheets/sheet3.xml');
  assert.ok(sheet.includes('<sheetPr filterMode="1"/><sheetData>') || sheet.includes('<sheetPr filterMode="1"/><dimension'), sheet.slice(0, 500));
  assert.ok(
    sheet.includes(
      '<autoFilter ref="A1:B3"><filterColumn colId="1"><filters blank="1"><filter val="2"/></filters></filterColumn><sortState ref="A2:B3"><sortCondition ref="A2:A3"/></sortState></autoFilter><mergeCells',
    ),
    sheet,
  );
  const reopened = await load(saved);
  assert.deepEqual(hidden(reopened.wb.sheets[2]), [2]);
  assert.deepEqual(reopened.wb.sheets[2].autoFilter!.sortState!.conditions, [{ range: { top: 1, left: 0, bottom: 2, right: 0 }, descending: false }]);
  book = await partText(saved, 'xl/workbook.xml');
  assert.equal(book.split('_xlnm._FilterDatabase').length - 1, 2);
});

test('filter xlsx: a workbook without defined names gets the list where the schema has it', async () => {
  const bytes = await buildXlsx({
    workbookTail: '<calcPr calcId="191029"/>',
    sheets: [{ name: 'A B', xml: worksheetXml(`<sheetData><row r="1">${str('A1', 'h')}</row><row r="2">${num('A2', 2)}</row></sheetData>`) }],
  });
  const doc = await load(bytes);
  run(doc, setAutoFilter(doc.wb, 0, { top: 0, left: 0, bottom: 1, right: 0 }));
  const book = await partText(await save(doc), 'xl/workbook.xml');
  assert.ok(book.includes("</sheets><definedNames><definedName name=\"_xlnm._FilterDatabase\" localSheetId=\"0\" hidden=\"1\">'A B'!$A$1:$A$2</definedName></definedNames><calcPr"), book);
});

test('filter xlsx: moved columns - the copied filter and the model agree, and a criterion read from the file keeps its place', async () => {
  // Not changed otherwise: the source element is adjusted.
  const copied = await load(await filtered());
  run(copied, [{ type: 'insert', sheet: 0, axis: 'col', at: 1, count: 1 }]);
  run(copied, [{ type: 'delete', sheet: 0, axis: 'col', at: 0, count: 1 }]);
  const fromSource = (await load(await save(copied))).wb.sheets[0].autoFilter!;
  const model = copied.wb.sheets[0].autoFilter!;
  assert.equal(rangeName(fromSource.range), 'A1:D7');
  assert.equal(rangeName(model.range), 'A1:D7');
  const view = (columns: typeof model.columns) => columns.map((column) => `${column.colId}:${column.kind}`);
  assert.deepEqual(view(fromSource.columns), ['1:color', '2:custom', '3:values']);
  assert.deepEqual(view(model.columns), view(fromSource.columns));
  assert.deepEqual(model.sortState!.conditions.map((condition) => rangeName(condition.range)), ['C2:C7']);
  assert.deepEqual(fromSource.sortState!.conditions.map((condition) => rangeName(condition.range)), ['C2:C7']);

  // Changed as well: written from the model, the kept elements at their new columns.
  const changed = await load(await filtered());
  run(changed, [{ type: 'insert', sheet: 0, axis: 'col', at: 1, count: 2 }]);
  run(changed, setFilterColumn(changed.wb, 0, 0, undefined));
  const sheet = await partText(await save(changed), 'xl/worksheets/sheet1.xml');
  assert.ok(sheet.includes('<autoFilter ref="A1:F7" '), sheet);
  assert.ok(!sheet.includes('colId="0"'));
  assert.ok(sheet.includes('<filterColumn colId="3" hiddenButton="1"><colorFilter dxfId="0" cellColor="0"/></filterColumn><filterColumn colId="4"><customFilters and="1">'));
  assert.ok(sheet.includes('<sortState ref="A2:F7"><sortCondition descending="1" ref="E2:E7"/><sortCondition ref="A2:A7" sortBy="cellColor" dxfId="1"/></sortState>'));
});

// --- openpyxl ----------------------------------------------------------------

const hasOpenpyxl = () => spawnSync('python3', ['-c', 'import openpyxl'], { encoding: 'utf8' }).status === 0;

/** Runs a python script with one file argument in a scratch folder inside the repository (ignored by git). */
function python(script: string, input: Uint8Array | undefined, output: boolean): { stdout: string; file?: Uint8Array } {
  const folder = join(dirname(fileURLToPath(import.meta.url)), '..', `tmp-filter-openpyxl-${process.pid}`);
  mkdirSync(folder, { recursive: true });
  try {
    const file = join(folder, 'book.xlsx');
    if (input) writeFileSync(file, input);
    const result = spawnSync('python3', ['-c', script, file], { encoding: 'utf8' });
    assert.equal(result.status, 0, result.stderr);
    return { stdout: result.stdout, file: output ? new Uint8Array(readFileSync(file)) : undefined };
  } finally {
    rmSync(folder, { recursive: true, force: true });
  }
}

const PYTHON_READ = `
import json, sys, zipfile
import xml.etree.ElementTree as ET
import openpyxl
path = sys.argv[1]
wb = openpyxl.load_workbook(path)
out = {}
for ws in wb.worksheets:
    af = ws.auto_filter
    cols = []
    for fc in af.filterColumn:
        col = {'colId': fc.colId, 'hiddenButton': bool(fc.hiddenButton)}
        if fc.filters is not None:
            col['values'] = list(fc.filters.filter)
            col['blank'] = bool(fc.filters.blank)
        if fc.customFilters is not None:
            col['and'] = bool(fc.customFilters._and)
            col['custom'] = [[c.operator or 'equal', c.val] for c in fc.customFilters.customFilter]
        if fc.colorFilter is not None:
            col['color'] = True
        cols.append(col)
    sort = None
    if af.sortState is not None:
        sort = {'ref': af.sortState.ref, 'conditions': [[c.ref, bool(c.descending)] for c in af.sortState.sortCondition]}
    out[ws.title] = {'ref': af.ref, 'columns': cols, 'sort': sort,
                     'hidden': sorted(r for r, d in ws.row_dimensions.items() if d.hidden)}
ns = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
root = ET.fromstring(zipfile.ZipFile(path).read('xl/workbook.xml'))
out['names'] = sorted([n.get('name'), n.get('localSheetId'), n.get('hidden'), n.text] for n in root.findall('m:definedNames/m:definedName', ns))
# What openpyxl itself writes as the name of each filter.
wb.save(path)
root = ET.fromstring(zipfile.ZipFile(path).read('xl/workbook.xml'))
out['theirNames'] = sorted([n.get('name'), n.get('localSheetId'), n.get('hidden'), n.text] for n in root.findall('m:definedNames/m:definedName', ns) if n.get('name') == '_xlnm._FilterDatabase')
print(json.dumps(out))
`;

test('filter xlsx: openpyxl reads the filter, the sort state and the hidden rows this engine writes', async (t) => {
  if (!hasOpenpyxl()) {
    t.skip('python3 with openpyxl is not available');
    return;
  }
  const doc = await load(await filtered());
  run(doc, setFilterColumn(doc.wb, 0, 2, { kind: 'custom', and: false, conditions: [{ operator: 'lessThan', value: '8' }, { operator: 'equal', value: '4*' }] }));
  run(doc, setFilterColumn(doc.wb, 0, 3, undefined));
  run(doc, setAutoFilter(doc.wb, 2, { top: 0, left: 0, bottom: 2, right: 1 }));
  run(doc, setFilterColumn(doc.wb, 2, 0, { kind: 'values', values: ['1'], blank: true }));
  run(doc, sortRange(doc.wb, 2, { top: 0, left: 0, bottom: 2, right: 1 }, [{ col: 1, descending: true }, { col: 0 }], { header: true }));
  const theirs = JSON.parse(python(PYTHON_READ, await save(doc), false).stdout);

  assert.deepEqual(theirs.Fruit, {
    ref: 'A1:D7',
    columns: [
      { colId: 0, hiddenButton: false, values: ['Apple', 'Cherry'], blank: true },
      { colId: 1, hiddenButton: true, color: true },
      { colId: 2, hiddenButton: false, and: false, custom: [['lessThan', '8'], ['equal', '4*']] },
    ],
    sort: { ref: 'A2:D7', conditions: [['C2:C7', true], ['A2:A7', false]] },
    hidden: hidden(doc.wb.sheets[0]),
  });
  assert.deepEqual(theirs.Plain, {
    ref: 'A1:B3',
    columns: [{ colId: 0, hiddenButton: false, values: ['1'], blank: true }],
    sort: { ref: 'A2:B3', conditions: [['B2:B3', true], ['A2:A3', false]] },
    hidden: hidden(doc.wb.sheets[2]),
  });
  // The second sheet's filter came without a name and was not changed: nothing is added for it.
  assert.deepEqual(theirs.names, [
    ['Total', null, null, 'Fruit!$C$2'],
    ['_xlnm._FilterDatabase', '0', '1', 'Fruit!$A$1:$D$7'],
    ['_xlnm._FilterDatabase', '2', '1', 'Plain!$A$1:$B$3'],
  ]);
  // openpyxl names the ranges of the same filters the same way (it quotes every sheet name).
  const unquote = (text: string) => text.replace(/^'([^']*)'!/, '$1!');
  const ours = (theirs.names as (string | null)[][]).filter((name) => name[0] === '_xlnm._FilterDatabase').map((name) => [name[1], unquote(name[3] as string)]);
  const written = (theirs.theirNames as (string | null)[][]).filter((name) => name[1] !== '1').map((name) => [name[1], unquote(name[3] as string)]);
  assert.deepEqual(ours, written);
});

const PYTHON_WRITE = `
import sys
from openpyxl import Workbook
from openpyxl.worksheet.filters import FilterColumn, CustomFilters, CustomFilter, Top10, DynamicFilter, ColorFilter
wb = Workbook()
ws = wb.active
ws.title = 'Made by openpyxl'
ws.append(['Name', 'Colour', 'Amount', 'Rank', 'Mean', 'Tint'])
rows = [['Apple', 'red', 10, 1, 5, 'x'], ['Banana', 'yellow', 25, 2, 6, 'y'], ['Cherry', 'red', 5, 3, 7, 'z'], [None, 'green', 7, 4, 8, 'w']]
for row in rows:
    ws.append(row)
ws.auto_filter.ref = 'A1:F5'
ws.auto_filter.add_filter_column(0, ['Apple', 'Cherry'], blank=True)
ws.auto_filter.add_filter_column(1, ['red'])
ws.auto_filter.filterColumn.append(FilterColumn(colId=2, customFilters=CustomFilters(_and=True, customFilter=[CustomFilter(operator='greaterThan', val='4'), CustomFilter(operator='lessThanOrEqual', val='10')])))
ws.auto_filter.filterColumn.append(FilterColumn(colId=3, top10=Top10(top=True, val=3, filterVal=2)))
ws.auto_filter.filterColumn.append(FilterColumn(colId=4, dynamicFilter=DynamicFilter(type='belowAverage', val=6.5)))
ws.auto_filter.filterColumn.append(FilterColumn(colId=5, colorFilter=ColorFilter(dxfId=0), hiddenButton=True))
ws.auto_filter.add_sort_condition('C2:C5', descending=True)
ws.row_dimensions[3].hidden = True
ws.row_dimensions[5].hidden = True
wb.save(sys.argv[1])
`;

test('filter xlsx: a filter written by openpyxl is read, can be changed and is read back by openpyxl', async (t) => {
  if (!hasOpenpyxl()) {
    t.skip('python3 with openpyxl is not available');
    return;
  }
  const made = python(PYTHON_WRITE, undefined, true).file as Uint8Array;
  const doc = await load(made);
  const sheet = doc.wb.sheets[0];
  const filter = sheet.autoFilter!;
  assert.equal(rangeName(filter.range), 'A1:F5');
  assert.deepEqual(
    filter.columns.map((column) => ({ ...column, raw: undefined })),
    [
      { colId: 0, kind: 'values', values: ['Apple', 'Cherry'], blank: true, raw: undefined },
      { colId: 1, kind: 'values', values: ['red'], raw: undefined },
      { colId: 2, kind: 'custom', and: true, conditions: [{ operator: 'greaterThan', value: '4' }, { operator: 'lessThanOrEqual', value: '10' }], raw: undefined },
      { colId: 3, kind: 'top10', top: true, percent: false, value: 3, raw: undefined },
      { colId: 4, kind: 'dynamic', dynamicType: 'belowAverage', raw: undefined },
      { colId: 5, kind: 'color', hiddenButton: true, raw: undefined },
    ],
  );
  // openpyxl states the range of the filter as the sorted range.
  assert.equal(rangeName(filter.sortState!.range), 'A1:F5');
  assert.deepEqual(filter.sortState!.conditions, [{ range: { top: 1, left: 2, bottom: 4, right: 2 }, descending: true }]);
  assert.deepEqual(hidden(sheet), [3, 5]);

  // Unchanged, the sheet part goes back byte for byte.
  const untouched = await save(await load(made));
  assert.deepEqual(await ZipArchive.read(untouched).read('xl/worksheets/sheet1.xml'), await ZipArchive.read(made).read('xl/worksheets/sheet1.xml'));

  // Criteria are removed and changed; what openpyxl then reads has to be the model.
  run(doc, setFilterColumn(doc.wb, 0, 3, undefined));
  run(doc, setFilterColumn(doc.wb, 0, 4, undefined));
  run(doc, setFilterColumn(doc.wb, 0, 1, { kind: 'values', values: ['red', 'green'] }));
  const theirs = JSON.parse(python(PYTHON_READ, await save(doc), false).stdout)['Made by openpyxl'];
  assert.equal(theirs.ref, 'A1:F5');
  assert.deepEqual(theirs.columns, [
    { colId: 0, hiddenButton: false, values: ['Apple', 'Cherry'], blank: true },
    { colId: 1, hiddenButton: false, values: ['red', 'green'], blank: false },
    { colId: 2, hiddenButton: false, and: true, custom: [['greaterThan', '4'], ['lessThanOrEqual', '10']] },
    { colId: 5, hiddenButton: true, color: true },
  ]);
  assert.deepEqual(theirs.sort, { ref: 'A1:F5', conditions: [['C2:C5', true]] });
  assert.deepEqual(theirs.hidden, hidden(sheet));
});

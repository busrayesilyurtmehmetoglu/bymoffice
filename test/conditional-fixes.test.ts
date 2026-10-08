// Conditional formats: the defects a review found, each pinned by a test that failed before
// its fix, and tests for the mutations the first suite let through.

import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { CalcEngine } from '../src/calc/engine';
import { ConditionalFormats, type CellConditional, type ConditionalOptions } from '../src/format/conditional';
import { parseStyles } from '../src/formats/xlsx/styles';
import { parseWorksheet } from '../src/formats/xlsx/worksheet';
import { writeXlsx } from '../src/formats/xlsx/write';
import { open } from '../src/index';
import { MAX_COLS, MAX_ROWS, cellName, createSheet, parseCellRef, type Cell, type CellRange, type ConditionalRule, type Sheet, type Workbook } from '../src/model/sheet/types';
import { applyOp, type Op } from '../src/ops/ops';
import { parseXml } from '../src/xml';
import { buildXlsx, worksheetXml } from './fixtures';

// --- helpers ---------------------------------------------------------------

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
const X14 = 'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"';
const utf8 = (text: string) => new TextEncoder().encode(text);

type Value = number | string | boolean | null;

const colName = (col: number) => String.fromCharCode(65 + col);

function sheetData(rows: Value[][]): string {
  let out = '<sheetData>';
  rows.forEach((row, r) => {
    out += `<row r="${r + 1}">`;
    row.forEach((value, c) => {
      const ref = `${colName(c)}${r + 1}`;
      if (value === null) return;
      if (typeof value === 'number') out += `<c r="${ref}"><v>${value}</v></c>`;
      else if (typeof value === 'boolean') out += `<c r="${ref}" t="b"><v>${value ? 1 : 0}</v></c>`;
      else out += `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${value}</t></is></c>`;
    });
    out += '</row>';
  });
  return out + '</sheetData>';
}

const column = (values: Value[]) => values.map((value) => [value]);

/** Differential format 0: a red fill. */
const DXFS = '<dxf><fill><patternFill><bgColor rgb="FFFF0000"/></patternFill></fill></dxf>';

function book(body: string): { workbook: Workbook; sheet: Sheet } {
  const sheet = createSheet({ name: 'Data', sheetId: 1, state: 'visible', kind: 'worksheet', part: '/xl/worksheets/sheet1.xml' });
  parseWorksheet(utf8(`<worksheet ${NS} ${X14}>${body}</worksheet>`), [], sheet);
  const styles = parseStyles(parseXml(`<styleSheet ${NS}><dxfs>${DXFS}</dxfs></styleSheet>`), undefined);
  return { workbook: { sheets: [sheet], styles, sharedStrings: [], definedNames: [], date1904: false, activeSheet: 0 }, sheet };
}

const cf = (sqref: string, rules: string) => `<conditionalFormatting sqref="${sqref}">${rules}</conditionalFormatting>`;
const expression = (formula: string) => `<cfRule type="expression" dxfId="0" priority="1"><formula>${formula}</formula></cfRule>`;

/** Serial 45000 is Wednesday 15 March 2023. */
const NOW = 45000.5;

function withEngine(workbook: Workbook): ConditionalOptions {
  const engine = new CalcEngine(workbook, { now: () => NOW });
  return { engine: () => engine };
}

function at(formats: ConditionalFormats, ref: string): CellConditional | undefined {
  const cell = parseCellRef(ref)!;
  return formats.resolve(cell.row, cell.col);
}

function filled(formats: ConditionalFormats, refs: string[]): string[] {
  return refs.filter((ref) => {
    const found = at(formats, ref);
    return !!found && !!found.style && found.style.fill !== undefined;
  });
}

const refs = (col: string, count: number) => Array.from({ length: count }, (_, i) => `${col}${i + 1}`);

function setValue(sheet: Sheet, ref: string, value: number | string): void {
  const { row, col } = parseCellRef(ref)!;
  const cell: Cell = typeof value === 'number' ? { kind: 'number', value, style: 0 } : { kind: 'string', value, style: 0 };
  let found = sheet.rows.get(row);
  if (!found) {
    found = { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells: [] };
    sheet.rows.set(row, found);
  }
  found.cells[col] = cell;
}

const BASE = { stopIfTrue: false, percent: false, bottom: false, aboveAverage: true, equalAverage: false, formulas: [] as string[], priority: 1, dxfId: 0 };

/** The cells of column A a rule given as an object fills, for values put in A1.. */
function ruled(values: Value[], rule: Partial<ConditionalRule> & { type: string }, options?: (workbook: Workbook) => ConditionalOptions, ranges?: CellRange[]): string[] {
  const { workbook, sheet } = book(sheetData(column(values)));
  sheet.conditionalFormats = [{ ranges: ranges || [{ top: 0, left: 0, bottom: values.length - 1, right: 0 }], rules: [{ ...BASE, ...rule }] }];
  return filled(new ConditionalFormats(workbook, sheet, options ? options(workbook) : {}), refs('A', values.length));
}

const elapsed = (run: () => void): number => {
  const start = Date.now();
  run();
  return Date.now() - start;
};

// --- 1. whole columns and rows through structural edits ------------------------

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const STYLES =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="${NS_MAIN}">` +
  '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
  '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
  '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
  `<dxfs count="1">${DXFS}</dxfs></styleSheet>`;

/** Six rows of six numbers, positive except on one diagonal pattern. */
const GRID = sheetData(Array.from({ length: 6 }, (_, r) => Array.from({ length: 6 }, (_, c) => ((r + c) % 3 === 0 ? -1 : 1))));

/** Which cells of the top-left corner of the first sheet are filled, and by nothing else than the rules. */
function snapshot(workbook: Workbook): string[] {
  const engine = new CalcEngine(workbook);
  const formats = new ConditionalFormats(workbook, workbook.sheets[0], { engine: () => engine });
  const out: string[] = [];
  for (let r = 0; r < 12; r++) {
    for (let c = 0; c < 10; c++) {
      const found = formats.resolve(r, c);
      if (found && found.style && found.style.fill !== undefined) out.push(cellName(r, c));
    }
  }
  // The far corner of the sheet, where whole columns and rows end.
  for (const [r, c] of [[MAX_ROWS - 1, 1], [MAX_ROWS - 1, 2], [2, MAX_COLS - 1], [3, MAX_COLS - 1]]) if (formats.resolve(r, c)) out.push(cellName(r, c));
  return out;
}

interface EditCase {
  name: string;
  sqref: string;
  formula: string;
  op: Op;
  /** What the file has to say afterwards; '' when the rule is gone. */
  sqrefAfter: string;
  formulaAfter: string;
}

const rowsOp = (type: 'insert' | 'delete', at: number, count: number): Op => ({ type, sheet: 0, axis: 'row', at, count });
const colsOp = (type: 'insert' | 'delete', at: number, count: number): Op => ({ type, sheet: 0, axis: 'col', at, count });

const EDIT_CASES: EditCase[] = [
  { name: 'whole column, column inserted before it', sqref: 'B:B', formula: 'B1&gt;0', op: colsOp('insert', 0, 1), sqrefAfter: 'C:C', formulaAfter: 'C1>0' },
  { name: 'whole row, row inserted above it', sqref: '3:3', formula: 'A3&gt;0', op: rowsOp('insert', 0, 1), sqrefAfter: '4:4', formulaAfter: 'A4>0' },
  { name: 'range and whole column, column inserted', sqref: 'D1:D4 B:B', formula: 'D1&gt;0', op: colsOp('insert', 0, 1), sqrefAfter: 'E1:E4 C:C', formulaAfter: 'E1>0' },
  { name: 'range and whole column, row inserted', sqref: 'D1:D4 B:B', formula: 'D1&gt;0', op: rowsOp('insert', 0, 1), sqrefAfter: 'D2:D5 B:B', formulaAfter: 'D2>0' },
  { name: 'whole column, row inserted', sqref: 'B:B', formula: 'B1&gt;=$A$2', op: rowsOp('insert', 0, 1), sqrefAfter: 'B:B', formulaAfter: 'B1>=$A$3' },
  { name: 'whole column, rows deleted', sqref: 'B:B', formula: 'B1&gt;0', op: rowsOp('delete', 0, 2), sqrefAfter: 'B:B', formulaAfter: 'B1>0' },
  { name: 'whole row, column inserted', sqref: '3:3', formula: 'A3&gt;0', op: colsOp('insert', 0, 2), sqrefAfter: '3:3', formulaAfter: 'A3>0' },
  { name: 'whole column written as a range, column inserted', sqref: 'B1:B1048576', formula: 'B1&gt;0', op: colsOp('insert', 0, 1), sqrefAfter: 'C1:C1048576', formulaAfter: 'C1>0' },
  { name: 'whole column written as a range, row inserted', sqref: 'B1:B1048576', formula: 'B1&gt;0', op: rowsOp('insert', 0, 1), sqrefAfter: 'B1:B1048576', formulaAfter: 'B1>0' },
  { name: 'whole row written as a range, row inserted', sqref: 'A3:XFD3', formula: 'A3&gt;0', op: rowsOp('insert', 0, 1), sqrefAfter: 'A4:XFD4', formulaAfter: 'A4>0' },
  { name: 'two whole columns, one deleted', sqref: 'B:C', formula: 'B1&gt;0', op: colsOp('delete', 1, 1), sqrefAfter: 'B:B', formulaAfter: 'B1>0' },
  { name: 'two whole rows, one deleted', sqref: '2:3', formula: 'A2&gt;0', op: rowsOp('delete', 1, 1), sqrefAfter: '2:2', formulaAfter: 'A2>0' },
  { name: 'whole column deleted, a range left', sqref: 'B:B D1:D3', formula: 'B1&gt;0', op: colsOp('delete', 1, 1), sqrefAfter: 'C1:C3', formulaAfter: 'C1>0' },
  { name: 'whole column deleted, nothing left', sqref: 'B:B', formula: 'B1&gt;0', op: colsOp('delete', 1, 1), sqrefAfter: '', formulaAfter: '' },
  // The cell above the first row of the range is the last row of the sheet.
  { name: 'a reference above the range wraps', sqref: 'B2:B5', formula: 'B1&gt;0', op: rowsOp('delete', 0, 1), sqrefAfter: 'B1:B4', formulaAfter: 'B1048576>0' },
  { name: 'a reference left of the range wraps', sqref: 'B2:B5', formula: 'A2&lt;=0', op: colsOp('delete', 0, 1), sqrefAfter: 'A2:A5', formulaAfter: 'XFD2<=0' },
];

interface Edited {
  live: string[];
  saved: Uint8Array;
  reopened: Workbook;
}

async function editAndSave(item: EditCase): Promise<Edited> {
  const bytes = await buildXlsx({ sheets: [{ name: 'Data', xml: worksheetXml(GRID + cf(item.sqref, expression(item.formula))) }], styles: STYLES });
  const file = await open(bytes);
  const workbook = await file.workbook();
  applyOp(workbook, item.op);
  const live = snapshot(workbook);
  await writeXlsx(file.package!, workbook);
  const saved = await file.package!.save();
  return { live, saved, reopened: await (await open(saved)).workbook() };
}

test('conditional fixes: after rows or columns move, the view and the saved file agree on whole columns and rows', async () => {
  for (const item of EDIT_CASES) {
    const { live, reopened } = await editAndSave(item);
    const formats = reopened.sheets[0].conditionalFormats || [];
    if (item.sqrefAfter === '') {
      assert.equal(formats.length, 0, item.name);
      assert.deepEqual(live, [], item.name);
      continue;
    }
    assert.equal(formats.length, 1, `${item.name}: the rule is still in the file`);
    assert.equal(formats[0].ranges.length, item.sqrefAfter.split(' ').length, `${item.name}: ranges`);
    assert.deepEqual(formats[0].rules[0].formulas, [item.formulaAfter], item.name);
    assert.ok(live.length > 0, `${item.name}: the rule fills something`);
    assert.deepEqual(snapshot(reopened), live, `${item.name}: reopened file against the live view`);
  }
});

test('conditional fixes: the view after an edit, cell by cell', async () => {
  // B:B with B1>0 and a column inserted before it: the numbers of old column B are in C now.
  const column = await editAndSave(EDIT_CASES[0]);
  assert.deepEqual(column.live, ['C1', 'C2', 'C4', 'C5']);
  // D1:D4 B:B, a row inserted above: the column still starts at row 1, the range moved down.
  const mixed = await editAndSave(EDIT_CASES[3]);
  assert.deepEqual(mixed.live, ['B2', 'B3', 'D3', 'D4', 'B5', 'B6']);
  assert.deepEqual(mixed.reopened.sheets[0].conditionalFormats![0].ranges, [
    { top: 1, left: 3, bottom: 4, right: 3 },
    { top: 0, left: 1, bottom: MAX_ROWS - 1, right: 1 },
  ]);
  // B2:B5 "the cell above is positive", row 1 deleted: B1 looks at the last row of the sheet (empty), B2 at B1.
  const wrapped = await editAndSave(EDIT_CASES[14]);
  assert.deepEqual(wrapped.live, ['B2', 'B4']);
});

// openpyxl 3.1 does not take a whole column or row as the range of a rule (it drops the rule with a
// warning), so those files are read with Python's own zip and XML modules instead: still a reader
// that shares no code with this project.
const PYTHON_READ = `
import json, sys, warnings, zipfile
import xml.etree.ElementTree as ET
import openpyxl
warnings.simplefilter('ignore')
NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
out = []
for path in sys.argv[1:]:
    root = ET.fromstring(zipfile.ZipFile(path).read('xl/worksheets/sheet1.xml'))
    raw = [{'sqref': cf.get('sqref'), 'formulas': [f.text for f in cf.iter(NS + 'formula')]} for cf in root.iter(NS + 'conditionalFormatting')]
    ws = openpyxl.load_workbook(path).active
    found = []
    for cf in ws.conditional_formatting:
        found.append({'sqref': str(cf.sqref), 'formulas': [str(f) for r in cf.rules for f in r.formula]})
    out.append({'raw': raw, 'openpyxl': found})
print(json.dumps(out))
`;

test('conditional fixes: openpyxl reads the ranges and formulas written after an edit', async (t) => {
  const probe = spawnSync('python3', ['-c', 'import openpyxl'], { encoding: 'utf8' });
  if (probe.status !== 0) {
    t.skip('python3 with openpyxl is not available');
    return;
  }
  const folder = join(dirname(fileURLToPath(import.meta.url)), '..', `tmp-cffix-openpyxl-${process.pid}`);
  mkdirSync(folder, { recursive: true });
  type Read = { sqref: string; formulas: string[] }[];
  let theirs: { raw: Read; openpyxl: Read }[];
  try {
    const files: string[] = [];
    for (let i = 0; i < EDIT_CASES.length; i++) {
      const file = join(folder, `case${i}.xlsx`);
      writeFileSync(file, (await editAndSave(EDIT_CASES[i])).saved);
      files.push(file);
    }
    const run = spawnSync('python3', ['-c', PYTHON_READ, ...files], { encoding: 'utf8' });
    assert.equal(run.status, 0, run.stderr);
    theirs = JSON.parse(run.stdout);
  } finally {
    rmSync(folder, { recursive: true, force: true });
  }
  EDIT_CASES.forEach((item, i) => {
    const expected = item.sqrefAfter === '' ? [] : [{ sqref: item.sqrefAfter, formulas: [item.formulaAfter] }];
    assert.deepEqual(theirs[i].raw, expected, item.name);
    // Ranges made of cells only are ranges openpyxl reads.
    const cellsOnly = item.sqrefAfter.split(' ').every((ref) => ref.split(':').every((side) => parseCellRef(side) !== undefined));
    if (cellsOnly) assert.deepEqual(theirs[i].openpyxl, expected, `${item.name} (openpyxl)`);
  });
  assert.ok(EDIT_CASES.filter((item) => item.sqrefAfter !== '' && !/(^| )(\\d+:|[A-Z]+:[A-Z]+( |$))/.test(item.sqrefAfter)).length >= 5);
});

// --- 2. hostile texts ----------------------------------------------------------

test('conditional fixes: long runs of digits and megabyte attributes are read in linear time', () => {
  const digits = '1'.repeat(40000) + 'x';
  // A formula and a threshold that are almost a number.
  const hostile = book(
    sheetData(column([1, 2, 3])) +
      cf('A1:A3', `<cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>${digits}</formula></cfRule>`) +
      cf('A1:A3', `<cfRule type="colorScale" priority="2"><colorScale><cfvo type="num" val="${digits}"/><cfvo type="max"/><color rgb="FF000000"/><color rgb="FFFFFFFF"/></colorScale></cfRule>`),
  );
  for (const options of [{}, withEngine(hostile.workbook)]) {
    const formats = new ConditionalFormats(hostile.workbook, hostile.sheet, options);
    assert.ok(elapsed(() => assert.deepEqual(filled(formats, refs('A', 3)), [])) < 500, 'a formula of 40,000 digits');
  }
  // A structural edit looks at the thresholds as well.
  hostile.sheet.structuralEdits = [{ kind: 'insert', axis: 'row', at: 0, count: 1 }];
  const moved = new ConditionalFormats(hostile.workbook, hostile.sheet);
  assert.ok(elapsed(() => moved.resolve(1, 0)) < 500, 'thresholds after an edit');

  const big = 1 << 20;
  const inputs: [string, string][] = [
    ['digits in a cell value', `<sheetData><row r="1"><c r="A1"><v>${'9'.repeat(big)}x</v></c></row></sheetData>`],
    ['digits as a formula', sheetData([[1]]) + cf('A1', `<cfRule type="cellIs" dxfId="0" priority="1" operator="equal"><formula>${'9'.repeat(big)}x</formula></cfRule>`)],
    ['digits and dots as a formula', sheetData([[1]]) + cf('A1', expression('1.'.repeat(big / 2)))],
    ['digits as a threshold', sheetData([[1]]) + cf('A1', `<cfRule type="iconSet" priority="1"><iconSet><cfvo type="num" val="${'9'.repeat(big)}e"/><cfvo type="num" val="${'9'.repeat(big)}"/><cfvo type="percent" val="${' '.repeat(big)}"/></iconSet></cfRule>`)],
    ['letters as sqref', sheetData([[1]]) + cf('A'.repeat(big), expression('TRUE'))],
    ['digits as sqref', sheetData([[1]]) + cf('1'.repeat(big / 2) + ':' + '1'.repeat(big / 2), expression('TRUE'))],
    ['colons as sqref', sheetData([[1]]) + cf('A:'.repeat(big / 2), expression('TRUE'))],
    ['spaces as sqref', sheetData([[1]]) + cf(' '.repeat(big) + 'A1', expression('TRUE'))],
    ['a colour', sheetData([[1]]) + cf('A1', `<cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="${'F'.repeat(big)}"/><color theme="${'1'.repeat(big)}"/></colorScale></cfRule>`)],
    ['bar lengths', sheetData([[1]]) + cf('A1', `<cfRule type="dataBar" priority="1"><dataBar minLength="${'1'.repeat(big)}x" maxLength="${'0.'.repeat(big / 2)}"><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule>`)],
    ['a text to look for', sheetData([['a'.repeat(40000)]]) + cf('A1', `<cfRule type="containsText" dxfId="0" priority="1" text="${'a*?~'.repeat(big / 4)}"/>`)],
    ['quotes as a formula', sheetData([[1]]) + cf('A1', expression('"' + '""'.repeat(big / 2)))],
  ];
  for (const [label, body] of inputs) {
    const time = elapsed(() => {
      let read: { workbook: Workbook; sheet: Sheet } | undefined;
      try {
        read = book(body);
      } catch (error) {
        // Refusing the sheet is an answer too, as long as it comes quickly.
        assert.ok(error instanceof Error, label);
      }
      if (read) {
        const formats = new ConditionalFormats(read.workbook, read.sheet, withEngine(read.workbook));
        formats.resolve(0, 0);
      }
    });
    assert.ok(time < 15000, `${label}: ${time} ms`);
  }
});

// --- 3. blanks -----------------------------------------------------------------

test('conditional fixes: only the space character is trimmed when looking for blanks', () => {
  // Excel asks LEN(TRIM(A1))=0, and TRIM takes away nothing but U+0020.
  const { workbook, sheet } = book(sheetData(column([1, 1, 1, 1, 1, 1, 1])));
  ['', ' ', '   ', ' ', '\t', '\n', '   '].forEach((text, i) => setValue(sheet, `A${i + 1}`, text));
  const run = (type: string) => {
    sheet.conditionalFormats = [{ ranges: [{ top: 0, left: 0, bottom: 7, right: 0 }], rules: [{ ...BASE, type }] }];
    return filled(new ConditionalFormats(workbook, sheet), refs('A', 8));
  };
  // A8 has no cell at all.
  assert.deepEqual(run('containsBlanks'), ['A1', 'A2', 'A3', 'A8']);
  assert.deepEqual(run('notContainsBlanks'), ['A4', 'A5', 'A6', 'A7']);
});

// --- 4. and 5. data bars -------------------------------------------------------

const bars = (values: number[], attributes: string, cfvos = '<x14:cfvo type="min"/><x14:cfvo type="max"/>') => {
  const { workbook, sheet } = book(
    sheetData(column(values)) +
      cf(`A1:A${values.length}`, '<cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"><x14:id>{1}</x14:id></ext></extLst></cfRule>') +
      `<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="dataBar" id="{1}"><x14:dataBar ${attributes}>${cfvos}</x14:dataBar></x14:cfRule><xm:sqref>A1:A${values.length}</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>`,
  );
  const formats = new ConditionalFormats(workbook, sheet);
  return { sheet, bars: refs('A', values.length).map((ref) => at(formats, ref)!.bar!) };
};

test('conditional fixes: a data bar of the 2010 extension without lengths runs from 10 to 90 percent', () => {
  // UNSURE: the schema default of x14:dataBar is minLength 10 and maxLength 90; Excel itself always
  // writes both attributes (0 and 100), so only files of other producers get here.
  const implied = bars([0, 50, 100], '');
  assert.equal(implied.sheet.conditionalFormats![0].rules[0].dataBar!.minLength, 10);
  assert.equal(implied.sheet.conditionalFormats![0].rules[0].dataBar!.maxLength, 90);
  assert.deepEqual(implied.bars.map((bar) => [bar.from, Math.round(bar.to * 1000) / 1000]), [[0, 0.1], [0, 0.5], [0, 0.9]]);
  // Stated lengths are taken as they are.
  const stated = bars([0, 50, 100], 'minLength="0" maxLength="100"');
  assert.deepEqual(stated.bars.map((bar) => bar.to), [0, 0.5, 1]);
});

test('conditional fixes: an axis in the middle stays there when every value is positive', () => {
  // UNSURE: Excel is believed to keep the axis at the cell midpoint for axisPosition="middle" whatever
  // the signs of the values; positive bars then use the right half only.
  const found = bars([50, 100, 25], 'minLength="0" maxLength="100" axisPosition="middle"', '<x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:axisColor rgb="FF123456"/>').bars;
  assert.deepEqual(found.map((bar) => [bar.axis, bar.from, bar.to, bar.negative]), [
    [0.5, 0.5, 0.75, false],
    [0.5, 0.5, 1, false],
    [0.5, 0.5, 0.625, false],
  ]);
  assert.equal(found[0].axisColor, '#123456');
  // Without an axis position that asks for it, positive values have no axis.
  const automatic = bars([50, 100], 'minLength="0" maxLength="100"', '<x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/>').bars;
  assert.deepEqual(automatic.map((bar) => [bar.axis, bar.from, bar.to]), [[undefined, 0, 0.5], [undefined, 0, 1]]);
});

// --- 6. repeated values --------------------------------------------------------

test('conditional fixes: a cell named by two ranges of a rule counts once', () => {
  const twice: CellRange[] = [{ top: 0, left: 0, bottom: 2, right: 0 }, { top: 0, left: 0, bottom: 2, right: 0 }];
  assert.deepEqual(ruled([1, 2, 3], { type: 'duplicateValues' }, undefined, twice), []);
  assert.deepEqual(ruled([1, 2, 3], { type: 'uniqueValues' }, undefined, twice), ['A1', 'A2', 'A3']);
  const partly: CellRange[] = [{ top: 0, left: 0, bottom: 1, right: 0 }, { top: 1, left: 0, bottom: 3, right: 0 }, { top: 0, left: 0, bottom: 3, right: 0 }];
  assert.deepEqual(ruled([1, 2, 3, 1], { type: 'duplicateValues' }, undefined, partly), ['A1', 'A4']);
  assert.deepEqual(ruled([1, 2, 3, 1], { type: 'uniqueValues' }, undefined, partly), ['A2', 'A3']);
  // The same through the file form of the range list.
  const read = book(sheetData(column([1, 2, 3])) + cf('A1:A3 A1:A3 A2', '<cfRule type="duplicateValues" dxfId="0" priority="1"/>'));
  assert.deepEqual(filled(new ConditionalFormats(read.workbook, read.sheet), refs('A', 3)), []);
  // Averages and ranks see each cell once as well: the mean of 0, 5, 9 is 4.67, not (0+5+9+9)/4 = 5.75.
  const overlap: CellRange[] = [{ top: 0, left: 0, bottom: 2, right: 0 }, { top: 2, left: 0, bottom: 2, right: 0 }];
  assert.deepEqual(ruled([0, 5, 9], { type: 'aboveAverage' }, undefined, overlap), ['A2', 'A3']);
  // Two of four values (9, 9) would be the top two if A3 counted twice.
  assert.deepEqual(ruled([0, 5, 9], { type: 'top10', rank: 2 }, undefined, overlap), ['A2', 'A3']);
});

test('conditional fixes: a number and a text that reads as that number are the same value when counting repeats', () => {
  // UNSURE: Excel is believed to count as COUNTIF does, where the criterion 1 matches the text "1".
  const values: Value[] = [1, '1', 2, '2.0', 'x', 'X', '3', true];
  assert.deepEqual(ruled(values, { type: 'duplicateValues' }), ['A1', 'A2', 'A3', 'A4', 'A5', 'A6']);
  assert.deepEqual(ruled(values, { type: 'uniqueValues' }), ['A7', 'A8']);
});

// --- 7. text rules -------------------------------------------------------------

test('conditional fixes: when the text attribute and the formula of a text rule disagree, the formula decides', () => {
  const values: Value[] = ['xabx', 'zz', 'ab', 'zzab'];
  const text = (type: string, formula: string) => `<cfRule type="${type}" dxfId="0" priority="1" operator="${type}" text="ab"><formula>${formula}</formula></cfRule>`;
  const run = (rule: string, engine: boolean) => {
    const { workbook, sheet } = book(sheetData(column(values)) + cf('A1:A4', rule));
    return filled(new ConditionalFormats(workbook, sheet, engine ? withEngine(workbook) : {}), refs('A', 4));
  };
  // The formula looks for "zz"; the attribute still says "ab".
  assert.deepEqual(run(text('containsText', 'NOT(ISERROR(SEARCH("zz",A1)))'), true), ['A2', 'A4']);
  assert.deepEqual(run(text('notContainsText', 'ISERROR(SEARCH("zz",A1))'), true), ['A1', 'A3']);
  assert.deepEqual(run(text('beginsWith', 'LEFT(A1,2)="zz"'), true), ['A2', 'A4']);
  assert.deepEqual(run(text('endsWith', 'RIGHT(A1,2)="zz"'), true), ['A2']);
  // A formula about another cell: every cell follows A1.
  assert.deepEqual(run(text('containsText', 'NOT(ISERROR(SEARCH("ab",$A$1)))'), true), ['A1', 'A2', 'A3', 'A4']);
  // Without an engine the attribute is all there is.
  assert.deepEqual(run(text('containsText', 'NOT(ISERROR(SEARCH("zz",A1)))'), false), ['A1', 'A3', 'A4']);
  assert.deepEqual(run(text('beginsWith', 'LEFT(A1,2)="zz"'), false), ['A3']);
});

test('conditional fixes: a text rule whose formula is the standard one for its text never asks the engine', () => {
  const values: Value[] = ['xabx', 'zz', 'ab', 'zzab', 'say "ab"'];
  let calls = 0;
  const counting = { evaluateText: () => (calls++, false) } as unknown as CalcEngine;
  const run = (type: string, wanted: string, formula: string | undefined) => {
    const { workbook, sheet } = book(sheetData(column(values)));
    sheet.conditionalFormats = [{ ranges: [{ top: 0, left: 0, bottom: 4, right: 0 }], rules: [{ ...BASE, type, text: wanted, formulas: formula === undefined ? [] : [formula] }] }];
    return filled(new ConditionalFormats(workbook, sheet, { engine: () => counting }), refs('A', 5));
  };
  assert.deepEqual(run('containsText', 'ab', 'NOT(ISERROR(SEARCH("ab",A1)))'), ['A1', 'A3', 'A4', 'A5']);
  assert.deepEqual(run('notContainsText', 'ab', 'ISERROR(SEARCH("ab",A1))'), ['A2']);
  assert.deepEqual(run('beginsWith', 'ab', 'LEFT(A1,2)="ab"'), ['A3']);
  assert.deepEqual(run('beginsWith', 'ab', 'LEFT(A1,LEN("ab"))="ab"'), ['A3']);
  assert.deepEqual(run('endsWith', 'ab', 'RIGHT(A1,2)="ab"'), ['A3', 'A4']);
  assert.deepEqual(run('endsWith', 'ab', 'RIGHT(A1,LEN("ab"))="ab"'), ['A3', 'A4']);
  // A quote inside the text is doubled in the formula.
  assert.deepEqual(run('containsText', '"ab"', 'NOT(ISERROR(SEARCH("""ab""",A1)))'), ['A5']);
  // No formula at all: nothing contradicts the attribute.
  assert.deepEqual(run('containsText', 'ab', undefined), ['A1', 'A3', 'A4', 'A5']);
  assert.equal(calls, 0);
  // Anything else goes to the engine, once per cell.
  assert.deepEqual(run('beginsWith', 'ab', 'LEFT(A1,3)="ab"'), []);
  assert.equal(calls, 5);
});

test('conditional fixes: text rules fold case without a locale', () => {
  // UNSURE: Excel's SEARCH folds case with a simple, locale-independent mapping. Here two texts are
  // equal when either their upper-case or their lower-case forms are, letter by letter: i = I = ı
  // (all three are I in upper case), while İ and ß, which have no single-letter counterpart, only
  // meet themselves.
  const contains = (wanted: string, values: string[]) => ruled(values, { type: 'containsText', text: wanted });
  assert.deepEqual(contains('i', ['İstanbul', 'ISPARTA', 'ılık', 'abc']), ['A2', 'A3']);
  assert.deepEqual(contains('İ', ['istanbul', 'İstanbul', 'ISPARTA']), ['A2']);
  assert.deepEqual(contains('STRASSE', ['strasse', 'straße']), ['A1']);
  assert.deepEqual(contains('ß', ['STRASSE', 'Straße']), ['A2']);
  assert.deepEqual(ruled(['ırmak', 'Irmak', 'irmak', 'İrmak'], { type: 'beginsWith', text: 'I' }), ['A1', 'A2', 'A3']);
  assert.deepEqual(ruled(['ırmak', 'Irmak', 'irmak', 'İrmak'], { type: 'beginsWith', text: 'İ' }), ['A4']);
  assert.deepEqual(ruled(['kapı', 'KAPI', 'kapi', 'KAPİ'], { type: 'endsWith', text: 'pı' }), ['A1', 'A2', 'A3']);
  // Wildcards go through the same folding.
  assert.deepEqual(contains('k?p*I', ['kapı', 'KOPARTI', 'kap']), ['A1', 'A2']);
});

// --- 8. the anchor of a rule ---------------------------------------------------

test('conditional fixes: a relative reference that leaves the sheet comes back in at the other edge', () => {
  // "The cell above is 5": for A1 that is the last row of the sheet.
  const above = book(sheetData(column([5, 5, 1, 0])) + cf('A1:A4', expression('A1048576=5')));
  assert.deepEqual(filled(new ConditionalFormats(above.workbook, above.sheet, withEngine(above.workbook)), refs('A', 4)), ['A2', 'A3']);
  setValue(above.sheet, `A${MAX_ROWS}`, 5);
  assert.deepEqual(filled(new ConditionalFormats(above.workbook, above.sheet, withEngine(above.workbook)), refs('A', 4)), ['A1', 'A2', 'A3']);
  // "The cell to the left is 5": for column A that is the last column.
  const left = book(sheetData([[5, 5, 1, 0]]) + cf('A1:D1', expression('XFD1=5')));
  assert.deepEqual(filled(new ConditionalFormats(left.workbook, left.sheet, withEngine(left.workbook)), ['A1', 'B1', 'C1', 'D1']), ['B1', 'C1']);
  // Downwards as well: a rule on the last rows that looks one row further.
  const below = book(sheetData([[7]]) + cf(`A${MAX_ROWS - 1}:A${MAX_ROWS}`, expression(`A${MAX_ROWS}=7`)));
  assert.deepEqual(filled(new ConditionalFormats(below.workbook, below.sheet, withEngine(below.workbook)), [`A${MAX_ROWS - 1}`, `A${MAX_ROWS}`]), [`A${MAX_ROWS}`]);
  // An engine without evaluateShifted (the older interface) gets the wrapped text.
  const asked: string[] = [];
  const older = { evaluateText: (text: string) => (asked.push(text), true) } as unknown as CalcEngine;
  const formats = new ConditionalFormats(above.workbook, above.sheet, { engine: () => older });
  formats.resolve(0, 0);
  formats.resolve(1, 0);
  assert.deepEqual(asked, ['A1048576=5', 'A1=5']);
});

test('conditional fixes: formulas are written for the top-left cell of the first range, wherever it lies', () => {
  // The first range is not the leftmost one: C1 is the anchor, and A1 looks at itself through it.
  const { workbook, sheet } = book(sheetData([[1, 0, -1], [-1, 0, 1], [1, 0, -1]]) + cf('C1:C3 A1:A3', expression('C1&gt;0')));
  assert.deepEqual(filled(new ConditionalFormats(workbook, sheet, withEngine(workbook)), ['A1', 'A2', 'A3', 'C1', 'C2', 'C3']), ['A1', 'A3', 'C2']);
});

// --- 9. batches ----------------------------------------------------------------

test('conditional fixes: formulas go to the engine unshifted, inside a batch', () => {
  const { workbook, sheet } = book(sheetData([[1, 2], [3, 4]]) + cf('B1:B2', expression('A1&gt;$A$1')));
  const log: string[] = [];
  let depth = 0;
  const engine = {
    batch<T>(run: () => T): T {
      depth++;
      log.push('batch');
      try {
        return run();
      } finally {
        depth--;
      }
    },
    evaluateShifted(text: string, anchor: { sheet: number; row: number; col: number }, cell: { row: number; col: number }) {
      log.push(`${depth}:${text}@${anchor.sheet},${anchor.row},${anchor.col}->${cell.row},${cell.col}`);
      return cell.row === 1;
    },
    evaluateText() {
      throw new Error('not used');
    },
  } as unknown as CalcEngine;
  const formats = new ConditionalFormats(workbook, sheet, { engine: () => engine });
  assert.deepEqual(filled(formats, ['B1', 'B2']), ['B2']);
  assert.deepEqual(log, ['batch', '1:A1>$A$1@0,0,1->0,1', 'batch', '1:A1>$A$1@0,0,1->1,1']);
  // A whole frame as one batch: the batches of the cells nest in it.
  log.length = 0;
  formats.invalidate();
  assert.deepEqual(formats.batch(() => filled(formats, ['B1', 'B2'])), ['B2']);
  assert.equal(log.filter((entry) => entry.charAt(0) === '2').length, 2);
  // Without an engine, or with one that has no batches, batch() just runs.
  assert.equal(new ConditionalFormats(workbook, sheet).batch(() => 7), 7);
  // The real engine gives the same answers both ways.
  const real = new ConditionalFormats(workbook, sheet, withEngine(workbook));
  assert.deepEqual(real.batch(() => filled(real, ['B1', 'B2'])), ['B2']);
});

test('conditional fixes: a rule formula follows rows inserted into another sheet', () => {
  const { workbook, sheet } = book(sheetData([[1], [1]]) + cf('A1:A2', expression('Other!$A$2&gt;0')) + cf('B1', expression('Other!A2&gt;0')));
  const other = createSheet({ name: 'Other', sheetId: 2, state: 'visible', kind: 'worksheet', part: '/xl/worksheets/sheet2.xml' });
  setValue(other, 'A2', 1);
  workbook.sheets.push(other);
  const formats = new ConditionalFormats(workbook, sheet, withEngine(workbook));
  assert.deepEqual(filled(formats, ['A1', 'A2', 'B1']), ['A1', 'A2', 'B1']);
  // A row inserted above row 1 of the other sheet, as the operation leaves the model.
  const row = other.rows.get(1)!;
  other.rows.clear();
  other.rows.set(2, row);
  other.structuralEdits = [{ kind: 'insert', axis: 'row', at: 0, count: 1 }];
  formats.invalidate();
  // The anchored reference follows its cell to A3. The relative one is "this far from the rule's
  // cell" and stays A2, which is empty now.
  assert.deepEqual(filled(formats, ['A1', 'A2', 'B1']), ['A1', 'A2']);
  // Without the edit in the log the formulas name row 2 again, which is empty.
  other.structuralEdits = undefined;
  sheet.structuralEdits = [];
  formats.invalidate();
  assert.deepEqual(filled(formats, ['A1', 'A2', 'B1']), []);
});

// --- 10. what the first suite let through -----------------------------------------

test('conditional fixes: the last seven days are seven', () => {
  const days = [45001, 45000, 44995, 44994, 44993, 44992];
  const period = (timePeriod: string) => ruled(days, { type: 'timePeriod', timePeriod }, () => ({ now: () => NOW }));
  // Today and the six days before it.
  assert.deepEqual(period('last7Days'), ['A2', 'A3', 'A4']);
});

test('conditional fixes: time periods in the 1904 date system', () => {
  // The same Wednesday 15 March 2023 is 43538 there; its week starts on Sunday 12 (43535).
  const shift = 1462;
  const days = [44996, 44997, 45003, 45004, 44985, 44986, 44957, 44958].map((day) => day - shift);
  const { workbook, sheet } = book(sheetData(column(days)));
  workbook.date1904 = true;
  const period = (timePeriod: string) => {
    sheet.conditionalFormats = [{ ranges: [{ top: 0, left: 0, bottom: 7, right: 0 }], rules: [{ ...BASE, type: 'timePeriod', timePeriod }] }];
    return filled(new ConditionalFormats(workbook, sheet, { now: () => NOW - shift }), refs('A', 8));
  };
  assert.deepEqual(period('thisWeek'), ['A2', 'A3']);
  assert.deepEqual(period('lastWeek'), ['A1']);
  assert.deepEqual(period('nextWeek'), ['A4']);
  // February 2023 is 44958..44985.
  assert.deepEqual(period('lastMonth'), ['A5', 'A8']);
  assert.deepEqual(period('thisMonth'), ['A1', 'A2', 'A3', 'A4', 'A6']);
});

test('conditional fixes: standard deviations around the average are those of the whole range', () => {
  // UNSURE: Excel is believed to use the population deviation (STDEV.P) here, not the sample one.
  // 2, 4, 4, 4, 5, 5, 7, 9: mean 5, population deviation exactly 2 (sample: 2.138).
  const classic = [2, 4, 4, 4, 5, 5, 7, 9];
  // The limit itself (7, and 3 below) never counts, with or without equalAverage.
  assert.deepEqual(ruled(classic, { type: 'aboveAverage', stdDev: 1 }), ['A8']);
  assert.deepEqual(ruled(classic, { type: 'aboveAverage', stdDev: 1, equalAverage: true }), ['A8']);
  assert.deepEqual(ruled(classic, { type: 'aboveAverage', stdDev: 1, aboveAverage: false }), ['A1']);
  assert.deepEqual(ruled(classic, { type: 'aboveAverage', stdDev: 1, aboveAverage: false, equalAverage: true }), ['A1']);
  assert.deepEqual(ruled(classic, { type: 'aboveAverage', stdDev: 2 }), []);
  // With 7.1 for the 7 the population limit is 7.025 and the sample limit 7.164: 7.1 lies between.
  assert.deepEqual(ruled([2, 4, 4, 4, 5, 5, 7.1, 9], { type: 'aboveAverage', stdDev: 1 }), ['A7', 'A8']);
  // Without deviations equalAverage takes the average in: the mean of 1, 2, 3 is 2.
  assert.deepEqual(ruled([1, 2, 3], { type: 'aboveAverage', equalAverage: true }), ['A2', 'A3']);
  assert.deepEqual(ruled([1, 2, 3], { type: 'aboveAverage' }), ['A3']);
  assert.deepEqual(ruled([1, 2, 3], { type: 'aboveAverage', aboveAverage: false, equalAverage: true }), ['A1', 'A2']);
  assert.deepEqual(ruled([1, 2, 3], { type: 'aboveAverage', aboveAverage: false }), ['A1']);
  // stdDev="0" is no deviation at all.
  assert.deepEqual(ruled([1, 2, 3], { type: 'aboveAverage', stdDev: 0, equalAverage: true }), ['A2', 'A3']);
});

test('conditional fixes: a percent threshold is a share of the distance from the smallest to the largest value', () => {
  const { workbook, sheet } = book(
    sheetData(column([100, 150, 200, 125])) +
      cf('A1:A4', '<cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="percent" val="50"/><cfvo type="max"/><color rgb="FF000000"/><color rgb="FF646464"/><color rgb="FFC8C8C8"/></colorScale></cfRule>') +
      cf('A1:A4', '<cfRule type="iconSet" priority="2"><iconSet iconSet="3Arrows"><cfvo type="percent" val="0"/><cfvo type="percent" val="30"/><cfvo type="percent" val="60"/></iconSet></cfRule>'),
  );
  const formats = new ConditionalFormats(workbook, sheet);
  // 50 percent is 150: the middle colour exactly; 125 is half way to it.
  assert.deepEqual(refs('A', 4).map((ref) => at(formats, ref)!.fill), ['#000000', '#646464', '#C8C8C8', '#323232']);
  // 30 percent is 130 and 60 percent is 160.
  assert.deepEqual(refs('A', 4).map((ref) => at(formats, ref)!.icon!.index), [0, 1, 2, 0]);
});

test('conditional fixes: invalidate() forgets what day it is', () => {
  const { workbook, sheet } = book(sheetData(column([45000, 45001])));
  let now = 45000.25;
  sheet.conditionalFormats = [{ ranges: [{ top: 0, left: 0, bottom: 1, right: 0 }], rules: [{ ...BASE, type: 'timePeriod', timePeriod: 'today' }] }];
  const formats = new ConditionalFormats(workbook, sheet, { now: () => now });
  assert.deepEqual(filled(formats, ['A1', 'A2']), ['A1']);
  now = 45001.25;
  // Until told otherwise the day is the one first asked for.
  assert.deepEqual(filled(formats, ['A1', 'A2']), ['A1']);
  formats.invalidate();
  assert.deepEqual(filled(formats, ['A1', 'A2']), ['A2']);
});

test('conditional fixes: the extension of a data bar - whose colour, whose showValue, and a border that is switched off', () => {
  const read = (main: string, mainColor: string, extension: string, inner: string) =>
    book(
      sheetData(column([1, 2])) +
        cf('A1:A2', `<cfRule type="dataBar" priority="1"><dataBar ${main}><cfvo type="min"/><cfvo type="max"/>${mainColor}</dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"><x14:id>{1}</x14:id></ext></extLst></cfRule>`) +
        `<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="dataBar" id="{1}"><x14:dataBar ${extension}><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/>${inner}</x14:dataBar></x14:cfRule><xm:sqref>A1:A2</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>`,
    ).sheet.conditionalFormats![0].rules[0].dataBar!;
  // Both state a colour: the main rule's.
  let bar = read('', '<color rgb="FF111111"/>', 'border="0"', '<x14:fillColor rgb="FF222222"/><x14:borderColor rgb="FF333333"/>');
  assert.deepEqual(bar.color, { rgb: 'FF111111' });
  assert.equal(bar.showValue, true);
  assert.equal(bar.extended, true);
  // border="0": the colour of the border is not taken, though it is there.
  assert.equal(bar.borderColor, undefined);
  // Only the extension states one.
  bar = read('showValue="0"', '', 'border="1"', '<x14:fillColor rgb="FF222222"/><x14:borderColor rgb="FF333333"/>');
  assert.deepEqual(bar.color, { rgb: 'FF222222' });
  assert.deepEqual(bar.borderColor, { rgb: 'FF333333' });
  // The value is hidden when either of the two says so.
  assert.equal(bar.showValue, false);
  assert.equal(read('', '<color rgb="FF111111"/>', 'showValue="0"', '').showValue, false);
  assert.equal(read('showValue="1"', '<color rgb="FF111111"/>', 'showValue="1"', '').showValue, true);
  // No border attribute at all is no border.
  assert.equal(read('', '<color rgb="FF111111"/>', '', '<x14:borderColor rgb="FF333333"/>').borderColor, undefined);
});

test('conditional fixes: the list of evaluated cells is emptied when it is full, and answers stay right', () => {
  const { workbook, sheet } = book(sheetData([[1]]));
  sheet.conditionalFormats = [{ ranges: [{ top: 0, left: 0, bottom: MAX_ROWS - 1, right: MAX_COLS - 1 }], rules: [{ ...BASE, type: 'containsBlanks' }] }];
  const formats = new ConditionalFormats(workbook, sheet);
  const kept = () => (formats as unknown as { cells: Map<number, unknown> }).cells.size;
  let most = 0;
  let matched = 0;
  for (let i = 0; i < 200001 + 500; i++) {
    if (formats.resolve(i, 3)) matched++;
    if (kept() > most) most = kept();
  }
  assert.equal(matched, 200501);
  assert.equal(most, 200000);
  // The list was emptied once and filled again from there.
  assert.equal(kept(), 501);
  assert.equal(formats.resolve(0, 0), undefined);
  assert.ok(formats.resolve(200400, 3));
});

test('conditional fixes: texts and patterns are cut to the lengths Excel allows', () => {
  // A cell holds 32,767 characters at most; a longer text (from a damaged file) is looked at up to there.
  const long = 'x'.repeat(32767 - 3);
  const values = [long + 'abc', long + 'xabc', 'a'.repeat(255), 'a'.repeat(254)];
  assert.deepEqual(ruled(values, { type: 'containsText', text: 'abc' }), ['A1']);
  assert.deepEqual(ruled(values, { type: 'notContainsText', text: 'abc' }), ['A2', 'A3', 'A4']);
  // A pattern is 255 characters at most: 300 are looked for as their first 255.
  assert.deepEqual(ruled(values, { type: 'containsText', text: 'a'.repeat(300) }), ['A3']);
  assert.deepEqual(ruled(values, { type: 'containsText', text: 'a'.repeat(254) + '?' + 'b'.repeat(40) }), ['A3']);
});

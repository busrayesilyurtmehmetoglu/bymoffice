// Data validations in the file: they are written back as they were read,
// follow moved rows, and are read the way openpyxl reads them (skipped where
// python3 or openpyxl is not installed).

import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { CalcEngine } from '../src/calc/engine';
import { DataValidations } from '../src/format/validation';
import { writeXlsx } from '../src/formats/xlsx/write';
import { ZipArchive, open, rangeName } from '../src/index';
import type { DataValidation, OfficeFile, Workbook } from '../src/index';
import { setCellInput } from '../src/ops/commands';
import { History, type Op } from '../src/ops/ops';
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

const VALIDATIONS =
  '<dataValidations count="4" xr:uid="{V}" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">' +
  '<dataValidation type="whole" operator="notBetween" allowBlank="1" showInputMessage="1" showErrorMessage="1" errorStyle="warning" errorTitle="Whole" error="Not 3 to 5" promptTitle="Number" prompt="Outside 3-5" sqref="A2:A9" xr:uid="{1}"><formula1>3</formula1><formula2>5</formula2></dataValidation>' +
  '<dataValidation type="list" showDropDown="1" showErrorMessage="1" sqref="B2:B9 D2"><formula1>"x,y,z"</formula1></dataValidation>' +
  '<dataValidation type="custom" showErrorMessage="1" imeMode="off" sqref="C2:C9"><formula1>C2&gt;$A$2</formula1></dataValidation>' +
  '<dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="E2:E9"><formula1>$G$2:$G$4</formula1></dataValidation>' +
  '</dataValidations>';

const EXTENSION =
  '<extLst><ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}">' +
  '<x14:dataValidations xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main" count="1">' +
  '<x14:dataValidation type="list" showErrorMessage="1"><x14:formula1><xm:f>Other!$A$1:$A$2</xm:f></x14:formula1><xm:sqref>F2:F9</xm:sqref></x14:dataValidation>' +
  '</x14:dataValidations></ext></extLst>';

const workbook = () =>
  buildXlsx({
    sheets: [
      {
        name: 'Checked',
        xml: worksheetXml(
          '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>n</t></is></c></row><row r="2"><c r="A2"><v>1</v></c><c r="G2" t="inlineStr"><is><t>one</t></is></c></row></sheetData>' + VALIDATIONS + EXTENSION,
        ),
      },
      { name: 'Other', xml: worksheetXml('<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>p</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>q</t></is></c></row></sheetData>') },
    ],
  });

test('validation xlsx: nothing changed - the parts are byte for byte what they were; a cell edit keeps the rules as written', async () => {
  const bytes = await workbook();
  const untouched = await save(await load(bytes));
  assert.deepEqual(await ZipArchive.read(untouched).read('xl/worksheets/sheet1.xml'), await ZipArchive.read(bytes).read('xl/worksheets/sheet1.xml'));

  const doc = await load(bytes);
  run(doc, setCellInput(doc.wb, 0, 2, 0, '7'));
  const sheet = await partText(await save(doc), 'xl/worksheets/sheet1.xml');
  assert.ok(sheet.includes(VALIDATIONS), 'the rules are the source text');
  assert.ok(sheet.includes(EXTENSION), 'the extension list is the source text');
  assert.ok(sheet.includes('<c r="A3"><v>7</v></c>'));
});

test('validation xlsx: after rows and columns moved, the file says what the view applies', async () => {
  const doc = await load(await workbook());
  run(doc, [{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 1 }]);
  run(doc, [{ type: 'insert', sheet: 0, axis: 'col', at: 1, count: 1 }]);
  run(doc, [{ type: 'delete', sheet: 1, axis: 'row', at: 0, count: 1 }]);
  const engine = new CalcEngine(doc.wb);
  const view = new DataValidations(doc.wb, doc.wb.sheets[0], { engine: () => engine });
  const applied = (row: number, col: number) => {
    const rule = view.ruleAt(row, col) as DataValidation;
    return { ranges: rule.ranges.map(rangeName).join(' '), formula1: rule.formula1, formula2: rule.formula2 };
  };
  const saved = await load(await save(doc));
  const written = saved.wb.sheets[0].dataValidations!.map((rule) => ({ ranges: rule.ranges.map(rangeName).join(' '), formula1: rule.formula1, formula2: rule.formula2 }));
  assert.deepEqual(written, [
    { ranges: 'A3:A10', formula1: '3', formula2: '5' },
    { ranges: 'C3:C10 E3', formula1: '"x,y,z"', formula2: undefined },
    { ranges: 'D3:D10', formula1: 'D3>$A$3', formula2: undefined },
    { ranges: 'F3:F10', formula1: '$H$3:$H$5', formula2: undefined },
    { ranges: 'G3:G10', formula1: 'Other!$A$1:$A$1', formula2: undefined },
  ]);
  assert.deepEqual([applied(2, 0), applied(2, 2), applied(2, 3), applied(2, 5), applied(2, 6)], written);
  // What the list on the other sheet offers now.
  assert.deepEqual(view.list(2, 6), { items: ['q'], truncated: false });
});

// --- openpyxl ----------------------------------------------------------------

const hasOpenpyxl = () => spawnSync('python3', ['-c', 'import openpyxl'], { encoding: 'utf8' }).status === 0;

function python(script: string, input: Uint8Array | undefined, output: boolean): { stdout: string; file?: Uint8Array } {
  const folder = join(dirname(fileURLToPath(import.meta.url)), '..', `tmp-validation-openpyxl-${process.pid}`);
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
import json, sys
import openpyxl
wb = openpyxl.load_workbook(sys.argv[1])
ws = wb.worksheets[0]
out = []
for dv in ws.data_validations.dataValidation:
    out.append({
        'ranges': str(dv.sqref), 'type': dv.type or 'none', 'operator': dv.operator or 'between',
        'formula1': dv.formula1, 'formula2': dv.formula2,
        'allowBlank': bool(dv.allowBlank), 'hideDropDown': bool(dv.showDropDown),
        'showInputMessage': bool(dv.showInputMessage), 'showErrorMessage': bool(dv.showErrorMessage),
        'errorStyle': dv.errorStyle or 'stop', 'errorTitle': dv.errorTitle, 'error': dv.error,
        'promptTitle': dv.promptTitle, 'prompt': dv.prompt,
    })
print(json.dumps(out))
`;

const orNull = <T>(value: T | undefined): T | null => (value === undefined ? null : value);

const ours = (rules: DataValidation[]) =>
  rules
    .filter((rule) => !rule.extended)
    .map((rule) => ({
      ranges: rule.ranges.map(rangeName).join(' '),
      type: rule.type,
      operator: rule.operator,
      formula1: orNull(rule.formula1),
      formula2: orNull(rule.formula2),
      allowBlank: rule.allowBlank,
      hideDropDown: rule.hideDropDown,
      showInputMessage: rule.showInputMessage,
      showErrorMessage: rule.showErrorMessage,
      errorStyle: rule.errorStyle,
      errorTitle: orNull(rule.errorTitle),
      error: orNull(rule.error),
      promptTitle: orNull(rule.promptTitle),
      prompt: orNull(rule.prompt),
    }));

test('validation xlsx: the rules of a synthetic workbook read the same as openpyxl reads them, also after an edit and a save', async (t) => {
  if (!hasOpenpyxl()) {
    t.skip('python3 with openpyxl is not available');
    return;
  }
  const bytes = await workbook();
  const doc = await load(bytes);
  assert.equal(doc.wb.sheets[0].dataValidations!.length, 5);
  assert.deepEqual(ours(doc.wb.sheets[0].dataValidations!), JSON.parse(python(PYTHON_READ, bytes, false).stdout));

  // Rows inserted and a cell written: openpyxl reads from the saved file what the reopened model holds.
  run(doc, [{ type: 'insert', sheet: 0, axis: 'row', at: 1, count: 3 }]);
  run(doc, setCellInput(doc.wb, 0, 0, 1, 'edited'));
  const saved = await save(doc);
  const reopened = await load(saved);
  const theirs = JSON.parse(python(PYTHON_READ, saved, false).stdout);
  assert.deepEqual(ours(reopened.wb.sheets[0].dataValidations!), theirs);
  assert.deepEqual(theirs.map((rule: { ranges: string }) => rule.ranges), ['A5:A12', 'B5:B12 D5', 'C5:C12', 'E5:E12']);
  assert.equal(theirs[2].formula1, 'C5>$A$5');
});

const PYTHON_WRITE = `
import sys
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
wb = Workbook()
ws = wb.active
ws.title = 'Made by openpyxl'
for i, v in enumerate(['red', 'green', 'blue'], start=1):
    ws.cell(row=i, column=8, value=v)
a = DataValidation(type='list', formula1='"Dog,Cat,Bat"', allow_blank=True, showDropDown=True)
a.error = 'Not in the list'; a.errorTitle = 'Invalid'; a.prompt = 'Choose an animal'; a.promptTitle = 'Animal'
a.showInputMessage = True; a.showErrorMessage = True
a.add('A1:A5')
b = DataValidation(type='whole', operator='between', formula1='1', formula2='10', errorStyle='information', showErrorMessage=True)
b.add('B1:B5'); b.add('D1')
c = DataValidation(type='list', formula1='$H$1:$H$3', showErrorMessage=True)
c.add('C1:C1048576')
d = DataValidation(type='custom', formula1='LEN(E1)<4', showErrorMessage=True, errorStyle='warning')
d.add('E1:E5')
e = DataValidation(type='decimal', operator='lessThan', formula1='0.5')
e.add('F1')
for dv in (a, b, c, d, e):
    ws.add_data_validation(dv)
wb.save(sys.argv[1])
`;

test('validation xlsx: rules written by openpyxl are read and applied', async (t) => {
  if (!hasOpenpyxl()) {
    t.skip('python3 with openpyxl is not available');
    return;
  }
  const made = python(PYTHON_WRITE, undefined, true).file as Uint8Array;
  const doc = await load(made);
  const rules = doc.wb.sheets[0].dataValidations!;
  assert.deepEqual(ours(rules), JSON.parse(python(PYTHON_READ, made, false).stdout));
  assert.equal(rules.length, 5);
  assert.equal(rules[0].hideDropDown, true);
  assert.deepEqual(rules[1].ranges.map(rangeName), ['B1:B5', 'D1']);
  assert.equal(rangeName(rules[2].ranges[0]), 'C1:C1048576');

  const engine = new CalcEngine(doc.wb);
  const view = new DataValidations(doc.wb, doc.wb.sheets[0], { engine: () => engine });
  const text = (value: string) => ({ kind: 'string' as const, value, style: 0 });
  const number = (value: number) => ({ kind: 'number' as const, value, style: 0 });
  assert.deepEqual(view.check(0, 0, text('Emu'), 'Emu'), { ok: false, style: 'stop', title: 'Invalid', message: 'Not in the list' });
  assert.deepEqual(view.check(0, 0, text('Cat'), 'Cat'), { ok: true });
  assert.equal(view.hasList(0, 0), false);
  assert.deepEqual(view.prompt(0, 0), { title: 'Animal', text: 'Choose an animal' });
  assert.deepEqual(view.check(0, 3, number(11), '11'), { ok: false, style: 'information' });
  assert.deepEqual(view.check(4, 1, number(10), '10'), { ok: true });
  assert.deepEqual(view.list(500000, 2), { items: ['red', 'green', 'blue'], truncated: false });
  assert.equal(view.hasList(500000, 2), true);
  assert.equal(view.check(9, 2, text('BLUE'), 'BLUE').ok, true);
  assert.equal(view.check(9, 2, text('pink'), 'pink').ok, false);
  // No error alert: nothing is refused.
  assert.equal(view.check(0, 5, number(3), '3').ok, true);

  // Unchanged, the sheet goes back byte for byte.
  const untouched = await save(await load(made));
  assert.deepEqual(await ZipArchive.read(untouched).read('xl/worksheets/sheet1.xml'), await ZipArchive.read(made).read('xl/worksheets/sheet1.xml'));
});

// Merging cells: what a merge keeps and what it reports it discards, "merge and center", what it
// refuses, that undo gives every value and format back, and what the writer makes of it -
// checked again by openpyxl where python3 has it.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { getCell, open, resolveCellStyle } from '../src/index';
import { setCellInput } from '../src/ops/commands';
import { mergeProblem, mergeRange, mergeReport, unmergeRange } from '../src/ops/format';
import { setComment } from '../src/ops/notes';
import { apply, hasPython, load, packageProblems, partText, runPython, sameBytes, save } from './edit-commands-fixtures';
import { buildXlsx, worksheetXml } from './fixtures';

const range = (top: number, left: number, bottom: number, right: number) => ({ top, left, bottom, right });

const STYLES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>' +
  '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
  '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>' +
  '</styleSheet>';

/**
 * A1 "head" (bold), B1 2, C1 =B1*2, A2 4 (bold), B2 empty but bold; a shared formula E1:E3 (=D1+1);
 * an array formula G1:G2; and a sheet beside it.
 */
const book = () =>
  buildXlsx({
    styles: STYLES,
    sharedStrings: '<si><t>head</t></si>',
    sheets: [
      {
        name: 'Data',
        xml: worksheetXml(
          '<sheetData>' +
            '<row r="1"><c r="A1" s="1" t="s"><v>0</v></c><c r="B1"><v>2</v></c><c r="C1"><f>B1*2</f><v>4</v></c><c r="D1"><v>1</v></c>' +
            '<c r="E1"><f t="shared" ref="E1:E3" si="0">D1+1</f><v>2</v></c><c r="G1"><f t="array" ref="G1:G2">D1:D2*2</f><v>2</v></c></row>' +
            '<row r="2"><c r="A2" s="1"><v>4</v></c><c r="B2" s="1"/><c r="D2"><v>5</v></c><c r="E2"><f t="shared" si="0"/><v>6</v></c><c r="G2"><v>10</v></c></row>' +
            '<row r="3"><c r="E3"><f t="shared" si="0"/><v>1</v></c></row>' +
            '</sheetData>',
        ),
      },
      { name: 'Other', xml: worksheetXml('<sheetData/>') },
    ],
  });

const valueAt = (wb: { sheets: { rows: Map<number, unknown> }[] }, sheet: number, row: number, col: number) => {
  const cell = getCell(wb.sheets[sheet] as never, row, col);
  return cell ? cell.value : undefined;
};

test('merge: only the top-left value stays, the others are reported and go, every cell keeps its format, and undo brings all back', async () => {
  const doc = await load(await book());
  const sheet = doc.wb.sheets[0];
  const report = mergeReport(doc.wb, 0, range(0, 0, 1, 2));
  assert.deepEqual(report, {
    losses: [
      { row: 0, col: 1, text: '2' },
      { row: 0, col: 2, text: '=B1*2' },
      { row: 1, col: 0, text: '4' },
    ],
    count: 3,
    notes: 0,
  });
  apply(doc, mergeRange(doc.wb, 0, range(0, 0, 1, 2)));
  assert.deepEqual(sheet.merges, [range(0, 0, 1, 2)]);
  assert.equal(valueAt(doc.wb, 0, 0, 0), 'head');
  assert.deepEqual([getCell(sheet, 0, 1), getCell(sheet, 0, 2)], [undefined, undefined]);
  // A2 was bold: it is emptied, and stays bold (an empty cell with a format of its own).
  assert.deepEqual(getCell(sheet, 1, 0), { kind: 'blank', value: null, style: 1 });
  assert.equal(getCell(sheet, 1, 1)!.style, 1);
  assert.deepEqual(mergeReport(doc.wb, 0, range(0, 0, 1, 2)).count, 0, 'nothing left to lose');
  doc.history.undo();
  assert.deepEqual(sheet.merges, []);
  assert.deepEqual([valueAt(doc.wb, 0, 0, 1), getCell(sheet, 0, 2)!.formula!.text, valueAt(doc.wb, 0, 1, 0), getCell(sheet, 1, 0)!.style], [2, 'B1*2', 4, 1]);
  assert.ok(sameBytes(await save(doc), doc.bytes), 'merge, undo, save: the file as it was');
});

test('merge and center: the cells are centred in the same step, the emptied ones too, and undo takes both back', async () => {
  const doc = await load(await book());
  const sheet = doc.wb.sheets[0];
  const before = doc.wb.styles.cellXfs.length;
  apply(doc, mergeRange(doc.wb, 0, range(0, 0, 1, 1), { center: true }));
  assert.equal(doc.history.canUndo, true);
  for (const [row, col] of [[0, 0], [0, 1], [1, 0], [1, 1]]) {
    const cell = getCell(sheet, row, col);
    assert.ok(cell, `${row},${col} has a cell`);
    assert.equal(resolveCellStyle(doc.wb.styles, cell!.style).alignment.horizontal, 'center', `${row},${col}`);
  }
  // Bold stays bold and centred; the value of A1 stays, the others go.
  assert.equal(resolveCellStyle(doc.wb.styles, getCell(sheet, 1, 0)!.style).font.bold, true);
  assert.deepEqual([valueAt(doc.wb, 0, 0, 0), valueAt(doc.wb, 0, 0, 1), valueAt(doc.wb, 0, 1, 0)], ['head', null, null]);
  assert.ok(doc.wb.styles.cellXfs.length > before);
  doc.history.undo();
  assert.deepEqual([valueAt(doc.wb, 0, 0, 1), valueAt(doc.wb, 0, 1, 0), getCell(sheet, 0, 0)!.style, getCell(sheet, 0, 1)!.style], [2, 4, 1, 0]);
  assert.deepEqual(sheet.merges, []);
});

test('merge: a shared formula whose master is emptied leaves the other cells of its group their formulas', async () => {
  const doc = await load(await book());
  const sheet = doc.wb.sheets[0];
  apply(doc, mergeRange(doc.wb, 0, range(0, 3, 0, 4)));
  assert.equal(getCell(sheet, 0, 4), undefined);
  assert.equal(getCell(sheet, 1, 4)!.formula!.text, 'D2+1');
  assert.equal(getCell(sheet, 2, 4)!.formula!.text, 'D3+1');
  doc.history.undo();
  assert.equal(getCell(sheet, 0, 4)!.formula!.text, 'D1+1');
});

test('merge: what is refused, and why', async () => {
  const doc = await load(await book());
  const wb = doc.wb;
  assert.match(mergeProblem(wb, 0, range(0, 6, 0, 7))!, /array formula in G1:G2/);
  assert.match(mergeProblem(wb, 0, range(1, 5, 4, 6))!, /array formula/);
  assert.throws(() => mergeRange(wb, 0, range(0, 5, 1, 6)), /array formula/);
  assert.throws(() => mergeRange(wb, 0, range(0, 0, 1048575, 1)), /Whole rows or columns/);
  assert.throws(() => mergeRange(wb, 0, range(0, 0, 0, 16383)), /Whole rows or columns/);
  assert.throws(() => mergeRange(wb, 0, range(2, 0, 1, 0)), /not valid/);
  assert.throws(() => mergeRange(wb, 9, range(0, 0, 1, 1)), /No sheet/);
  assert.deepEqual(mergeRange(wb, 0, range(3, 3, 3, 3)), [], 'one cell: nothing to merge');
  wb.sheets[1].anchored = ['table'];
  assert.match(mergeProblem(wb, 1, range(0, 0, 1, 1))!, /table/);
  // A filter that hides a row of the range.
  wb.sheets[1].anchored = undefined;
  apply(doc, setCellInput(wb, 1, 0, 0, 'h'));
  wb.sheets[1].autoFilter = { range: range(0, 0, 5, 0), columns: [{ colId: 0, kind: 'values', values: ['x'], blank: false } as never] };
  wb.sheets[1].rows.set(2, { customHeight: false, hidden: true, outlineLevel: 0, collapsed: false, cells: [] });
  assert.match(mergeProblem(wb, 1, range(1, 0, 3, 1))!, /filter hides/);
  assert.equal(mergeProblem(wb, 1, range(1, 0, 1, 1)), undefined, 'one row: nothing hidden in it');
  assert.equal(mergeProblem(wb, 1, range(3, 0, 4, 1)), undefined);
});

test('merge: a note on a cell that is merged under another is counted, and kept', async () => {
  const doc = await load(await book());
  apply(doc, setComment(doc.wb, 0, 1, 1, 'under'));
  const report = mergeReport(doc.wb, 0, range(0, 0, 1, 1));
  assert.equal(report.notes, 1);
  apply(doc, mergeRange(doc.wb, 0, range(0, 0, 1, 1)));
  assert.equal((doc.wb.sheets[0].comments || []).length, 1);
});

test('merge: the report lists at most a thousand cells, and counts them all', async () => {
  const doc = await load(await buildXlsx({ sheets: [{ name: 'S', xml: worksheetXml('<sheetData/>') }] }));
  const sheet = doc.wb.sheets[0];
  for (let r = 0; r < 60; r++) {
    const cells = [];
    for (let c = 0; c < 30; c++) cells.push({ kind: 'number' as const, value: r * 30 + c, style: 0 });
    sheet.rows.set(r, { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells });
  }
  const report = mergeReport(doc.wb, 0, range(0, 0, 59, 29));
  assert.equal(report.count, 1799);
  assert.equal(report.losses.length, 1000);
});

test('merge: the file says what the model says, unmerge takes the merge away again', async () => {
  const doc = await load(await book());
  apply(doc, mergeRange(doc.wb, 0, range(4, 0, 5, 2), { center: true }));
  apply(doc, mergeRange(doc.wb, 0, range(7, 1, 7, 3)));
  let saved = await save(doc);
  assert.match(await partText(saved, 'xl/worksheets/sheet1.xml'), /<mergeCells count="2"><mergeCell ref="A5:C6"\/><mergeCell ref="B8:D8"\/><\/mergeCells>/);
  const again = await (await open(saved)).workbook();
  assert.deepEqual(again.sheets[0].merges, [range(4, 0, 5, 2), range(7, 1, 7, 3)]);
  apply(doc, unmergeRange(doc.wb, 0, range(7, 2, 7, 2)));
  saved = await save(doc);
  assert.match(await partText(saved, 'xl/worksheets/sheet1.xml'), /<mergeCells count="1"><mergeCell ref="A5:C6"\/><\/mergeCells>/);
});

const OPENPYXL = String.raw`
import json, os, sys
from openpyxl import load_workbook
folder = sys.argv[1]
out = {}
for name in sorted(os.listdir(folder)):
    wb = load_workbook(os.path.join(folder, name))
    ws = wb.worksheets[0]
    out[name] = {
        'merged': sorted(str(r) for r in ws.merged_cells.ranges),
        'A1': ws['A1'].value,
        'A1h': ws['A1'].alignment.horizontal,
        'B1': ws['B1'].value,
        'C1': ws['C1'].value,
    }
print(json.dumps(out))
`;

test('merge: openpyxl reads the merged cells, their values and the centring, and the package is sound', async (t) => {
  if (!hasPython(['openpyxl'])) {
    t.skip('python3 with openpyxl is not available');
    return;
  }
  const doc = await load(await book());
  apply(doc, mergeRange(doc.wb, 0, range(0, 0, 0, 2), { center: true }));
  apply(doc, mergeRange(doc.wb, 0, range(4, 1, 6, 3)));
  const files = { 'merged.xlsx': await save(doc) };
  const read = runPython<{ [name: string]: { merged: string[]; A1: unknown; A1h: string | null; B1: unknown; C1: unknown } }>(files, OPENPYXL);
  assert.deepEqual(read['merged.xlsx'], { merged: ['A1:C1', 'B5:D7'], A1: 'head', A1h: 'center', B1: null, C1: null });
  assert.deepEqual(packageProblems(files), []);
});

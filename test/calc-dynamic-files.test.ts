// Spills next to what a file holds: formulas the engine cannot calculate, cells that are not empty though
// they look so, rows of the file, `ref`s of a damaged file, spills of the file larger than a formula's share.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CalcEngine, type CalcEngineOptions } from '../src/calc/engine';
import { getCell, open, writeXlsx, type Workbook } from '../src/index';
import { setCellInput } from '../src/ops/commands';
import { History } from '../src/ops/ops';
import { ZipArchive } from '../src/zip';
import { createBlank } from '../src/formats/blank';
import { buildXlsx, worksheetXml } from './fixtures';

const METADATA =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">' +
  '<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>' +
  '<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>' +
  '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>';

const decoder = new TextDecoder();

async function opened(sheetData: string, options: CalcEngineOptions = {}, sharedStrings?: string) {
  const bytes = await buildXlsx({ sheets: [{ name: 'S', xml: worksheetXml(`<sheetData>${sheetData}</sheetData>`) }], metadata: METADATA, sharedStrings });
  const file = await open(bytes);
  const workbook = await file.workbook();
  const history = new History(workbook);
  const engine = new CalcEngine(workbook, options);
  engine.recalculate(engine.uncalculated());
  const type = (row: number, col: number, text: string) => {
    history.apply(setCellInput(workbook, 0, row, col, text), undefined, undefined);
    engine.recalculate([{ sheet: 0, row, col }]);
  };
  const undo = (row: number, col: number) => {
    history.undo();
    engine.recalculate([{ sheet: 0, row, col }]);
  };
  const sheetXml = async () => {
    await writeXlsx(file.package!, workbook);
    const saved = await file.package!.save();
    return decoder.decode(await ZipArchive.read(saved).read('xl/worksheets/sheet1.xml'));
  };
  const original = decoder.decode(await ZipArchive.read(bytes).read('xl/worksheets/sheet1.xml'));
  return { workbook, engine, type, undo, sheetXml, original };
}

const value = (workbook: Workbook, row: number, col: number) => {
  const cell = getCell(workbook.sheets[0], row, col);
  return cell ? cell.value : undefined;
};

test('the cells of an array formula the engine cannot calculate are in the way of a spill (Ö1)', async () => {
  for (const formula of ['_xlfn.NOSUCHFN(1)', 'Sheet9!A1:A3', '[1]S!A1:A3']) {
    const t = await opened(`<row r="1"><c r="B1" cm="1"><f t="array" ref="B1:B3">${formula}</f><v>10</v></c></row><row r="2"><c r="B2"><v>20</v></c></row><row r="3"><c r="B3"><v>30</v></c></row>`);
    t.type(1, 0, '=SEQUENCE(1,2)');
    assert.equal(`${value(t.workbook, 1, 0)}`, '#SPILL!', formula);
    assert.deepEqual([value(t.workbook, 0, 1), value(t.workbook, 1, 1), value(t.workbook, 2, 1)], [10, 20, 30], formula);
    // What reads them reads the values of the file.
    t.type(4, 0, '=SUM(B1:B3)');
    assert.equal(value(t.workbook, 4, 0), 60, formula);
  }
});

test('only an empty cell is open to a spill: an empty shared string, value metadata, an empty inline string, a phonetic mark are not (Ö2)', async () => {
  for (const b2 of ['<c r="B2" t="s"><v>0</v></c>', '<c r="B2" vm="1"/>', '<c r="B2" t="inlineStr"><is><t></t></is></c>', '<c r="B2" t="s" ph="1"><v>0</v></c>']) {
    const t = await opened(`<row r="1"><c r="A1"><v>1</v></c></row><row r="2">${b2}</row>`, {}, '<si><t></t></si>');
    t.type(0, 1, '=SEQUENCE(3)');
    assert.equal(`${value(t.workbook, 0, 1)}`, '#SPILL!', b2);
    t.undo(0, 1);
    assert.equal(await t.sheetXml(), t.original, b2);
  }
});

test('a row of the file stays when a spill that went through it goes, though it holds nothing (Ö2)', async () => {
  const t = await opened('<row r="1"><c r="A1"><v>1</v></c></row><row r="2"/><row r="3" ht="20" customHeight="1"/>');
  t.type(0, 1, '=SEQUENCE(4)');
  assert.equal(value(t.workbook, 3, 1), 4);
  t.undo(0, 1);
  assert.equal(await t.sheetXml(), t.original);
});

test('the `ref` of a damaged file does not make the engine fill millions of cells (Ö4)', async () => {
  for (const legacy of [false, true]) {
    const count = 300;
    let cells = '';
    const name = (k: number) => (k < 26 ? '' : String.fromCharCode(64 + Math.floor(k / 26))) + String.fromCharCode(65 + (k % 26));
    for (let k = 0; k < count; k++) {
      cells += legacy ? `<c r="${name(k)}1"><f t="array" ref="${name(k)}1:${name(k)}1000000">1</f></c>` : `<c r="${name(k)}1" cm="1"><f t="array" ref="${name(k)}1:${name(k)}1000000">_xlfn.SEQUENCE(1)</f></c>`;
    }
    let rows = `<row r="1">${cells}</row>`;
    for (let r = 2; r <= 200; r++) rows += `<row r="${r}"><c r="A${r}"><v>1</v></c></row>`;
    const before = process.memoryUsage().heapUsed;
    const started = Date.now();
    const bytes = await buildXlsx({ sheets: [{ name: 'S', xml: worksheetXml(`<sheetData>${rows}</sheetData>`) }], metadata: METADATA });
    const workbook = await (await open(bytes)).workbook();
    const engine = new CalcEngine(workbook, { budget: { milliseconds: 250 } });
    engine.recalculate(engine.uncalculated());
    for (let k = 0; k < 2000 && engine.incomplete; k++) engine.recalculate([]);
    const took = Date.now() - started;
    const grew = process.memoryUsage().heapUsed - before;
    assert.equal(engine.incomplete, false);
    assert.equal(value(workbook, 0, 5), 1);
    // Guards, not measures: before, 10 such formulas took 2.9 GB and 8 seconds, 2000 overflowed the heap.
    assert.ok(took < 30000, `${legacy ? 'array' : 'dynamic'}: ${took} ms`);
    assert.ok(grew < 1.5e9, `${legacy ? 'array' : 'dynamic'}: ${Math.round(grew / 1e6)} MB`);
  }
});

test('a spill the file gives stays whatever the budget while it does not change; one whose cells the file lacks does not (K3)', async () => {
  // A budget of 5 cells: A1 fills 6 cells of the file, C1 wants 2 the file did not give it, E1 holds one value.
  const first = '<c r="A1" cm="1"><f t="array" ref="A1:A6">_xlfn.SEQUENCE(6)</f><v>1</v></c><c r="C1" cm="1"><f t="array" ref="C1">_xlfn.SEQUENCE(2)</f><v>1</v></c><c r="E1" cm="1"><f>_xlfn.SEQUENCE(1)</f><v>1</v></c>';
  const sheetData = `<row r="1">${first}</row><row r="2"><c r="A2"><v>2</v></c></row><row r="3"><c r="A3"><v>3</v></c></row><row r="4"><c r="A4"><v>4</v></c></row><row r="5"><c r="A5"><v>5</v></c></row><row r="6"><c r="A6"><v>6</v></c></row>`;
  const t = await opened(sheetData, { limits: { workbookSpillCells: 5 } });
  t.engine.invalidate();
  t.engine.recalculate(undefined, { overwrite: true });
  // A1 keeps the six cells of the file; they use up the budget, so C1, after it on the sheet, is #SPILL!.
  assert.deepEqual([value(t.workbook, 0, 0), value(t.workbook, 5, 0), `${value(t.workbook, 0, 2)}`, value(t.workbook, 0, 4)], [1, 6, '#SPILL!', 1]);
  // Grown, it asks for cells the file did not give it: refused, and C1 gets its two.
  t.type(0, 0, '=SEQUENCE(7)');
  assert.deepEqual([`${value(t.workbook, 0, 0)}`, value(t.workbook, 0, 2), value(t.workbook, 1, 2)], ['#SPILL!', 1, 2]);
  // A file that names six cells and has none of them: no spill of the file; A1 is refused, C1 spills.
  const damaged = await opened(`<row r="1">${first}</row>`, { limits: { workbookSpillCells: 5 } });
  damaged.engine.invalidate();
  damaged.engine.recalculate(undefined, { overwrite: true });
  assert.deepEqual([`${value(damaged.workbook, 0, 0)}`, value(damaged.workbook, 1, 2)], ['#SPILL!', 2]);
});

test('two spills that read each other: undoing one leaves the other as an engine built afresh would (Ö6)', async () => {
  const file = await open(await createBlank('xlsx'));
  const workbook = await file.workbook();
  const history = new History(workbook);
  const engine = new CalcEngine(workbook);
  const type = (row: number, col: number, text: string) => {
    history.apply(setCellInput(workbook, 0, row, col, text), undefined, undefined);
    engine.recalculate([{ sheet: 0, row, col }]);
  };
  const state = (book: Workbook) => [0, 1].map((col) => [0, 1, 2].map((row) => `${value(book, row, col)}`).join(',')).join(' ');
  const fresh = () => {
    const copy: Workbook = { ...workbook, sheets: workbook.sheets.map((s) => ({ ...s, rows: new Map(Array.from(s.rows, ([k, r]) => [k, { ...r, cells: r.cells.slice() }])) })) };
    new CalcEngine(copy).recalculate(undefined, { overwrite: true });
    return state(copy);
  };
  type(0, 1, '=SEQUENCE(COUNT(A2:A10)+2)');
  type(0, 0, '=SEQUENCE(COUNT(B2:B10)+2)');
  assert.equal(state(workbook), fresh());
  history.undo();
  engine.recalculate([{ sheet: 0, row: 0, col: 0 }]);
  assert.equal(state(workbook), 'undefined,undefined,undefined 1,2,undefined');
  assert.equal(state(workbook), fresh());
});

test('"C1#" names what the formula in C1 fills when it is a dynamic one, calculated here or not; anything else is #REF! (K-D)', async () => {
  for (const [c1, expected] of [['=NOSUCH(1)', 1], ['=[1]S!A1:A3', 1], ['=1/0', '#REF!'], ['=SEQUENCE(2)', 2], ['5', '#REF!']] as [string, unknown][]) {
    const file = await open(await createBlank('xlsx'));
    const workbook = await file.workbook();
    const history = new History(workbook);
    const engine = new CalcEngine(workbook);
    const type = (book: Workbook, h: History, e: CalcEngine, row: number, col: number, text: string) => {
      h.apply(setCellInput(book, 0, row, col, text), undefined, undefined);
      e.recalculate([{ sheet: 0, row, col }]);
    };
    type(workbook, history, engine, 0, 2, c1);
    type(workbook, history, engine, 0, 3, '=ROWS(C1#)+A5');
    await writeXlsx(file.package!, workbook);
    const joined = await (await open(await file.package!.save())).workbook();
    const joinedHistory = new History(joined);
    const joinedEngine = new CalcEngine(joined);
    joinedEngine.recalculate(joinedEngine.uncalculated());
    type(workbook, history, engine, 4, 0, '0');
    type(joined, joinedHistory, joinedEngine, 4, 0, '0');
    assert.equal(`${value(workbook, 0, 3)}`, `${expected}`, `member ${c1}`);
    assert.equal(`${value(joined, 0, 3)}`, `${expected}`, `joiner ${c1}`);
  }
});

test('a formula of the file stays an ordinary one: "SUM(COUNTIF(A:A,A:A))" keeps its implicit intersection, typed again too', async () => {
  // A1:A4 = 1, 1, 2, 1; B1 and B3 hold formulas the file wrote without cell metadata (ordinary ones).
  const rows =
    '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>SUM(COUNTIF(A:A,A:A))</f><v>3</v></c><c r="C1"><f>MAX(IF(A:A=1,ROW(A:A)))</f><v>1</v></c></row>' +
    '<row r="2"><c r="A2"><v>1</v></c></row><row r="3"><c r="A3"><v>2</v></c><c r="B3"><f>SUM(LEN(A:A&amp;"x"))</f><v>2</v></c></row><row r="4"><c r="A4"><v>1</v></c></row>';
  const t = await opened(rows);
  const cell = (row: number, col: number) => getCell(t.workbook.sheets[0], row, col);
  assert.equal(cell(0, 1)?.formula?.dynamic, undefined);
  // Opened: the values of the file stand.
  assert.deepEqual([value(t.workbook, 0, 1), value(t.workbook, 0, 2), value(t.workbook, 2, 1)], [3, 1, 2]);
  // Something they read changes: calculated as ordinary formulas, a whole column meets the formula's own row.
  t.type(4, 0, '1');
  // COUNTIF(A:A, A1): the 1s of A1:A5; IF(A1=1, ROW(A1)); LEN(A3&"x").
  assert.deepEqual([value(t.workbook, 0, 1), value(t.workbook, 0, 2), value(t.workbook, 2, 1)], [4, 1, 2]);
  assert.equal(cell(0, 1)?.formula?.ref, undefined);
  // Typed again as it is (F2 + Enter), it stays ordinary.
  t.type(0, 1, '=SUM(COUNTIF(A:A,A:A))');
  assert.deepEqual([cell(0, 1)?.formula?.dynamic, value(t.workbook, 0, 1)], [undefined, 4]);
  // Typed in another cell, it is a dynamic one: every row of A is counted (4 x 4 + 1 x 1).
  t.type(0, 4, '=SUM(COUNTIF(A:A,A:A))');
  assert.deepEqual([cell(0, 4)?.formula?.dynamic, value(t.workbook, 0, 4)], [true, 17]);
});

test('a circle that runs through a spill: who opens the saved file works out what the member does when it changes (s9)', async () => {
  // B4 spills B1:C2 over B4:C5; B2 adds up column C, where B4 would spill: a circle. C6 holds 5. Both are circular:
  // 0, and B4 spills nowhere; the circle still runs through the cells it would spill into. Then one edit, made by
  // the member (whose engine found the circle) and by one who opened the file saved in that state, who finds it
  // again (rebuildForced()) without calculating anything.
  const cases: [number, number, string][] = [
    [3, 1, '=SUM(C:C)'],
    [3, 1, '7'],
    [3, 1, '=B1:B2'],
    [3, 1, ''],
    [5, 2, '9'],
    [1, 1, '=SUM(C:C)+1'],
    [0, 1, '2'],
    [3, 2, '1'],
  ];
  const state = (book: Workbook) => ['B2', 'B4', 'B5', 'C4', 'C5', 'E3'].map((a) => `${a}=${value(book, Number(a.slice(1)) - 1, a.charCodeAt(0) - 65)}`).join(' ');
  for (const [row, col, text] of cases) {
    const file = await open(await createBlank('xlsx'));
    const workbook = await file.workbook();
    const history = new History(workbook);
    const engine = new CalcEngine(workbook);
    const type = (book: Workbook, h: History, e: CalcEngine, r: number, c: number, input: string) => {
      h.apply(setCellInput(book, 0, r, c, input), undefined, undefined);
      e.recalculate([{ sheet: 0, row: r, col: c }]);
    };
    type(workbook, history, engine, 5, 2, '5');
    type(workbook, history, engine, 1, 1, '=SUM(C:C)');
    type(workbook, history, engine, 3, 1, '=B1:C2');
    type(workbook, history, engine, 2, 4, '=B2*2');
    assert.equal(state(workbook), 'B2=0 B4=0 B5=undefined C4=undefined C5=undefined E3=0', 'the circle');
    await writeXlsx(file.package!, workbook);
    const joined = await (await open(await file.package!.save())).workbook();
    const joinedHistory = new History(joined);
    const joinedEngine = new CalcEngine(joined);
    joinedEngine.recalculate(joinedEngine.uncalculated());
    type(workbook, history, engine, row, col, text);
    type(joined, joinedHistory, joinedEngine, row, col, text);
    assert.equal(state(joined), state(workbook), `joiner and member, R${row + 1}C${col + 1} = ${text}`);
    if (row === 3 && col === 1 && text === '=SUM(C:C)') assert.equal(state(workbook), 'B2=5 B4=5 B5=undefined C4=undefined C5=undefined E3=10');
  }
});

test('a formula of a circle the file was saved with is calculated again only when what it reads changes: an unrelated edit and its undo leave the file as it was (Y2)', async () => {
  const t = await opened('<row r="1"><c r="A1" cm="1"><f t="array" ref="A1">_xlfn.SEQUENCE(3)</f><v>0</v></c><c r="B1"><f>SUM(A1:A3)</f><v>0</v></c></row><row r="5"/><row r="6"><c r="D6"><v>1</v></c></row>');
  t.type(19, 6, '5');
  t.undo(19, 6);
  assert.equal(await t.sheetXml(), t.original);
  assert.deepEqual([value(t.workbook, 0, 0), value(t.workbook, 0, 1)], [0, 0]);
});

test('the limit on circles put back from a file counts the formulas that would be put back, not every dynamic one (B4)', async () => {
  const file = await open(await createBlank('xlsx'));
  const workbook = await file.workbook();
  const history = new History(workbook);
  const engine = new CalcEngine(workbook);
  const type = (book: Workbook, h: History, e: CalcEngine, r: number, c: number, input: string) => {
    h.apply(setCellInput(book, 0, r, c, input), undefined, undefined);
    e.recalculate([{ sheet: 0, row: r, col: c }]);
  };
  type(workbook, history, engine, 5, 2, '5');
  type(workbook, history, engine, 1, 1, '=SUM(C:C)');
  type(workbook, history, engine, 3, 1, '=B1:C2');
  type(workbook, history, engine, 2, 4, '=B2*2');
  type(workbook, history, engine, 0, 9, '1');
  // 1,001 dynamic formulas of one value each: none of them is to be put back.
  for (let k = 0; k < 1001; k++) type(workbook, history, engine, 20 + k, 8, '=XLOOKUP(1,J1:J2,J1:J2)');
  await writeXlsx(file.package!, workbook);
  const joined = await (await open(await file.package!.save())).workbook();
  const joinedHistory = new History(joined);
  const joinedEngine = new CalcEngine(joined);
  joinedEngine.recalculate(joinedEngine.uncalculated());
  type(workbook, history, engine, 3, 1, '=SUM(C:C)');
  type(joined, joinedHistory, joinedEngine, 3, 1, '=SUM(C:C)');
  const state = (book: Workbook) => [[1, 1], [3, 1], [2, 4]].map(([r, c]) => `${value(book, r, c)}`).join(' ');
  assert.equal(state(workbook), '5 5 10');
  assert.equal(state(joined), state(workbook));
});

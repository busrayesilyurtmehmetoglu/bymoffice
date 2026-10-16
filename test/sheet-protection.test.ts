// The protection of sheets and of the structure of a workbook, applied: what a protected
// sheet refuses (protectionProblem() over the operations of each kind of command), how the
// flags of <sheetProtection> are read (their defaults), locked and unlocked cells (of the
// cell, the row, the column, an open range), the password of a protection, protecting and
// unprotecting ("setProtection") and how the writer puts that into the file, xls records,
// and <workbookProtection>. Every file is built here.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { open, applyOp, buildXlsx, type Op, type Sheet, type Workbook } from '../src/index';
import { readXlsStream } from '../src/formats/xls/workbook';
import { clearRange, copyRange, cutRange, pasteAt, setCellInput } from '../src/ops/commands';
import { formatRange, mergeRange, setColumnWidth, setRowHeight } from '../src/ops/format';
import { deleteRows, setAutoFilter, setFilterColumn } from '../src/ops/filter';
import { sortRange } from '../src/ops/sort';
import { setComment } from '../src/ops/notes';
import { setValidation } from '../src/ops/rules';
import { freezePanes } from '../src/ops/panes';
import { canDeleteSheet, canInsertSheet, canMoveSheet, canRenameSheet, deleteSheet, insertSheet, moveSheet, renameSheet, setSheetState, setTabColor } from '../src/ops/sheets';
import { replaceAllOps } from '../src/ops/find';
import {
  FLAG_DEFAULTS,
  PasswordMismatch,
  ProtectedError,
  actionProblem,
  cellProtected,
  checkProtectionPassword,
  defaultProtectionAttrs,
  formulaHidden,
  legacyPasswordHash,
  passwordAttrs,
  protectedCellIn,
  protectionForbids,
  protectionHasPassword,
  protectionProblem,
  setProtectionOps,
  sheetProtected,
  structureLocked,
  structureProblem,
  unprotectSheetOps,
} from '../src/ops/protection';
import { buildXlsx as makeXlsx, worksheetXml } from './fixtures';
import { cat, rec, u16, u32, workbookStream, xf, font, number, xlsFile } from './xls-fixtures';
import { bookOf, put, sheetOf } from './xlsx-build-fixtures';

const text = (bytes: Uint8Array) => new TextDecoder().decode(bytes);

/** A workbook of one sheet A1:C3 with numbers; style 1 is unlocked, style 2 hides formulas. */
function book(attrs: [string, string][] | null = [['sheet', '1']]): Workbook {
  const sheet = sheetOf('S');
  const workbook = bookOf(sheet);
  const base = workbook.styles.cellXfs[0];
  workbook.styles.cellXfs.push({ ...base, protection: { locked: false } }, { ...base, protection: { hidden: true } });
  for (let r = 0; r < 3; r++) for (let c = 0; c < 3; c++) put(sheet, r, c, { kind: 'number', value: r * 3 + c + 1 });
  if (attrs) sheet.protection = { attrs };
  return workbook;
}

const UNLOCKED = 1;
const HIDDEN = 2;

function unlock(workbook: Workbook, cells: [number, number][]): void {
  const sheet = workbook.sheets[0];
  for (const [r, c] of cells) {
    const row = sheet.rows.get(r);
    const cell = row && row.cells[c];
    if (cell) row!.cells[c] = { ...cell, style: UNLOCKED };
    else put(sheet, r, c, { kind: 'blank', value: null, style: UNLOCKED });
  }
}

const reasonOf = (workbook: Workbook, ops: Op[]) => {
  const refusal = protectionProblem(workbook, ops);
  return refusal ? refusal.reason : undefined;
};

// ---------------------------------------------------------------------------- flags

test('a sheet is protected by sheet="1" (or "true") alone', () => {
  assert.equal(sheetProtected(book([['sheet', '1']]).sheets[0]), true);
  assert.equal(sheetProtected(book([['sheet', 'true']]).sheets[0]), true);
  assert.equal(sheetProtected(book([['sheet', '0']]).sheets[0]), false);
  assert.equal(sheetProtected(book([['objects', '1']]).sheets[0]), false);
  assert.equal(sheetProtected(book(null).sheets[0]), false);
  // Ranges alone protect nothing.
  const ranges = book(null);
  ranges.sheets[0].protection = { ranges: [{ attrs: [['sqref', 'A1'], ['name', 'R']] }] };
  assert.equal(sheetProtected(ranges.sheets[0]), false);
});

test('flags left out take the defaults of the schema: formats, rows, columns, sort, filter, pivots forbidden; objects, scenarios and selection free', () => {
  const sheet = book([['sheet', '1']]).sheets[0];
  for (const flag of Object.keys(FLAG_DEFAULTS) as (keyof typeof FLAG_DEFAULTS)[]) assert.equal(protectionForbids(sheet, flag), FLAG_DEFAULTS[flag], flag);
  assert.equal(protectionForbids(sheet, 'formatCells'), true);
  assert.equal(protectionForbids(sheet, 'deleteRows'), true);
  assert.equal(protectionForbids(sheet, 'objects'), false);
  assert.equal(protectionForbids(sheet, 'selectLockedCells'), false);
  const open = book([['sheet', '1'], ['formatCells', '0'], ['objects', 'true'], ['sort', 'false'], ['autoFilter', 'maybe']]).sheets[0];
  assert.equal(protectionForbids(open, 'formatCells'), false);
  assert.equal(protectionForbids(open, 'objects'), true);
  assert.equal(protectionForbids(open, 'sort'), false);
  // Neither true nor false: the default.
  assert.equal(protectionForbids(open, 'autoFilter'), true);
  // A sheet that is not protected forbids nothing.
  assert.equal(protectionForbids(book([['sheet', '0'], ['formatCells', '1']]).sheets[0], 'formatCells'), false);
});

// ---------------------------------------------------------------------------- locked cells

test('cells are locked unless their format, their row or their column says locked="0", or an open range holds them', () => {
  const workbook = book();
  const sheet = workbook.sheets[0];
  assert.equal(cellProtected(workbook, sheet, 0, 0), true);
  unlock(workbook, [[0, 0]]);
  assert.equal(cellProtected(workbook, sheet, 0, 0), false);
  // An empty position: its row's format, else its column's, else the default (locked).
  assert.equal(cellProtected(workbook, sheet, 10, 10), true);
  sheet.rows.set(10, { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells: [], style: UNLOCKED });
  assert.equal(cellProtected(workbook, sheet, 10, 10), false);
  sheet.columns = [{ min: 20, max: 22, customWidth: false, hidden: false, style: UNLOCKED, outlineLevel: 0, collapsed: false }];
  assert.equal(cellProtected(workbook, sheet, 50, 21), false);
  assert.equal(cellProtected(workbook, sheet, 50, 23), true);
  // A range open to everyone; one with a password is not.
  sheet.protection = { attrs: [['sheet', '1']], ranges: [{ attrs: [['sqref', 'E5:F6 H8'], ['name', 'Open']] }, { attrs: [['sqref', 'J10'], ['name', 'Closed'], ['password', 'CE88']] }] };
  assert.equal(cellProtected(workbook, sheet, 4, 4), false);
  assert.equal(cellProtected(workbook, sheet, 7, 7), false);
  assert.equal(cellProtected(workbook, sheet, 9, 9), true);
  // Not protected: nothing is.
  sheet.protection = undefined;
  assert.equal(cellProtected(workbook, sheet, 0, 1), false);
});

test('a locked position of a range is found among cells, empty positions, rows and column formats', () => {
  const workbook = book();
  const sheet = workbook.sheets[0];
  unlock(workbook, [[0, 0], [0, 1], [1, 0], [1, 1]]);
  assert.equal(protectedCellIn(workbook, sheet, { top: 0, left: 0, bottom: 1, right: 1 }), undefined);
  assert.deepEqual(protectedCellIn(workbook, sheet, { top: 0, left: 0, bottom: 1, right: 2 }), { row: 0, col: 2 });
  // Empty rows below take the default format.
  assert.deepEqual(protectedCellIn(workbook, sheet, { top: 5, left: 0, bottom: 6, right: 0 }), { row: 5, col: 0 });
  // Whole columns: every position of them, without looking at a million rows.
  sheet.columns = [{ min: 5, max: 6, customWidth: false, hidden: false, style: UNLOCKED, outlineLevel: 0, collapsed: false }];
  assert.equal(protectedCellIn(workbook, sheet, { top: 0, left: 5, bottom: 1048575, right: 6 }), undefined);
  put(sheet, 100, 5, { kind: 'number', value: 1, style: 0 });
  assert.deepEqual(protectedCellIn(workbook, sheet, { top: 0, left: 5, bottom: 1048575, right: 6 }), { row: 100, col: 5 });
});

test('a formula is hidden on a protected sheet when its format says hidden="1"; its value is not', () => {
  const workbook = book();
  const sheet = workbook.sheets[0];
  const row = sheet.rows.get(0)!;
  row.cells[0] = { kind: 'number', value: 3, style: HIDDEN, formula: { text: '1+2', type: 'normal' } };
  row.cells[1] = { kind: 'number', value: 3, style: HIDDEN };
  assert.equal(formulaHidden(workbook, sheet, 0, 0), true);
  assert.equal(formulaHidden(workbook, sheet, 0, 1), false);
  sheet.protection = undefined;
  assert.equal(formulaHidden(workbook, sheet, 0, 0), false);
});

// ---------------------------------------------------------------------------- what each command may do

test('typing: a locked cell is refused, an unlocked one is written (with the format a date brings)', () => {
  const workbook = book();
  assert.equal(reasonOf(workbook, setCellInput(workbook, 0, 0, 0, '42')), 'cells');
  unlock(workbook, [[0, 0]]);
  assert.equal(reasonOf(workbook, setCellInput(workbook, 0, 0, 0, '42')), undefined);
  assert.equal(reasonOf(workbook, setCellInput(workbook, 0, 0, 0, '2026-10-16')), undefined);
  // Not protected: anything goes.
  workbook.sheets[0].protection = undefined;
  assert.equal(reasonOf(workbook, setCellInput(workbook, 0, 1, 1, 'x')), undefined);
});

test('Delete: clearing a range with a locked cell is refused; clearing unlocked cells is not', () => {
  const workbook = book();
  assert.equal(reasonOf(workbook, clearRange(workbook, 0, { top: 0, left: 0, bottom: 2, right: 2 })), 'cells');
  unlock(workbook, [[2, 0], [2, 1], [2, 2]]);
  assert.equal(reasonOf(workbook, clearRange(workbook, 0, { top: 2, left: 0, bottom: 2, right: 2 })), undefined);
});

test('paste: all of the range or nothing - one locked cell refuses it, formats come along into unlocked cells', () => {
  const workbook = book();
  unlock(workbook, [[0, 0], [0, 1], [1, 0], [1, 1]]);
  const copied = copyRange(workbook, 0, { top: 0, left: 0, bottom: 0, right: 1 });
  assert.equal(reasonOf(workbook, pasteAt(workbook, 0, { top: 1, left: 0, bottom: 1, right: 0 }, copied.text, copied).ops), undefined);
  // Into B2:C2, where C2 is locked.
  assert.equal(reasonOf(workbook, pasteAt(workbook, 0, { top: 1, left: 1, bottom: 1, right: 1 }, copied.text, copied).ops), 'cells');
  // A copy of locked cells brings their (locked) format into the unlocked cells it lands on, as Excel does.
  const locked = copyRange(workbook, 0, { top: 2, left: 0, bottom: 2, right: 1 });
  assert.equal(reasonOf(workbook, pasteAt(workbook, 0, { top: 0, left: 0, bottom: 0, right: 0 }, locked.text, locked).ops), undefined);
});

test('cut: refused when a cut cell is locked', () => {
  const workbook = book();
  assert.equal(reasonOf(workbook, cutRange(workbook, 0, { top: 0, left: 0, bottom: 0, right: 0 })), 'cells');
  unlock(workbook, [[0, 0]]);
  assert.equal(reasonOf(workbook, cutRange(workbook, 0, { top: 0, left: 0, bottom: 0, right: 0 })), undefined);
});

test('Replace All: refused when one of the cells it changes is locked, done when all of them are unlocked', () => {
  const workbook = book();
  unlock(workbook, [[0, 0]]);
  // "1" is in A1 (unlocked) alone: replaced. "2" is in B1 (locked) as well: refused as a whole.
  assert.equal(reasonOf(workbook, replaceAllOps(workbook, '1', '7').ops), undefined);
  const ops = replaceAllOps(workbook, '2', '7').ops;
  assert.ok(ops.length > 0);
  assert.equal(reasonOf(workbook, ops), 'cells');
});

test('formats: refused when formatCells is forbidden (the default), allowed on locked cells when it is not; locking never', () => {
  const workbook = book();
  const range = { top: 0, left: 0, bottom: 1, right: 1 };
  assert.equal(reasonOf(workbook, formatRange(workbook, 0, range, { bold: true })), 'formatCells');
  unlock(workbook, [[0, 0]]);
  // Unlocked cells are not formatted either.
  assert.equal(reasonOf(workbook, formatRange(workbook, 0, { top: 0, left: 0, bottom: 0, right: 0 }, { bold: true })), 'formatCells');
  workbook.sheets[0].protection = { attrs: [['sheet', '1'], ['formatCells', '0']] };
  assert.equal(reasonOf(workbook, formatRange(workbook, 0, range, { bold: true })), undefined);
  // Whole columns: the format of the columns too.
  assert.equal(reasonOf(workbook, formatRange(workbook, 0, { top: 0, left: 4, bottom: 1048575, right: 4 }, { italic: true })), undefined);
  // A format that unlocks a cell (or hides its formula) is refused even then.
  const ops: Op[] = [{ type: 'setCells', sheet: 0, cells: [{ row: 1, col: 1, cell: { kind: 'number', value: 5, style: UNLOCKED } }] }];
  assert.equal(reasonOf(workbook, ops), 'lock');
  assert.equal(reasonOf(workbook, [{ type: 'setColumns', sheet: 0, columns: [{ min: 0, max: 0, customWidth: false, hidden: false, style: HIDDEN, outlineLevel: 0, collapsed: false }] }]), 'lock');
});

test('column widths and row heights need formatColumns and formatRows', () => {
  const workbook = book();
  assert.equal(reasonOf(workbook, setColumnWidth(workbook, 0, 0, 0, 20)), 'formatColumns');
  assert.equal(reasonOf(workbook, setRowHeight(0, 0, 0, 30)), 'formatRows');
  workbook.sheets[0].protection = { attrs: [['sheet', '1'], ['formatColumns', '0'], ['formatRows', '0']] };
  assert.equal(reasonOf(workbook, setColumnWidth(workbook, 0, 0, 0, 20)), undefined);
  assert.equal(reasonOf(workbook, setRowHeight(0, 0, 0, 30)), undefined);
});

test('rows and columns: inserting needs its flag, deleting needs its flag and no locked position in the rows', () => {
  const workbook = book();
  const insertRow: Op[] = [{ type: 'insert', sheet: 0, axis: 'row', at: 1, count: 1 }];
  const insertCol: Op[] = [{ type: 'insert', sheet: 0, axis: 'col', at: 1, count: 1 }];
  assert.equal(reasonOf(workbook, insertRow), 'insertRows');
  assert.equal(reasonOf(workbook, insertCol), 'insertColumns');
  assert.equal(reasonOf(workbook, deleteRows(workbook, 0, 1, 1)), 'deleteRows');
  workbook.sheets[0].protection = { attrs: [['sheet', '1'], ['insertRows', '0'], ['insertColumns', '0'], ['deleteRows', '0'], ['deleteColumns', '0']] };
  assert.equal(reasonOf(workbook, insertRow), undefined);
  assert.equal(reasonOf(workbook, insertCol), undefined);
  // Whole rows hold locked cells (the default format) unless their own format is unlocked.
  assert.equal(reasonOf(workbook, deleteRows(workbook, 0, 5, 1)), 'deleteRows');
  assert.equal(reasonOf(workbook, [{ type: 'delete', sheet: 0, axis: 'col', at: 5, count: 1 }]), 'deleteColumns');
  const sheet = workbook.sheets[0];
  sheet.rows.set(5, { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells: [], style: UNLOCKED });
  assert.equal(reasonOf(workbook, deleteRows(workbook, 0, 5, 1)), undefined);
  // A locked cell of its own in that row.
  put(sheet, 5, 3, { kind: 'number', value: 1, style: 0 });
  assert.equal(reasonOf(workbook, deleteRows(workbook, 0, 5, 1)), 'deleteRows');
  assert.equal(actionProblem(workbook, 0, 'deleteRows', { top: 5, left: 0, bottom: 5, right: 0 }) !== undefined, true);
});

test('sort: needs sort and unlocked cells; the state of the sort it writes into the filter comes with it', () => {
  const workbook = book();
  const range = { top: 0, left: 0, bottom: 2, right: 0 };
  const ops = sortRange(workbook, 0, range, [{ col: 0, descending: true }]);
  assert.equal(reasonOf(workbook, ops), 'sort');
  workbook.sheets[0].protection = { attrs: [['sheet', '1'], ['sort', '0']] };
  assert.equal(reasonOf(workbook, ops), 'sort');
  unlock(workbook, [[0, 0], [1, 0], [2, 0]]);
  assert.equal(reasonOf(workbook, sortRange(workbook, 0, range, [{ col: 0, descending: true }])), undefined);
});

test('filter: never added or removed on a protected sheet; its criteria need autoFilter', () => {
  const workbook = book(null);
  const added = setAutoFilter(workbook, 0, { top: 0, left: 0, bottom: 2, right: 2 });
  for (const op of added) applyOp(workbook, op);
  workbook.sheets[0].protection = { attrs: [['sheet', '1']] };
  assert.equal(reasonOf(workbook, setAutoFilter(workbook, 0, undefined)), 'autoFilter');
  const criterion = setFilterColumn(workbook, 0, 0, { kind: 'values', values: ['1'] });
  assert.equal(reasonOf(workbook, criterion), 'autoFilter');
  workbook.sheets[0].protection = { attrs: [['sheet', '1'], ['autoFilter', '0']] };
  assert.equal(reasonOf(workbook, criterion), undefined);
  assert.equal(reasonOf(workbook, setAutoFilter(workbook, 0, undefined)), 'autoFilter');
  assert.equal(actionProblem(workbook, 0, 'filterOnOff') !== undefined, true);
});

test('merging needs formatCells; notes and drawings need objects; rules are not changed', () => {
  const workbook = book();
  unlock(workbook, [[0, 0], [0, 1]]);
  assert.equal(reasonOf(workbook, mergeRange(workbook, 0, { top: 0, left: 0, bottom: 0, right: 1 })), 'merge');
  // A note: objects left out leaves them free; objects="1" (what Excel writes) forbids.
  assert.equal(reasonOf(workbook, setComment(workbook, 0, 0, 0, 'Hello')), undefined);
  workbook.sheets[0].protection = { attrs: [['sheet', '1'], ['objects', '1']] };
  assert.equal(reasonOf(workbook, setComment(workbook, 0, 0, 0, 'Hello')), 'objects');
  assert.equal(reasonOf(workbook, [{ type: 'setDrawings', sheet: 0, drawings: [] }]), 'objects');
  assert.equal(reasonOf(workbook, setValidation(workbook, 0, { top: 0, left: 0, bottom: 0, right: 0 }, { type: 'whole', operator: 'greaterThan', formula1: '0' })), 'rules');
  assert.equal(reasonOf(workbook, [{ type: 'setConditionalFormats', sheet: 0, formats: [] }]), 'rules');
});

test('frozen panes, the page setup and the protection itself are free on a protected sheet', () => {
  const workbook = book();
  assert.equal(reasonOf(workbook, freezePanes(workbook, 0, 1, 0)), undefined);
  assert.equal(reasonOf(workbook, [{ type: 'setPageSetup', sheet: 0, setup: { page: { orientation: 'landscape' } } }]), undefined);
  assert.equal(reasonOf(workbook, setProtectionOps(workbook, 0, undefined)), undefined);
});

test('a step on a sheet that is not protected is never refused, also when another sheet is', () => {
  const workbook = book();
  const other = sheetOf('T', 2);
  workbook.sheets.push(other);
  put(other, 0, 0, { kind: 'number', value: 1 });
  assert.equal(reasonOf(workbook, setCellInput(workbook, 1, 0, 0, '2')), undefined);
  assert.equal(reasonOf(workbook, formatRange(workbook, 1, { top: 0, left: 0, bottom: 0, right: 0 }, { bold: true })), undefined);
});

// ---------------------------------------------------------------------------- the structure of the workbook

test('a locked structure refuses every change of the list of sheets, and says so before', () => {
  const workbook = book(null);
  workbook.sheets.push(sheetOf('T', 2));
  workbook.protection = { attrs: [['lockStructure', '1']] };
  assert.equal(structureLocked(workbook), true);
  for (const ops of [insertSheet(workbook, 2), renameSheet(workbook, 0, 'X'), deleteSheet(workbook, 1), moveSheet(workbook, 0, 1), setSheetState(workbook, 1, 'hidden'), setTabColor(workbook, 0, '#FF0000')]) {
    assert.equal(reasonOf(workbook, ops), 'structure');
  }
  assert.equal(canRenameSheet(workbook, 0), structureProblem(workbook));
  assert.equal(canDeleteSheet(workbook, 1), structureProblem(workbook));
  assert.equal(canInsertSheet(workbook, 0), structureProblem(workbook));
  assert.equal(canMoveSheet(workbook, 0, 1), structureProblem(workbook));
  // Cells are not the structure.
  assert.equal(reasonOf(workbook, setCellInput(workbook, 0, 0, 0, '5')), undefined);
  workbook.protection = { attrs: [['lockStructure', '0'], ['lockWindows', '1']] };
  assert.equal(structureLocked(workbook), false);
  assert.equal(reasonOf(workbook, renameSheet(workbook, 0, 'X')), undefined);
});

test('the refusal is a ProtectedError (a RangeError) with the code "protected" and the sheet', () => {
  const workbook = book();
  const refusal = protectionProblem(workbook, setCellInput(workbook, 0, 0, 0, '9'));
  assert.ok(refusal instanceof ProtectedError);
  assert.ok(refusal instanceof RangeError);
  assert.equal(refusal.code, 'protected');
  assert.equal(refusal.sheet, 0);
});

// ---------------------------------------------------------------------------- passwords

test('the legacy 16-bit hash of a password is the one openpyxl computes', () => {
  // Values of openpyxl.utils.protection.hash_password.
  assert.equal(legacyPasswordHash('secret'), 'DAA7');
  assert.equal(legacyPasswordHash('password'), '83AF');
  assert.equal(legacyPasswordHash('a'), 'CE88');
  assert.equal(legacyPasswordHash('Deneme123'), '921E');
});

test('a password is checked against the legacy hash and against the SHA hash, either round order', async () => {
  assert.equal(await checkProtectionPassword([['sheet', '1']], 'anything'), true);
  assert.equal(await checkProtectionPassword([['sheet', '1'], ['password', 'DAA7']], 'secret'), true);
  assert.equal(await checkProtectionPassword([['sheet', '1'], ['password', 'daa7']], 'secret'), true);
  assert.equal(await checkProtectionPassword([['sheet', '1'], ['password', 'DAA7']], 'Secret'), false);
  // A hash of 1000 rounds made by Python's hashlib: H(salt + password), then H(hash + round), the round after.
  const attrs: [string, string][] = [
    ['algorithmName', 'SHA-512'],
    ['hashValue', 'NKulX3iIHYXWyPrvf72PReBaPGA+mRJDgiuhqO2ya8vjKp7E1WKQUF7JBoL8aQRO5SLHtktNc13MLTIh2VouPA=='],
    ['saltValue', 'AAECAwQFBgcICQoLDA0ODw=='],
    ['spinCount', '1000'],
    ['sheet', '1'],
  ];
  assert.equal(await checkProtectionPassword(attrs, 'Deneme123'), true);
  assert.equal(await checkProtectionPassword(attrs, 'deneme123'), false);
  const made = await passwordAttrs('Şifre-İğü', new Uint8Array(16).fill(7), 50);
  assert.equal(await checkProtectionPassword(made, 'Şifre-İğü'), true);
  assert.equal(await checkProtectionPassword(made, 'Sifre-Igu'), false);
  await assert.rejects(checkProtectionPassword([['algorithmName', 'GOST'], ['hashValue', 'AA=='], ['saltValue', 'AA=='], ['spinCount', '1']], 'x'), RangeError);
  await assert.rejects(checkProtectionPassword([['algorithmName', 'SHA-512'], ['hashValue', 'AA=='], ['saltValue', 'AA=='], ['spinCount', '99999999']], 'x'), RangeError);
});

test('unprotecting checks the password; protecting sets the attributes Excel writes', async () => {
  const workbook = book([['sheet', '1'], ['objects', '1'], ['scenarios', '1'], ['password', 'DAA7']]);
  assert.equal(protectionHasPassword(workbook.sheets[0]), true);
  await assert.rejects(unprotectSheetOps(workbook, 0, 'wrong'), PasswordMismatch);
  const ops = await unprotectSheetOps(workbook, 0, 'secret');
  const undo = applyOp(workbook, ops[0]);
  assert.equal(sheetProtected(workbook.sheets[0]), false);
  applyOp(workbook, undo);
  assert.equal(sheetProtected(workbook.sheets[0]), true);
  assert.deepEqual(defaultProtectionAttrs(), [['sheet', '1'], ['objects', '1'], ['scenarios', '1']]);
  assert.deepEqual(defaultProtectionAttrs({ formatCells: true, selectLockedCells: false }), [['sheet', '1'], ['objects', '1'], ['scenarios', '1'], ['formatCells', '0'], ['selectLockedCells', '1']]);
});

// ---------------------------------------------------------------------------- files

const STYLES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
  '<borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1"><protection locked="0"/></xf></cellXfs>' +
  '</styleSheet>';

const PROTECTION = '<sheetProtection algorithmName="SHA-512" hashValue="NKulX3iIHYXWyPrvf72PReBaPGA+mRJDgiuhqO2ya8vjKp7E1WKQUF7JBoL8aQRO5SLHtktNc13MLTIh2VouPA==" saltValue="AAECAwQFBgcICQoLDA0ODw==" spinCount="1000" sheet="1" objects="1" scenarios="1"/>';

const fileWith = (protection: string, head = '') =>
  makeXlsx({
    styles: STYLES,
    workbookHead: head,
    sheets: [{ name: 'S', xml: worksheetXml(`<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1" s="1"><v>2</v></c></row></sheetData>${protection}<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>`) }],
  });

const sheetPart = async (bytes: Uint8Array) => text(await (await open(bytes)).package!.readPart('/xl/worksheets/sheet1.xml'));

test('xlsx: the protection and the unlocked cells are read; an untouched save keeps <sheetProtection> byte for byte', async () => {
  const bytes = await fileWith(PROTECTION, '<workbookProtection lockStructure="1" workbookPassword="CE88"/>');
  const file = await open(bytes);
  const workbook = await file.workbook();
  const sheet = workbook.sheets[0];
  assert.equal(sheetProtected(sheet), true);
  assert.equal(cellProtected(workbook, sheet, 0, 0), true);
  assert.equal(cellProtected(workbook, sheet, 0, 1), false);
  assert.equal(structureLocked(workbook), true);
  assert.deepEqual(workbook.protection, { attrs: [['lockStructure', '1'], ['workbookPassword', 'CE88']] });
  // Typing into B1 is fine, into A1 not.
  assert.equal(reasonOf(workbook, setCellInput(workbook, 0, 0, 0, '5')), 'cells');
  for (const op of setCellInput(workbook, 0, 0, 1, '5')) applyOp(workbook, op);
  const saved = await file.save();
  assert.ok((await sheetPart(saved)).indexOf(PROTECTION) >= 0);
  assert.ok(text(await (await open(saved)).package!.readPart('/xl/workbook.xml')).indexOf('<workbookProtection lockStructure="1" workbookPassword="CE88"/>') >= 0);
});

test('xlsx: unprotecting takes <sheetProtection> out, protecting puts one where the schema has it, undoing both gives the file back', async () => {
  const bytes = await fileWith(PROTECTION);
  const file = await open(bytes);
  const workbook = await file.workbook();
  const ops = await unprotectSheetOps(workbook, 0, 'Deneme123');
  const undo = applyOp(workbook, ops[0]);
  const unprotected = await sheetPart(await file.save());
  assert.equal(unprotected.indexOf('sheetProtection'), -1);
  assert.ok(unprotected.indexOf('</sheetData><pageMargins') >= 0);
  const reread = await (await open(await file.save())).workbook();
  assert.equal(sheetProtected(reread.sheets[0]), false);
  applyOp(workbook, undo);
  assert.deepEqual(await file.save(), bytes);

  // A sheet without protection gets one between <sheetData> and <pageMargins>.
  const plain = await fileWith('');
  const other = await open(plain);
  const book2 = await other.workbook();
  const undo2 = applyOp(book2, setProtectionOps(book2, 0, defaultProtectionAttrs())[0]);
  const part = await sheetPart(await other.save());
  assert.ok(part.indexOf('</sheetData><sheetProtection sheet="1" objects="1" scenarios="1"/><pageMargins') >= 0, part);
  const again = await (await open(await other.save())).workbook();
  assert.equal(sheetProtected(again.sheets[0]), true);
  applyOp(book2, undo2);
  assert.deepEqual(await other.save(), plain);
});

test('xlsx: a sheet whose cells and protection both changed is written with both', async () => {
  const file = await open(await fileWith(''));
  const workbook = await file.workbook();
  for (const op of setCellInput(workbook, 0, 2, 2, 'x')) applyOp(workbook, op);
  applyOp(workbook, setProtectionOps(workbook, 0, [['sheet', '1'], ['formatCells', '0']])[0]);
  const reread = await (await open(await file.save())).workbook();
  assert.equal(sheetProtected(reread.sheets[0]), true);
  assert.equal(protectionForbids(reread.sheets[0], 'formatCells'), false);
  assert.equal(reread.sheets[0].rows.get(2)!.cells[2].value, 'x');
});

// ---------------------------------------------------------------------------- xls

const protectRec = (id: number, on: number) => rec(id, u16(on));

function xlsBook(sheetRecords: Uint8Array, globals: Uint8Array = new Uint8Array(0)) {
  const head = cat(font(), font(), font(), font(), ...Array.from({ length: 15 }, () => xf({ style: true })), xf({ parent: 0 }), xf({ parent: 0, locked: false }), globals);
  return workbookStream({ head, sheets: [{ name: 'S', records: sheetRecords }, { name: 'T' }] });
}

test('xls: PROTECT, OBJPROTECT, SCENPROTECT and PASSWORD of a sheet become <sheetProtection>; an unlocked XF unlocks its cells', () => {
  const stream = xlsBook(cat(number(0, 0, 1, 15), number(0, 1, 2, 16), protectRec(0x0012, 1), protectRec(0x0063, 1), protectRec(0x00dd, 1), rec(0x0013, u16(0xdaa7))));
  const { workbook } = readXlsStream(stream);
  const sheet = workbook.sheets[0];
  assert.deepEqual(sheet.protection, { attrs: [['password', 'DAA7'], ['sheet', '1'], ['objects', '1'], ['scenarios', '1']] });
  assert.equal(cellProtected(workbook, sheet, 0, 0), true);
  assert.equal(cellProtected(workbook, sheet, 0, 1), false);
  // An Excel 97 protection: formats and rows forbidden, selection free.
  assert.equal(protectionForbids(sheet, 'formatCells'), true);
  assert.equal(protectionForbids(sheet, 'selectLockedCells'), false);
  assert.equal(workbook.sheets[1].protection, undefined);
});

test('xls: the enhanced protection of FEATHEADR says what is allowed; PROTECT 0 protects nothing', () => {
  // Allowed: formatCells (bit 2), insertRows (bit 6), sort (bit 11), selectLockedCells (bit 10) and selectUnlockedCells (bit 14).
  const flags = (1 << 2) | (1 << 6) | (1 << 10) | (1 << 11) | (1 << 14);
  const feat = rec(0x0867, u16(0x0867), u16(0), [0, 0, 0, 0, 0, 0, 0, 0], u16(2), [0], u32(4), u32(flags));
  const { workbook } = readXlsStream(xlsBook(cat(protectRec(0x0012, 1), feat)));
  const sheet = workbook.sheets[0];
  assert.equal(protectionForbids(sheet, 'formatCells'), false);
  assert.equal(protectionForbids(sheet, 'insertRows'), false);
  assert.equal(protectionForbids(sheet, 'sort'), false);
  assert.equal(protectionForbids(sheet, 'formatRows'), true);
  assert.equal(protectionForbids(sheet, 'selectLockedCells'), false);
  assert.equal(protectionForbids(sheet, 'selectUnlockedCells'), false);
  const off = readXlsStream(xlsBook(protectRec(0x0012, 0))).workbook;
  assert.equal(off.sheets[0].protection, undefined);
});

test('xls: PROTECT of the globals locks the structure; converted to xlsx, both protections are written', async () => {
  const stream = xlsBook(cat(number(0, 0, 1, 15), protectRec(0x0012, 1)), cat(protectRec(0x0012, 1), protectRec(0x0019, 0), rec(0x0013, u16(0xce88))));
  const { workbook } = readXlsStream(stream);
  assert.deepEqual(workbook.protection, { attrs: [['lockStructure', '1'], ['workbookPassword', 'CE88']] });
  assert.equal(structureLocked(workbook), true);
  const file = await open(xlsFile(stream));
  const xlsx = await file.saveAsXlsx!();
  const back = await (await open(xlsx)).workbook();
  assert.equal(structureLocked(back), true);
  assert.equal(sheetProtected(back.sheets[0]), true);
  assert.equal(cellProtected(back, back.sheets[0], 0, 0), true);
  // Built from the model as well.
  const built = await (await open(await buildXlsx(workbook))).workbook();
  assert.equal(sheetProtected(built.sheets[0]), true);
});

test('a refusal leaves the workbook as it was: the check runs before any operation is applied', () => {
  const workbook = book();
  const before = JSON.stringify([...workbook.sheets[0].rows.entries()]);
  const ops = setCellInput(workbook, 0, 0, 0, '9');
  assert.ok(protectionProblem(workbook, ops));
  assert.equal(JSON.stringify([...workbook.sheets[0].rows.entries()]), before);
});

test('a sheet with a record for each of its 16,384 columns is judged in one walk, not one per column', () => {
  // A guard against getting stuck, not a measure of speed: on an idle machine this takes some 20 ms.
  const workbook = book([['sheet', '1'], ['formatColumns', '0'], ['deleteRows', '0']]);
  const sheet = workbook.sheets[0];
  sheet.columns = Array.from({ length: 16384 }, (_value, c) => ({ min: c, max: c, customWidth: false, hidden: false, style: UNLOCKED, outlineLevel: 0, collapsed: false }));
  for (let r = 10; r < 400; r++) {
    const cells = [];
    for (let c = 0; c < 16384; c += 8) cells[c] = { kind: 'number' as const, value: c, style: UNLOCKED };
    sheet.rows.set(r, { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells });
  }
  const started = Date.now();
  const widths = sheet.columns.map((column) => ({ ...column, width: 12, customWidth: true }));
  assert.equal(reasonOf(workbook, [{ type: 'setColumns', sheet: 0, columns: widths }]), undefined);
  assert.equal(reasonOf(workbook, [{ type: 'delete', sheet: 0, axis: 'row', at: 10, count: 390 }]), undefined);
  assert.ok(Date.now() - started < 5000);
});

// Keeps a reference to Sheet for the type-only import.
export type _Sheet = Sheet;

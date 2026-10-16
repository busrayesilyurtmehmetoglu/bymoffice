// What the review of the protection of sheets found: a refused step left the formats its
// building made in the style sheet (the saved file was no longer the one that was read); the
// flags of the enhanced protection of xls files were lost when the header gave its size as -1;
// a hidden formula was found by Find; the legacy hash of a password longer than 15 characters
// was not openpyxl's; protectedCellIn() could name a position of an open range. Every file and
// workbook is built here.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { open, type Workbook } from '../src/index';
import { readXlsStream } from '../src/formats/xls/workbook';
import { setCellInput } from '../src/ops/commands';
import { findNextCell, replaceAllOps } from '../src/ops/find';
import { ProtectedError, cellProtected, checkProtectionPassword, legacyPasswordHash, protectedCellIn, protectionForbids } from '../src/ops/protection';
import { WorkbookView } from '../src/ui/sheet-view';
import { buildXlsx as makeXlsx, worksheetXml } from './fixtures';
import { cat, font, rec, u16, u32, workbookStream, xf } from './xls-fixtures';
import { bookOf, put, sheetOf } from './xlsx-build-fixtures';

// ---------------------------------------------------------------------------- B1: formats of a refused step

const lockedText = () =>
  makeXlsx({ sheets: [{ name: 'S', xml: worksheetXml('<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>x</t></is></c></row></sheetData><sheetProtection sheet="1" objects="1" scenarios="1"/>') }] });

/** The view's way of taking a step, on a view that has nothing but its workbook: refused steps stop before anything else is asked of it. */
function stepper(workbook: Workbook): { applyOps(step: unknown): boolean; refuseProtected(ops: unknown, mark?: unknown): void } {
  const view = Object.create(WorkbookView.prototype);
  view.workbook = workbook;
  return view;
}

test('a step that the protection refuses takes back the formats its building added: the file saves as it was read', async () => {
  const bytes = await lockedText();
  const file = await open(bytes);
  const workbook = await file.workbook();
  const formats = workbook.styles.cellXfs.length;
  // "50%" asks for a percentage format; the cell is locked.
  assert.throws(() => stepper(workbook).applyOps(() => setCellInput(workbook, 0, 0, 0, '50%')), ProtectedError);
  assert.equal(workbook.styles.cellXfs.length, formats);
  assert.equal(workbook.styles.modified, undefined);
  assert.deepEqual(await file.save(), bytes);
});

test('Replace All that is refused takes back the formats the job added (through refuseProtected with its mark)', async () => {
  const bytes = await lockedText();
  const file = await open(bytes);
  const workbook = await file.workbook();
  const view = stepper(workbook);
  const { markStyles } = await import('../src/ops/ops');
  const mark = markStyles(workbook.styles);
  const result = replaceAllOps(workbook, 'x', '1.5.2024', { locale: 'tr' } as never);
  assert.ok(result.ops.length > 0);
  assert.throws(() => view.refuseProtected(result.ops, mark), ProtectedError);
  assert.deepEqual(await file.save(), bytes);
});

// ---------------------------------------------------------------------------- B2: FEATHEADR with cbHdrData -1

test('xls: the enhanced protection is read whether the header gives its size as 4 or as -1 (as Excel and LibreOffice write it)', () => {
  const head = cat(font(), font(), font(), font(), ...Array.from({ length: 15 }, () => xf({ style: true })), xf({ parent: 0 }), xf({ parent: 0, locked: false }));
  const flags = (1 << 2) | (1 << 10) | (1 << 11) | (1 << 14);
  for (const size of [4, 0xffffffff]) {
    const feat = rec(0x0867, u16(0x0867), u16(0), [0, 0, 0, 0, 0, 0, 0, 0], u16(2), [1], u32(size), u32(flags));
    const { workbook } = readXlsStream(workbookStream({ head, sheets: [{ name: 'S', records: cat(rec(0x0012, u16(1)), feat) }] }));
    const sheet = workbook.sheets[0];
    assert.equal(protectionForbids(sheet, 'formatCells'), false, `size ${size}`);
    assert.equal(protectionForbids(sheet, 'sort'), false, `size ${size}`);
    assert.equal(protectionForbids(sheet, 'formatRows'), true, `size ${size}`);
  }
});

// ---------------------------------------------------------------------------- B3: hidden formulas and Find

function hiddenBook(): Workbook {
  const sheet = sheetOf('S');
  const workbook = bookOf(sheet);
  workbook.styles.cellXfs.push({ ...workbook.styles.cellXfs[0], protection: { hidden: true, locked: false } });
  put(sheet, 0, 0, { kind: 'number', value: 3, style: 1, formula: { text: 'SUM(7,8)', type: 'normal' } });
  sheet.protection = { attrs: [['sheet', '1']] };
  return workbook;
}

test('Find does not look into a formula the protection hides; its value is still found, and Replace leaves it alone', () => {
  const workbook = hiddenBook();
  const start = { sheet: 0, row: -1, col: -1 };
  assert.equal(findNextCell(workbook, 'SUM', start, { lookIn: 'formulas' }), undefined);
  assert.deepEqual(findNextCell(workbook, '3', start, { lookIn: 'values' }), { sheet: 0, row: 0, col: 0 });
  assert.equal(replaceAllOps(workbook, 'SUM', 'MAX').ops.length, 0);
  // Not protected: the formula is searched again.
  workbook.sheets[0].protection = undefined;
  assert.deepEqual(findNextCell(workbook, 'SUM', start, { lookIn: 'formulas' }), { sheet: 0, row: 0, col: 0 });
});

// ---------------------------------------------------------------------------- B4: legacy hash of long passwords

test('the legacy hash takes the whole password, as openpyxl does; a file hashed from its first 15 characters is opened too', async () => {
  // openpyxl.utils.protection.hash_password('a' * 20) and of 'a' * 15.
  assert.equal(legacyPasswordHash('a'.repeat(20)), 'B9DE');
  assert.equal(legacyPasswordHash('a'.repeat(15)), 'B1BB');
  assert.equal(legacyPasswordHash('secret'), 'DAA7');
  assert.equal(await checkProtectionPassword([['sheet', '1'], ['password', 'B9DE']], 'a'.repeat(20)), true);
  assert.equal(await checkProtectionPassword([['sheet', '1'], ['password', 'B1BB']], 'a'.repeat(20)), true);
  assert.equal(await checkProtectionPassword([['sheet', '1'], ['password', 'B9DE']], 'a'.repeat(19)), false);
  // Beyond the bits of a number: openpyxl's values, without rounding.
  assert.equal(legacyPasswordHash('Z'.repeat(60)), 'D80000000CE1B');
  assert.equal(legacyPasswordHash('a'.repeat(40)), '83FF4E22');
  assert.equal(legacyPasswordHash('\u015e'.repeat(20)), 'FD4B');
  assert.equal(legacyPasswordHash('xyz'.repeat(11)), '14BCE38');
});

// ---------------------------------------------------------------------------- B5: the position protectedCellIn() names

test('protectedCellIn() names a position that is locked, never one of an open range (random sheets against every position)', () => {
  let seed = 12345;
  const random = (n: number) => {
    seed = (seed * 1103515245 + 12345) & 0x7fffffff;
    return seed % n;
  };
  let named = 0;
  for (let iteration = 0; iteration < 3000; iteration++) {
    const sheet = sheetOf('S');
    const workbook = bookOf(sheet);
    workbook.styles.cellXfs.push({ ...workbook.styles.cellXfs[0], protection: { locked: false } });
    let at = 0;
    while (at < 12 && random(2)) {
      const min = at + random(3);
      const max = min + random(3);
      sheet.columns.push({ min, max, customWidth: false, hidden: false, outlineLevel: 0, collapsed: false, style: random(3) === 0 ? undefined : random(2) });
      at = max + 1;
    }
    for (let r = 0; r < 10; r++) {
      if (random(3) === 0) continue;
      for (let c = 0; c < 12; c++) if (random(3) === 0) put(sheet, r, c, { kind: 'number', value: 1, style: random(2) });
      if (!sheet.rows.get(r)) sheet.rows.set(r, { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells: [] });
      if (random(3) === 0) sheet.rows.get(r)!.style = random(2);
    }
    const ranges = [];
    for (let q = 0; q < 3; q++) {
      const top = random(10);
      const left = random(12);
      ranges.push({ attrs: [['sqref', `${String.fromCharCode(65 + left)}${top + 1}:${String.fromCharCode(65 + left + random(3))}${top + 1 + random(4)}`], ['name', 'x']] as [string, string][] });
    }
    sheet.protection = { attrs: [['sheet', '1']], ranges };
    const top = random(12);
    const left = random(14);
    const range = { top, left, bottom: top + random(5), right: left + random(5) };
    const got = protectedCellIn(workbook, sheet, range);
    let any = false;
    for (let r = range.top; r <= range.bottom && !any; r++) for (let c = range.left; c <= range.right && !any; c++) if (cellProtected(workbook, sheet, r, c)) any = true;
    assert.equal(!!got, any, `iteration ${iteration}`);
    if (got) {
      named++;
      assert.ok(got.row >= range.top && got.row <= range.bottom && got.col >= range.left && got.col <= range.right, `iteration ${iteration}: outside the range`);
      assert.ok(cellProtected(workbook, sheet, got.row, got.col), `iteration ${iteration}: ${JSON.stringify(got)} is not locked`);
    }
  }
  assert.ok(named > 1000);
});

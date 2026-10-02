// Operations, second round: formats that typed text asks for can be taken back
// when nothing is written after all, and references that name a sheet stay put
// when a formula moves with its row.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { LOCALE_TR } from '../src/format/locale';
import { createSheet, emptyStyleSheet, type Workbook } from '../src/model/sheet/types';
import { cellFromInput } from '../src/ops/commands';
import { markStyles, resetStyles, shiftUnqualified } from '../src/ops/ops';

function workbook(): Workbook {
  const sheet = createSheet({ name: 'Data', sheetId: 1, state: 'visible', kind: 'worksheet', part: '/xl/worksheets/sheet1.xml' });
  return { sheets: [sheet], styles: emptyStyleSheet(), sharedStrings: [], definedNames: [], date1904: false, activeSheet: 0 };
}

const snapshot = (wb: Workbook) => JSON.stringify({ ...wb.styles, numFmts: Array.from(wb.styles.numFmts.entries()) });

test('ops: formats added for typed text are taken back as if they had never been asked for', () => {
  const wb = workbook();
  const before = snapshot(wb);
  const mark = markStyles(wb.styles);
  // A date, a percentage and an amount each ask for a format of their own.
  const date = cellFromInput(wb, wb.sheets[0], 0, 0, '01.02.2024', { locale: LOCALE_TR });
  cellFromInput(wb, wb.sheets[0], 0, 1, '%12,5', { locale: LOCALE_TR });
  cellFromInput(wb, wb.sheets[0], 0, 2, '1.234,50 ₺', { locale: LOCALE_TR });
  assert.ok(date && date.style > 0);
  assert.equal(wb.styles.modified, true);
  assert.notEqual(snapshot(wb), before);
  resetStyles(wb.styles, mark);
  assert.equal(snapshot(wb), before);
  assert.equal(wb.styles.modified, undefined);
  assert.equal('modified' in wb.styles, false);

  // Formats that were there when the mark was made stay, and so does the note that the styles changed.
  cellFromInput(wb, wb.sheets[0], 0, 0, '01.02.2024', { locale: LOCALE_TR });
  const kept = snapshot(wb);
  const later = markStyles(wb.styles);
  cellFromInput(wb, wb.sheets[0], 0, 2, '1.234,50 ₺', { locale: LOCALE_TR });
  resetStyles(wb.styles, later);
  assert.equal(snapshot(wb), kept);
  assert.equal(wb.styles.modified, true);
  // Nothing added: nothing happens.
  resetStyles(wb.styles, markStyles(wb.styles));
  assert.equal(snapshot(wb), kept);
});

test('ops: moving a formula with its row leaves references that name a sheet alone', () => {
  const cases: [string, number, string][] = [
    ['D6+Other!A1', 16, 'D22+Other!A1'],
    ['D6+Other!A1', -5, 'D1+Other!A1'],
    ['D6+Other!A1', -6, '#REF!+Other!A1'],
    ["A1+'My data'!A1:B2+'It''s'!$A1", 2, "A3+'My data'!A1:B2+'It''s'!$A1"],
    ['Data!D5*D5', 1, 'Data!D5*D6'],
    ['SUM(Other!A:A,A:A,Other!1:2,1:2)', 3, 'SUM(Other!A:A,A:A,Other!1:2,4:5)'],
    ['First:Last!A1+A1', 1, 'First:Last!A1+A2'],
    ['[1]Sheet1!A1+[2]Sheet2!$B$2+A1', 1, '[1]Sheet1!A1+[2]Sheet2!$B$2+A2'],
    // Text and the names of sheets are never read as references.
    ['"Other!A1"&A1&"it""s!B2"', 1, '"Other!A1"&A2&"it""s!B2"'],
    ["'A1!B2'!C3+C3", 1, "'A1!B2'!C3+C4"],
    // A name of a sheet's own is not a cell; an error constant ends with "!" too.
    ['Other!Total+A1', 1, 'Other!Total+A2'],
    ['IF(A1=#REF!,#N/A,#VALUE!+A1)', 1, 'IF(A2=#REF!,#N/A,#VALUE!+A2)'],
    ['Other!A1:INDEX(B:B,2)+B2', 1, 'Other!A1:INDEX(B:B,2)+B3'],
    ['Table1[[#This Row],[a!b1]]+B2', 1, 'Table1[[#This Row],[a!b1]]+B3'],
    // Without a sheet name everything moves, as it always did; anchored parts stay.
    ['A1+$A1+A$1+$A$1', 2, 'A3+$A3+A$1+$A$1'],
    ['Other!A1', 0, 'Other!A1'],
  ];
  for (const [text, rows, expected] of cases) assert.equal(shiftUnqualified(text, rows), expected, `${text} by ${rows}`);
  // Odd text comes back as text, whatever it is.
  for (const odd of ['', '!', '!!', "'", '"', '[', 'A1!', '!A1', "'x'!", '#REF!A1', 'Other!A1048576']) assert.equal(typeof shiftUnqualified(odd, 3), 'string', odd);
  assert.equal(shiftUnqualified('!A1+A1', 1), '!A1+A2');
});

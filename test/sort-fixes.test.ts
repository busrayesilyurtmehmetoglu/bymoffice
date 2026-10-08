// Sorting, second round: the order of accented letters against an independent
// collator, case folding by locale, references that name a sheet, whether a
// block has a header row, and what a sort notes in the filter.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { LOCALE_EN, LOCALE_TR } from '../src/format/locale';
import { createSheet, emptyStyleSheet, getCell, type Cell, type Row, type Sheet, type Workbook } from '../src/model/sheet/types';
import { History, applyOp, type Op } from '../src/ops/ops';
import { collationKey, foldCase, hasHeaderRow, sortRange, sortValueOf, SORT_NUMBER, SORT_TEXT } from '../src/ops/sort';

type Value = number | string | boolean | null | { f: string } | { date: string } | { s: number; v: number | string };

function cellOf(value: Value): Cell | undefined {
  if (value === null) return undefined;
  if (typeof value === 'number') return { kind: 'number', value, style: 0 };
  if (typeof value === 'string') return { kind: 'string', value, style: 0 };
  if (typeof value === 'boolean') return { kind: 'boolean', value, style: 0 };
  if ('date' in value) return { kind: 'date', value: value.date, style: 0 };
  if ('s' in value) return { kind: typeof value.v === 'number' ? 'number' : 'string', value: value.v, style: value.s };
  return { kind: 'blank', value: null, style: 0, formula: { text: value.f, type: 'normal' } };
}

function sheetOf(grid: Value[][], name = 'Data', sheetId = 1): Sheet {
  const sheet = createSheet({ name, sheetId, state: 'visible', kind: 'worksheet', part: `/xl/worksheets/sheet${sheetId}.xml` });
  grid.forEach((values, r) => {
    const cells: Cell[] = [];
    values.forEach((value, c) => {
      const cell = cellOf(value);
      if (cell) {
        cells[c] = cell;
        sheet.colCount = Math.max(sheet.colCount, c + 1);
      }
    });
    if (cells.length > 0) {
      sheet.rows.set(r, { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells });
      sheet.rowCount = r + 1;
    }
  });
  return sheet;
}

const bookOf = (...sheets: Sheet[]): Workbook => ({ sheets, styles: emptyStyleSheet(), sharedStrings: [], definedNames: [], date1904: false, activeSheet: 0 });

const column = (sheet: Sheet, col: number, top: number, bottom: number) => {
  const out: unknown[] = [];
  for (let r = top; r <= bottom; r++) {
    const cell = getCell(sheet, r, col);
    out.push(cell ? (cell.formula ? `=${cell.formula.text}` : cell.value) : null);
  }
  return out;
};

const apply = (workbook: Workbook, ops: Op[]): Op[] => ops.map((op) => applyOp(workbook, op)).reverse();

/** A generator of numbers in [0, 1) that gives the same run every time. */
function seeded(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0;
    return state / 4294967296;
  };
}

const byKey = (locale: typeof LOCALE_EN) => (a: string, b: string) => {
  const ka = collationKey(a, locale);
  const kb = collationKey(b, locale);
  return ka < kb ? -1 : ka > kb ? 1 : 0;
};

// Letters of Latin-1 Supplement and Latin Extended-A, lower and upper case. Left out are the
// letters the reference collator gives a place of their own in the alphabet, where this engine
// folds them to a base letter as the Windows sort of Excel does: thorn, eng, t with stroke,
// kra, n with apostrophe (and, in English only, the dotless i).
const OWN_PLACE = 'þÞŋŊŧŦĸŉ';
function latinLetters(skip: string): string {
  let out = '';
  for (let code = 0xc0; code <= 0x17f; code++) {
    const char = String.fromCharCode(code);
    if (code === 0xd7 || code === 0xf7 || skip.indexOf(char) >= 0) continue;
    out += char;
  }
  return out;
}

function randomWords(count: number, alphabet: string, seed: number): string[] {
  const random = seeded(seed);
  const pick = (from: string) => from.charAt(Math.floor(random() * from.length));
  const words: string[] = [];
  // A small set of base letters makes words that differ by accents or case only come up often.
  const plain = 'aceinosuAEIOU z9';
  for (let i = 0; i < count; i++) {
    const length = 1 + Math.floor(random() * 5);
    let word = '';
    for (let k = 0; k < length; k++) word += random() < 0.55 ? pick(plain) : pick(alphabet);
    // Leading and trailing spaces say nothing about letters.
    words.push(word.trim() === '' ? 'a' : word.trim());
  }
  return words;
}

test('sort: English text - 2,000 random words come out in the order of an independent collator', () => {
  const reference = new Intl.Collator('en', { sensitivity: 'accent' });
  const alphabet = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 ' + latinLetters(OWN_PLACE + 'ı');
  const words = randomWords(2000, alphabet, 20240117);
  const ours = words.slice().sort(byKey(LOCALE_EN));
  const theirs = words.slice().sort(reference.compare);
  assert.deepEqual(ours, theirs);
  // The letters left out of the comparison are folded to a base letter too.
  for (const [word, low, high] of [
    ['ırmak', 'irmak', 'is'],
    ['þorn', 'thorn', 'ti'],
    ['ŋa', 'na', 'nb'],
    ['ŧa', 'ta', 'tb'],
  ]) {
    assert.ok(collationKey(low, LOCALE_EN) < collationKey(word, LOCALE_EN), `${low} < ${word}`);
    assert.ok(collationKey(word, LOCALE_EN) < collationKey(high, LOCALE_EN), `${word} < ${high}`);
  }
});

test('sort: Turkish text - 2,000 random words come out in the order of an independent collator', () => {
  const reference = new Intl.Collator('tr', { sensitivity: 'accent' });
  const alphabet = 'abcçdefgğhıijklmnoöpqrsştuüvwxyzABCÇDEFGĞHIİJKLMNOÖPQRSŞTUÜVWXYZ0123456789 ' + latinLetters(OWN_PLACE);
  const words = randomWords(2000, alphabet, 19231029);
  const ours = words.slice().sort(byKey(LOCALE_TR));
  const theirs = words.slice().sort(reference.compare);
  assert.deepEqual(ours, theirs);
});

test('sort: accented letters sort with their base letter, the plain one first', () => {
  const sorted = (words: string[], locale: typeof LOCALE_EN) => words.slice().sort(byKey(locale));
  // Turkish: letters that are not in the alphabet follow their base letter, they do not go after z.
  assert.deepEqual(sorted(['kas', 'zam', 'kâr', 'kara', 'hâl', 'kar', 'hala', 'hal'], LOCALE_TR), ['hal', 'hâl', 'hala', 'kar', 'kâr', 'kara', 'kas', 'zam']);
  assert.ok(collationKey('a', LOCALE_TR) < collationKey('â', LOCALE_TR));
  assert.ok(collationKey('â', LOCALE_TR) < collationKey('b', LOCALE_TR));
  assert.ok(collationKey('é', LOCALE_TR) < collationKey('f', LOCALE_TR));
  // English: every letter outside ASCII is folded to its base letter.
  assert.deepEqual(sorted(['zebra', 'çay', 'Şule', 'ırmak', 'Öz', 'cay', 'sule', 'oz', 'irmak', 'ğ', 'üst', 'g', 'ust'], LOCALE_EN), [
    'cay',
    'çay',
    'g',
    'ğ',
    'irmak',
    'ırmak',
    'oz',
    'Öz',
    'sule',
    'Şule',
    'ust',
    'üst',
    'zebra',
  ]);
  assert.deepEqual(sorted(['résumé', 'resumes', 'resume', 'Resumé'], LOCALE_EN), ['resume', 'Resumé', 'résumé', 'resumes']);
  // Plain ASCII keeps the key it always had.
  assert.equal(collationKey('ABC def', LOCALE_EN), 'abc def');
  // A sort of cells uses the same order.
  const sheet = sheetOf([['zam'], ['kâr'], ['kas'], ['kar']]);
  const workbook = bookOf(sheet);
  apply(workbook, sortRange(workbook, 0, { top: 0, left: 0, bottom: 3, right: 0 }, [{ col: 0 }], { locale: LOCALE_TR }));
  assert.deepEqual(column(sheet, 0, 0, 3), ['kar', 'kâr', 'kas', 'zam']);
});

test('sort: case folding follows the locale - Turkish pairs I with ı and İ with i, others I with i', () => {
  assert.equal(foldCase('KIL', LOCALE_TR), foldCase('kıl', LOCALE_TR));
  assert.equal(foldCase('KİL', LOCALE_TR), foldCase('kil', LOCALE_TR));
  assert.notEqual(foldCase('kıl', LOCALE_TR), foldCase('kil', LOCALE_TR));
  assert.equal(foldCase('ISTANBUL', LOCALE_TR), 'ıstanbul');
  assert.equal(foldCase('İSTANBUL', LOCALE_TR), 'istanbul');
  // Elsewhere "I" is the capital of "i"; the two Turkish letters stand for themselves.
  assert.equal(foldCase('KIL', LOCALE_EN), 'kil');
  assert.equal(foldCase('KIL'), 'kil');
  assert.notEqual(foldCase('kıl', LOCALE_EN), foldCase('kil', LOCALE_EN));
  assert.notEqual(foldCase('KİL', LOCALE_EN), foldCase('kil', LOCALE_EN));
  assert.notEqual(foldCase('KİL', LOCALE_EN), foldCase('kıl', LOCALE_EN));
  // One character stays one character, so that "?" in a pattern still stands for it.
  assert.equal(foldCase('İ', LOCALE_EN).length, 1);
  assert.equal(foldCase('ÇĞÖŞÜ', LOCALE_EN), 'çğöşü');
  // A locale built by hand counts as Turkish by its month names (Azerbaijani has the same pairs).
  const azeri = { ...LOCALE_TR, monthNames: ['Yanvar'].concat(LOCALE_TR.monthNames.slice(1)) };
  assert.equal(foldCase('KIL', azeri), 'kıl');
});

test('sort: a date cell counts as the number it stands for', () => {
  assert.deepEqual(sortValueOf({ kind: 'date', value: '2024-03-15T00:00:00', style: 0 }, false), { rank: SORT_NUMBER, number: 45366, text: '' });
  assert.equal(sortValueOf({ kind: 'date', value: 'not a date', style: 0 }, false).rank, SORT_TEXT);
  // Dates given as text sort among the numbers, by their day, and before any text.
  const sheet = sheetOf([[{ date: '2024-03-15' }], ['text'], [45000], [{ date: '2023-01-01' }], [46000]]);
  const workbook = bookOf(sheet);
  apply(workbook, sortRange(workbook, 0, { top: 0, left: 0, bottom: 4, right: 0 }, [{ col: 0 }]));
  assert.deepEqual(column(sheet, 0, 0, 4), ['2023-01-01', 45000, '2024-03-15', 46000, 'text']);
});

test('sort: tables and pivot tables in the sheet, and comments the model does not hold, refuse a sort, each by name', () => {
  for (const part of ['comments', 'table', 'pivotTable']) {
    const sheet = sheetOf([[2], [1]]);
    sheet.anchored = ['drawing', part];
    // Notes go with their rows; comments the reader could not show would stay behind.
    if (part === 'comments') sheet.warnings = ['Comments: 1 comment(s) could not be read and are not shown'];
    const workbook = bookOf(sheet);
    const ops = sortRange(workbook, 0, { top: 0, left: 0, bottom: 1, right: 0 }, [{ col: 0 }]);
    assert.throws(() => apply(workbook, ops), new RegExp(`^RangeError: Cannot sort in a sheet with ${part}: not supported yet$`));
    assert.deepEqual(column(sheet, 0, 0, 1), [2, 1]);
  }
  // Pictures alone are no obstacle.
  const sheet = sheetOf([[2], [1]]);
  sheet.anchored = ['drawing', 'image'];
  const workbook = bookOf(sheet);
  apply(workbook, sortRange(workbook, 0, { top: 0, left: 0, bottom: 1, right: 0 }, [{ col: 0 }]));
  assert.deepEqual(column(sheet, 0, 0, 1), [1, 2]);
});

test('sort: references that name a sheet stay as they are; references without one move with the row', () => {
  const data = sheetOf(
    [
      [3, { f: 'A1+Other!A1' }, { f: "SUM('My Sheet'!A1:B2,Data!A1,$A1,A$1)" }],
      [1, { f: 'A2+Other!A2' }, { f: 'Other!$A2&"Other!A2"&A2' }],
      [2, { f: 'A3+Other!A3+Other!A:A+Other!3:3' }, { f: 'First:Last!A3+[1]Ext!A3+B3' }],
    ],
    'Data',
  );
  const workbook = bookOf(data, sheetOf([[10], [20], [30]], 'Other', 2), sheetOf([[1, 2], [3, 4]], 'My Sheet', 3));
  const history = new History(workbook);
  history.apply(sortRange(workbook, 0, { top: 0, left: 0, bottom: 2, right: 2 }, [{ col: 0 }]), undefined, undefined);
  assert.deepEqual(column(data, 0, 0, 2), [1, 2, 3]);
  // Row 2 went to row 1, row 3 to row 2, row 1 to row 3.
  assert.deepEqual(column(data, 1, 0, 2), ['=A1+Other!A2', '=A2+Other!A3+Other!A:A+Other!3:3', '=A3+Other!A1']);
  assert.deepEqual(column(data, 2, 0, 2), ['=Other!$A2&"Other!A2"&A1', '=First:Last!A3+[1]Ext!A3+B2', "=SUM('My Sheet'!A1:B2,Data!A1,$A3,A$1)"]);
  history.undo();
  assert.deepEqual(column(data, 1, 0, 2), ['=A1+Other!A1', '=A2+Other!A2', '=A3+Other!A3+Other!A:A+Other!3:3']);
});

test('sort: a row moved up does not turn a reference to another sheet into #REF!', () => {
  const data = sheetOf([[9, { f: 'A1' }], [8, { f: 'A2' }], [7, { f: 'A3' }], [1, { f: 'A4+Other!A1+A1' }]]);
  const workbook = bookOf(data, sheetOf([[10]], 'Other', 2));
  apply(workbook, sortRange(workbook, 0, { top: 0, left: 0, bottom: 3, right: 1 }, [{ col: 0 }]));
  // The reference without a sheet name that would leave the sheet still does, as when a cell is copied.
  assert.equal(getCell(data, 0, 1)!.formula!.text, 'A1+Other!A1+#REF!');
});

test('sort: whether the first row of a block is a header - by the kinds of its values, or by bold type', () => {
  const none = bookOf(sheetOf([[3], [1], [2]]));
  assert.equal(hasHeaderRow(none, none.sheets[0], { top: 0, left: 0, bottom: 2, right: 0 }), false);
  const titled = bookOf(sheetOf([['Amount'], [1], [2]]));
  assert.equal(hasHeaderRow(titled, titled.sheets[0], { top: 0, left: 0, bottom: 2, right: 0 }), true);
  // One column that differs is enough; a column of text under text says nothing.
  const mixed = bookOf(sheetOf([['Name', 'Amount'], ['b', 2], ['a', 1]]));
  assert.equal(hasHeaderRow(mixed, mixed.sheets[0], { top: 0, left: 0, bottom: 2, right: 1 }), true);
  const text = bookOf(sheetOf([['pear'], ['apple'], ['fig']]));
  assert.equal(hasHeaderRow(text, text.sheets[0], { top: 0, left: 0, bottom: 2, right: 0 }), false);
  // The kind most of the cells below have counts; dates are numbers; empty cells say nothing.
  const mostly = bookOf(sheetOf([[2024], ['n/a'], [1], [2], [null], [3]]));
  assert.equal(hasHeaderRow(mostly, mostly.sheets[0], { top: 0, left: 0, bottom: 5, right: 0 }), false);
  const dates = bookOf(sheetOf([[{ date: '2024-01-01' }], [45000], [45001]]));
  assert.equal(hasHeaderRow(dates, dates.sheets[0], { top: 0, left: 0, bottom: 2, right: 0 }), false);
  const gaps = bookOf(sheetOf([[null, 'x'], [1, 'y'], [2, 'z']]));
  assert.equal(hasHeaderRow(gaps, gaps.sheets[0], { top: 0, left: 0, bottom: 2, right: 1 }), false);
  // A single row has nothing under it.
  assert.equal(hasHeaderRow(titled, titled.sheets[0], { top: 0, left: 0, bottom: 0, right: 0 }), false);

  // A first row in bold above rows that are not is a header, whatever it holds.
  const bold = bookOf(sheetOf([[{ s: 1, v: 'pear' }], ['apple'], ['fig']]));
  const styles = bold.styles;
  styles.fonts.push({ ...styles.fonts[0], bold: true });
  styles.cellXfs.push({ ...styles.cellXfs[0], fontId: styles.fonts.length - 1 });
  assert.equal(hasHeaderRow(bold, bold.sheets[0], { top: 0, left: 0, bottom: 2, right: 0 }), true);
  const allBold = bookOf(sheetOf([[{ s: 1, v: 'pear' }], [{ s: 1, v: 'apple' }], [{ s: 1, v: 'fig' }]]));
  allBold.styles.fonts.push({ ...allBold.styles.fonts[0], bold: true });
  allBold.styles.cellXfs.push({ ...allBold.styles.cellXfs[0], fontId: allBold.styles.fonts.length - 1 });
  assert.equal(hasHeaderRow(allBold, allBold.sheets[0], { top: 0, left: 0, bottom: 2, right: 0 }), false);
});

test('sort: the filter notes the range that was sorted, not all of its rows', () => {
  const sheet = sheetOf([['Name', 'n'], ['c', 3], ['a', 1], ['b', 2], ['z', 9], ['y', 8]]);
  sheet.autoFilter = { range: { top: 0, left: 0, bottom: 5, right: 1 }, columns: [] };
  const workbook = bookOf(sheet);
  const history = new History(workbook);
  // The first three data rows only.
  history.apply(sortRange(workbook, 0, { top: 1, left: 0, bottom: 3, right: 1 }, [{ col: 1 }]), undefined, undefined);
  assert.deepEqual(column(sheet, 0, 0, 5), ['Name', 'a', 'b', 'c', 'z', 'y']);
  assert.deepEqual(sheet.autoFilter!.sortState, {
    range: { top: 1, left: 0, bottom: 3, right: 1 },
    conditions: [{ range: { top: 1, left: 1, bottom: 3, right: 1 }, descending: false }],
  });
  // A hidden row inside the filter stays where it is and is still part of what was sorted.
  (sheet.rows.get(4) as Row).hidden = true;
  history.apply(sortRange(workbook, 0, { top: 0, left: 0, bottom: 5, right: 1 }, [{ col: 0, descending: true }], { header: true }), undefined, undefined);
  assert.deepEqual(column(sheet, 0, 0, 5), ['Name', 'y', 'c', 'b', 'z', 'a']);
  assert.deepEqual(sheet.autoFilter!.sortState!.range, { top: 1, left: 0, bottom: 5, right: 1 });
});

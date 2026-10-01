import assert from 'node:assert/strict';
import { test } from 'node:test';
import { DEFAULT_INDEXED_COLORS, formatCodeOf, resolveCellStyle, resolveColor } from '../src/model/sheet/color';
import { emptyStyleSheet } from '../src/model/sheet/types';
import type { StyleSheet } from '../src/model/sheet/types';

// The Office 2013+ theme in spreadsheet index order (lt1, dk1, lt2, dk2, accents, links).
const OFFICE_THEME = [
  'FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31',
  'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72',
];

function themed(): StyleSheet {
  const styles = emptyStyleSheet();
  styles.themeColors = OFFICE_THEME.slice();
  return styles;
}

/** Asserts that every channel of "#RRGGBB" is within one step of the expected colour. */
function assertNear(actual: string | undefined, expected: string): void {
  assert.ok(actual !== undefined && /^#[0-9A-F]{6}$/.test(actual), `not a colour: ${actual}`);
  for (let i = 0; i < 3; i++) {
    const a = parseInt(actual.slice(1 + i * 2, 3 + i * 2), 16);
    const e = parseInt(expected.slice(i * 2, i * 2 + 2), 16);
    assert.ok(Math.abs(a - e) <= 1, `${actual} is not within 1 of #${expected}`);
  }
}

test('color: default indexed palette has the 64 standard entries', () => {
  assert.equal(DEFAULT_INDEXED_COLORS.length, 64);
  const known: Record<number, string> = {
    0: '000000', 1: 'FFFFFF', 2: 'FF0000', 8: '000000', 9: 'FFFFFF', 10: 'FF0000', 22: 'C0C0C0', 63: '333333',
  };
  for (const index of Object.keys(known)) assert.equal(DEFAULT_INDEXED_COLORS[Number(index)], known[Number(index)]);
  for (const entry of DEFAULT_INDEXED_COLORS) assert.match(entry, /^[0-9A-F]{6}$/);
});

test('color: rgb resolves, dropping alpha and normalising case', () => {
  const styles = emptyStyleSheet();
  assert.equal(resolveColor({ rgb: 'FF4472C4' }, styles), '#4472C4');
  assert.equal(resolveColor({ rgb: '4472C4' }, styles), '#4472C4');
  assert.equal(resolveColor({ rgb: '00ff00aa' }, styles), '#FF00AA');
  assert.equal(resolveColor({ rgb: 'nothex' }, styles), undefined);
  assert.equal(resolveColor({ rgb: '' }, styles), undefined);
});

test('color: undefined, automatic and empty colours do not resolve', () => {
  const styles = themed();
  assert.equal(resolveColor(undefined, styles), undefined);
  assert.equal(resolveColor({ auto: true }, styles), undefined);
  assert.equal(resolveColor({}, styles), undefined);
});

test('color: indexed uses the default palette, and 64/65 are automatic', () => {
  const styles = emptyStyleSheet();
  assert.equal(resolveColor({ indexed: 2 }, styles), '#FF0000');
  assert.equal(resolveColor({ indexed: 22 }, styles), '#C0C0C0');
  assert.equal(resolveColor({ indexed: 63 }, styles), '#333333');
  assert.equal(resolveColor({ indexed: 64 }, styles), undefined);
  assert.equal(resolveColor({ indexed: 65 }, styles), undefined);
  assert.equal(resolveColor({ indexed: 5000 }, styles), undefined);
  assert.equal(resolveColor({ indexed: -1 }, styles), undefined);
});

test('color: a custom palette overrides the default one', () => {
  const styles = emptyStyleSheet();
  styles.indexedColors = ['111111', '222222', '333333'];
  assert.equal(resolveColor({ indexed: 2 }, styles), '#333333');
  // Beyond a short custom palette the standard entry still answers.
  assert.equal(resolveColor({ indexed: 22 }, styles), '#C0C0C0');
  assert.equal(resolveColor({ indexed: 64 }, styles), undefined);
  styles.indexedColors = [];
  assert.equal(resolveColor({ indexed: 2 }, styles), '#FF0000');
});

test('color: theme colours resolve by index; a missing theme colour is undefined', () => {
  const styles = themed();
  assert.equal(resolveColor({ theme: 0 }, styles), '#FFFFFF');
  assert.equal(resolveColor({ theme: 1 }, styles), '#000000');
  assert.equal(resolveColor({ theme: 4 }, styles), '#4472C4');
  assert.equal(resolveColor({ theme: 12 }, styles), undefined);
  assert.equal(resolveColor({ theme: 4 }, emptyStyleSheet()), undefined);
  assert.equal(resolveColor({ theme: 4, tint: 0.5 }, emptyStyleSheet()), undefined);
  styles.themeColors[4] = '';
  assert.equal(resolveColor({ theme: 4 }, styles), undefined);
});

test('color: tint matches the colours Excel shows', () => {
  const styles = themed();
  // Accent 1 "Lighter 80%", "Lighter 60%", "Lighter 40%", "Darker 25%", "Darker 50%".
  assertNear(resolveColor({ theme: 4, tint: 0.7999816888943144 }, styles), 'DAE3F3');
  assertNear(resolveColor({ theme: 4, tint: 0.5999938962981048 }, styles), 'B4C7E7');
  assertNear(resolveColor({ theme: 4, tint: 0.3999755851924192 }, styles), '8FAADC');
  assertNear(resolveColor({ theme: 4, tint: -0.249977111117893 }, styles), '2F5597');
  assertNear(resolveColor({ theme: 4, tint: -0.499984740745262 }, styles), '203864');
  // White "Darker 15%" and black "Lighter 50%" (greys).
  assertNear(resolveColor({ theme: 0, tint: -0.1499984740745262 }, styles), 'D9D9D9');
  assertNear(resolveColor({ theme: 1, tint: 0.499984740745262 }, styles), '7F7F7F');
});

test('color: tint limits and odd values', () => {
  const styles = themed();
  assert.equal(resolveColor({ theme: 4, tint: 0 }, styles), '#4472C4');
  assert.equal(resolveColor({ theme: 4, tint: 1 }, styles), '#FFFFFF');
  assert.equal(resolveColor({ theme: 4, tint: -1 }, styles), '#000000');
  assert.equal(resolveColor({ theme: 4, tint: 7 }, styles), '#FFFFFF');
  assert.equal(resolveColor({ theme: 4, tint: NaN }, styles), '#4472C4');
  // Tint applies to explicit and indexed colours too.
  assert.equal(resolveColor({ rgb: 'FF000000', tint: 1 }, styles), '#FFFFFF');
  assertNear(resolveColor({ indexed: 2, tint: -0.5 }, styles), '800000');
});

test('numFmt: built-in codes', () => {
  const styles = emptyStyleSheet();
  const expected: Record<number, string> = {
    0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00',
    12: '# ?/?', 13: '# ??/??', 14: 'mm-dd-yy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM',
    19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'm/d/yy h:mm', 37: '#,##0_);(#,##0)',
    38: '#,##0_);[Red](#,##0)', 39: '#,##0.00_);(#,##0.00)', 40: '#,##0.00_);[Red](#,##0.00)', 45: 'mm:ss',
    46: '[h]:mm:ss', 47: 'mmss.0', 48: '##0.0E+0', 49: '@',
  };
  for (const id of Object.keys(expected)) assert.equal(formatCodeOf(styles, Number(id)), expected[Number(id)]);
  assert.equal(formatCodeOf(styles, 5), '$#,##0_);($#,##0)');
  assert.equal(formatCodeOf(styles, 8), '$#,##0.00_);[Red]($#,##0.00)');
  assert.equal(formatCodeOf(styles, 44), '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)');
  for (let id = 5; id <= 8; id++) assert.notEqual(formatCodeOf(styles, id), 'General');
  for (let id = 41; id <= 44; id++) assert.notEqual(formatCodeOf(styles, id), 'General');
});

test('numFmt: unknown ids read as General, custom formats win over built-in ones', () => {
  const styles = emptyStyleSheet();
  assert.equal(formatCodeOf(styles, 23), 'General');
  assert.equal(formatCodeOf(styles, 164), 'General');
  assert.equal(formatCodeOf(styles, -3), 'General');
  assert.equal(formatCodeOf(styles, NaN), 'General');
  styles.numFmts.set(164, 'dd.mm.yyyy');
  styles.numFmts.set(14, 'd.m.yyyy');
  assert.equal(formatCodeOf(styles, 164), 'dd.mm.yyyy');
  assert.equal(formatCodeOf(styles, 14), 'd.m.yyyy');
});

test('resolveCellStyle: follows the indexes of a cell format', () => {
  const styles = emptyStyleSheet();
  styles.fonts.push({ name: 'Arial', size: 14, bold: true, italic: false, strike: false, underline: 'none' });
  styles.fills.push({ type: 'pattern', pattern: 'solid', fg: { rgb: 'FFFFFF00' } });
  styles.borders.push({ diagonalUp: false, diagonalDown: false, top: { style: 'thin' } });
  styles.numFmts.set(164, '0.000');
  styles.cellXfs.push({
    numFmtId: 164, fontId: 1, fillId: 1, borderId: 1, xfId: 0, quotePrefix: false,
    alignment: { horizontal: 'center', wrapText: true },
  });

  const resolved = resolveCellStyle(styles, 1);
  assert.equal(resolved.font, styles.fonts[1]);
  assert.equal(resolved.fill, styles.fills[1]);
  assert.equal(resolved.border, styles.borders[1]);
  assert.deepEqual(resolved.alignment, { horizontal: 'center', wrapText: true });
  assert.equal(resolved.numFmt, '0.000');

  const first = resolveCellStyle(styles, 0);
  assert.equal(first.font.name, 'Calibri');
  assert.deepEqual(first.alignment, {});
  assert.equal(first.numFmt, 'General');
});

test('resolveCellStyle: out-of-range indexes fall back instead of throwing', () => {
  const styles = emptyStyleSheet();
  for (const index of [1, 99999, -1, NaN]) {
    const resolved = resolveCellStyle(styles, index);
    assert.equal(resolved.font, styles.fonts[0]);
    assert.deepEqual(resolved.alignment, {});
  }

  // A hand-built sheet whose format points outside its tables.
  styles.cellXfs[0] = { numFmtId: 0, fontId: 7, fillId: 7, borderId: 7, quotePrefix: false };
  const damaged = resolveCellStyle(styles, 0);
  assert.equal(damaged.font, styles.fonts[0]);
  assert.equal(damaged.fill, styles.fills[0]);
  assert.equal(damaged.border, styles.borders[0]);

  // Even entirely empty tables give something drawable.
  const empty: StyleSheet = { ...emptyStyleSheet(), fonts: [], fills: [], borders: [], cellXfs: [] };
  const fallback = resolveCellStyle(empty, 3);
  assert.equal(fallback.font.bold, false);
  assert.deepEqual(fallback.fill, { type: 'pattern', pattern: 'none' });
  assert.equal(fallback.numFmt, 'General');
});

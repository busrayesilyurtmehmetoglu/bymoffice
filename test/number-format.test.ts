import assert from 'node:assert/strict';
import { test } from 'node:test';
import { getLocale, LOCALE_EN, LOCALE_TR, normalizeLocale, type FormatLocale } from '../src/format/locale';
import { compiledFormatCount, formatValue, isDateFormat, serialToDate } from '../src/format/number-format';

type Value = number | string | boolean | null;
/** value, format code, locale, expected text */
type Row = [Value, string, 'en' | 'tr', string];

/** Fraction of a day for a time of day. */
function time(hour: number, minute: number, second: number, millisecond = 0): number {
  return (hour * 3600 + minute * 60 + second + millisecond / 1000) / 86400;
}

/** 15 March 2024, a Friday. */
const DAY = 45366;
const AFTERNOON = DAY + time(13, 45, 10);
const MORNING = DAY + time(9, 5, 7);

const ACCOUNTING_41 = '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)';
const ACCOUNTING_42 = '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)';
const ACCOUNTING_43 = '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)';
const ACCOUNTING_44 = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)';

function check(rows: Row[]): void {
  for (const [value, code, locale, expected] of rows) {
    const actual = formatValue(value, code, { locale: locale === 'tr' ? LOCALE_TR : LOCALE_EN });
    assert.equal(actual.text, expected, `${JSON.stringify(value)} with ${JSON.stringify(code)} (${locale})`);
  }
}

test('number format: built-in codes 0-22 and 37-49', () => {
  check([
    // 0 General
    [1234.5, 'General', 'en', '1234.5'],
    [1234.5, 'General', 'tr', '1234,5'],
    [-0.25, 'General', 'tr', '-0,25'],
    [1e11, 'General', 'en', '1E+11'],
    [1.23456789e15, 'General', 'tr', '1,23457E+15'],
    // 1, 2
    [1234.5, '0', 'en', '1235'],
    [-1234.5, '0', 'en', '-1235'],
    [0.4, '0', 'en', '0'],
    [1234.5, '0.00', 'en', '1234.50'],
    [1234.5, '0.00', 'tr', '1234,50'],
    [-7, '0.00', 'en', '-7.00'],
    // 3, 4
    [1234.5, '#,##0', 'en', '1,235'],
    [1234.5, '#,##0', 'tr', '1.235'],
    [999, '#,##0', 'en', '999'],
    [0, '#,##0', 'en', '0'],
    [1234567.891, '#,##0.00', 'en', '1,234,567.89'],
    [1234567.891, '#,##0.00', 'tr', '1.234.567,89'],
    [-1234567.891, '#,##0.00', 'tr', '-1.234.567,89'],
    // 9, 10
    [0.256, '0%', 'en', '26%'],
    [1, '0%', 'en', '100%'],
    [-0.5, '0%', 'tr', '-50%'],
    [0.2567, '0.00%', 'en', '25.67%'],
    [0.2567, '0.00%', 'tr', '25,67%'],
    // 11
    [12345.678, '0.00E+00', 'en', '1.23E+04'],
    [12345.678, '0.00E+00', 'tr', '1,23E+04'],
    [-0.00012345, '0.00E+00', 'en', '-1.23E-04'],
    [0, '0.00E+00', 'en', '0.00E+00'],
    // 12, 13
    [2.5, '# ?/?', 'en', '2 1/2'],
    [-2.5, '# ?/?', 'en', '-2 1/2'],
    [5.25, '# ?/?', 'en', '5 1/4'],
    [5.25, '# ??/??', 'en', '5  1/4 '],
    [5.3, '# ??/??', 'en', '5  3/10'],
    // 14: the code as written; the system short date needs the built-in id (tested below)
    [DAY, 'mm-dd-yy', 'en', '03-15-24'],
    [DAY, 'mm-dd-yy', 'tr', '03-15-24'],
    // 15, 16, 17
    [DAY, 'd-mmm-yy', 'en', '15-Mar-24'],
    [45505, 'd-mmm-yy', 'tr', '1-Ağu-24'],
    [DAY, 'd-mmm', 'en', '15-Mar'],
    [45505, 'd-mmm', 'tr', '1-Ağu'],
    [DAY, 'mmm-yy', 'en', 'Mar-24'],
    [45505, 'mmm-yy', 'tr', 'Ağu-24'],
    // 18, 19
    [AFTERNOON, 'h:mm AM/PM', 'en', '1:45 PM'],
    [MORNING, 'h:mm AM/PM', 'en', '9:05 AM'],
    [AFTERNOON, 'h:mm AM/PM', 'tr', '1:45 ÖS'],
    [AFTERNOON, 'h:mm:ss AM/PM', 'en', '1:45:10 PM'],
    [MORNING, 'h:mm:ss AM/PM', 'tr', '9:05:07 ÖÖ'],
    // 20, 21
    [AFTERNOON, 'h:mm', 'en', '13:45'],
    [MORNING, 'h:mm', 'en', '9:05'],
    [AFTERNOON, 'h:mm:ss', 'en', '13:45:10'],
    [MORNING, 'h:mm:ss', 'tr', '9:05:07'],
    // 22: the code as written
    [AFTERNOON, 'm/d/yy h:mm', 'en', '3/15/24 13:45'],
    [AFTERNOON, 'm/d/yy h:mm', 'tr', '3.15.24 13:45'],
    // 37, 38, 39, 40
    [1234.5, '#,##0_);(#,##0)', 'en', '1,235 '],
    [-1234.5, '#,##0_);(#,##0)', 'en', '(1,235)'],
    [0, '#,##0_);(#,##0)', 'en', '0 '],
    [-1234.5, '#,##0_);(#,##0)', 'tr', '(1.235)'],
    [1234.5, '#,##0_);[Red](#,##0)', 'en', '1,235 '],
    [-1234.5, '#,##0_);[Red](#,##0)', 'en', '(1,235)'],
    [1234.5, '#,##0.00_);(#,##0.00)', 'en', '1,234.50 '],
    [-1234.5, '#,##0.00_);(#,##0.00)', 'en', '(1,234.50)'],
    [-1234.5, '#,##0.00_);[Red](#,##0.00)', 'tr', '(1.234,50)'],
    // 41-44 accounting. "*x" fill emits nothing, "_x" one space.
    [1234.5, ACCOUNTING_41, 'en', ' 1,235 '],
    [-1234.5, ACCOUNTING_41, 'en', ' (1,235)'],
    [0, ACCOUNTING_41, 'en', ' - '],
    ['note', ACCOUNTING_41, 'en', ' note '],
    [1234.5, ACCOUNTING_42, 'en', ' $1,235 '],
    [-1234.5, ACCOUNTING_42, 'en', ' $(1,235)'],
    [0, ACCOUNTING_42, 'en', ' $- '],
    [1234.5, ACCOUNTING_43, 'en', ' 1,234.50 '],
    [-1234.5, ACCOUNTING_43, 'tr', ' (1.234,50)'],
    [0, ACCOUNTING_43, 'en', ' -   '],
    [1234.5, ACCOUNTING_44, 'en', ' $1,234.50 '],
    [-1234.5, ACCOUNTING_44, 'en', ' $(1,234.50)'],
    [0, ACCOUNTING_44, 'en', ' $-   '],
    ['note', ACCOUNTING_44, 'tr', ' note '],
    // 45, 46, 47
    [time(13, 45, 10), 'mm:ss', 'en', '45:10'],
    [1 + time(13, 45, 10), '[h]:mm:ss', 'en', '37:45:10'],
    [time(0, 5, 3), '[h]:mm:ss', 'en', '0:05:03'],
    [time(13, 45, 10, 400), 'mmss.0', 'en', '4510.4'],
    [time(13, 45, 10, 400), 'mmss.0', 'tr', '4510,4'],
    // 48
    [12345, '##0.0E+0', 'en', '12.3E+3'],
    [123456, '##0.0E+0', 'en', '123.5E+3'],
    [1234567, '##0.0E+0', 'en', '1.2E+6'],
    [1, '##0.0E+0', 'en', '1.0E+0'],
    [0.00012, '##0.0E+0', 'en', '120.0E-6'],
    // 49
    [12, '@', 'en', '12'],
    [1.5, '@', 'tr', '1,5'],
    ['abc', '@', 'en', 'abc'],
  ]);
});

test('number format: digit placeholders and rounding', () => {
  check([
    [12, '00000', 'en', '00012'],
    [123456, '00000', 'en', '123456'],
    [12, '???', 'en', ' 12'],
    [0, '#', 'en', ''],
    [0, '#.##', 'en', '.'],
    [5, '#.##', 'en', '5.'],
    [5.5, '#.##', 'en', '5.5'],
    [0.5, '#.00', 'en', '.50'],
    [0.5, '#.00', 'tr', ',50'],
    [1.5, '.00', 'en', '1.50'],
    [1234.5, '0.0??', 'en', '1234.5  '],
    [1234.5678, '#,##0.0?', 'en', '1,234.57'],
    [1.2, '0.0#', 'en', '1.2'],
    [1.25, '0.0#', 'en', '1.25'],
    [1, '0.0#', 'en', '1.0'],
    // Half up on the decimal digits, not on the binary value.
    [1.005, '0.00', 'en', '1.01'],
    [0.285, '0.00', 'en', '0.29'],
    [2.675, '0.00', 'en', '2.68'],
    [1.115, '0.00', 'en', '1.12'],
    [0.5, '0', 'en', '1'],
    [1.5, '0', 'en', '2'],
    [2.5, '0', 'en', '3'],
    [-2.5, '0', 'en', '-3'],
    [0.1 + 0.2, '0.0000000000', 'en', '0.3000000000'],
    [0.07, '0%', 'en', '7%'],
    [0.285, '0%', 'en', '29%'],
    // Unverified: every "%" multiplies by a hundred.
    [0.5, '0%%', 'en', '5000%%'],
    [9.995, '0.00', 'en', '10.00'],
    [999.9996, '#,##0.000', 'en', '1,000.000'],
    [999999.5, '#,##0', 'tr', '1.000.000'],
    // Negative values that round to zero lose the sign.
    [-0.001, '0.00', 'en', '-0.00'],
    [-0.4, '0', 'en', '-0'],
    [-0, '0.0', 'en', '0.0'],
    [-0.001, '0.00;(0.00)', 'en', '(0.00)'],
    [-0.004, '0%', 'en', '-0%'],
    // Very large and very small numbers.
    [1e20, '#,##0', 'en', '100,000,000,000,000,000,000'],
    [1.5e15, '0', 'en', '1500000000000000'],
    [123456789012345680000, '0', 'en', '123456789012346000000'],
    [1e-20, '0.00', 'en', '0.00'],
    [1e-7, '0.0000000', 'en', '0.0000001'],
    [4.9e-324, '0.00', 'en', '0.00'],
    [0.000123456, '0.000000', 'tr', '0,000123'],
    // Decimals beyond 15 significant digits are zeros.
    [1 / 3, '0.00000000000000000000', 'en', '0.33333333333333300000'],
  ]);
});

test('number format: thousands and scaling commas', () => {
  check([
    [1234567, '#,##0,', 'en', '1,235'],
    [1234567, '#,##0,', 'tr', '1.235'],
    [1234567, '0.0,,', 'en', '1.2'],
    [1234567, '0.0,,', 'tr', '1,2'],
    [1234567890, '#,##0.0,,"M"', 'en', '1,234.6M'],
    [1500, '0,', 'en', '2'],
    [499, '0,', 'en', '0'],
    [1234567, '#,###', 'en', '1,234,567'],
    [0, '#,###', 'en', ''],
    [12, '#,##0', 'en', '12'],
    [1234, '0,000', 'en', '1,234'],
    [12, '0,000', 'en', '0,012'],
    [1234567, '#,##0 "pcs"', 'tr', '1.234.567 pcs'],
    // A comma that touches no digit placeholder is plain text.
    [5, '"a", 0', 'en', 'a, 5'],
  ]);
});

test('number format: scientific', () => {
  check([
    [1234, '0.0E+0', 'en', '1.2E+3'],
    [1234, '0.0e+0', 'en', '1.2E+3'],
    [1234, '0.0E-0', 'en', '1.2E3'],
    [0.0012, '0.0E-0', 'en', '1.2E-3'],
    [0.0012, '0.0E+0', 'en', '1.2E-3'],
    [9.996, '0.00E+00', 'en', '1.00E+01'],
    [999999, '0.0E+0', 'en', '1.0E+6'],
    [1e100, '0.00E+00', 'en', '1.00E+100'],
    [1.5e-10, '0.00E+00', 'tr', '1,50E-10'],
    [123456, '0E+0', 'en', '1E+5'],
    [-123456, '0.###E+0', 'en', '-1.235E+5'],
    [100000, '0.###E+0', 'en', '1.E+5'],
    [999960, '##0.0E+0', 'en', '1.0E+6'],
    // Unverified: with "00" before the point the exponent is taken in steps of two.
    [12345, '00.0E+00', 'en', '01.2E+04'],
  ]);
});

test('number format: fractions', () => {
  check([
    // Unverified in detail: the blank padding around a missing whole part or fraction.
    [0.5, '# ?/?', 'en', ' 1/2'],
    [2, '# ?/?', 'en', '2    '],
    [0, '# ?/?', 'en', '0    '],
    // One digit denominators: the closest fraction to 0.3 is 2/7, not 1/3.
    [0.3, '# ?/?', 'en', ' 2/7'],
    [3.14159, '# ?/?', 'en', '3 1/7'],
    [3.14159, '# ??/??', 'en', '3 14/99'],
    [3.14159, '# ???/???', 'en', '3  16/113'],
    [0.999, '# ?/?', 'en', '1    '],
    [1.5, '?/?', 'en', '3/2'],
    [0, '?/?', 'en', '0/1'],
    [-0.75, '?/?', 'en', '-3/4'],
    [0.375, '# ?/8', 'en', ' 3/8'],
    [2.3, '# ?/8', 'en', '2 2/8'],
    [2.96, '# ?/8', 'en', '3    '],
    [1.25, '# ??/16', 'en', '1  4/16'],
    [3.14159, '# ??/100', 'en', '3 14/100'],
    [0.5, '0 ?/?', 'en', '0 1/2'],
    [1.75, '# ?/? "in"', 'en', '1 3/4 in'],
    [1234.5, '#,##0 ?/?', 'tr', '1.234 1/2'],
  ]);
});

test('number format: literal text', () => {
  check([
    [12, '0 "kg"', 'en', '12 kg'],
    [12, '"No. "0', 'en', 'No. 12'],
    [12, '\\x0\\y', 'en', 'x12y'],
    [12, '0\\ \\%', 'en', '12 %'],
    [12, '$0.00', 'en', '$12.00'],
    [-12, '$0.00', 'en', '-$12.00'],
    [12, '+0', 'en', '+12'],
    [12, '(0)', 'en', '(12)'],
    [12, '0 !', 'en', '12 !'],
    [12, "0'", 'en', "12'"],
    [12, '{0}', 'en', '{12}'],
    [12, '0_)', 'en', '12 '],
    [12, '_(0', 'en', ' 12'],
    [12, '*-0', 'en', '12'],
    [12, '0*x', 'en', '12'],
    [12, '"Total:"* 0', 'en', 'Total:12'],
    [123, '000-00-0000', 'en', '000-00-0123'],
    [123456789, '000-00-0000', 'en', '123-45-6789'],
    [5551234567, '(000) 000-0000', 'en', '(555) 123-4567'],
    [15551234567, '(000) 000-0000', 'en', '(1555) 123-4567'],
    [1234, '0"h" 00"m"', 'en', '12h 34m'],
    [12.5, '0"."00', 'en', '0.13'],
    [12, '"abc"', 'en', 'abc'],
    [-12, '"abc"', 'en', '-abc'],
    [1234.5, '#,##0.00\\ "₺"', 'tr', '1.234,50 ₺'],
    [1234.5, '#,##0.00\\ "₺"', 'en', '1,234.50 ₺'],
    [-1234.5, '#,##0.00\\ "₺"', 'tr', '-1.234,50 ₺'],
    [12, '"a""b"0', 'en', 'ab12'],
  ]);
});

test('number format: currency and locale codes', () => {
  check([
    [1234.5, '#,##0.00 [$₺-41F]', 'tr', '1.234,50 ₺'],
    [1234.5, '#,##0.00 [$₺-41F]', 'en', '1,234.50 ₺'],
    [1234.5, '[$€-2] #,##0.00', 'en', '€ 1,234.50'],
    [1234.5, '[$USD] 0', 'en', 'USD 1235'],
    [1234.5, '[$$-409]#,##0.00', 'en', '$1,234.50'],
    [-1234.5, '[$$-409]#,##0.00;[Red]-[$$-409]#,##0.00', 'en', '-$1,234.50'],
    [1234.5, '[$-409]0.0', 'tr', '1234,5'],
    [DAY, '[$-409]d-mmm-yyyy', 'en', '15-Mar-2024'],
    // The LCID picks the language of month and day names.
    [45505, '[$-409]d mmmm yyyy', 'tr', '1 August 2024'],
    [45505, '[$-41F]d mmmm yyyy', 'en', '1 Ağustos 2024'],
    [45505, '[$-41F]dddd', 'en', 'Perşembe'],
    // System long date and time; the rest of the section is not used.
    [DAY, '[$-F800]dddd, mmmm dd, yyyy', 'en', 'Friday, March 15, 2024'],
    [DAY, '[$-F800]dddd, mmmm dd, yyyy', 'tr', '15 Mart 2024 Cuma'],
    [AFTERNOON, '[$-F400]h:mm:ss AM/PM', 'en', '1:45:10 PM'],
    [AFTERNOON, '[$-F400]h:mm:ss AM/PM', 'tr', '13:45:10'],
    [-1, '[$-F800]dddd, mmmm dd, yyyy', 'en', '####'],
    // Unknown bracket contents have no visible effect.
    [12, '[DBNum1]0', 'en', '12'],
  ]);
});

test('number format: sections and conditions', () => {
  check([
    [5, '0.0;(0.0)', 'en', '5.0'],
    [-5, '0.0;(0.0)', 'en', '(5.0)'],
    [0, '0.0;(0.0)', 'en', '0.0'],
    [5, '0;-0;"zero"', 'en', '5'],
    [-5, '0;-0;"zero"', 'en', '-5'],
    [0, '0;-0;"zero"', 'en', 'zero'],
    [5, '"pos";"neg";"zero";"text"', 'en', 'pos'],
    [-5, '"pos";"neg";"zero";"text"', 'en', 'neg'],
    [0, '"pos";"neg";"zero";"text"', 'en', 'zero'],
    ['x', '"pos";"neg";"zero";"text"', 'en', 'text'],
    ['x', '0;-0;0;"<"@">"', 'en', '<x>'],
    ['x', '0;-0;0;@@', 'en', 'xx'],
    ['x', '0.00', 'en', 'x'],
    ['x', '0;-0', 'en', 'x'],
    ['x', 'yyyy-mm-dd', 'en', 'x'],
    ['x', 'General', 'en', 'x'],
    [5, ';;;', 'en', ''],
    [-5, ';;;', 'en', ''],
    ['x', ';;;', 'en', ''],
    [5, '0;;', 'en', '5'],
    [-5, '0;;', 'en', ''],
    [0, '0;;', 'en', ''],
    [1.5, '0.0"a";@', 'en', '1.5a'],
    ['t', '0.0"a";@" b"', 'en', 't b'],
    [-1.5, 'General;"neg "General', 'tr', 'neg 1,5'],
    [1.5, '"x"General"y"', 'en', 'x1.5y'],
    [150, '[>100]"big";"small"', 'en', 'big'],
    [100, '[>100]"big";"small"', 'en', 'small'],
    [100, '[>=100]"big";"small"', 'en', 'big'],
    [1, '[=1]"one";[=2]"two";"many"', 'en', 'one'],
    [2, '[=1]"one";[=2]"two";"many"', 'en', 'two'],
    [3, '[=1]"one";[=2]"two";"many"', 'en', 'many'],
    [2, '[<>2]"other";"two"', 'en', 'two'],
    [7, '[<>2]"other";"two"', 'en', 'other'],
    [0, '[<=0]"none";0', 'en', 'none'],
    [999, '[<1000]0;[<1000000]0.0,"K";0.0,,"M"', 'en', '999'],
    [45600, '[<1000]0;[<1000000]0.0,"K";0.0,,"M"', 'en', '45.6K'],
    [7800000, '[<1000]0;[<1000000]0.0,"K";0.0,,"M"', 'en', '7.8M'],
    [0.5, '[<1]0.00;0', 'en', '0.50'],
    // Unverified: the second of two sections takes what the first refused and keeps the sign.
    [-5, '[>100]0.0;0.00', 'en', '-5.00'],
    // Unverified: a section whose condition only negative numbers meet drops the sign.
    [-5, '[<0]"neg "0;0', 'en', 'neg 5'],
  ]);
});

test('number format: colours', () => {
  const colors: [number, string, string | undefined][] = [
    [5, '[Red]0', '#FF0000'],
    [5, '[red]0', '#FF0000'],
    [5, '[BLUE]0', '#0000FF'],
    [5, '[Black]0', '#000000'],
    [5, '[Cyan]0', '#00FFFF'],
    [5, '[Green]0', '#00FF00'],
    [5, '[Magenta]0', '#FF00FF'],
    [5, '[White]0', '#FFFFFF'],
    [5, '[Yellow]0', '#FFFF00'],
    [5, '[Color 1]0', '#000000'],
    [5, '[Color 3]0', '#FF0000'],
    [5, '[Color10]0', '#008000'],
    [5, '[color 56]0', '#333333'],
    [5, '[Color 41]0', '#3366FF'],
    [5, '[Color 57]0', undefined],
    [5, '[Color 0]0', undefined],
    [5, '0', undefined],
    [5, '#,##0_);[Red](#,##0)', undefined],
    [-5, '#,##0_);[Red](#,##0)', '#FF0000'],
    [-5, '[Blue]0;[Red]-0;[Green]0', '#FF0000'],
    [0, '[Blue]0;[Red]-0;[Green]0', '#00FF00'],
    [5, '[Red][>100]0;[Blue]0', '#0000FF'],
    [500, '[Red][>100]0;[Blue]0', '#FF0000'],
    [5, '[Red]General', '#FF0000'],
    [5, '[Red]', '#FF0000'],
    [DAY, '[Blue]yyyy', '#0000FF'],
  ];
  for (const [value, code, color] of colors) {
    assert.equal(formatValue(value, code).color, color, code);
  }
  assert.deepEqual(formatValue('x', '0;0;0;[Red]@'), { text: 'x', color: '#FF0000' });
  assert.deepEqual(formatValue(5, '[Red]'), { text: '5', color: '#FF0000' });
  // No colour property at all when the section names none.
  assert.deepEqual(formatValue(5, '0'), { text: '5' });
});

test('number format: dates', () => {
  check([
    [DAY, 'yyyy-mm-dd', 'en', '2024-03-15'],
    [DAY, 'yyyy-mm-dd', 'tr', '2024-03-15'],
    [DAY, 'dd/mm/yyyy', 'en', '15/03/2024'],
    [DAY, 'dd/mm/yyyy', 'tr', '15.03.2024'],
    [DAY, 'dd.mm.yyyy', 'en', '15.03.2024'],
    [DAY, 'd/m/yy', 'en', '15/3/24'],
    [DAY, 'm/d/yyyy', 'tr', '3.15.2024'],
    [DAY, 'yy', 'en', '24'],
    [DAY, 'y', 'en', '24'],
    [DAY, 'yyy', 'en', '2024'],
    [DAY, 'yyyy', 'en', '2024'],
    [DAY, 'm', 'en', '3'],
    [DAY, 'mm', 'en', '03'],
    [DAY, 'mmm', 'en', 'Mar'],
    [DAY, 'mmmm', 'en', 'March'],
    [DAY, 'mmmmm', 'en', 'M'],
    [DAY, 'mmmm', 'tr', 'Mart'],
    [45337, 'mmm', 'tr', 'Şub'],
    [45337, 'mmmm', 'tr', 'Şubat'],
    [45337, 'mmmmm', 'tr', 'Ş'],
    [45658, 'mmmm yyyy', 'tr', 'Ocak 2025'],
    [45657, 'd mmmm yyyy', 'tr', '31 Aralık 2024'],
    [DAY, 'd', 'en', '15'],
    [45352, 'd', 'en', '1'],
    [45352, 'dd', 'en', '01'],
    [DAY, 'ddd', 'en', 'Fri'],
    [DAY, 'dddd', 'en', 'Friday'],
    [DAY, 'ddd', 'tr', 'Cum'],
    [DAY, 'dddd', 'tr', 'Cuma'],
    [DAY + 1, 'dddd', 'tr', 'Cumartesi'],
    [DAY + 2, 'ddd', 'tr', 'Paz'],
    [DAY + 3, 'ddd', 'tr', 'Pzt'],
    [DAY + 5, 'dddd', 'tr', 'Çarşamba'],
    [DAY, 'dddd, d mmmm yyyy', 'en', 'Friday, 15 March 2024'],
    [DAY, 'YYYY-MM-DD', 'en', '2024-03-15'],
    [DAY, 'Dddd', 'en', 'Friday'],
    [DAY, 'yyyy"y"mm"m"', 'en', '2024y03m'],
    [DAY, '"day "d', 'en', 'day 15'],
    [45292, 'yyyy-mm-dd', 'en', '2024-01-01'],
    [45351, 'yyyy-mm-dd', 'en', '2024-02-29'],
    [36526, 'yyyy-mm-dd', 'en', '2000-01-01'],
    [36585, 'yyyy-mm-dd', 'en', '2000-02-29'],
    [73050, 'yyyy-mm-dd', 'en', '2099-12-31'],
    [73051, 'yyyy-mm-dd', 'en', '2100-01-01'],
    [73110, 'yyyy-mm-dd', 'en', '2100-03-01'],
  ]);
});

test('number format: times', () => {
  check([
    [time(9, 5, 7), 'h:mm:ss', 'en', '9:05:07'],
    [time(9, 5, 7), 'hh:mm:ss', 'en', '09:05:07'],
    [time(9, 5, 7), 'h:m:s', 'en', '9:5:7'],
    [time(9, 5, 7), 'hh', 'en', '09'],
    [time(9, 5, 7), 'ss', 'en', '07'],
    [time(0, 0, 0), 'h:mm AM/PM', 'en', '12:00 AM'],
    [time(12, 0, 0), 'h:mm AM/PM', 'en', '12:00 PM'],
    [time(23, 59, 0), 'hh:mm AM/PM', 'en', '11:59 PM'],
    [time(15, 0, 0), 'h A/P', 'en', '3 P'],
    [time(3, 0, 0), 'h A/P', 'en', '3 A'],
    [time(15, 0, 0), 'h a/p', 'en', '3 p'],
    // Lower case "am/pm" still gives the locale's upper case text.
    [time(15, 0, 0), 'h:mm am/pm', 'en', '3:00 PM'],
    [time(3, 0, 0), 'h:mm Am/Pm', 'tr', '3:00 ÖÖ'],
    [time(15, 0, 0), 'h:mm AM/PM', 'tr', '3:00 ÖS'],
    // "m" is minutes after hours or before seconds, the month otherwise.
    [AFTERNOON, 'yyyy-mm-dd hh:mm', 'en', '2024-03-15 13:45'],
    [AFTERNOON, 'm/d/yyyy h:mm', 'en', '3/15/2024 13:45'],
    [AFTERNOON, 'mm:ss', 'en', '45:10'],
    [AFTERNOON, 'm:ss', 'en', '45:10'],
    [AFTERNOON, 'h "h" m "min"', 'en', '13 h 45 min'],
    [AFTERNOON, 'mm', 'en', '03'],
    [AFTERNOON, 'hh:mm dd/mm', 'en', '13:45 15/03'],
    [AFTERNOON, 'h:mm AM/PM mmm', 'en', '1:45 PM Mar'],
    // Separators follow the locale.
    [AFTERNOON, 'dd/mm/yyyy hh:mm:ss', 'tr', '15.03.2024 13:45:10'],
    [AFTERNOON, 'dd-mm-yyyy', 'tr', '15-03-2024'],
    // Fractions of a second.
    [time(13, 45, 10, 123), 'h:mm:ss.0', 'en', '13:45:10.1'],
    [time(13, 45, 10, 123), 'h:mm:ss.00', 'en', '13:45:10.12'],
    [time(13, 45, 10, 123), 'h:mm:ss.000', 'en', '13:45:10.123'],
    [time(13, 45, 10, 123), 'h:mm:ss.000', 'tr', '13:45:10,123'],
    [time(13, 45, 10, 960), 'h:mm:ss.0', 'en', '13:45:11.0'],
    [time(13, 45, 10, 960), 'ss.00', 'en', '10.96'],
    // Without fractions the time is rounded to the second.
    [time(13, 45, 10, 600), 'h:mm:ss', 'en', '13:45:11'],
    [time(13, 45, 10, 400), 'h:mm:ss', 'en', '13:45:10'],
    // Unverified: also when the seconds are not shown.
    [time(13, 45, 59, 700), 'h:mm', 'en', '13:46'],
    // Rounding carries into the next day.
    [DAY + time(23, 59, 59, 999.5), 'yyyy-mm-dd hh:mm:ss.000', 'en', '2024-03-16 00:00:00.000'],
    [DAY + time(23, 59, 59, 600), 'yyyy-mm-dd hh:mm:ss', 'en', '2024-03-16 00:00:00'],
    [DAY + time(23, 59, 59, 400), 'yyyy-mm-dd hh:mm:ss', 'en', '2024-03-15 23:59:59'],
    [DAY + time(23, 59, 59, 600), 'dddd', 'en', 'Saturday'],
    // Elapsed time.
    [1.5, '[h]:mm', 'en', '36:00'],
    [1.5, '[hh]:mm', 'en', '36:00'],
    [time(5, 0, 0), '[hh]:mm', 'en', '05:00'],
    [time(2, 57, 42), '[m]:ss', 'en', '177:42'],
    [time(2, 57, 42), '[mm]:ss', 'en', '177:42'],
    [time(0, 3, 9), '[mm]:ss', 'en', '03:09'],
    [time(2, 57, 42), '[s]', 'en', '10662'],
    [time(0, 0, 5), '[ss]', 'en', '05'],
    [2, '[s].00', 'en', '172800.00'],
    [1000.25, '[h]:mm:ss', 'en', '24006:00:00'],
    [-0.5, '[h]:mm', 'en', '####'],
  ]);
});

test('number format: date boundaries and overflow', () => {
  check([
    [0, 'dd/mm/yyyy', 'tr', '00.01.1900'],
    [0, 'dddd', 'en', 'Saturday'],
    [0.5, 'dd/mm/yyyy hh:mm', 'tr', '00.01.1900 12:00'],
    [1, 'dd/mm/yyyy', 'tr', '01.01.1900'],
    [1, 'dddd', 'en', 'Sunday'],
    [59, 'dd/mm/yyyy', 'tr', '28.02.1900'],
    [60, 'dd/mm/yyyy', 'tr', '29.02.1900'],
    [60, 'dddd', 'en', 'Wednesday'],
    [61, 'dd/mm/yyyy', 'tr', '01.03.1900'],
    [61, 'dddd', 'en', 'Thursday'],
    [366, 'yyyy-mm-dd', 'en', '1900-12-31'],
    [367, 'yyyy-mm-dd', 'en', '1901-01-01'],
    [2958465, 'dd/mm/yyyy', 'tr', '31.12.9999'],
    [2958465.5, 'yyyy-mm-dd hh:mm', 'en', '9999-12-31 12:00'],
    [2958466, 'dd/mm/yyyy', 'tr', '####'],
    [2958465.999999, 'yyyy-mm-dd', 'en', '####'],
    [1e300, 'yyyy-mm-dd', 'en', '####'],
    [-1, 'dd/mm/yyyy', 'tr', '####'],
    [-0.25, 'h:mm', 'en', '####'],
    [-1, 'mm-dd-yy', 'en', '####'],
  ]);
  assert.deepEqual(formatValue(-1, 'yyyy-mm-dd'), { text: '####', overflow: true });
  assert.deepEqual(formatValue(2958466, 'yyyy-mm-dd'), { text: '####', overflow: true });
  assert.equal(formatValue(1, 'yyyy-mm-dd').overflow, undefined);
  // A negative number in the negative section of a date format is still not a date.
  assert.equal(formatValue(-1, 'yyyy;yyyy').overflow, true);

  // 1904 date system.
  const options = { date1904: true };
  assert.equal(formatValue(0, 'dd/mm/yyyy', options).text, '01/01/1904');
  assert.equal(formatValue(0, 'dddd', options).text, 'Friday');
  assert.equal(formatValue(59, 'yyyy-mm-dd', options).text, '1904-02-29');
  assert.equal(formatValue(60, 'yyyy-mm-dd', options).text, '1904-03-01');
  assert.equal(formatValue(DAY - 1462, 'yyyy-mm-dd dddd', options).text, '2024-03-15 Friday');
  assert.equal(formatValue(2957003, 'yyyy-mm-dd', options).text, '9999-12-31');
  assert.equal(formatValue(2957004, 'yyyy-mm-dd', options).overflow, true);
  // The 1904 system shows negative dates and durations with a minus sign.
  assert.deepEqual(formatValue(-1, 'yyyy-mm-dd', options), { text: '-1904-01-02' });
  assert.equal(formatValue(-0.5, '[h]:mm', options).text, '-12:00');
  assert.equal(formatValue(-0.25, 'h:mm', options).text, '-6:00');
  assert.equal(formatValue(-1, '[$-F800]', options).text, '-Saturday, January 2, 1904');
  // In the negative section the sign is the section's business.
  assert.equal(formatValue(-1, 'yyyy;"neg "yyyy', options).text, 'neg 1904');
  assert.equal(formatValue(-2957004, 'yyyy-mm-dd', options).overflow, true);
  // The 1900 system has no negative dates.
  assert.equal(formatValue(-0.5, '[h]:mm', { date1904: false }).overflow, true);
});

test('number format: serialToDate', () => {
  assert.deepEqual(serialToDate(DAY + time(13, 45, 10, 250)), {
    year: 2024, month: 3, day: 15, hour: 13, minute: 45, second: 10, millisecond: 250, weekday: 5,
  });
  assert.deepEqual(serialToDate(60), {
    year: 1900, month: 2, day: 29, hour: 0, minute: 0, second: 0, millisecond: 0, weekday: 3,
  });
  assert.deepEqual(serialToDate(0), {
    year: 1900, month: 1, day: 0, hour: 0, minute: 0, second: 0, millisecond: 0, weekday: 6,
  });
  assert.equal(serialToDate(1).day, 1);
  assert.equal(serialToDate(59).day, 28);
  assert.equal(serialToDate(61).month, 3);
  assert.deepEqual(serialToDate(0, true), {
    year: 1904, month: 1, day: 1, hour: 0, minute: 0, second: 0, millisecond: 0, weekday: 5,
  });
  const last = serialToDate(2958465);
  assert.deepEqual([last.year, last.month, last.day], [9999, 12, 31]);
  // Rounding to the millisecond carries into the next day.
  const carried = serialToDate(DAY + 0.99999999999);
  assert.deepEqual([carried.day, carried.hour, carried.millisecond], [16, 0, 0]);
  for (const bad of [-1, NaN, Infinity, 2958466]) {
    assert.ok(Number.isNaN(serialToDate(bad).year), String(bad));
  }

  // Every day from 1 March 1900 to the end of 2100 agrees with the platform calendar.
  for (let serial = 61; serial <= 73415; serial++) {
    const expected = new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
    const actual = serialToDate(serial);
    if (
      actual.year !== expected.getUTCFullYear() || actual.month !== expected.getUTCMonth() + 1 ||
      actual.day !== expected.getUTCDate() || actual.weekday !== expected.getUTCDay()
    ) {
      assert.fail(`serial ${serial}: ${JSON.stringify(actual)}`);
    }
  }
});

test('number format: booleans, blanks and odd values', () => {
  assert.deepEqual(formatValue(true, '0.00'), { text: 'TRUE' });
  assert.deepEqual(formatValue(false, 'General'), { text: 'FALSE' });
  assert.deepEqual(formatValue(true, '0.00', { locale: LOCALE_TR }), { text: 'DOĞRU' });
  assert.deepEqual(formatValue(false, '@', { locale: LOCALE_TR }), { text: 'YANLIŞ' });
  assert.deepEqual(formatValue(null, '0.00'), { text: '' });
  assert.deepEqual(formatValue(null, 'yyyy-mm-dd'), { text: '' });
  assert.deepEqual(formatValue('', '0.00'), { text: '' });
  for (const code of ['0.00', 'General', 'yyyy-mm-dd', '# ?/?', '0.00E+00', '@', '"x"', '[abc']) {
    assert.deepEqual(formatValue(NaN, code), { text: '#NUM!' }, code);
    assert.deepEqual(formatValue(Infinity, code), { text: '#NUM!' }, code);
    assert.deepEqual(formatValue(-Infinity, code, { locale: LOCALE_TR }), { text: '#NUM!' }, code);
  }
  assert.equal(formatValue(5, '').text, '5');
  assert.equal(formatValue('x', '').text, 'x');
  assert.equal(formatValue(-0, 'General').text, '0');
});

test('number format: isDateFormat', () => {
  const dates = [
    'yyyy-mm-dd', '[h]:mm', '[$-409]d-mmm', 'mm-dd-yy', 'd-mmm-yy', 'h:mm AM/PM', 'mm:ss', 'mmss.0',
    'm/d/yy h:mm', '[$-F800]dddd, mmmm dd, yyyy', '[$-F400]h:mm:ss AM/PM', '[Red]d/m/yyyy', 'YYYY',
    '0;yyyy', '"on" d mmmm',
  ];
  const others = [
    '[Red]0.00', '"day" 0', 'General', '0.00E+00', '0', '#,##0.00', '0%', '# ?/?', '@', '', '[Red]',
    '\\d0', '0 "h"', '[$-409]0.00', '[$d-409]0', '#,##0_);(#,##0)', ACCOUNTING_44, '0;0;0;"date "@', '[abc',
    '[>=100]0;0', '_d0', '*m0',
  ];
  for (const code of dates) assert.equal(isDateFormat(code), true, code);
  for (const code of others) assert.equal(isDateFormat(code), false, code);
});

test('number format: locales', () => {
  assert.equal(getLocale('tr'), LOCALE_TR);
  assert.equal(getLocale('tr-TR'), LOCALE_TR);
  assert.equal(getLocale('TR_tr'), LOCALE_TR);
  assert.equal(getLocale('en'), LOCALE_EN);
  assert.equal(getLocale('en-GB'), LOCALE_EN);
  assert.equal(getLocale('tru'), LOCALE_EN);
  assert.equal(getLocale(''), LOCALE_EN);
  assert.equal(getLocale(undefined), LOCALE_EN);
  for (const locale of [LOCALE_EN, LOCALE_TR]) {
    assert.equal(locale.monthNames.length, 12);
    assert.equal(locale.shortMonthNames.length, 12);
    assert.equal(locale.dayNames.length, 7);
    assert.equal(locale.shortDayNames.length, 7);
  }
  assert.deepEqual(LOCALE_TR.shortMonthNames, [
    'Oca', 'Şub', 'Mar', 'Nis', 'May', 'Haz', 'Tem', 'Ağu', 'Eyl', 'Eki', 'Kas', 'Ara',
  ]);
  assert.deepEqual(LOCALE_TR.shortDayNames, ['Paz', 'Pzt', 'Sal', 'Çar', 'Per', 'Cum', 'Cmt']);
  // The default locale is English.
  assert.equal(formatValue(1234.5, '#,##0.00').text, '1,234.50');
  // A custom locale without the optional date codes falls back to the English ones.
  const custom = { ...LOCALE_TR, shortDate: undefined, longDate: undefined, longTime: undefined };
  assert.equal(formatValue(DAY, 'mm-dd-yy', { locale: custom, builtinId: 14 }).text, '3.15.2024');
  assert.equal(formatValue(DAY, '[$-F800]', { locale: custom }).text, 'Cuma, Mart 15, 2024');
});

test('number format: malformed codes never throw', () => {
  const codes = [
    '"abc', '[abc', '[', ']', '[]', '[[Red]]0', '[Red', '0;[', '0*', '0_', '0\\', '*', '_', '\\', '"',
    '[>abc]0', '[>]0', '[$', '[$]', '[$-]0', '[$-ZZZZ]0', '[Color 999]0', 'E+', '0E+', 'E+0', '/', '0/',
    '/0', '0/0', '?/', '.', ',', '%', '..', '0.0.0', '0,,,,', ',,,0', '@@@', 'General General', 'Generalx',
    ';', ';;', ';;;;', ';;;;;;;;', 'yyyy0', 'h.0', 's.', 'AM/PM', 'A/P', 'am/p', 'e', 'b', 'g', 'ggg',
    '[h', '[h][m][s]', '0\ud83d', '_😀 0', '\\😀 0', '# ?/0', '# ?/000', '# 0/#',
    '# ?/99999999999999999999', '?'.repeat(40) + '/' + '?'.repeat(40), '0.' + '0'.repeat(400),
    '0'.repeat(1024), '#,'.repeat(500), '['.repeat(1000), '"'.repeat(999), ';'.repeat(5000),
    '0'.repeat(100000), '[Red]'.repeat(200) + '0', '@'.repeat(1024), 'y'.repeat(1024), '\u0000', '\n0',
  ];
  const values: Value[] = [0, 1234.5, -1234.5, 1e300, -1e-300, 0.5, DAY, 'text', true, null, NaN, -Infinity];
  for (const code of codes) {
    for (const value of values) {
      for (const locale of [LOCALE_EN, LOCALE_TR]) {
        const result = formatValue(value, code, { locale });
        assert.equal(typeof result.text, 'string', code.slice(0, 30));
        assert.ok(result.text.length <= 4096, `${code.slice(0, 30)}: ${result.text.length} characters`);
      }
    }
    assert.equal(typeof isDateFormat(code), 'boolean');
  }

  // What cannot be compiled shows the value as General.
  assert.equal(formatValue(1234.5, '[abc').text, '1234.5');
  assert.equal(formatValue(1234.5, '[abc', { locale: LOCALE_TR }).text, '1234,5');
  assert.equal(formatValue(1234.5, '0'.repeat(100000)).text, '1234.5');
  assert.equal(formatValue(1234.5, ';'.repeat(5000)).text, '1234.5');
  assert.equal(formatValue(1234.5, '0;0;0;@;0').text, '1234.5');
  // Trailing fill, pad and escape characters are ignored.
  assert.equal(formatValue(1234.5, '0*').text, '1235');
  assert.equal(formatValue(1234.5, '0_').text, '1235');
  assert.equal(formatValue(1234.5, '0\\').text, '1235');
  // An unclosed quote runs to the end of the code.
  assert.equal(formatValue(5, '0 "abc').text, '5 abc');
  // A long text through many "@" stays bounded.
  const long = formatValue('x'.repeat(30000), '@'.repeat(1000)).text;
  assert.ok(long.length <= 100000, String(long.length));
  // Not a string at all (untyped callers).
  assert.equal(formatValue(5, undefined as unknown as string).text, '5');
});

test('number format: cache keeps working past its size limit', () => {
  for (let i = 0; i < 1500; i++) {
    assert.equal(formatValue(7, `0" unit${i}"`).text, `7 unit${i}`);
  }
  assert.equal(formatValue(7, '0" unit0"').text, '7 unit0');
  assert.ok(compiledFormatCount() <= 500, String(compiledFormatCount()));
  assert.ok(compiledFormatCount() > 0);
  assert.equal(formatValue(1234.5, '#,##0.00').text, '1,234.50');
});

test('number format: one million calls stay fast', () => {
  const codes = ['#,##0.00', '0%', 'General', 'yyyy-mm-dd', '0.00E+00', '#,##0_);(#,##0)', '0', 'h:mm:ss'];
  let characters = 0;
  const start = process.hrtime();
  for (let i = 0; i < 1000000; i++) {
    characters += formatValue(i * 1.37 - 1000, codes[i & 7]).text.length;
  }
  const elapsed = process.hrtime(start);
  const milliseconds = elapsed[0] * 1000 + elapsed[1] / 1e6;
  assert.ok(characters > 0);
  // The target is about a second; the limit leaves room for slow, busy machines.
  assert.ok(milliseconds < 4000, `1,000,000 calls took ${milliseconds.toFixed(0)} ms`);
});

test('number format: locale ids in brackets are never shown', () => {
  check([
    [DAY, '[$-tr-TR]d mmmm yyyy', 'en', '15 Mart 2024'],
    [DAY, '[$-tr]dddd', 'en', 'Cuma'],
    [DAY, '[$-TR-tr]mmm', 'en', 'Mar'],
    [45505, '[$-en-US]d mmmm yyyy', 'tr', '1 August 2024'],
    [45505, '[$-en-GB]dddd', 'tr', 'Thursday'],
    [45505, '[$-809]mmmm', 'tr', 'August'],
    [45505, '[$-de-DE]mmmm', 'tr', 'Ağustos'],
    [45000.6, '[$$-en-US]#,##0.00', 'en', '$45,000.60'],
    [45000.6, '[$₺-tr-TR]#,##0.00', 'tr', '₺45.000,60'],
    [45000.6, '#,##0.00 [$TL-tr-TR]', 'en', '45,000.60 TL'],
    [5, '[$-409-x]0', 'en', '5'],
    [5, '[$-x-y-z]0', 'en', '5'],
    [5, '[$€-]0', 'en', '€5'],
    [DAY, '[$-x-sysdate]dddd, mmmm dd, yyyy', 'tr', '15 Mart 2024 Cuma'],
    [AFTERNOON, '[$-x-systime]h:mm:ss AM/PM', 'tr', '13:45:10'],
  ]);
  assert.equal(isDateFormat('[$-tr-TR]d mmmm yyyy'), true);
  assert.equal(isDateFormat('[$$-en-US]#,##0.00'), false);
});

test('number format: unquoted letters next to digit placeholders are text', () => {
  check([
    [1500, '0 kg', 'en', '1500 kg'],
    [1.5, '0.00 EUR', 'en', '1.50 EUR'],
    [1.5, '0.00 EUR', 'tr', '1,50 EUR'],
    [1.5, '0 adet', 'en', '2 adet'],
    [12, '0 ms', 'en', '12 ms'],
    [1500, '#,##0 gün', 'tr', '1.500 gün'],
    [5, '0 AM/PM', 'en', '5 AM/PM'],
    [5, '0 a/p', 'en', '5 a/p'],
    [3, 'd0', 'en', 'd3'],
    [2024, 'yyyy0', 'en', 'yyyy2024'],
    [-1500, '0 kg', 'en', '-1500 kg'],
    [0.5, '0% hedef', 'en', '50% hedef'],
    // Real date and time codes keep their meaning.
    [time(13, 45, 10, 400), 'mmss.0', 'en', '4510.4'],
    [time(13, 45, 10, 123), 'h:mm:ss.000', 'en', '13:45:10.123'],
    [1.5, '[h]:mm:ss', 'en', '36:00:00'],
    [2, '[s].00', 'en', '172800.00'],
    [DAY, 'yyyy-mm-dd', 'en', '2024-03-15'],
    // Era letters show nothing in a date and do not make one.
    [DAY, 'gg yyyy', 'en', ' 2024'],
    [time(13, 45, 10), 'g mm:ss', 'en', ' 45:10'],
  ]);
  for (const code of ['0 kg', '0.00 EUR', '0 adet', '0 ms', '#,##0 gün', '0 AM/PM', 'd0', 'h.0', '0 "x";0 ms']) {
    assert.equal(isDateFormat(code), false, code);
  }
  for (const code of ['mmss.0', 'h:mm:ss.000', '[h]:mm:ss', 'yyyy-mm-dd', '[s].00', 'ss.00', '0;h:mm']) {
    assert.equal(isDateFormat(code), true, code);
  }
});

test('number format: built-in ids 14 and 22 use the system short date', () => {
  const tr = LOCALE_TR;
  assert.equal(formatValue(DAY, 'mm-dd-yy', { builtinId: 14 }).text, '3/15/2024');
  assert.equal(formatValue(DAY, 'mm-dd-yy', { builtinId: 14, locale: tr }).text, '15.03.2024');
  assert.equal(formatValue(45292, 'mm-dd-yy', { builtinId: 14 }).text, '1/1/2024');
  assert.equal(formatValue(45292, 'mm-dd-yy', { builtinId: 14, locale: tr }).text, '01.01.2024');
  // The stored code does not matter for these ids.
  assert.equal(formatValue(DAY, 'dd/mm/yyyy', { builtinId: 14 }).text, '3/15/2024');
  assert.equal(formatValue(DAY, '', { builtinId: 14, locale: tr }).text, '15.03.2024');
  assert.equal(formatValue(AFTERNOON, 'm/d/yy h:mm', { builtinId: 22 }).text, '3/15/2024 13:45');
  assert.equal(formatValue(AFTERNOON, 'm/d/yy h:mm', { builtinId: 22, locale: tr }).text, '15.03.2024 13:45');
  assert.equal(formatValue(DAY + 0.25, 'm/d/yy h:mm', { builtinId: 22 }).text, '3/15/2024 6:00');
  assert.equal(formatValue(DAY + 0.25, 'm/d/yy h:mm', { builtinId: 22, locale: tr }).text, '15.03.2024 06:00');
  assert.equal(formatValue(-1, 'mm-dd-yy', { builtinId: 14 }).overflow, true);
  // Without the id, and for other ids, the code is applied as written.
  assert.equal(formatValue(DAY, 'mm-dd-yy').text, '03-15-24');
  assert.equal(formatValue(DAY, 'mm-dd-yy', { locale: tr }).text, '03-15-24');
  assert.equal(formatValue(DAY, 'mm-dd-yy', { builtinId: 15 }).text, '03-15-24');
  assert.equal(formatValue(1234.5, '0.00', { builtinId: 2 }).text, '1234.50');
  // Text and booleans are not dates.
  assert.equal(formatValue('x', 'mm-dd-yy', { builtinId: 14 }).text, 'x');
  assert.equal(formatValue(true, 'mm-dd-yy', { builtinId: 14 }).text, 'TRUE');
});

test('number format: incomplete locales are completed from English', () => {
  const empty = {} as FormatLocale;
  assert.equal(formatValue(45000, 'mmm d', { locale: empty }).text, 'Mar 15');
  assert.equal(formatValue(1.5, '0.0', { locale: empty }).text, '1.5');
  assert.equal(formatValue(true, '0', { locale: empty }).text, 'TRUE');
  assert.equal(formatValue(45000, '', { locale: empty, builtinId: 22 }).text, '3/15/2023 0:00');
  const partial = { decimal: ',', monthNames: ['only one'], shortDayNames: [1, 2, 3, 4, 5, 6, 7] };
  const loose = partial as unknown as FormatLocale;
  assert.equal(formatValue(1234.5, '#,##0.0', { locale: loose }).text, '1,234,5');
  assert.equal(formatValue(45000, 'ddd d mmmm', { locale: loose }).text, 'Wed 15 March');
  for (const bad of [null, 'tr', 42]) {
    assert.equal(formatValue(1.5, '0.0', { locale: bad as unknown as FormatLocale }).text, '1.5');
  }

  assert.equal(normalizeLocale(undefined), LOCALE_EN);
  assert.equal(normalizeLocale(null), LOCALE_EN);
  assert.equal(normalizeLocale(LOCALE_TR), LOCALE_TR);
  const complete = normalizeLocale(partial as unknown as Partial<FormatLocale>);
  assert.equal(complete.decimal, ',');
  assert.equal(complete.group, ',');
  assert.deepEqual(complete.monthNames, LOCALE_EN.monthNames);
  assert.deepEqual(complete.shortDayNames, LOCALE_EN.shortDayNames);
  assert.equal(complete.shortTime, 'h:mm');
  // The same input object gives the same result object: no allocation per call.
  assert.equal(normalizeLocale(partial as unknown as Partial<FormatLocale>), complete);

  // Separators other than "/" and ":".
  const dashed: FormatLocale = { ...LOCALE_EN, dateSeparator: '-', timeSeparator: '.' };
  assert.equal(formatValue(AFTERNOON, 'dd/mm/yyyy hh:mm:ss', { locale: dashed }).text, '15-03-2024 13.45.10');
  assert.equal(formatValue(time(1, 2, 3), '[h]:mm', { locale: dashed }).text, '1.02');
  // In a number both characters are plain text.
  assert.equal(formatValue(5, '0/"x":0', { locale: dashed }).text, '0/x:5');
});

test('number format: fractions with forced digits', () => {
  check([
    [3, '# 0/0', 'en', '3 0/1'],
    [0, '0 0/0', 'en', '0 0/1'],
    [1.5, '0 0/0', 'en', '1 1/2'],
    [0.5, '0 0/0', 'en', '0 1/2'],
    [3, '# ?/?', 'en', '3    '],
    [3, '# 00/8', 'en', '3 00/8'],
    [3.5, '# 00/8', 'en', '3 04/8'],
    // Unverified: "0" in a searched denominator pads it with zeros.
    [1.5, '# 0/00', 'en', '1 1/02'],
    // A fixed denominator of any length.
    [0.5, '# ?/999999999999', 'en', ' 500000000000/999999999999'],
  ]);
  for (const value of [0, 0.5, 1234.5, -2.25, 1e14]) {
    const text = formatValue(value, '# ?/' + '9'.repeat(40)).text;
    assert.ok(text.length < 100, text);
    assert.ok(text.indexOf('NaN') < 0 && text.indexOf('e+') < 0, text);
  }
});

test('number format: conditions that no section meets, and signs', () => {
  // One conditional section that does not apply: General.
  assert.deepEqual(formatValue(5, '[>100]0.00'), { text: '5' });
  assert.deepEqual(formatValue(5.5, '[Red][>100]0.00', { locale: LOCALE_TR }), { text: '5,5' });
  assert.deepEqual(formatValue(-5.5, '[>100]0.00'), { text: '-5.5' });
  assert.deepEqual(formatValue(150, '[Red][>100]0.00'), { text: '150.00', color: '#FF0000' });
  // Two conditional sections, neither applies: nothing can be shown.
  assert.deepEqual(formatValue(5, '[>100]0;[<-100]0'), { text: '####', overflow: true });
  assert.equal(formatValue(500, '[>100]0;[<-100]0').text, '500');
  assert.equal(formatValue(-500, '[>100]0;[<-100]0').text, '500');
  // With a third section there is always a place.
  assert.equal(formatValue(5, '[>100]0;[<-100]0;0.0').text, '5.0');

  check([
    // Only a condition that excludes zero and positives drops the minus sign.
    [-5, '[<0]"x"0;0', 'en', 'x5'],
    [-5, '[<=0]"x"0;0', 'en', '-x5'],
    [-5, '[<=-1]"x"0;0', 'en', 'x5'],
    [-5, '[<1]"x"0;0', 'en', '-x5'],
    [-5, '[<-1]"x"0;0', 'en', 'x5'],
    [-5, '[<>0]"x"0;0', 'en', '-x5'],
    // The same in the second section.
    [-10, '0;[<-5]"x"0;0', 'en', 'x10'],
    [-3, '0;[<-5]"x"0;0', 'en', '-3'],
    [-3, '0;[<=0]"x"0;0.0', 'en', '-x3'],
    [0, '0;[<=0]"x"0;0.0', 'en', 'x0'],
    [7, '0;[<-5]"x"0;0.0', 'en', '7'],
  ]);
});

test('number format: bracket names that collide with object members', () => {
  for (const name of ['constructor', '__proto__', 'toString', 'hasOwnProperty', 'valueOf']) {
    assert.deepEqual(formatValue(5, `[${name}]0`), { text: '5' }, name);
    assert.deepEqual(formatValue('x', `0;0;0;[${name}]@`), { text: 'x' }, name);
  }
});
